//! Plain-text extractor for the legacy binary `.doc` format.
//!
//! Built on top of [`lo_core::CfbFile`] for the Compound File Binary
//! container, then walks the WordDocument piece table (CLX/PlcPcd) to
//! decode each text piece (compressed CP1252 or uncompressed UTF-16LE).
//!
//! This is intentionally text-first rather than full layout/style parity,
//! but it is more robust than a single fixed-offset CLX lookup: it falls
//! back across `0Table`/`1Table`, scans for plausible CLX locations if the
//! canonical FIB offsets are stale, and normalizes common control markers
//! into paragraphs, line breaks, and tabs.

use lo_core::{CfbFile, LoError, Result};

#[derive(Clone, Debug)]
struct Piece {
    cp_start: usize,
    cp_end: usize,
    fc: u32,
    compressed: bool,
}

/// Extract normalized plain text from a legacy `.doc` byte stream.
pub fn extract_text_from_doc(bytes: &[u8]) -> Result<String> {
    let cfb = CfbFile::open(bytes)?;
    let word_document = cfb.read_stream("WordDocument")?;
    if word_document.len() < 0x20 {
        return Err(LoError::Parse(
            "DOC WordDocument stream is too small".to_string(),
        ));
    }

    let fib_flags = read_u16(&word_document, 0x0A)?;
    if (fib_flags & (1 << 8)) != 0 {
        return Err(LoError::Unsupported(
            "encrypted legacy DOC files are not supported by the pure-Rust parser"
                .to_string(),
        ));
    }

    let preferred_table = if (fib_flags & (1 << 9)) != 0 {
        "1Table"
    } else {
        "0Table"
    };
    let alternates = if preferred_table == "1Table" {
        ["1Table", "0Table"]
    } else {
        ["0Table", "1Table"]
    };

    let mut last_error: Option<LoError> = None;
    for table_name in alternates {
        let table = match cfb.read_stream(table_name) {
            Ok(stream) => stream,
            Err(err) => {
                last_error = Some(err);
                continue;
            }
        };
        let (fc_clx, lcb_clx) = match locate_clx(&word_document, &table) {
            Ok(value) => value,
            Err(err) => {
                last_error = Some(err);
                continue;
            }
        };
        let end = fc_clx
            .checked_add(lcb_clx)
            .ok_or_else(|| LoError::Parse("DOC CLX overflow".to_string()))?;
        if end > table.len() {
            last_error = Some(LoError::Parse(
                "DOC CLX extends past the table stream".to_string(),
            ));
            continue;
        }
        let pieces = match parse_clx(&table[fc_clx..end]) {
            Ok(pieces) => pieces,
            Err(err) => {
                last_error = Some(err);
                continue;
            }
        };
        let raw = extract_piece_text(&word_document, &pieces)?;
        return Ok(normalize_doc_text(&raw));
    }

    Err(last_error.unwrap_or_else(|| {
        LoError::Parse("failed to locate the DOC piece table (CLX)".to_string())
    }))
}

fn locate_clx(word_document: &[u8], table: &[u8]) -> Result<(usize, usize)> {
    if word_document.len() >= 0x1AA {
        let fc = read_u32(word_document, 0x1A2)? as usize;
        let lcb = read_u32(word_document, 0x1A6)? as usize;
        if looks_like_clx(table, fc, lcb) {
            return Ok((fc, lcb));
        }
    }

    let limit = word_document.len().min(0x400);
    for offset in (0x80..limit.saturating_sub(8)).step_by(2) {
        let fc = read_u32(word_document, offset)? as usize;
        let lcb = read_u32(word_document, offset + 4)? as usize;
        if looks_like_clx(table, fc, lcb) {
            return Ok((fc, lcb));
        }
    }

    Err(LoError::Parse(
        "failed to locate the DOC piece table (CLX)".to_string(),
    ))
}

fn looks_like_clx(table: &[u8], fc: usize, lcb: usize) -> bool {
    if lcb < 5 {
        return false;
    }
    let Some(end) = fc.checked_add(lcb) else {
        return false;
    };
    if end > table.len() {
        return false;
    }
    let clx = &table[fc..end];
    let mut index = 0usize;
    while index < clx.len() {
        match clx[index] {
            0x01 => {
                if index + 3 > clx.len() {
                    return false;
                }
                let size = u16::from_le_bytes([clx[index + 1], clx[index + 2]]) as usize;
                index += 3 + size;
            }
            0x02 => return index + 5 <= clx.len(),
            _ => return false,
        }
    }
    false
}

fn parse_clx(clx: &[u8]) -> Result<Vec<Piece>> {
    let mut index = 0usize;
    while index < clx.len() {
        match clx[index] {
            0x01 => {
                if index + 3 > clx.len() {
                    return Err(LoError::Parse(
                        "DOC CLX has a truncated RgPrc".to_string(),
                    ));
                }
                let size = u16::from_le_bytes([clx[index + 1], clx[index + 2]]) as usize;
                index += 3 + size;
            }
            0x02 => {
                let lcb = read_u32(clx, index + 1)? as usize;
                let start = index + 5;
                let end = start
                    .checked_add(lcb)
                    .ok_or_else(|| LoError::Parse("DOC PlcPcd overflow".to_string()))?;
                if end > clx.len() {
                    return Err(LoError::Parse(
                        "DOC PlcPcd exceeds CLX bounds".to_string(),
                    ));
                }
                return parse_plcpcd(&clx[start..end]);
            }
            other => {
                return Err(LoError::Parse(format!(
                    "unexpected DOC CLX tag 0x{other:02X}"
                )))
            }
        }
    }
    Err(LoError::Parse(
        "DOC CLX did not contain a piece table".to_string(),
    ))
}

fn parse_plcpcd(bytes: &[u8]) -> Result<Vec<Piece>> {
    if bytes.len() < 4 || (bytes.len() - 4) % 12 != 0 {
        return Err(LoError::Parse(
            "DOC PlcPcd has an invalid size".to_string(),
        ));
    }
    let piece_count = (bytes.len() - 4) / 12;
    let cp_count = piece_count + 1;
    let pcd_offset = cp_count * 4;
    let mut pieces = Vec::new();
    for index in 0..piece_count {
        let cp_start = read_u32(bytes, index * 4)? as usize;
        let cp_end = read_u32(bytes, (index + 1) * 4)? as usize;
        let raw_fc = read_u32(bytes, pcd_offset + index * 8 + 2)?;
        pieces.push(Piece {
            cp_start,
            cp_end,
            fc: raw_fc & 0x3FFF_FFFF,
            compressed: (raw_fc & 0x4000_0000) != 0,
        });
    }
    Ok(pieces)
}

fn extract_piece_text(word_document: &[u8], pieces: &[Piece]) -> Result<String> {
    let mut out = String::new();
    for piece in pieces {
        if piece.cp_end <= piece.cp_start {
            continue;
        }
        let chars = piece.cp_end - piece.cp_start;
        if piece.compressed {
            let start = (piece.fc / 2) as usize;
            let end = start.checked_add(chars).ok_or_else(|| {
                LoError::Parse("DOC compressed piece overflow".to_string())
            })?;
            if end > word_document.len() {
                return Err(LoError::Parse(
                    "DOC compressed piece exceeds WordDocument bounds".to_string(),
                ));
            }
            out.push_str(&decode_compressed_text(&word_document[start..end]));
        } else {
            let start = piece.fc as usize;
            let byte_len = chars.checked_mul(2).ok_or_else(|| {
                LoError::Parse("DOC unicode piece overflow".to_string())
            })?;
            let end = start.checked_add(byte_len).ok_or_else(|| {
                LoError::Parse("DOC unicode piece overflow".to_string())
            })?;
            if end > word_document.len() {
                return Err(LoError::Parse(
                    "DOC unicode piece exceeds WordDocument bounds".to_string(),
                ));
            }
            out.push_str(&decode_utf16le(&word_document[start..end]));
        }
    }
    Ok(out)
}

fn decode_utf16le(bytes: &[u8]) -> String {
    let mut words = Vec::new();
    for chunk in bytes.chunks_exact(2) {
        words.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }
    String::from_utf16_lossy(&words)
}

fn decode_compressed_text(bytes: &[u8]) -> String {
    bytes.iter().copied().map(decode_cp1252).collect()
}

fn normalize_doc_text(raw: &str) -> String {
    let mut out = String::new();
    let mut last_was_space = false;
    for ch in raw.chars() {
        match ch {
            '\r' => {
                if !out.ends_with("\n\n") {
                    out.push_str("\n\n");
                }
                last_was_space = false;
            }
            '\n' | '\u{000B}' => {
                if !out.ends_with('\n') {
                    out.push('\n');
                }
                last_was_space = false;
            }
            '\u{000C}' => {
                if !out.ends_with("\n\n") {
                    out.push_str("\n\n");
                }
                last_was_space = false;
            }
            '\u{0007}' => {
                if !out.ends_with('\t') {
                    out.push('\t');
                }
                last_was_space = false;
            }
            '\u{0013}' | '\u{0014}' | '\u{0015}' | '\u{0000}' => {}
            ch if ch.is_control() && ch != '\t' => {}
            ch if ch.is_whitespace() => {
                if !last_was_space {
                    out.push(' ');
                    last_was_space = true;
                }
            }
            other => {
                out.push(other);
                last_was_space = false;
            }
        }
    }

    let mut cleaned = String::new();
    let mut blank_run = 0usize;
    for line in out.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() {
            blank_run += 1;
            if blank_run <= 1 && !cleaned.is_empty() && !cleaned.ends_with("\n\n") {
                cleaned.push_str("\n\n");
            }
        } else {
            blank_run = 0;
            if !cleaned.is_empty() && !cleaned.ends_with("\n\n") {
                cleaned.push_str("\n\n");
            }
            cleaned.push_str(trimmed);
        }
    }
    cleaned.trim().to_string()
}

fn read_u16(bytes: &[u8], offset: usize) -> Result<u16> {
    let slice = bytes
        .get(offset..offset + 2)
        .ok_or_else(|| LoError::Parse(format!("DOC read_u16 out of bounds at {offset}")))?;
    Ok(u16::from_le_bytes([slice[0], slice[1]]))
}

fn read_u32(bytes: &[u8], offset: usize) -> Result<u32> {
    let slice = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| LoError::Parse(format!("DOC read_u32 out of bounds at {offset}")))?;
    Ok(u32::from_le_bytes([slice[0], slice[1], slice[2], slice[3]]))
}

fn decode_cp1252(byte: u8) -> char {
    match byte {
        0x80 => '\u{20AC}',
        0x82 => '\u{201A}',
        0x83 => '\u{0192}',
        0x84 => '\u{201E}',
        0x85 => '\u{2026}',
        0x86 => '\u{2020}',
        0x87 => '\u{2021}',
        0x88 => '\u{02C6}',
        0x89 => '\u{2030}',
        0x8A => '\u{0160}',
        0x8B => '\u{2039}',
        0x8C => '\u{0152}',
        0x8E => '\u{017D}',
        0x91 => '\u{2018}',
        0x92 => '\u{2019}',
        0x93 => '\u{201C}',
        0x94 => '\u{201D}',
        0x95 => '\u{2022}',
        0x96 => '\u{2013}',
        0x97 => '\u{2014}',
        0x98 => '\u{02DC}',
        0x99 => '\u{2122}',
        0x9A => '\u{0161}',
        0x9B => '\u{203A}',
        0x9C => '\u{0153}',
        0x9E => '\u{017E}',
        0x9F => '\u{0178}',
        value => value as char,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn single_piece_plcpcd(cp_end: u32, raw_fc: u32) -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&cp_end.to_le_bytes());
        bytes.extend_from_slice(&[0, 0]);
        bytes.extend_from_slice(&raw_fc.to_le_bytes());
        bytes.extend_from_slice(&[0, 0]);
        bytes
    }

    #[test]
    fn parse_clx_skips_prc_and_reads_piece_table() {
        let mut clx = vec![0x01, 0x02, 0x00, 0xAA, 0xBB, 0x02];
        let plcpcd = single_piece_plcpcd(5, 0x4000_0000);
        clx.extend_from_slice(&(plcpcd.len() as u32).to_le_bytes());
        clx.extend_from_slice(&plcpcd);

        let pieces = parse_clx(&clx).expect("pieces");
        assert_eq!(pieces.len(), 1);
        assert_eq!(pieces[0].cp_start, 0);
        assert_eq!(pieces[0].cp_end, 5);
        assert!(pieces[0].compressed);
        assert_eq!(pieces[0].fc, 0);
    }

    #[test]
    fn locate_clx_falls_back_to_scanning_header() {
        let mut word = vec![0u8; 0x200];
        let mut table = vec![0u8; 64];
        let mut clx = vec![0x02];
        let plcpcd = single_piece_plcpcd(3, 0x4000_0000);
        clx.extend_from_slice(&(plcpcd.len() as u32).to_le_bytes());
        clx.extend_from_slice(&plcpcd);

        let fc = 12u32;
        let lcb = clx.len() as u32;
        word[0x80..0x84].copy_from_slice(&fc.to_le_bytes());
        word[0x84..0x88].copy_from_slice(&lcb.to_le_bytes());
        table[12..12 + clx.len()].copy_from_slice(&clx);

        let (found_fc, found_lcb) = locate_clx(&word, &table).expect("locate");
        assert_eq!(found_fc, 12);
        assert_eq!(found_lcb, clx.len());
    }

    #[test]
    fn extract_piece_text_handles_compressed_and_utf16_pieces() {
        let mut word = Vec::new();
        word.extend_from_slice(b"Hello");
        while word.len() < 16 {
            word.push(0);
        }
        word.extend_from_slice(&0x03A9u16.to_le_bytes()); // Ω

        let pieces = vec![
            Piece {
                cp_start: 0,
                cp_end: 5,
                fc: 0,
                compressed: true,
            },
            Piece {
                cp_start: 5,
                cp_end: 6,
                fc: 16,
                compressed: false,
            },
        ];

        let text = extract_piece_text(&word, &pieces).expect("text");
        assert_eq!(text, "HelloΩ");
    }

    #[test]
    fn normalize_doc_text_turns_control_markers_into_paragraphs() {
        let text = normalize_doc_text("One\rTwo\u{0007}Tab\u{000C}Three\u{0013}\u{0015}");
        assert_eq!(text, "One\n\nTwo\tTab\n\nThree");
    }
}
