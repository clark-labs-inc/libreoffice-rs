//! ZIP archive reader supporting stored (method 0) and deflate (method 8)
//! compression. Used by importers that need to parse OOXML/ODF packages.

use std::collections::BTreeMap;

use lo_core::{LoError, Result};

#[derive(Clone, Debug)]
pub struct ZipEntryMeta {
    pub name: String,
    pub method: u16,
    pub compressed_size: u32,
    pub uncompressed_size: u32,
    pub local_header_offset: u32,
}

#[derive(Clone, Debug)]
pub struct ZipArchive {
    bytes: Vec<u8>,
    entries: BTreeMap<String, ZipEntryMeta>,
}

impl ZipArchive {
    pub fn new(bytes: &[u8]) -> Result<Self> {
        let eocd = find_eocd(bytes)
            .ok_or_else(|| LoError::Parse("zip end-of-central-directory not found".to_string()))?;
        if eocd + 22 > bytes.len() {
            return Err(LoError::Parse(
                "truncated zip end-of-central-directory".to_string(),
            ));
        }
        let total_entries = read_u16(bytes, eocd + 10)? as usize;
        let central_directory_size = read_u32(bytes, eocd + 12)? as usize;
        let central_directory_offset = read_u32(bytes, eocd + 16)? as usize;
        if central_directory_offset + central_directory_size > bytes.len() {
            return Err(LoError::Parse(
                "zip central directory is out of bounds".to_string(),
            ));
        }

        let mut entries = BTreeMap::new();
        let mut offset = central_directory_offset;
        for _ in 0..total_entries {
            if read_u32(bytes, offset)? != 0x0201_4b50 {
                return Err(LoError::Parse(format!(
                    "bad central directory signature at {offset}"
                )));
            }
            let method = read_u16(bytes, offset + 10)?;
            let compressed_size = read_u32(bytes, offset + 20)?;
            let uncompressed_size = read_u32(bytes, offset + 24)?;
            let file_name_length = read_u16(bytes, offset + 28)? as usize;
            let extra_length = read_u16(bytes, offset + 30)? as usize;
            let comment_length = read_u16(bytes, offset + 32)? as usize;
            let local_header_offset = read_u32(bytes, offset + 42)?;
            let name_start = offset + 46;
            let name_end = name_start + file_name_length;
            if name_end > bytes.len() {
                return Err(LoError::Parse(
                    "zip entry name is out of bounds".to_string(),
                ));
            }
            let name = String::from_utf8(bytes[name_start..name_end].to_vec())
                .map_err(|err| LoError::Parse(format!("zip entry name is not utf-8: {err}")))?;
            let normalized = normalize_zip_path(&name);
            entries.insert(
                normalized.clone(),
                ZipEntryMeta {
                    name: normalized,
                    method,
                    compressed_size,
                    uncompressed_size,
                    local_header_offset,
                },
            );
            offset = name_end + extra_length + comment_length;
        }

        Ok(Self {
            bytes: bytes.to_vec(),
            entries,
        })
    }

    pub fn entries(&self) -> Vec<&str> {
        self.entries.keys().map(String::as_str).collect()
    }

    pub fn contains(&self, path: &str) -> bool {
        self.entries.contains_key(&normalize_zip_path(path))
    }

    pub fn read(&self, path: &str) -> Result<Vec<u8>> {
        let key = normalize_zip_path(path);
        let entry = self
            .entries
            .get(&key)
            .ok_or_else(|| LoError::InvalidInput(format!("zip entry not found: {path}")))?;
        let local_offset = entry.local_header_offset as usize;
        if read_u32(&self.bytes, local_offset)? != 0x0403_4b50 {
            return Err(LoError::Parse(format!(
                "bad local file header for {}",
                entry.name
            )));
        }
        let flags = read_u16(&self.bytes, local_offset + 6)?;
        let method = read_u16(&self.bytes, local_offset + 8)?;
        let file_name_length = read_u16(&self.bytes, local_offset + 26)? as usize;
        let extra_length = read_u16(&self.bytes, local_offset + 28)? as usize;
        let data_start = local_offset + 30 + file_name_length + extra_length;
        let data_end = data_start + entry.compressed_size as usize;
        if data_end > self.bytes.len() {
            return Err(LoError::Parse(format!(
                "zip data out of bounds for {}",
                entry.name
            )));
        }
        let compressed = &self.bytes[data_start..data_end];
        let out = match method {
            0 => compressed.to_vec(),
            8 => inflate_deflate(compressed, entry.uncompressed_size as usize)?,
            other => {
                return Err(LoError::Unsupported(format!(
                    "zip compression method {other} for {}",
                    entry.name
                )))
            }
        };
        if flags & 0x0008 == 0 && out.len() != entry.uncompressed_size as usize {
            return Err(LoError::Parse(format!(
                "zip size mismatch for {}: expected {}, got {}",
                entry.name,
                entry.uncompressed_size,
                out.len()
            )));
        }
        Ok(out)
    }

    pub fn read_string(&self, path: &str) -> Result<String> {
        let data = self.read(path)?;
        String::from_utf8(data)
            .map_err(|err| LoError::Parse(format!("zip entry is not utf-8: {err}")))
    }
}

pub fn normalize_zip_path(path: &str) -> String {
    let normalized = path.replace('\\', "/");
    let mut parts: Vec<&str> = Vec::new();
    for part in normalized.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            other => parts.push(other),
        }
    }
    parts.join("/")
}

pub fn resolve_part_target(base_part: &str, target: &str) -> String {
    if target.starts_with('/') {
        return normalize_zip_path(target.trim_start_matches('/'));
    }
    let base_dir = base_part.rsplit_once('/').map(|(dir, _)| dir).unwrap_or("");
    if base_dir.is_empty() {
        normalize_zip_path(target)
    } else {
        normalize_zip_path(&format!("{base_dir}/{target}"))
    }
}

pub fn rels_path_for(part: &str) -> String {
    let normalized = normalize_zip_path(part);
    if let Some((dir, file)) = normalized.rsplit_once('/') {
        format!("{dir}/_rels/{file}.rels")
    } else {
        format!("_rels/{normalized}.rels")
    }
}

fn find_eocd(bytes: &[u8]) -> Option<usize> {
    let start = bytes.len().saturating_sub(66_000);
    let sig = 0x0605_4b50u32.to_le_bytes();
    (start..bytes.len().saturating_sub(3))
        .rev()
        .find(|&idx| bytes[idx..].starts_with(&sig))
}

fn read_u16(bytes: &[u8], offset: usize) -> Result<u16> {
    let slice = bytes
        .get(offset..offset + 2)
        .ok_or_else(|| LoError::Parse(format!("zip read_u16 out of bounds at {offset}")))?;
    Ok(u16::from_le_bytes([slice[0], slice[1]]))
}

fn read_u32(bytes: &[u8], offset: usize) -> Result<u32> {
    let slice = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| LoError::Parse(format!("zip read_u32 out of bounds at {offset}")))?;
    Ok(u32::from_le_bytes([slice[0], slice[1], slice[2], slice[3]]))
}

struct BitReader<'a> {
    data: &'a [u8],
    bit_pos: usize,
}

impl<'a> BitReader<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, bit_pos: 0 }
    }

    fn read_bits(&mut self, count: u8) -> Result<u32> {
        if count == 0 {
            return Ok(0);
        }
        let mut out = 0u32;
        for bit_index in 0..count {
            let byte_index = self.bit_pos / 8;
            let bit_in_byte = self.bit_pos % 8;
            let byte = *self
                .data
                .get(byte_index)
                .ok_or_else(|| LoError::Parse("unexpected end of deflate stream".to_string()))?;
            let bit = (byte >> bit_in_byte) & 1;
            out |= (bit as u32) << bit_index;
            self.bit_pos += 1;
        }
        Ok(out)
    }

    fn align_byte(&mut self) {
        self.bit_pos = (self.bit_pos + 7) & !7;
    }

    fn read_aligned_bytes(&mut self, len: usize) -> Result<&'a [u8]> {
        self.align_byte();
        let start = self.bit_pos / 8;
        let end = start + len;
        let slice = self
            .data
            .get(start..end)
            .ok_or_else(|| LoError::Parse("unexpected end of aligned deflate block".to_string()))?;
        self.bit_pos += len * 8;
        Ok(slice)
    }
}

#[derive(Clone, Debug)]
struct Huffman {
    max_len: u8,
    table: BTreeMap<(u8, u16), u16>,
}

impl Huffman {
    fn from_code_lengths(lengths: &[u8]) -> Result<Self> {
        let max_len = *lengths.iter().max().unwrap_or(&0);
        if max_len == 0 {
            return Err(LoError::Parse("empty huffman table".to_string()));
        }
        let mut bl_count = vec![0u16; max_len as usize + 1];
        for &len in lengths {
            if len > 0 {
                bl_count[len as usize] += 1;
            }
        }
        let mut next_code = vec![0u16; max_len as usize + 1];
        let mut code = 0u16;
        for bits in 1..=max_len as usize {
            code = (code + bl_count[bits - 1]) << 1;
            next_code[bits] = code;
        }
        let mut table = BTreeMap::new();
        for (symbol, &len) in lengths.iter().enumerate() {
            if len == 0 {
                continue;
            }
            let canonical = next_code[len as usize];
            next_code[len as usize] += 1;
            let reversed = reverse_bits(canonical, len);
            table.insert((len, reversed), symbol as u16);
        }
        Ok(Self { max_len, table })
    }

    fn decode_symbol(&self, bits: &mut BitReader<'_>) -> Result<u16> {
        let mut code = 0u16;
        for len in 1..=self.max_len {
            let bit = bits.read_bits(1)? as u16;
            code |= bit << (len - 1);
            if let Some(symbol) = self.table.get(&(len, code)) {
                return Ok(*symbol);
            }
        }
        Err(LoError::Parse("invalid huffman code".to_string()))
    }
}

fn reverse_bits(mut code: u16, len: u8) -> u16 {
    let mut out = 0u16;
    for _ in 0..len {
        out = (out << 1) | (code & 1);
        code >>= 1;
    }
    out
}

fn inflate_deflate(data: &[u8], expected_len: usize) -> Result<Vec<u8>> {
    let mut reader = BitReader::new(data);
    let mut out = Vec::with_capacity(expected_len.max(256));
    loop {
        let is_final = reader.read_bits(1)? != 0;
        let block_type = reader.read_bits(2)? as u8;
        match block_type {
            0 => read_stored_block(&mut reader, &mut out)?,
            1 => {
                let litlen = fixed_literal_huffman()?;
                let dist = fixed_distance_huffman()?;
                read_huffman_block(&mut reader, &litlen, &dist, &mut out)?;
            }
            2 => {
                let (litlen, dist) = read_dynamic_huffman_tables(&mut reader)?;
                read_huffman_block(&mut reader, &litlen, &dist, &mut out)?;
            }
            3 => return Err(LoError::Parse("reserved deflate block type".to_string())),
            _ => unreachable!(),
        }
        if is_final {
            break;
        }
    }
    Ok(out)
}

fn read_stored_block(reader: &mut BitReader<'_>, out: &mut Vec<u8>) -> Result<()> {
    reader.align_byte();
    let header = reader.read_aligned_bytes(4)?;
    let len = u16::from_le_bytes([header[0], header[1]]);
    let nlen = u16::from_le_bytes([header[2], header[3]]);
    if len != !nlen {
        return Err(LoError::Parse(
            "stored deflate block length checksum mismatch".to_string(),
        ));
    }
    let bytes = reader.read_aligned_bytes(len as usize)?;
    out.extend_from_slice(bytes);
    Ok(())
}

fn read_dynamic_huffman_tables(reader: &mut BitReader<'_>) -> Result<(Huffman, Huffman)> {
    let hlit = reader.read_bits(5)? as usize + 257;
    let hdist = reader.read_bits(5)? as usize + 1;
    let hclen = reader.read_bits(4)? as usize + 4;

    let order = [
        16usize, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15,
    ];
    let mut code_lengths = vec![0u8; 19];
    for i in 0..hclen {
        code_lengths[order[i]] = reader.read_bits(3)? as u8;
    }
    let code_length_huffman = Huffman::from_code_lengths(&code_lengths)?;

    let total = hlit + hdist;
    let mut lengths = Vec::with_capacity(total);
    while lengths.len() < total {
        match code_length_huffman.decode_symbol(reader)? {
            symbol @ 0..=15 => lengths.push(symbol as u8),
            16 => {
                let repeat = reader.read_bits(2)? as usize + 3;
                let previous = *lengths.last().ok_or_else(|| {
                    LoError::Parse("repeat code without previous code length".to_string())
                })?;
                lengths.extend(std::iter::repeat_n(previous, repeat));
            }
            17 => {
                let repeat = reader.read_bits(3)? as usize + 3;
                lengths.extend(std::iter::repeat_n(0u8, repeat));
            }
            18 => {
                let repeat = reader.read_bits(7)? as usize + 11;
                lengths.extend(std::iter::repeat_n(0u8, repeat));
            }
            other => {
                return Err(LoError::Parse(format!(
                    "invalid code length symbol {other}"
                )))
            }
        }
    }

    let litlen = Huffman::from_code_lengths(&lengths[..hlit])?;
    let dist_lengths = &lengths[hlit..hlit + hdist];
    let dist = if dist_lengths.iter().all(|&len| len == 0) {
        Huffman::from_code_lengths(&[1])?
    } else {
        Huffman::from_code_lengths(dist_lengths)?
    };
    Ok((litlen, dist))
}

fn read_huffman_block(
    reader: &mut BitReader<'_>,
    litlen: &Huffman,
    dist: &Huffman,
    out: &mut Vec<u8>,
) -> Result<()> {
    loop {
        let symbol = litlen.decode_symbol(reader)?;
        match symbol {
            0..=255 => out.push(symbol as u8),
            256 => return Ok(()),
            257..=285 => {
                let length = decode_length(reader, symbol)?;
                let distance_symbol = dist.decode_symbol(reader)?;
                let distance = decode_distance(reader, distance_symbol)?;
                if distance == 0 || distance > out.len() {
                    return Err(LoError::Parse(
                        "invalid deflate back-reference distance".to_string(),
                    ));
                }
                let start = out.len() - distance;
                for i in 0..length {
                    let byte = out[start + (i % distance)];
                    out.push(byte);
                }
            }
            other => {
                return Err(LoError::Parse(format!(
                    "invalid deflate literal/length symbol {other}"
                )))
            }
        }
    }
}

fn fixed_literal_huffman() -> Result<Huffman> {
    let mut lengths = vec![0u8; 288];
    for symbol in 0..=143 {
        lengths[symbol] = 8;
    }
    for symbol in 144..=255 {
        lengths[symbol] = 9;
    }
    for symbol in 256..=279 {
        lengths[symbol] = 7;
    }
    for symbol in 280..=287 {
        lengths[symbol] = 8;
    }
    Huffman::from_code_lengths(&lengths)
}

fn fixed_distance_huffman() -> Result<Huffman> {
    Huffman::from_code_lengths(&[5u8; 32])
}

fn decode_length(reader: &mut BitReader<'_>, symbol: u16) -> Result<usize> {
    const BASES: [usize; 29] = [
        3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115,
        131, 163, 195, 227, 258,
    ];
    const EXTRA: [u8; 29] = [
        0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0,
    ];
    if !(257..=285).contains(&symbol) {
        return Err(LoError::Parse(format!(
            "invalid deflate length symbol {symbol}"
        )));
    }
    if symbol == 285 {
        return Ok(258);
    }
    let index = (symbol - 257) as usize;
    let extra_bits = EXTRA[index];
    let extra = if extra_bits == 0 {
        0
    } else {
        reader.read_bits(extra_bits)? as usize
    };
    Ok(BASES[index] + extra)
}

fn decode_distance(reader: &mut BitReader<'_>, symbol: u16) -> Result<usize> {
    const BASES: [usize; 30] = [
        1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537,
        2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577,
    ];
    const EXTRA: [u8; 30] = [
        0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12,
        13, 13,
    ];
    let index = symbol as usize;
    if index >= BASES.len() {
        return Err(LoError::Parse(format!(
            "invalid deflate distance symbol {symbol}"
        )));
    }
    let extra_bits = EXTRA[index];
    let extra = if extra_bits == 0 {
        0
    } else {
        reader.read_bits(extra_bits)? as usize
    };
    Ok(BASES[index] + extra)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normalize_path_collapses_segments() {
        assert_eq!(
            normalize_zip_path("word/../word/document.xml"),
            "word/document.xml"
        );
    }

    #[test]
    fn rels_path_for_root_part() {
        assert_eq!(
            rels_path_for("word/document.xml"),
            "word/_rels/document.xml.rels"
        );
        assert_eq!(
            rels_path_for("[Content_Types].xml"),
            "_rels/[Content_Types].xml.rels"
        );
    }
}
