use std::fs::File;
use std::io::{Cursor, Read, Seek, Write};
use std::path::Path;

use lo_core::{LoError, Result};

pub mod archive;
pub use archive::{
    normalize_zip_path, rels_path_for, resolve_part_target, ZipArchive, ZipEntryMeta,
};

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ZipEntry {
    pub name: String,
    pub data: Vec<u8>,
}

impl ZipEntry {
    pub fn new(name: impl Into<String>, data: impl Into<Vec<u8>>) -> Self {
        Self {
            name: name.into(),
            data: data.into(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CentralEntry {
    pub name: String,
    pub compressed_size: u32,
    pub uncompressed_size: u32,
    pub local_header_offset: u32,
}

fn write_u16_le<W: Write>(writer: &mut W, value: u16) -> Result<()> {
    writer.write_all(&value.to_le_bytes())?;
    Ok(())
}

fn write_u32_le<W: Write>(writer: &mut W, value: u32) -> Result<()> {
    writer.write_all(&value.to_le_bytes())?;
    Ok(())
}

fn read_u16_le(slice: &[u8], offset: usize) -> Result<u16> {
    let bytes = slice
        .get(offset..offset + 2)
        .ok_or_else(|| LoError::Parse("unexpected end of ZIP file".to_string()))?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32_le(slice: &[u8], offset: usize) -> Result<u32> {
    let bytes = slice
        .get(offset..offset + 4)
        .ok_or_else(|| LoError::Parse("unexpected end of ZIP file".to_string()))?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

pub fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xffff_ffffu32;
    for &byte in bytes {
        let mut x = (crc ^ (byte as u32)) & 0xff;
        for _ in 0..8 {
            x = if x & 1 != 0 {
                0xedb8_8320u32 ^ (x >> 1)
            } else {
                x >> 1
            };
        }
        crc = (crc >> 8) ^ x;
    }
    !crc
}

pub fn write_zip<W: Write + Seek>(writer: &mut W, entries: &[ZipEntry]) -> Result<()> {
    struct CentralRecord {
        name: String,
        crc: u32,
        size: u32,
        offset: u32,
    }

    let mut records = Vec::with_capacity(entries.len());

    for entry in entries {
        let offset = writer.stream_position()? as u32;
        let name_bytes = entry.name.as_bytes();
        let data = &entry.data;
        let size = u32::try_from(data.len())
            .map_err(|_| LoError::InvalidInput("ZIP entry too large".to_string()))?;
        let crc = crc32(data);

        write_u32_le(writer, 0x0403_4b50)?;
        write_u16_le(writer, 20)?;
        write_u16_le(writer, 0)?;
        write_u16_le(writer, 0)?;
        write_u16_le(writer, 0)?;
        write_u16_le(writer, 0)?;
        write_u32_le(writer, crc)?;
        write_u32_le(writer, size)?;
        write_u32_le(writer, size)?;
        write_u16_le(
            writer,
            u16::try_from(name_bytes.len())
                .map_err(|_| LoError::InvalidInput("ZIP entry name too long".to_string()))?,
        )?;
        write_u16_le(writer, 0)?;
        writer.write_all(name_bytes)?;
        writer.write_all(data)?;

        records.push(CentralRecord {
            name: entry.name.clone(),
            crc,
            size,
            offset,
        });
    }

    let central_directory_offset = writer.stream_position()? as u32;

    for record in &records {
        let name_bytes = record.name.as_bytes();
        write_u32_le(writer, 0x0201_4b50)?;
        write_u16_le(writer, 20)?;
        write_u16_le(writer, 20)?;
        write_u16_le(writer, 0)?;
        write_u16_le(writer, 0)?;
        write_u16_le(writer, 0)?;
        write_u16_le(writer, 0)?;
        write_u32_le(writer, record.crc)?;
        write_u32_le(writer, record.size)?;
        write_u32_le(writer, record.size)?;
        write_u16_le(
            writer,
            u16::try_from(name_bytes.len())
                .map_err(|_| LoError::InvalidInput("ZIP entry name too long".to_string()))?,
        )?;
        write_u16_le(writer, 0)?;
        write_u16_le(writer, 0)?;
        write_u16_le(writer, 0)?;
        write_u16_le(writer, 0)?;
        write_u32_le(writer, 0)?;
        write_u32_le(writer, record.offset)?;
        writer.write_all(name_bytes)?;
    }

    let central_directory_size = (writer.stream_position()? as u32) - central_directory_offset;

    write_u32_le(writer, 0x0605_4b50)?;
    write_u16_le(writer, 0)?;
    write_u16_le(writer, 0)?;
    write_u16_le(
        writer,
        u16::try_from(records.len())
            .map_err(|_| LoError::InvalidInput("too many ZIP entries".to_string()))?,
    )?;
    write_u16_le(
        writer,
        u16::try_from(records.len())
            .map_err(|_| LoError::InvalidInput("too many ZIP entries".to_string()))?,
    )?;
    write_u32_le(writer, central_directory_size)?;
    write_u32_le(writer, central_directory_offset)?;
    write_u16_le(writer, 0)?;
    Ok(())
}

pub fn write_zip_file(path: impl AsRef<Path>, entries: &[ZipEntry]) -> Result<()> {
    let mut file = File::create(path)?;
    write_zip(&mut file, entries)
}

/// Serialize a list of ZIP entries into an in-memory `Vec<u8>`.
///
/// Convenient for callers that want to package an OOXML/ODF document and
/// then hand the bytes to a downstream consumer (HTTP, embedding, hashing).
pub fn write_zip_to_vec(entries: &[ZipEntry]) -> Result<Vec<u8>> {
    let mut cursor = Cursor::new(Vec::new());
    write_zip(&mut cursor, entries)?;
    Ok(cursor.into_inner())
}

/// Build an OOXML package (`.docx`/`.xlsx`/`.pptx`) from a list of entries.
///
/// OOXML packages are just plain ZIPs without the special "mimetype must be
/// the first stored file" requirement that ODF imposes, so this is a thin
/// wrapper around [`write_zip_to_vec`] for readability at call sites.
pub fn ooxml_package(entries: &[ZipEntry]) -> Result<Vec<u8>> {
    write_zip_to_vec(entries)
}

/// Build an ODF package (`.odt`/`.ods`/`.odp`/`.odg`/`.odb`/`.odf`).
///
/// ODF requires the `mimetype` entry to be the first file in the archive
/// and stored uncompressed; we already store everything uncompressed so we
/// just need to make sure `mimetype` comes first.
pub fn odf_package(mimetype: &str, mut entries: Vec<ZipEntry>) -> Result<Vec<u8>> {
    let mut all = Vec::with_capacity(entries.len() + 1);
    all.push(ZipEntry::new("mimetype", mimetype.as_bytes().to_vec()));
    all.append(&mut entries);
    write_zip_to_vec(&all)
}

pub fn list_entries(path: impl AsRef<Path>) -> Result<Vec<CentralEntry>> {
    let mut file = File::open(path)?;
    let mut data = Vec::new();
    file.read_to_end(&mut data)?;

    let signature = [0x50, 0x4b, 0x05, 0x06];
    let search_start = data.len().saturating_sub(66_000);
    let eocd = (search_start..data.len().saturating_sub(3))
        .rev()
        .find(|&idx| data.get(idx..idx + 4) == Some(&signature))
        .ok_or_else(|| LoError::Parse("could not find ZIP central directory".to_string()))?;

    let entry_count = read_u16_le(&data, eocd + 10)? as usize;
    let central_offset = read_u32_le(&data, eocd + 16)? as usize;

    let mut cursor = central_offset;
    let mut entries = Vec::with_capacity(entry_count);
    for _ in 0..entry_count {
        let sig = read_u32_le(&data, cursor)?;
        if sig != 0x0201_4b50 {
            return Err(LoError::Parse(
                "invalid central directory header".to_string(),
            ));
        }
        let compressed_size = read_u32_le(&data, cursor + 20)?;
        let uncompressed_size = read_u32_le(&data, cursor + 24)?;
        let name_len = read_u16_le(&data, cursor + 28)? as usize;
        let extra_len = read_u16_le(&data, cursor + 30)? as usize;
        let comment_len = read_u16_le(&data, cursor + 32)? as usize;
        let local_header_offset = read_u32_le(&data, cursor + 42)?;
        let name_start = cursor + 46;
        let name_end = name_start + name_len;
        let name_bytes = data
            .get(name_start..name_end)
            .ok_or_else(|| LoError::Parse("invalid central directory name".to_string()))?;
        let name = String::from_utf8_lossy(name_bytes).to_string();
        entries.push(CentralEntry {
            name,
            compressed_size,
            uncompressed_size,
            local_header_offset,
        });
        cursor = name_end + extra_len + comment_len;
    }

    Ok(entries)
}

pub fn read_entry(path: impl AsRef<Path>, entry_name: &str) -> Result<Vec<u8>> {
    let mut file = File::open(path)?;
    let mut data = Vec::new();
    file.read_to_end(&mut data)?;

    for entry in list_entries_from_bytes(&data)? {
        if entry.name == entry_name {
            let offset = entry.local_header_offset as usize;
            let sig = read_u32_le(&data, offset)?;
            if sig != 0x0403_4b50 {
                return Err(LoError::Parse("invalid local ZIP header".to_string()));
            }
            let name_len = read_u16_le(&data, offset + 26)? as usize;
            let extra_len = read_u16_le(&data, offset + 28)? as usize;
            let body_start = offset + 30 + name_len + extra_len;
            let body_end = body_start + entry.uncompressed_size as usize;
            return data
                .get(body_start..body_end)
                .map(|slice| slice.to_vec())
                .ok_or_else(|| LoError::Parse("invalid ZIP body range".to_string()));
        }
    }

    Err(LoError::InvalidInput(format!(
        "ZIP entry not found: {entry_name}"
    )))
}

fn list_entries_from_bytes(data: &[u8]) -> Result<Vec<CentralEntry>> {
    let signature = [0x50, 0x4b, 0x05, 0x06];
    let search_start = data.len().saturating_sub(66_000);
    let eocd = (search_start..data.len().saturating_sub(3))
        .rev()
        .find(|&idx| data.get(idx..idx + 4) == Some(&signature))
        .ok_or_else(|| LoError::Parse("could not find ZIP central directory".to_string()))?;

    let entry_count = read_u16_le(data, eocd + 10)? as usize;
    let central_offset = read_u32_le(data, eocd + 16)? as usize;

    let mut cursor = central_offset;
    let mut entries = Vec::with_capacity(entry_count);
    for _ in 0..entry_count {
        let sig = read_u32_le(data, cursor)?;
        if sig != 0x0201_4b50 {
            return Err(LoError::Parse(
                "invalid central directory header".to_string(),
            ));
        }
        let compressed_size = read_u32_le(data, cursor + 20)?;
        let uncompressed_size = read_u32_le(data, cursor + 24)?;
        let name_len = read_u16_le(data, cursor + 28)? as usize;
        let extra_len = read_u16_le(data, cursor + 30)? as usize;
        let comment_len = read_u16_le(data, cursor + 32)? as usize;
        let local_header_offset = read_u32_le(data, cursor + 42)?;
        let name_start = cursor + 46;
        let name_end = name_start + name_len;
        let name_bytes = data
            .get(name_start..name_end)
            .ok_or_else(|| LoError::Parse("invalid central directory name".to_string()))?;
        entries.push(CentralEntry {
            name: String::from_utf8_lossy(name_bytes).to_string(),
            compressed_size,
            uncompressed_size,
            local_header_offset,
        });
        cursor = name_end + extra_len + comment_len;
    }
    Ok(entries)
}

#[cfg(test)]
mod tests {
    use super::{
        crc32, list_entries, ooxml_package, read_entry, write_zip_file, write_zip_to_vec, ZipEntry,
    };
    use std::time::{SystemTime, UNIX_EPOCH};

    #[test]
    fn crc32_matches_known_value() {
        assert_eq!(crc32(b"123456789"), 0xcbf4_3926);
    }

    #[test]
    fn zip_roundtrip_works() {
        let ts = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system time before unix epoch")
            .as_nanos();
        let path = std::env::temp_dir().join(format!("lo_zip_{ts}.zip"));
        write_zip_file(
            &path,
            &[
                ZipEntry::new("a.txt", b"hello".to_vec()),
                ZipEntry::new("dir/b.txt", b"world".to_vec()),
            ],
        )
        .expect("write zip");
        let entries = list_entries(&path).expect("list entries");
        assert_eq!(entries.len(), 2);
        assert_eq!(read_entry(&path, "a.txt").expect("read entry"), b"hello");
        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn write_to_vec_starts_with_pk_signature() {
        let bytes =
            write_zip_to_vec(&[ZipEntry::new("a.txt", b"hi".to_vec())]).expect("write zip to vec");
        assert!(bytes.starts_with(b"PK"));
    }

    #[test]
    fn ooxml_package_is_valid_zip() {
        let bytes = ooxml_package(&[
            ZipEntry::new("[Content_Types].xml", b"<Types/>".to_vec()),
            ZipEntry::new("word/document.xml", b"<doc/>".to_vec()),
        ])
        .expect("build ooxml");
        assert!(bytes.starts_with(b"PK"));
    }
}
