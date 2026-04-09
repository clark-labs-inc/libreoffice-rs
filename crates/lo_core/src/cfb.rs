//! Compound File Binary (CFB) reader.
//!
//! CFB is the Microsoft container format used by legacy Office files
//! (`.doc`, `.xls`, `.ppt`, `.msg`, …). This module exposes a small
//! `CfbFile` type that opens a byte slice and lets you list and read
//! its embedded streams; higher-level decoders (Word piece tables,
//! Excel BIFF, …) live in their own crates and call into here.

use std::collections::BTreeMap;

use crate::{LoError, Result};

const FREE_SECT: u32 = 0xFFFF_FFFF;
const END_OF_CHAIN: u32 = 0xFFFF_FFFE;
const FAT_SECT: u32 = 0xFFFF_FFFD;
const DIFAT_SECT: u32 = 0xFFFF_FFFC;

#[derive(Clone, Debug)]
pub struct CfbEntry {
    pub name: String,
    pub object_type: u8,
    pub start_sector: u32,
    pub stream_size: u64,
}

#[derive(Clone, Debug)]
pub struct CfbFile {
    bytes: Vec<u8>,
    sector_size: usize,
    mini_sector_size: usize,
    fat: Vec<u32>,
    mini_fat: Vec<u32>,
    entries: Vec<CfbEntry>,
    entries_by_name: BTreeMap<String, usize>,
    mini_stream: Vec<u8>,
    mini_stream_cutoff: usize,
}

impl CfbFile {
    pub fn open(bytes: &[u8]) -> Result<Self> {
        if bytes.len() < 512 {
            return Err(LoError::Parse("cfb header truncated".to_string()));
        }
        let magic = &bytes[..8];
        if magic != [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1] {
            return Err(LoError::Parse("invalid cfb signature".to_string()));
        }
        let sector_shift = read_u16(bytes, 30)? as usize;
        let mini_sector_shift = read_u16(bytes, 32)? as usize;
        let sector_size = 1usize << sector_shift;
        let mini_sector_size = 1usize << mini_sector_shift;
        let num_fat_sectors = read_u32(bytes, 44)?;
        let first_dir_sector = read_u32(bytes, 48)?;
        let mini_stream_cutoff = read_u32(bytes, 56)? as usize;
        let first_mini_fat_sector = read_u32(bytes, 60)?;
        let num_mini_fat_sectors = read_u32(bytes, 64)?;
        let first_difat_sector = read_u32(bytes, 68)?;
        let num_difat_sectors = read_u32(bytes, 72)?;

        let mut difat = Vec::new();
        for index in 0..109usize {
            let value = read_u32(bytes, 76 + index * 4)?;
            if value != FREE_SECT {
                difat.push(value);
            }
        }
        let mut next_difat = first_difat_sector;
        for _ in 0..num_difat_sectors {
            if next_difat == END_OF_CHAIN || next_difat == FREE_SECT {
                break;
            }
            let sector = read_sector(bytes, sector_size, next_difat)?;
            let entries_per_difat = sector_size / 4 - 1;
            for index in 0..entries_per_difat {
                let value = read_u32(sector, index * 4)?;
                if value != FREE_SECT {
                    difat.push(value);
                }
            }
            next_difat = read_u32(sector, sector_size - 4)?;
        }
        if difat.len() < num_fat_sectors as usize {
            return Err(LoError::Parse(
                "cfb DIFAT smaller than FAT sector count".to_string(),
            ));
        }
        difat.truncate(num_fat_sectors as usize);

        let mut fat = Vec::new();
        for sector_id in &difat {
            let sector = read_sector(bytes, sector_size, *sector_id)?;
            for chunk in sector.chunks_exact(4) {
                fat.push(u32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]));
            }
        }

        let dir_stream = read_chain(bytes, sector_size, &fat, first_dir_sector)?;
        let mut entries = Vec::new();
        let mut entries_by_name = BTreeMap::new();
        for entry_bytes in dir_stream.chunks_exact(128) {
            let name_len = read_u16(entry_bytes, 64)? as usize;
            if name_len < 2 {
                continue;
            }
            let object_type = entry_bytes[66];
            let name = decode_utf16_name(&entry_bytes[..name_len.saturating_sub(2)])?;
            let start_sector = read_u32(entry_bytes, 116)?;
            let stream_size = read_u64(entry_bytes, 120)?;
            entries_by_name.insert(name.clone(), entries.len());
            entries.push(CfbEntry {
                name,
                object_type,
                start_sector,
                stream_size,
            });
        }

        let root_entry = entries
            .iter()
            .find(|entry| entry.object_type == 5)
            .ok_or_else(|| LoError::Parse("cfb root entry not found".to_string()))?;
        let mini_stream = if root_entry.start_sector != END_OF_CHAIN && root_entry.stream_size > 0 {
            let mut bytes_out = read_chain(bytes, sector_size, &fat, root_entry.start_sector)?;
            bytes_out.truncate(root_entry.stream_size as usize);
            bytes_out
        } else {
            Vec::new()
        };

        let mini_fat_bytes = if num_mini_fat_sectors > 0 && first_mini_fat_sector != END_OF_CHAIN {
            read_chain(bytes, sector_size, &fat, first_mini_fat_sector)?
        } else {
            Vec::new()
        };
        let mut mini_fat = Vec::new();
        for chunk in mini_fat_bytes.chunks_exact(4) {
            mini_fat.push(u32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]));
        }

        Ok(Self {
            bytes: bytes.to_vec(),
            sector_size,
            mini_sector_size,
            fat,
            mini_fat,
            entries,
            entries_by_name,
            mini_stream,
            mini_stream_cutoff,
        })
    }

    pub fn list_streams(&self) -> Vec<&str> {
        self.entries
            .iter()
            .filter(|entry| entry.object_type == 2 || entry.object_type == 5)
            .map(|entry| entry.name.as_str())
            .collect()
    }

    pub fn read_stream(&self, name: &str) -> Result<Vec<u8>> {
        let entry_index = self
            .entries_by_name
            .get(name)
            .or_else(|| {
                self.entries_by_name
                    .iter()
                    .find(|(key, _)| key.eq_ignore_ascii_case(name))
                    .map(|(_, idx)| idx)
            })
            .copied()
            .ok_or_else(|| LoError::InvalidInput(format!("cfb stream not found: {name}")))?;
        let entry = &self.entries[entry_index];
        if entry.stream_size == 0 {
            return Ok(Vec::new());
        }
        let mut data =
            if entry.stream_size as usize >= self.mini_stream_cutoff || entry.object_type == 5 {
                read_chain(&self.bytes, self.sector_size, &self.fat, entry.start_sector)?
            } else {
                self.read_mini_chain(entry.start_sector)?
            };
        data.truncate(entry.stream_size as usize);
        Ok(data)
    }

    pub fn has_stream(&self, name: &str) -> bool {
        self.entries_by_name.contains_key(name)
            || self
                .entries_by_name
                .keys()
                .any(|key| key.eq_ignore_ascii_case(name))
    }

    fn read_mini_chain(&self, start_sector: u32) -> Result<Vec<u8>> {
        let mut out = Vec::new();
        let mut current = start_sector;
        let mut hops = 0usize;
        while current != END_OF_CHAIN {
            let current_index = current as usize;
            let start = current_index
                .checked_mul(self.mini_sector_size)
                .ok_or_else(|| LoError::Parse("mini stream offset overflow".to_string()))?;
            let end = start + self.mini_sector_size;
            let chunk = self
                .mini_stream
                .get(start..end)
                .ok_or_else(|| LoError::Parse("mini stream sector out of bounds".to_string()))?;
            out.extend_from_slice(chunk);
            current = *self
                .mini_fat
                .get(current_index)
                .ok_or_else(|| LoError::Parse("mini FAT index out of bounds".to_string()))?;
            hops += 1;
            if hops > self.mini_fat.len() + 1 {
                return Err(LoError::Parse("mini FAT cycle detected".to_string()));
            }
        }
        Ok(out)
    }
}

fn read_chain(bytes: &[u8], sector_size: usize, fat: &[u32], start_sector: u32) -> Result<Vec<u8>> {
    if start_sector == END_OF_CHAIN || start_sector == FREE_SECT {
        return Ok(Vec::new());
    }
    let mut out = Vec::new();
    let mut current = start_sector;
    let mut hops = 0usize;
    while current != END_OF_CHAIN {
        if matches!(current, FREE_SECT | FAT_SECT | DIFAT_SECT) {
            return Err(LoError::Parse(
                "unexpected special sector in chain".to_string(),
            ));
        }
        let sector = read_sector(bytes, sector_size, current)?;
        out.extend_from_slice(sector);
        current = *fat
            .get(current as usize)
            .ok_or_else(|| LoError::Parse("FAT index out of bounds".to_string()))?;
        hops += 1;
        if hops > fat.len() + 1 {
            return Err(LoError::Parse("FAT cycle detected".to_string()));
        }
    }
    Ok(out)
}

fn read_sector(bytes: &[u8], sector_size: usize, sector_id: u32) -> Result<&[u8]> {
    let start = (sector_id as usize + 1)
        .checked_mul(sector_size)
        .ok_or_else(|| LoError::Parse("cfb sector offset overflow".to_string()))?;
    let end = start + sector_size;
    bytes
        .get(start..end)
        .ok_or_else(|| LoError::Parse(format!("cfb sector {sector_id} out of bounds")))
}

fn read_u16(bytes: &[u8], offset: usize) -> Result<u16> {
    let slice = bytes
        .get(offset..offset + 2)
        .ok_or_else(|| LoError::Parse(format!("cfb read_u16 out of bounds at {offset}")))?;
    Ok(u16::from_le_bytes([slice[0], slice[1]]))
}

fn read_u32(bytes: &[u8], offset: usize) -> Result<u32> {
    let slice = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| LoError::Parse(format!("cfb read_u32 out of bounds at {offset}")))?;
    Ok(u32::from_le_bytes([slice[0], slice[1], slice[2], slice[3]]))
}

fn read_u64(bytes: &[u8], offset: usize) -> Result<u64> {
    let slice = bytes
        .get(offset..offset + 8)
        .ok_or_else(|| LoError::Parse(format!("cfb read_u64 out of bounds at {offset}")))?;
    Ok(u64::from_le_bytes([
        slice[0], slice[1], slice[2], slice[3], slice[4], slice[5], slice[6], slice[7],
    ]))
}

fn decode_utf16_name(bytes: &[u8]) -> Result<String> {
    if bytes.len() % 2 != 0 {
        return Err(LoError::Parse(
            "cfb directory name length is odd".to_string(),
        ));
    }
    let words = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    String::from_utf16(&words)
        .map(|s| s.trim_end_matches('\0').to_string())
        .map_err(|err| LoError::Parse(format!("cfb directory name decode failed: {err}")))
}
