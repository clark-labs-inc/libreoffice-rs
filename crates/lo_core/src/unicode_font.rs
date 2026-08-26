//! Unicode font discovery, subsetting, and PDF CID font metadata.

use std::collections::{BTreeMap, BTreeSet};
use std::path::Path;

use fontdb::Database;
use subsetter::{subset, GlyphRemapper};
use ttf_parser::{Face, GlyphId};

use crate::{LoError, Result};

const FONT_OVERRIDE_ENV: &str = "LIBREOFFICE_PURE_UNICODE_FONT";

pub(crate) struct EmbeddedUnicodeFont {
    pub(crate) base_name: String,
    pub(crate) bytes: Vec<u8>,
    pub(crate) is_true_type: bool,
    pub(crate) widths: Vec<u16>,
    pub(crate) ascent: i16,
    pub(crate) descent: i16,
    pub(crate) bbox: (i16, i16, i16, i16),
    pub(crate) italic_angle: f32,
    pub(crate) char_to_cid: BTreeMap<char, u16>,
}

impl EmbeddedUnicodeFont {
    pub(crate) fn encode_hex(&self, text: &str) -> Result<String> {
        let mut out = String::with_capacity(text.chars().count() * 4 + 2);
        out.push('<');
        for ch in text.chars() {
            let cid = self.char_to_cid.get(&ch).ok_or_else(|| {
                LoError::Unsupported(format!(
                    "Unicode PDF font does not contain U+{:04X}",
                    ch as u32
                ))
            })?;
            out.push_str(&format!("{cid:04X}"));
        }
        out.push('>');
        Ok(out)
    }

    pub(crate) fn widths_array(&self) -> String {
        self.widths
            .iter()
            .map(u16::to_string)
            .collect::<Vec<_>>()
            .join(" ")
    }

    pub(crate) fn to_unicode_cmap(&self) -> Vec<u8> {
        let mut mappings = self
            .char_to_cid
            .iter()
            .map(|(ch, cid)| (*cid, *ch))
            .collect::<Vec<_>>();
        mappings.sort_by_key(|(cid, _)| *cid);

        let mut cmap = String::from(
            "/CIDInit /ProcSet findresource begin\n\
             12 dict begin\n\
             begincmap\n\
             /CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n\
             /CMapName /LOUnicode def\n\
             /CMapType 2 def\n\
             1 begincodespacerange\n\
             <0000> <FFFF>\n\
             endcodespacerange\n",
        );
        for chunk in mappings.chunks(100) {
            cmap.push_str(&format!("{} beginbfchar\n", chunk.len()));
            for (cid, ch) in chunk {
                cmap.push_str(&format!("<{cid:04X}> <{}>\n", utf16be_hex(*ch)));
            }
            cmap.push_str("endbfchar\n");
        }
        cmap.push_str("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
        cmap.into_bytes()
    }
}

pub(crate) fn build_unicode_font(chars: &BTreeSet<char>) -> Result<EmbeddedUnicodeFont> {
    let mut database = Database::new();
    if let Some(path) = std::env::var_os(FONT_OVERRIDE_ENV) {
        database.load_font_file(Path::new(&path)).map_err(|error| {
            LoError::Io(format!(
                "loading Unicode PDF font {}: {error}",
                Path::new(&path).display()
            ))
        })?;
    } else {
        database.load_system_fonts();
    }

    let mut candidates = database.faces().collect::<Vec<_>>();
    candidates.sort_by_key(|face| std::cmp::Reverse(font_preference(face)));
    let selected = candidates.iter().find(|face| {
        database
            .with_face_data(face.id, |data, index| {
                Face::parse(data, index)
                    .ok()
                    .is_some_and(|parsed| chars.iter().all(|ch| parsed.glyph_index(*ch).is_some()))
            })
            .unwrap_or(false)
    });
    let selected = selected.ok_or_else(|| {
        let best = candidates.iter().max_by_key(|face| {
            database
                .with_face_data(face.id, |data, index| {
                    Face::parse(data, index)
                        .ok()
                        .map(|parsed| {
                            chars
                                .iter()
                                .filter(|ch| parsed.glyph_index(**ch).is_some())
                                .count()
                        })
                        .unwrap_or(0)
                })
                .unwrap_or(0)
        });
        let missing = best
            .and_then(|face| {
                database.with_face_data(face.id, |data, index| {
                    Face::parse(data, index)
                        .ok()
                        .map(|parsed| {
                            chars
                                .iter()
                                .filter(|ch| parsed.glyph_index(**ch).is_none())
                                .copied()
                                .collect::<Vec<_>>()
                        })
                        .unwrap_or_else(|| chars.iter().copied().collect())
                })
            })
            .unwrap_or_else(|| chars.iter().copied().collect());
        let sample = missing
            .iter()
            .take(8)
            .map(|ch| format!("U+{:04X}", *ch as u32))
            .collect::<Vec<_>>()
            .join(", ");
        LoError::Unsupported(format!(
            "no installed font can preserve all Unicode PDF text; missing {sample}; install Noto Sans CJK or set {FONT_OVERRIDE_ENV}"
        ))
    })?;

    database
        .with_face_data(selected.id, |data, index| {
            build_from_face(data, index, &selected.post_script_name, chars)
        })
        .ok_or_else(|| LoError::Io("reading selected Unicode PDF font".to_string()))?
}

fn build_from_face(
    data: &[u8],
    face_index: u32,
    post_script_name: &str,
    chars: &BTreeSet<char>,
) -> Result<EmbeddedUnicodeFont> {
    let face = Face::parse(data, face_index)
        .map_err(|_| LoError::Parse("selected Unicode PDF font is malformed".to_string()))?;
    let mut remapper = GlyphRemapper::new();
    let mut old_glyphs = BTreeMap::new();
    for ch in chars {
        let glyph = face.glyph_index(*ch).ok_or_else(|| {
            LoError::Unsupported(format!(
                "selected Unicode PDF font is missing U+{:04X}",
                *ch as u32
            ))
        })?;
        if let Some(previous) = old_glyphs.insert(glyph.0, *ch) {
            if previous != *ch {
                return Err(LoError::Unsupported(format!(
                    "Unicode PDF font aliases U+{:04X} and U+{:04X} to one glyph",
                    previous as u32, *ch as u32
                )));
            }
        }
        remapper.remap(glyph.0);
    }

    let units_per_em = face.units_per_em().max(1) as u32;
    let widths = remapper
        .remapped_gids()
        .map(|old_gid| {
            let advance = face.glyph_hor_advance(GlyphId(old_gid)).unwrap_or(units_per_em as u16);
            ((advance as u32 * 1000 + units_per_em / 2) / units_per_em).min(u16::MAX as u32)
                as u16
        })
        .collect::<Vec<_>>();
    let mut char_to_cid = BTreeMap::new();
    for ch in chars {
        let old_gid = face.glyph_index(*ch).expect("coverage checked").0;
        char_to_cid.insert(*ch, remapper.get(old_gid).expect("glyph remapped"));
    }
    let bbox = face.global_bounding_box();
    let scale_metric = |value: i16| -> i16 {
        ((value as i32 * 1000) / units_per_em as i32)
            .clamp(i16::MIN as i32, i16::MAX as i32) as i16
    };
    let subset = subset(data, face_index, &remapper)
        .map_err(|error| LoError::Unsupported(format!("subsetting Unicode PDF font: {error}")))?;
    let is_true_type = has_sfnt_table(&subset, b"glyf");
    if !is_true_type && !has_sfnt_table(&subset, b"CFF ") {
        return Err(LoError::Unsupported(
            "Unicode PDF font subset has neither TrueType nor CFF outlines".to_string(),
        ));
    }

    Ok(EmbeddedUnicodeFont {
        base_name: format!("LOUNIC+{}", pdf_name(post_script_name)),
        bytes: subset,
        is_true_type,
        widths,
        ascent: scale_metric(face.ascender()),
        descent: scale_metric(face.descender()),
        bbox: (
            scale_metric(bbox.x_min),
            scale_metric(bbox.y_min),
            scale_metric(bbox.x_max),
            scale_metric(bbox.y_max),
        ),
        italic_angle: face.italic_angle(),
        char_to_cid,
    })
}

fn font_preference(face: &fontdb::FaceInfo) -> u8 {
    let names = face
        .families
        .iter()
        .map(|(name, _)| name.to_ascii_lowercase())
        .chain(std::iter::once(face.post_script_name.to_ascii_lowercase()))
        .collect::<Vec<_>>()
        .join(" ");
    for (needle, score) in [
        ("noto sans cjk sc", 100),
        ("noto sans sc", 95),
        ("source han sans sc", 90),
        ("arial unicode", 85),
        ("pingfang sc", 80),
        ("hiragino sans gb", 75),
        ("droid sans fallback", 70),
        ("wenquanyi", 65),
    ] {
        if names.contains(needle) {
            return score;
        }
    }
    0
}

fn pdf_name(value: &str) -> String {
    let cleaned = value
        .chars()
        .filter(|ch| ch.is_ascii_alphanumeric() || matches!(ch, '-' | '_'))
        .collect::<String>();
    if cleaned.is_empty() {
        "UnicodeFont".to_string()
    } else {
        cleaned
    }
}

fn has_sfnt_table(bytes: &[u8], wanted: &[u8; 4]) -> bool {
    if bytes.len() < 12 {
        return false;
    }
    let count = u16::from_be_bytes([bytes[4], bytes[5]]) as usize;
    (0..count).any(|index| {
        let start = 12 + index * 16;
        bytes.get(start..start + 4) == Some(wanted.as_slice())
    })
}

fn utf16be_hex(ch: char) -> String {
    let mut units = [0_u16; 2];
    ch.encode_utf16(&mut units)
        .iter()
        .map(|unit| format!("{unit:04X}"))
        .collect::<String>()
}
