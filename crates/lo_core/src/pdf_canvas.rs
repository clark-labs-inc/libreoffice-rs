//! Multi-page PDF canvas built on top of [`crate::pdf::pdf_from_objects`].
//!
//! This layer is still intentionally small, but it now exposes the
//! primitives needed by Clark's visual Writer/Impress verification path:
//! colored text, filled/stroked rectangles, ellipses, line width, and a
//! larger set of base fonts.

use std::collections::BTreeSet;

use crate::pdf::{pdf_can_encode_win_ansi, pdf_escape_win_ansi, pdf_text_string};
use crate::pdf::pdf_from_binary_objects;
use crate::unicode_font::{build_unicode_font, EmbeddedUnicodeFont};
use crate::{LoError, Result};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PdfFont {
    Helvetica,
    HelveticaBold,
    HelveticaOblique,
    Courier,
    TimesRoman,
    TimesBold,
    TimesItalic,
}

impl PdfFont {
    fn resource_name(self) -> &'static str {
        match self {
            Self::Helvetica => "F1",
            Self::HelveticaBold => "F2",
            Self::HelveticaOblique => "F3",
            Self::Courier => "F4",
            Self::TimesRoman => "F5",
            Self::TimesBold => "F6",
            Self::TimesItalic => "F7",
        }
    }

    fn base_font(self) -> &'static str {
        match self {
            Self::Helvetica => "Helvetica",
            Self::HelveticaBold => "Helvetica-Bold",
            Self::HelveticaOblique => "Helvetica-Oblique",
            Self::Courier => "Courier",
            Self::TimesRoman => "Times-Roman",
            Self::TimesBold => "Times-Bold",
            Self::TimesItalic => "Times-Italic",
        }
    }

    fn is_bold(self) -> bool {
        matches!(self, Self::HelveticaBold | Self::TimesBold)
    }

    fn is_italic(self) -> bool {
        matches!(
            self,
            Self::HelveticaOblique | Self::TimesItalic
        )
    }
}

#[derive(Clone, Debug)]
pub struct PdfPage {
    pub width: f32,
    pub height: f32,
    commands: String,
    unicode_texts: Vec<String>,
    links: Vec<PdfLink>,
    tags: Vec<PdfTag>,
}

impl PdfPage {
    pub fn new(width: f32, height: f32) -> Self {
        Self {
            width,
            height,
            commands: String::new(),
            unicode_texts: Vec::new(),
            links: Vec::new(),
            tags: Vec::new(),
        }
    }

    pub fn text(&mut self, x: f32, y: f32, size: f32, font: PdfFont, text: &str) {
        self.text_rgb(x, y, size, font, text, 0.0, 0.0, 0.0);
    }

    pub fn text_rgb(
        &mut self,
        x: f32,
        y: f32,
        size: f32,
        font: PdfFont,
        text: &str,
        r: f32,
        g: f32,
        b: f32,
    ) {
        let unicode = !pdf_can_encode_win_ansi(text);
        let resource_name = if unicode { "FU" } else { font.resource_name() };
        self.commands.push_str("BT\n");
        self.commands
            .push_str(&format!("{r:.3} {g:.3} {b:.3} rg\n"));
        if unicode && font.is_bold() {
            self.commands.push_str("0.35 w\n2 Tr\n");
        }
        self.commands
            .push_str(&format!("/{resource_name} {size} Tf\n"));
        let shear = if unicode && font.is_italic() { 0.20 } else { 0.0 };
        self.commands
            .push_str(&format!("1 0 {shear:.2} 1 {x:.2} {y:.2} Tm\n"));
        if unicode {
            let index = self.unicode_texts.len();
            self.unicode_texts.push(text.to_string());
            self.commands
                .push_str(&format!("__LO_UNICODE_TEXT_{index}__ Tj\n"));
        } else {
            self.commands.push('(');
            self.commands.push_str(&pdf_escape_win_ansi(text));
            self.commands.push_str(") Tj\n");
        }
        if unicode && font.is_bold() {
            self.commands.push_str("0 Tr\n1 w\n");
        }
        self.commands.push_str("ET\n0 0 0 rg\n");
    }

    pub fn line(&mut self, x1: f32, y1: f32, x2: f32, y2: f32) {
        self.commands
            .push_str(&format!("{:.2} {:.2} m {:.2} {:.2} l S\n", x1, y1, x2, y2));
    }

    pub fn line_rgb(&mut self, x1: f32, y1: f32, x2: f32, y2: f32, r: f32, g: f32, b: f32) {
        self.commands.push_str(&format!(
            "{r:.3} {g:.3} {b:.3} RG\n{:.2} {:.2} m {:.2} {:.2} l S\n0 0 0 RG\n",
            x1, y1, x2, y2
        ));
    }

    pub fn line_width(&mut self, width: f32) {
        self.commands.push_str(&format!("{width:.2} w\n"));
    }

    pub fn rect_stroke(&mut self, x: f32, y: f32, width: f32, height: f32) {
        self.commands.push_str(&format!(
            "{:.2} {:.2} {:.2} {:.2} re S\n",
            x, y, width, height
        ));
    }

    pub fn rect_stroke_rgb(
        &mut self,
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        r: f32,
        g: f32,
        b: f32,
    ) {
        self.commands.push_str(&format!(
            "{r:.3} {g:.3} {b:.3} RG\n{:.2} {:.2} {:.2} {:.2} re S\n0 0 0 RG\n",
            x, y, width, height
        ));
    }

    pub fn rect_fill_rgb(
        &mut self,
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        r: f32,
        g: f32,
        b: f32,
    ) {
        self.commands.push_str(&format!(
            "{r:.3} {g:.3} {b:.3} rg\n{:.2} {:.2} {:.2} {:.2} re f\n0 0 0 rg\n",
            x, y, width, height
        ));
    }

    pub fn rect_fill_stroke_rgb(
        &mut self,
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        fill: (f32, f32, f32),
        stroke: (f32, f32, f32),
    ) {
        self.commands.push_str(&format!(
            "{:.3} {:.3} {:.3} rg\n{:.3} {:.3} {:.3} RG\n{:.2} {:.2} {:.2} {:.2} re B\n0 0 0 rg\n0 0 0 RG\n",
            fill.0, fill.1, fill.2, stroke.0, stroke.1, stroke.2, x, y, width, height
        ));
    }

    pub fn polygon_fill_stroke_rgb(
        &mut self,
        points: &[(f32, f32)],
        fill: (f32, f32, f32),
        stroke: (f32, f32, f32),
    ) {
        let Some((first, rest)) = points.split_first() else {
            return;
        };
        if points.len() < 3 {
            return;
        }
        self.commands.push_str(&format!(
            "{:.3} {:.3} {:.3} rg\n{:.3} {:.3} {:.3} RG\n{:.2} {:.2} m\n",
            fill.0, fill.1, fill.2, stroke.0, stroke.1, stroke.2, first.0, first.1
        ));
        for point in rest {
            self.commands
                .push_str(&format!("{:.2} {:.2} l\n", point.0, point.1));
        }
        self.commands.push_str("h B\n0 0 0 rg\n0 0 0 RG\n");
    }

    pub fn ellipse_stroke_rgb(
        &mut self,
        cx: f32,
        cy: f32,
        rx: f32,
        ry: f32,
        r: f32,
        g: f32,
        b: f32,
    ) {
        let kappa = 0.552_284_8_f32;
        let ox = rx * kappa;
        let oy = ry * kappa;
        self.commands.push_str(&format!(
            "{r:.3} {g:.3} {b:.3} RG\n{:.2} {:.2} m\n{:.2} {:.2} {:.2} {:.2} {:.2} {:.2} c\n{:.2} {:.2} {:.2} {:.2} {:.2} {:.2} c\n{:.2} {:.2} {:.2} {:.2} {:.2} {:.2} c\n{:.2} {:.2} {:.2} {:.2} {:.2} {:.2} c\nS\n0 0 0 RG\n",
            cx - rx,
            cy,
            cx - rx,
            cy + oy,
            cx - ox,
            cy + ry,
            cx,
            cy + ry,
            cx + ox,
            cy + ry,
            cx + rx,
            cy + oy,
            cx + rx,
            cy,
            cx + rx,
            cy - oy,
            cx + ox,
            cy - ry,
            cx,
            cy - ry,
            cx - ox,
            cy - ry,
            cx - rx,
            cy - oy,
            cx - rx,
            cy
        ));
    }

    pub fn ellipse_fill_stroke_rgb(
        &mut self,
        cx: f32,
        cy: f32,
        rx: f32,
        ry: f32,
        fill: (f32, f32, f32),
        stroke: (f32, f32, f32),
    ) {
        let kappa = 0.552_284_8_f32;
        let ox = rx * kappa;
        let oy = ry * kappa;
        self.commands.push_str(&format!(
            "{:.3} {:.3} {:.3} rg\n{:.3} {:.3} {:.3} RG\n{:.2} {:.2} m\n{:.2} {:.2} {:.2} {:.2} {:.2} {:.2} c\n{:.2} {:.2} {:.2} {:.2} {:.2} {:.2} c\n{:.2} {:.2} {:.2} {:.2} {:.2} {:.2} c\n{:.2} {:.2} {:.2} {:.2} {:.2} {:.2} c\nB\n0 0 0 rg\n0 0 0 RG\n",
            fill.0,
            fill.1,
            fill.2,
            stroke.0,
            stroke.1,
            stroke.2,
            cx - rx,
            cy,
            cx - rx,
            cy + oy,
            cx - ox,
            cy + ry,
            cx,
            cy + ry,
            cx + ox,
            cy + ry,
            cx + rx,
            cy + oy,
            cx + rx,
            cy,
            cx + rx,
            cy - oy,
            cx + ox,
            cy - ry,
            cx,
            cy - ry,
            cx - ox,
            cy - ry,
            cx - rx,
            cy - oy,
            cx - rx,
            cy
        ));
    }

    pub fn raw(&mut self, command: &str) {
        self.commands.push_str(command);
        if !command.ends_with('\n') {
            self.commands.push('\n');
        }
    }

    pub fn image(&mut self, resource_name: &str, x: f32, y: f32, width: f32, height: f32) {
        self.commands.push_str(&format!(
            "q\n{width:.2} 0 0 {height:.2} {x:.2} {y:.2} cm\n/{resource_name} Do\nQ\n"
        ));
    }

    pub fn link(&mut self, x: f32, y: f32, width: f32, height: f32, uri: impl Into<String>) {
        self.links.push(PdfLink {
            x,
            y,
            width,
            height,
            uri: uri.into(),
        });
    }

    /// Begin a semantically tagged marked-content sequence. Each call must be
    /// paired with [`PdfPage::end_tag`]. The role should be a standard PDF
    /// structure type such as `P`, `H1`, `LI`, `Table`, `Figure`, or `Code`.
    pub fn begin_tag(&mut self, role: &str, alt: Option<&str>) {
        let role = role
            .chars()
            .filter(|ch| ch.is_ascii_alphanumeric())
            .collect::<String>();
        let role = if role.is_empty() { "Span" } else { &role };
        let mcid = self.tags.len();
        self.tags.push(PdfTag {
            role: role.to_string(),
            alt: alt.map(str::to_string),
            mcid,
        });
        self.commands
            .push_str(&format!("/{role} <</MCID {mcid}>> BDC\n"));
    }

    pub fn end_tag(&mut self) {
        self.commands.push_str("EMC\n");
    }

    pub fn begin_artifact(&mut self) {
        self.commands.push_str("/Artifact BMC\n");
    }

    pub fn end_artifact(&mut self) {
        self.commands.push_str("EMC\n");
    }
}

#[derive(Clone, Debug)]
struct PdfLink {
    x: f32,
    y: f32,
    width: f32,
    height: f32,
    uri: String,
}

#[derive(Clone, Debug)]
struct PdfTag {
    role: String,
    alt: Option<String>,
    mcid: usize,
}

#[derive(Clone, Debug)]
struct PdfImage {
    width: u32,
    height: u32,
    bytes: Vec<u8>,
    jpeg: bool,
}

#[derive(Clone, Debug, Default)]
pub struct PdfDocument {
    pages: Vec<PdfPage>,
    images: Vec<PdfImage>,
}

impl PdfDocument {
    pub fn new() -> Self {
        Self {
            pages: Vec::new(),
            images: Vec::new(),
        }
    }

    pub fn add_page(&mut self, width: f32, height: f32) -> usize {
        self.pages.push(PdfPage::new(width, height));
        self.pages.len() - 1
    }

    pub fn page_mut(&mut self, index: usize) -> Result<&mut PdfPage> {
        self.pages
            .get_mut(index)
            .ok_or_else(|| LoError::InvalidInput(format!("pdf page index out of range: {index}")))
    }

    pub fn pages(&self) -> &[PdfPage] {
        &self.pages
    }

    pub fn add_rgb_image(&mut self, width: u32, height: u32, bytes: Vec<u8>) -> String {
        self.images.push(PdfImage {
            width,
            height,
            bytes,
            jpeg: false,
        });
        format!("Im{}", self.images.len())
    }

    pub fn add_jpeg_image(&mut self, width: u32, height: u32, bytes: Vec<u8>) -> String {
        self.images.push(PdfImage {
            width,
            height,
            bytes,
            jpeg: true,
        });
        format!("Im{}", self.images.len())
    }

    pub fn finish(self) -> Vec<u8> {
        self.try_finish()
            .expect("PDF text contains Unicode that no installed font can preserve")
    }

    pub fn try_finish(self) -> Result<Vec<u8>> {
        if self.pages.is_empty() {
            return empty_pdf();
        }
        let unicode_chars = self
            .pages
            .iter()
            .flat_map(|page| page.unicode_texts.iter())
            .flat_map(|text| text.chars())
            .collect::<BTreeSet<_>>();
        let unicode_font = if unicode_chars.is_empty() {
            None
        } else {
            Some(build_unicode_font(&unicode_chars)?)
        };
        let mut objects: Vec<Vec<u8>> = Vec::new();
        objects.push(Vec::new()); // catalog
        objects.push(Vec::new()); // pages tree
        for font in [
            PdfFont::Helvetica,
            PdfFont::HelveticaBold,
            PdfFont::HelveticaOblique,
            PdfFont::Courier,
            PdfFont::TimesRoman,
            PdfFont::TimesBold,
            PdfFont::TimesItalic,
        ] {
            objects.push(
                format!(
                    "<< /Type /Font /Subtype /Type1 /BaseFont /{} /Encoding /WinAnsiEncoding >>",
                    font.base_font()
                )
                .into_bytes(),
            );
        }
        let unicode_font_obj = unicode_font
            .as_ref()
            .map(|font| append_unicode_font_objects(&mut objects, font));
        let page_start = objects.len() + 1;
        let image_start = page_start + self.pages.len() * 2;
        let annotation_start = image_start + self.images.len();
        let annotation_count = self
            .pages
            .iter()
            .map(|page| page.links.len())
            .sum::<usize>();
        let structure_element_start = annotation_start + annotation_count;
        let structure_element_count = self.pages.iter().map(|page| page.tags.len()).sum::<usize>();
        let parent_tree_obj = structure_element_start + structure_element_count;
        let structure_tree_obj = parent_tree_obj + 1;
        let xobjects = if self.images.is_empty() {
            String::new()
        } else {
            format!(
                " /XObject << {} >>",
                self.images
                    .iter()
                    .enumerate()
                    .map(|(index, _)| format!("/Im{} {} 0 R", index + 1, image_start + index))
                    .collect::<Vec<_>>()
                    .join(" ")
            )
        };
        let mut kids = Vec::new();
        let mut annotations_before_page = 0usize;
        for (index, page) in self.pages.iter().enumerate() {
            let page_obj = page_start + index * 2;
            let content_obj = page_obj + 1;
            kids.push(format!("{} 0 R", page_obj));
            let unicode_resource = unicode_font_obj
                .map(|object| format!(" /FU {object} 0 R"))
                .unwrap_or_default();
            let resources = format!("<< /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R /F4 6 0 R /F5 7 0 R /F6 8 0 R /F7 9 0 R{unicode_resource} >>{xobjects} >>");
            let annotations = if page.links.is_empty() {
                String::new()
            } else {
                format!(
                    " /Annots [{}]",
                    (0..page.links.len())
                        .map(|offset| format!(
                            "{} 0 R",
                            annotation_start + annotations_before_page + offset
                        ))
                        .collect::<Vec<_>>()
                        .join(" ")
                )
            };
            objects.push(format!(
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {:.2} {:.2}] /Resources {} /Contents {} 0 R /StructParents {}{} >>",
                page.width, page.height, resources, content_obj, index, annotations
            ).into_bytes());
            let commands = render_page_commands(page, unicode_font.as_ref())?;
            objects.push(
                format!(
                    "<< /Length {} >>\nstream\n{}endstream",
                    commands.len(),
                    commands
                )
                .into_bytes(),
            );
            annotations_before_page += page.links.len();
        }
        for image in &self.images {
            let filter = if image.jpeg {
                " /Filter /DCTDecode"
            } else {
                ""
            };
            let mut object = format!(
                "<< /Type /XObject /Subtype /Image /Width {} /Height {} /ColorSpace /DeviceRGB /BitsPerComponent 8{filter} /Length {} >>\nstream\n",
                image.width, image.height, image.bytes.len()
            ).into_bytes();
            object.extend_from_slice(&image.bytes);
            object.extend_from_slice(b"\nendstream");
            objects.push(object);
        }
        for (page_index, page) in self.pages.iter().enumerate() {
            let page_obj = page_start + page_index * 2;
            for link in &page.links {
                let uri = pdf_text_string(&link.uri);
                objects.push(
                    format!(
                        "<< /Type /Annot /Subtype /Link /P {page_obj} 0 R /Rect [{:.2} {:.2} {:.2} {:.2}] /Border [0 0 0] /A << /S /URI /URI {uri} >> >>",
                        link.x,
                        link.y,
                        link.x + link.width,
                        link.y + link.height
                    )
                    .into_bytes(),
                );
            }
        }
        let mut structure_refs = Vec::new();
        let mut parent_tree_entries = Vec::new();
        let mut tags_before_page = 0usize;
        for (page_index, page) in self.pages.iter().enumerate() {
            let page_obj = page_start + page_index * 2;
            let refs = (0..page.tags.len())
                .map(|offset| structure_element_start + tags_before_page + offset)
                .collect::<Vec<_>>();
            parent_tree_entries.push(format!(
                "{page_index} [{}]",
                refs.iter()
                    .map(|object| format!("{object} 0 R"))
                    .collect::<Vec<_>>()
                    .join(" ")
            ));
            for (tag, object) in page.tags.iter().zip(refs) {
                let alt = tag
                    .alt
                    .as_ref()
                    .map_or_else(String::new, |value| format!(" /Alt {}", pdf_text_string(value)));
                objects.push(
                    format!(
                        "<< /Type /StructElem /S /{} /P {} 0 R /Pg {} 0 R /K {}{} >>",
                        tag.role, structure_tree_obj, page_obj, tag.mcid, alt
                    )
                    .into_bytes(),
                );
                structure_refs.push(format!("{object} 0 R"));
            }
            tags_before_page += page.tags.len();
        }
        objects.push(format!("<< /Nums [{}] >>", parent_tree_entries.join(" ")).into_bytes());
        objects.push(
            format!(
                "<< /Type /StructTreeRoot /K [{}] /ParentTree {} 0 R /ParentTreeNextKey {} >>",
                structure_refs.join(" "),
                parent_tree_obj,
                self.pages.len()
            )
            .into_bytes(),
        );
        let language = document_language(&unicode_chars);
        objects[0] = format!(
            "<< /Type /Catalog /Pages 2 0 R /Lang ({language}) /MarkInfo << /Marked true >> /StructTreeRoot {} 0 R >>",
            structure_tree_obj
        )
        .into_bytes();
        objects[1] = format!(
            "<< /Type /Pages /Count {} /Kids [{}] >>",
            self.pages.len(),
            kids.join(" ")
        )
        .into_bytes();
        Ok(pdf_from_binary_objects(&objects))
    }
}

fn empty_pdf() -> Result<Vec<u8>> {
    let mut doc = PdfDocument::new();
    let page = doc.add_page(595.0, 842.0);
    let _ = doc
        .page_mut(page)
        .map(|p| p.text(50.0, 792.0, 12.0, PdfFont::Helvetica, ""));
    doc.try_finish()
}

fn render_page_commands(
    page: &PdfPage,
    unicode_font: Option<&EmbeddedUnicodeFont>,
) -> Result<String> {
    let mut commands = page.commands.clone();
    for (index, text) in page.unicode_texts.iter().enumerate() {
        let font = unicode_font.ok_or_else(|| {
            LoError::InvalidInput("Unicode PDF text is missing its embedded font".to_string())
        })?;
        commands = commands.replace(
            &format!("__LO_UNICODE_TEXT_{index}__"),
            &font.encode_hex(text)?,
        );
    }
    Ok(commands)
}

fn append_unicode_font_objects(
    objects: &mut Vec<Vec<u8>>,
    font: &EmbeddedUnicodeFont,
) -> usize {
    let type0_obj = objects.len() + 1;
    let descendant_obj = type0_obj + 1;
    let descriptor_obj = type0_obj + 2;
    let font_file_obj = type0_obj + 3;
    let to_unicode_obj = type0_obj + 4;
    let descendant_subtype = if font.is_true_type {
        "CIDFontType2"
    } else {
        "CIDFontType0"
    };
    let font_file_key = if font.is_true_type {
        "FontFile2"
    } else {
        "FontFile3"
    };
    let font_file_extra = if font.is_true_type {
        format!(" /Length1 {}", font.bytes.len())
    } else {
        " /Subtype /OpenType".to_string()
    };
    objects.push(
        format!(
            "<< /Type /Font /Subtype /Type0 /BaseFont /{} /Encoding /Identity-H /DescendantFonts [{} 0 R] /ToUnicode {} 0 R >>",
            font.base_name, descendant_obj, to_unicode_obj
        )
        .into_bytes(),
    );
    objects.push(
        format!(
            "<< /Type /Font /Subtype /{descendant_subtype} /BaseFont /{} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor {} 0 R /W [0 [{}]]{} >>",
            font.base_name,
            descriptor_obj,
            font.widths_array(),
            if font.is_true_type { " /CIDToGIDMap /Identity" } else { "" }
        )
        .into_bytes(),
    );
    objects.push(
        format!(
            "<< /Type /FontDescriptor /FontName /{} /Flags 32 /FontBBox [{} {} {} {}] /ItalicAngle {:.2} /Ascent {} /Descent {} /CapHeight {} /StemV 80 /{} {} 0 R >>",
            font.base_name,
            font.bbox.0,
            font.bbox.1,
            font.bbox.2,
            font.bbox.3,
            font.italic_angle,
            font.ascent,
            font.descent,
            font.ascent,
            font_file_key,
            font_file_obj
        )
        .into_bytes(),
    );
    let mut font_stream = format!(
        "<< /Length {}{} >>\nstream\n",
        font.bytes.len(),
        font_file_extra
    )
    .into_bytes();
    font_stream.extend_from_slice(&font.bytes);
    font_stream.extend_from_slice(b"\nendstream");
    objects.push(font_stream);
    let cmap = font.to_unicode_cmap();
    let mut cmap_stream = format!("<< /Length {} >>\nstream\n", cmap.len()).into_bytes();
    cmap_stream.extend_from_slice(&cmap);
    cmap_stream.extend_from_slice(b"endstream");
    objects.push(cmap_stream);
    type0_obj
}

fn document_language(chars: &BTreeSet<char>) -> &'static str {
    if chars.iter().any(|ch| matches!(*ch as u32, 0x4E00..=0x9FFF | 0x3400..=0x4DBF)) {
        "zh-CN"
    } else if chars.iter().any(|ch| matches!(*ch as u32, 0x3040..=0x30FF)) {
        "ja"
    } else if chars.iter().any(|ch| matches!(*ch as u32, 0xAC00..=0xD7AF)) {
        "ko"
    } else if chars.is_empty() {
        "en-US"
    } else {
        "und"
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn finish_emits_valid_pdf() {
        let mut doc = PdfDocument::new();
        let p = doc.add_page(595.0, 842.0);
        doc.page_mut(p).unwrap().text_rgb(
            50.0,
            792.0,
            12.0,
            PdfFont::TimesRoman,
            "hello",
            1.0,
            0.0,
            0.0,
        );
        let bytes = doc.finish();
        assert!(bytes.starts_with(b"%PDF-1.4"));
        assert!(bytes.ends_with(b"%%EOF\n"));
    }

    #[test]
    fn finish_emits_clickable_link_annotation() {
        let mut doc = PdfDocument::new();
        let page = doc.add_page(300.0, 200.0);
        doc.page_mut(page)
            .unwrap()
            .link(20.0, 40.0, 80.0, 14.0, "https://example.com");
        let bytes = doc.finish();
        assert!(bytes
            .windows(b"/Subtype /Link".len())
            .any(|window| window == b"/Subtype /Link"));
        assert!(bytes
            .windows(b"https://example.com".len())
            .any(|window| window == b"https://example.com"));
    }

    #[test]
    fn finish_emits_tagged_pdf_structure_tree() {
        let mut doc = PdfDocument::new();
        let page = doc.add_page(300.0, 200.0);
        let page = doc.page_mut(page).unwrap();
        page.begin_tag("H1", Some("Accessible heading"));
        page.text(20.0, 160.0, 16.0, PdfFont::HelveticaBold, "Heading");
        page.end_tag();
        let bytes = doc.finish();
        for marker in [
            b"/StructTreeRoot".as_slice(),
            b"/MarkInfo".as_slice(),
            b"/Marked true".as_slice(),
            b"/StructParents 0".as_slice(),
            b"/S /H1".as_slice(),
            b"/MCID 0".as_slice(),
        ] {
            assert!(bytes.windows(marker.len()).any(|window| window == marker));
        }
    }
}
