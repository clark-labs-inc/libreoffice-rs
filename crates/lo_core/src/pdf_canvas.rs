//! Multi-page PDF canvas built on top of [`pdf_from_objects`].
//!
//! This layer is still intentionally small, but it now exposes the
//! primitives needed by Clark's visual Writer/Impress verification path:
//! colored text, filled/stroked rectangles, ellipses, line width, and a
//! larger set of base fonts.

use crate::pdf::pdf_from_objects;
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
}

#[derive(Clone, Debug)]
pub struct PdfPage {
    pub width: f32,
    pub height: f32,
    commands: String,
}

impl PdfPage {
    pub fn new(width: f32, height: f32) -> Self {
        Self {
            width,
            height,
            commands: String::new(),
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
        self.commands.push_str("BT\n");
        self.commands
            .push_str(&format!("{r:.3} {g:.3} {b:.3} rg\n"));
        self.commands
            .push_str(&format!("/{} {} Tf\n", font.resource_name(), size));
        self.commands
            .push_str(&format!("1 0 0 1 {:.2} {:.2} Tm\n", x, y));
        self.commands.push('(');
        self.commands.push_str(&pdf_escape(text));
        self.commands.push_str(") Tj\nET\n0 0 0 rg\n");
    }

    pub fn line(&mut self, x1: f32, y1: f32, x2: f32, y2: f32) {
        self.commands
            .push_str(&format!("{:.2} {:.2} m {:.2} {:.2} l S\n", x1, y1, x2, y2));
    }

    pub fn line_rgb(&mut self, x1: f32, y1: f32, x2: f32, y2: f32, r: f32, g: f32, b: f32) {
        self.commands
            .push_str(&format!("{r:.3} {g:.3} {b:.3} RG\n{:.2} {:.2} m {:.2} {:.2} l S\n0 0 0 RG\n", x1, y1, x2, y2));
    }

    pub fn line_width(&mut self, width: f32) {
        self.commands.push_str(&format!("{width:.2} w\n"));
    }

    pub fn rect_stroke(&mut self, x: f32, y: f32, width: f32, height: f32) {
        self.commands
            .push_str(&format!("{:.2} {:.2} {:.2} {:.2} re S\n", x, y, width, height));
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
}

#[derive(Clone, Debug, Default)]
pub struct PdfDocument {
    pages: Vec<PdfPage>,
}

impl PdfDocument {
    pub fn new() -> Self {
        Self { pages: Vec::new() }
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

    pub fn finish(self) -> Vec<u8> {
        if self.pages.is_empty() {
            return empty_pdf();
        }
        let mut objects = Vec::new();
        objects.push(String::new()); // catalog
        objects.push(String::new()); // pages tree
        for font in [
            PdfFont::Helvetica,
            PdfFont::HelveticaBold,
            PdfFont::HelveticaOblique,
            PdfFont::Courier,
            PdfFont::TimesRoman,
            PdfFont::TimesBold,
            PdfFont::TimesItalic,
        ] {
            objects.push(format!(
                "<< /Type /Font /Subtype /Type1 /BaseFont /{} >>",
                font.base_font()
            ));
        }
        let page_start = 10usize;
        let mut kids = Vec::new();
        for (index, page) in self.pages.iter().enumerate() {
            let page_obj = page_start + index * 2;
            let content_obj = page_obj + 1;
            kids.push(format!("{} 0 R", page_obj));
            let resources = "<< /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R /F4 6 0 R /F5 7 0 R /F6 8 0 R /F7 9 0 R >> >>";
            objects.push(format!(
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {:.2} {:.2}] /Resources {} /Contents {} 0 R >>",
                page.width, page.height, resources, content_obj
            ));
            objects.push(format!(
                "<< /Length {} >>\nstream\n{}endstream",
                page.commands.len(), page.commands
            ));
        }
        objects[0] = "<< /Type /Catalog /Pages 2 0 R >>".to_string();
        objects[1] = format!(
            "<< /Type /Pages /Count {} /Kids [{}] >>",
            self.pages.len(),
            kids.join(" ")
        );
        pdf_from_objects(&objects)
    }
}

fn pdf_escape(value: &str) -> String {
    // Transliterate the typographic Unicode punctuation real-world DOCX
    // files routinely use into the closest ASCII equivalents. The built-in
    // Type1 fonts used by `PdfDocument` (Helvetica/Times/Courier with
    // StandardEncoding) cannot represent these code points, so emitting
    // their UTF-8 bytes directly produces glyph holes that `pdftotext`
    // either drops or splits awkwardly. ASCII equivalents survive the
    // round-trip and match what the LibreOffice CLI ends up extracting.
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        let mapped: &str = match ch {
            '\u{2018}' | '\u{2019}' | '\u{201A}' | '\u{2032}' => "'",
            '\u{201C}' | '\u{201D}' | '\u{201E}' | '\u{2033}' => "\"",
            '\u{2013}' | '\u{2014}' | '\u{2212}' => "-",
            '\u{2026}' => "...",
            '\u{00A0}' | '\u{2007}' | '\u{202F}' => " ",
            // Tab characters get rendered as wide whitespace so
            // `pdftotext` reliably treats them as a word break
            // (PDF tab escape would be valid but is rendered as
            // zero-width by some viewers).
            '\t' => "    ",
            '\u{00AD}' => "",
            '\u{2022}' => "*",
            _ => {
                if ch == '\\' {
                    out.push('\\');
                    out.push('\\');
                } else if ch == '(' {
                    out.push('\\');
                    out.push('(');
                } else if ch == ')' {
                    out.push('\\');
                    out.push(')');
                } else if (ch as u32) < 0x80 {
                    out.push(ch);
                } else {
                    // Out-of-range chars: emit nothing rather than raw
                    // UTF-8 bytes that the StandardEncoding font can't
                    // map. The Markdown / text extraction paths still see
                    // the original character upstream of this layer.
                }
                continue;
            }
        };
        out.push_str(mapped);
    }
    out
}

fn empty_pdf() -> Vec<u8> {
    let mut doc = PdfDocument::new();
    let page = doc.add_page(595.0, 842.0);
    let _ = doc
        .page_mut(page)
        .map(|p| p.text(50.0, 792.0, 12.0, PdfFont::Helvetica, ""));
    doc.finish()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn finish_emits_valid_pdf() {
        let mut doc = PdfDocument::new();
        let p = doc.add_page(595.0, 842.0);
        doc.page_mut(p)
            .unwrap()
            .text_rgb(50.0, 792.0, 12.0, PdfFont::TimesRoman, "hello", 1.0, 0.0, 0.0);
        let bytes = doc.finish();
        assert!(bytes.starts_with(b"%PDF-1.4"));
        assert!(bytes.ends_with(b"%%EOF\n"));
    }
}
