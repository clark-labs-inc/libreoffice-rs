mod chart;
pub mod html;
pub mod import;
pub mod markdown;
pub mod pdf;
pub mod pptx;
pub mod raster;
pub mod svg;

pub use html::to_html;
pub use import::{from_odp_bytes, from_pptx_bytes, from_text_outline, load_bytes};
pub use markdown::to_markdown;
pub use pdf::to_pdf;
pub use pptx::to_pptx;
pub use raster::{render_jpeg_pages, render_pages, render_png_pages};
pub use svg::render_svg;

use std::path::Path;

use lo_core::{
    Length, LoError, Presentation, Rect, Result, Shape, ShapeKind, ShapeStyle, Slide, SlideElement,
    TextBox, TextBoxStyle,
};

pub struct ImpressBuilder {
    pub presentation: Presentation,
}

impl ImpressBuilder {
    pub fn new(title: impl Into<String>) -> Self {
        Self {
            presentation: Presentation::new(title),
        }
    }

    pub fn add_slide(&mut self, slide: Slide) -> &mut Self {
        self.presentation.slides.push(slide);
        self
    }

    pub fn add_title_slide(&mut self, title: &str, subtitle: &str) -> &mut Self {
        self.presentation
            .slides
            .push(Presentation::title_slide(title, subtitle));
        self
    }

    pub fn add_bullet_slide(&mut self, title: &str, bullets: &[String]) -> &mut Self {
        let body = bullets
            .iter()
            .map(|line| format!("• {line}"))
            .collect::<Vec<_>>()
            .join("\n");
        self.presentation.slides.push(Slide {
            name: title.to_string(),
            elements: vec![
                SlideElement::TextBox(TextBox {
                    frame: Rect::new(
                        Length::mm(15.0),
                        Length::mm(15.0),
                        Length::mm(240.0),
                        Length::mm(20.0),
                    ),
                    text: title.to_string(),
                    style: TextBoxStyle::default(),
                }),
                SlideElement::TextBox(TextBox {
                    frame: Rect::new(
                        Length::mm(20.0),
                        Length::mm(45.0),
                        Length::mm(220.0),
                        Length::mm(80.0),
                    ),
                    text: body,
                    style: TextBoxStyle::default(),
                }),
            ],
            notes: Vec::new(),
            chart_tokens: Vec::new(),
        });
        self
    }

    pub fn add_diagram_slide(&mut self, title: &str) -> &mut Self {
        self.presentation.slides.push(Slide {
            name: title.to_string(),
            elements: vec![
                SlideElement::TextBox(TextBox {
                    frame: Rect::new(
                        Length::mm(20.0),
                        Length::mm(15.0),
                        Length::mm(220.0),
                        Length::mm(20.0),
                    ),
                    text: title.to_string(),
                    style: TextBoxStyle::default(),
                }),
                SlideElement::Shape(Shape {
                    frame: Rect::new(
                        Length::mm(30.0),
                        Length::mm(50.0),
                        Length::mm(50.0),
                        Length::mm(25.0),
                    ),
                    style: ShapeStyle {
                        fill: "#d9eaf7".to_string(),
                        stroke: "#1f4e79".to_string(),
                        stroke_width_mm: 1,
                    },
                    kind: ShapeKind::Rectangle,
                }),
                SlideElement::Shape(Shape {
                    frame: Rect::new(
                        Length::mm(110.0),
                        Length::mm(50.0),
                        Length::mm(50.0),
                        Length::mm(25.0),
                    ),
                    style: ShapeStyle {
                        fill: "#e2f0d9".to_string(),
                        stroke: "#375623".to_string(),
                        stroke_width_mm: 1,
                    },
                    kind: ShapeKind::Rectangle,
                }),
                SlideElement::Shape(Shape {
                    frame: Rect::new(
                        Length::mm(80.0),
                        Length::mm(62.0),
                        Length::mm(30.0),
                        Length::mm(0.0),
                    ),
                    style: ShapeStyle {
                        fill: String::new(),
                        stroke: "#666666".to_string(),
                        stroke_width_mm: 1,
                    },
                    kind: ShapeKind::Line,
                }),
                SlideElement::TextBox(TextBox {
                    frame: Rect::new(
                        Length::mm(38.0),
                        Length::mm(57.0),
                        Length::mm(35.0),
                        Length::mm(12.0),
                    ),
                    text: "Input".to_string(),
                    style: TextBoxStyle::default(),
                }),
                SlideElement::TextBox(TextBox {
                    frame: Rect::new(
                        Length::mm(118.0),
                        Length::mm(57.0),
                        Length::mm(35.0),
                        Length::mm(12.0),
                    ),
                    text: "Output".to_string(),
                    style: TextBoxStyle::default(),
                }),
            ],
            notes: vec!["Auto-generated diagram slide".to_string()],
            chart_tokens: Vec::new(),
        });
        self
    }

    pub fn save_odp(&self, path: impl AsRef<Path>) -> Result<()> {
        lo_odf::save_presentation_document(path, &self.presentation)
    }
}

pub fn demo_presentation(title: &str) -> Presentation {
    let mut builder = ImpressBuilder::new(title);
    builder
        .add_title_slide(title, "Generated by libreoffice-rs")
        .add_bullet_slide(
            "Highlights",
            &[
                "Pure Rust workspace".to_string(),
                "ODF packaging".to_string(),
                "Spreadsheet formulas".to_string(),
            ],
        )
        .add_diagram_slide("Architecture");
    builder.presentation
}

pub fn save_odp(path: impl AsRef<Path>, presentation: &Presentation) -> Result<()> {
    lo_odf::save_presentation_document(path, presentation)
}

/// Render the presentation into bytes for the requested format.
///
/// Supported (case-insensitive): `html`, `md`, `svg`, `pdf`, `odp`, `pptx`.
/// Multi-slide raster output is exposed through [`render_png_pages`] and
/// [`render_jpeg_pages`].
pub fn save_as(presentation: &Presentation, format: &str) -> Result<Vec<u8>> {
    match format.to_ascii_lowercase().as_str() {
        "html" => Ok(to_html(presentation).into_bytes()),
        "md" | "markdown" => Ok(to_markdown(presentation).into_bytes()),
        "svg" => Ok(render_svg(presentation).into_bytes()),
        "pdf" => Ok(to_pdf(presentation)),
        "odp" => {
            let tmp = std::env::temp_dir().join(format!(
                "lo_impress_{}.odp",
                std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|d| d.as_nanos())
                    .unwrap_or(0)
            ));
            lo_odf::save_presentation_document(&tmp, presentation)?;
            let bytes = std::fs::read(&tmp)?;
            let _ = std::fs::remove_file(&tmp);
            Ok(bytes)
        }
        "pptx" => to_pptx(presentation),
        other => Err(LoError::Unsupported(format!(
            "impress format not supported: {other}"
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn pptx_export_is_a_zip_archive() {
        let mut b = ImpressBuilder::new("Demo");
        b.add_title_slide("Hello", "world")
            .add_bullet_slide("Items", &["one".to_string(), "two".to_string()]);
        let bytes = to_pptx(&b.presentation).expect("pptx");
        assert!(bytes.starts_with(b"PK"));
    }

    #[test]
    fn save_as_dispatches_by_format() {
        let p = demo_presentation("Demo");
        for fmt in ["html", "svg", "pdf", "odp", "pptx"] {
            let bytes = save_as(&p, fmt).unwrap_or_else(|e| panic!("{fmt}: {e}"));
            assert!(!bytes.is_empty(), "{fmt} produced empty output");
        }
        assert!(save_as(&p, "qqq").is_err());
    }
}
