pub mod import;
pub mod pdf;
pub mod svg;

pub use import::{from_odg_bytes, from_svg, load_bytes};
pub use pdf::to_pdf;
pub use svg::render_svg;

use std::path::Path;

use lo_core::{
    DrawElement, DrawPage, Drawing, Length, LoError, Rect, Result, Shape, ShapeKind, ShapeStyle,
    TextBox, TextBoxStyle,
};

pub struct DrawBuilder {
    pub drawing: Drawing,
}

impl DrawBuilder {
    pub fn new(title: impl Into<String>) -> Self {
        Self {
            drawing: Drawing::new(title),
        }
    }

    pub fn current_page_mut(&mut self) -> &mut DrawPage {
        if self.drawing.pages.is_empty() {
            self.drawing.pages.push(DrawPage::default());
        }
        self.drawing.pages.last_mut().expect("page exists")
    }

    pub fn add_page(&mut self, name: &str) -> &mut Self {
        self.drawing.pages.push(DrawPage {
            name: name.to_string(),
            elements: Vec::new(),
        });
        self
    }

    pub fn add_text_box(
        &mut self,
        text: &str,
        x_mm: f32,
        y_mm: f32,
        w_mm: f32,
        h_mm: f32,
    ) -> &mut Self {
        self.current_page_mut()
            .elements
            .push(DrawElement::TextBox(TextBox {
                frame: Rect::new(
                    Length::mm(x_mm),
                    Length::mm(y_mm),
                    Length::mm(w_mm),
                    Length::mm(h_mm),
                ),
                text: text.to_string(),
                style: TextBoxStyle::default(),
            }));
        self
    }

    pub fn add_rectangle(
        &mut self,
        x_mm: f32,
        y_mm: f32,
        w_mm: f32,
        h_mm: f32,
        fill: &str,
    ) -> &mut Self {
        self.current_page_mut()
            .elements
            .push(DrawElement::Shape(Shape {
                frame: Rect::new(
                    Length::mm(x_mm),
                    Length::mm(y_mm),
                    Length::mm(w_mm),
                    Length::mm(h_mm),
                ),
                style: ShapeStyle {
                    fill: fill.to_string(),
                    stroke: "#222222".to_string(),
                    stroke_width_mm: 1,
                },
                kind: ShapeKind::Rectangle,
            }));
        self
    }

    pub fn add_ellipse(
        &mut self,
        x_mm: f32,
        y_mm: f32,
        w_mm: f32,
        h_mm: f32,
        fill: &str,
    ) -> &mut Self {
        self.current_page_mut()
            .elements
            .push(DrawElement::Shape(Shape {
                frame: Rect::new(
                    Length::mm(x_mm),
                    Length::mm(y_mm),
                    Length::mm(w_mm),
                    Length::mm(h_mm),
                ),
                style: ShapeStyle {
                    fill: fill.to_string(),
                    stroke: "#222222".to_string(),
                    stroke_width_mm: 1,
                },
                kind: ShapeKind::Ellipse,
            }));
        self
    }

    pub fn add_line(&mut self, x1_mm: f32, y1_mm: f32, x2_mm: f32, y2_mm: f32) -> &mut Self {
        self.current_page_mut()
            .elements
            .push(DrawElement::Shape(Shape {
                frame: Rect::new(
                    Length::mm(x1_mm),
                    Length::mm(y1_mm),
                    Length::mm(x2_mm - x1_mm),
                    Length::mm(y2_mm - y1_mm),
                ),
                style: ShapeStyle {
                    fill: String::new(),
                    stroke: "#333333".to_string(),
                    stroke_width_mm: 1,
                },
                kind: ShapeKind::Line,
            }));
        self
    }

    pub fn save_odg(&self, path: impl AsRef<Path>) -> Result<()> {
        lo_odf::save_drawing_document(path, &self.drawing)
    }
}

pub fn demo_drawing(title: &str) -> Drawing {
    let mut builder = DrawBuilder::new(title);
    let _ = title;
    builder
        .add_text_box("libreoffice-rs", 20.0, 15.0, 120.0, 12.0)
        .add_rectangle(20.0, 40.0, 40.0, 20.0, "#d9eaf7")
        .add_text_box("Writer", 22.0, 45.0, 36.0, 10.0)
        .add_rectangle(80.0, 40.0, 40.0, 20.0, "#fce4d6")
        .add_text_box("Calc", 82.0, 45.0, 36.0, 10.0)
        .add_rectangle(140.0, 40.0, 40.0, 20.0, "#e2f0d9")
        .add_text_box("Impress", 142.0, 45.0, 36.0, 10.0)
        .add_line(60.0, 50.0, 80.0, 50.0)
        .add_line(120.0, 50.0, 140.0, 50.0);
    builder.drawing
}

pub fn save_odg(path: impl AsRef<Path>, drawing: &Drawing) -> Result<()> {
    lo_odf::save_drawing_document(path, drawing)
}

/// Render the drawing into bytes for the requested format.
///
/// Supported (case-insensitive): `svg`, `pdf`, `odg`.
pub fn save_as(drawing: &Drawing, format: &str) -> Result<Vec<u8>> {
    match format.to_ascii_lowercase().as_str() {
        "svg" => Ok(render_svg(drawing).into_bytes()),
        "pdf" => Ok(to_pdf(drawing)),
        "odg" => {
            let tmp = std::env::temp_dir().join(format!(
                "lo_draw_{}.odg",
                std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|d| d.as_nanos())
                    .unwrap_or(0)
            ));
            lo_odf::save_drawing_document(&tmp, drawing)?;
            let bytes = std::fs::read(&tmp)?;
            let _ = std::fs::remove_file(&tmp);
            Ok(bytes)
        }
        other => Err(LoError::Unsupported(format!(
            "draw format not supported: {other}"
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn save_as_dispatches_by_format() {
        let d = demo_drawing("Demo");
        for fmt in ["svg", "pdf", "odg"] {
            let bytes = save_as(&d, fmt).unwrap_or_else(|e| panic!("{fmt}: {e}"));
            assert!(!bytes.is_empty(), "{fmt} produced empty output");
        }
        assert!(save_as(&d, "qqq").is_err());
    }

    #[test]
    fn svg_render_includes_each_page() {
        let mut b = DrawBuilder::new("D");
        b.add_rectangle(10.0, 10.0, 30.0, 30.0, "#ff0000")
            .add_page("Two")
            .add_ellipse(20.0, 20.0, 40.0, 20.0, "#00ff00");
        let svg = render_svg(&b.drawing);
        assert!(svg.contains("<rect"));
        assert!(svg.contains("<ellipse"));
        assert!(svg.contains("Two"));
    }
}
