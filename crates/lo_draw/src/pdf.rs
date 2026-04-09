//! PDF export of a `Drawing`. Each page is summarized as text lines.

use lo_core::{units::Length, write_text_pdf, DrawElement, Drawing, ShapeKind};

pub fn to_pdf(drawing: &Drawing) -> Vec<u8> {
    let mut lines = Vec::new();
    lines.push(drawing.meta.title.clone());
    lines.push(String::new());
    for (idx, page) in drawing.pages.iter().enumerate() {
        lines.push(format!("Page {} — {}", idx + 1, page.name));
        for element in &page.elements {
            match element {
                DrawElement::TextBox(tb) => {
                    for line in tb.text.lines() {
                        lines.push(format!("  text: {line}"));
                    }
                }
                DrawElement::Shape(shape) => {
                    let kind = match shape.kind {
                        ShapeKind::Rectangle => "rect",
                        ShapeKind::Ellipse => "ellipse",
                        ShapeKind::Line => "line",
                    };
                    lines.push(format!(
                        "  shape: {kind} {}x{}",
                        shape.frame.size.width, shape.frame.size.height
                    ));
                }
                DrawElement::Image(image) => lines.push(format!("  image: {}", image.alt)),
            }
        }
        lines.push(String::new());
    }
    write_text_pdf(&lines, Length::pt(842.0), Length::pt(595.0))
}
