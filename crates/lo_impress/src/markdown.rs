use lo_core::{Presentation, ShapeKind, SlideElement};

use crate::chart::append_chart_rows_markdown;

pub fn to_markdown(presentation: &Presentation) -> String {
    let mut out = String::new();
    if !presentation.meta.title.trim().is_empty() {
        out.push_str(&format!("# {}\n\n", presentation.meta.title));
    }
    for (index, slide) in presentation.slides.iter().enumerate() {
        out.push_str(&format!("## Slide {}: {}\n\n", index + 1, slide.name));
        let mut bullets = Vec::new();
        for element in &slide.elements {
            match element {
                SlideElement::TextBox(text_box) => {
                    let text = text_box.text.trim();
                    if !text.is_empty() {
                        bullets.push(text.to_string());
                    }
                }
                SlideElement::Shape(shape) => bullets.push(format!(
                    "[shape: {}]",
                    match shape.kind {
                        ShapeKind::Rectangle => "rectangle",
                        ShapeKind::Ellipse => "ellipse",
                        ShapeKind::Line => "line",
                    }
                )),
                SlideElement::Image(image) => {
                    bullets.push(format!("![{}]({})", image.alt, image.name))
                }
            }
        }
        for bullet in bullets {
            out.push_str("- ");
            out.push_str(&bullet.replace('\n', " "));
            out.push('\n');
        }
        if !slide.notes.is_empty() {
            out.push_str("\n### Notes\n\n");
            for note in &slide.notes {
                out.push_str("- ");
                out.push_str(note);
                out.push('\n');
            }
        }
        if !slide.chart_tokens.is_empty() {
            append_chart_rows_markdown(&mut out, &slide.chart_tokens);
        }
        out.push('\n');
    }
    out.trim().to_string()
}
