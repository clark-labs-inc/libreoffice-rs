//! HTML export for `Presentation`. Each slide becomes a `<section>` and
//! shapes/text boxes are summarized as plain blocks. This is a simple
//! readable rendering, not a fully styled HTML5 slideshow.

use lo_core::{escape_text, html_escape, Presentation, SlideElement};

pub fn to_html(presentation: &Presentation) -> String {
    let mut body = String::new();
    for (idx, slide) in presentation.slides.iter().enumerate() {
        body.push_str(&format!(
            "<section class=\"slide\"><h2>Slide {} — {}</h2>\n",
            idx + 1,
            escape_text(&slide.name)
        ));
        for element in &slide.elements {
            match element {
                SlideElement::TextBox(tb) => {
                    body.push_str("<div class=\"text-box\">");
                    for line in tb.text.lines() {
                        body.push_str(&format!("<p>{}</p>", html_escape(line)));
                    }
                    body.push_str("</div>\n");
                }
                SlideElement::Shape(shape) => {
                    body.push_str(&format!(
                        "<div class=\"shape\">[{:?} {}×{}]</div>\n",
                        shape.kind, shape.frame.size.width, shape.frame.size.height
                    ));
                }
                SlideElement::Image(image) => {
                    body.push_str(&format!(
                        "<figure><figcaption>{}</figcaption></figure>\n",
                        html_escape(&image.alt)
                    ));
                }
            }
        }
        if !slide.notes.is_empty() {
            body.push_str("<aside><strong>Notes:</strong><ul>");
            for note in &slide.notes {
                body.push_str(&format!("<li>{}</li>", html_escape(note)));
            }
            body.push_str("</ul></aside>");
        }
        body.push_str("</section>\n");
    }
    format!(
        "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"utf-8\"/>\n<title>{}</title>\n</head>\n<body>\n{}</body>\n</html>\n",
        escape_text(&presentation.meta.title),
        body
    )
}
