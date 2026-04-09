//! PDF export of a `Presentation`.
//!
//! Each slide is rendered onto its own page using the shared pure-Rust
//! PDF canvas. The renderer now respects shape fill/stroke colors,
//! textbox padding/backgrounds, and wraps textbox content to the frame.

use lo_core::{PdfDocument, PdfFont, Presentation, ShapeKind, SlideElement, TextBoxStyle};

use crate::chart::render_chart_rows_pdf;

pub fn to_pdf(presentation: &Presentation) -> Vec<u8> {
    let slide_w = presentation.page_size.width.as_pt().max(320.0);
    let slide_h = presentation.page_size.height.as_pt().max(180.0);
    let mut pdf = PdfDocument::new();

    if presentation.slides.is_empty() {
        let page_index = pdf.add_page(slide_w, slide_h);
        if let Ok(page) = pdf.page_mut(page_index) {
            page.rect_fill_rgb(0.0, 0.0, slide_w, slide_h, 1.0, 1.0, 1.0);
            page.text(24.0, slide_h - 32.0, 24.0, PdfFont::HelveticaBold, &presentation.meta.title);
        }
        return pdf.finish();
    }

    for slide in &presentation.slides {
        let page_index = pdf.add_page(slide_w, slide_h);
        let page = pdf.page_mut(page_index).expect("slide page");
        page.rect_fill_rgb(0.0, 0.0, slide_w, slide_h, 1.0, 1.0, 1.0);

        for element in &slide.elements {
            match element {
                SlideElement::TextBox(text_box) => {
                    render_text_box(page, slide_h, text_box);
                }
                SlideElement::Shape(shape) => {
                    let x = shape.frame.origin.x.as_pt();
                    let y_top = shape.frame.origin.y.as_pt();
                    let w = shape.frame.size.width.as_pt().max(2.0);
                    let h = shape.frame.size.height.as_pt().max(2.0);
                    let y_bottom = slide_h - y_top - h;
                    let fill = parse_color(&shape.style.fill).unwrap_or((1.0, 1.0, 1.0));
                    let stroke = parse_color(&shape.style.stroke).unwrap_or((0.0, 0.0, 0.0));
                    let stroke_width = (shape.style.stroke_width_mm as f32 * 72.0 / 25.4).max(0.75);
                    page.line_width(stroke_width);
                    match shape.kind {
                        ShapeKind::Rectangle => {
                            page.rect_fill_stroke_rgb(x, y_bottom, w, h, fill, stroke);
                        }
                        ShapeKind::Ellipse => {
                            page.ellipse_fill_stroke_rgb(x + w / 2.0, y_bottom + h / 2.0, w / 2.0, h / 2.0, fill, stroke);
                        }
                        ShapeKind::Line => {
                            page.line_rgb(x, slide_h - y_top, x + w, slide_h - y_top - h, stroke.0, stroke.1, stroke.2);
                        }
                    }
                    page.line_width(1.0);
                }
                SlideElement::Image(image) => {
                    let x = image.frame.origin.x.as_pt();
                    let y_top = image.frame.origin.y.as_pt();
                    let w = image.frame.size.width.as_pt().max(32.0);
                    let h = image.frame.size.height.as_pt().max(24.0);
                    let y_bottom = slide_h - y_top - h;
                    page.rect_fill_stroke_rgb(x, y_bottom, w, h, (0.98, 0.98, 0.98), (0.55, 0.55, 0.55));
                    page.line_rgb(x, y_bottom, x + w, y_bottom + h, 0.70, 0.70, 0.70);
                    page.line_rgb(x, y_bottom + h, x + w, y_bottom, 0.70, 0.70, 0.70);
                    page.text_rgb(
                        x + 6.0,
                        y_bottom + 10.0,
                        11.0,
                        PdfFont::HelveticaOblique,
                        &format!("[image: {}]", image.alt),
                        0.30,
                        0.30,
                        0.30,
                    );
                }
            }
        }

        if !slide.chart_tokens.is_empty() {
            render_chart_rows_pdf(page, slide_h, slide_w, &slide.chart_tokens);
        }
        if !slide.notes.is_empty() {
            let band_h = 48.0f32.min(slide_h * 0.18);
            page.rect_fill_rgb(0.0, 0.0, slide_w, band_h, 0.96, 0.96, 0.96);
            let mut y = band_h - 14.0;
            for note in slide.notes.iter().take(3) {
                page.text_rgb(18.0, y, 10.5, PdfFont::HelveticaOblique, note, 0.20, 0.20, 0.20);
                y -= 12.0;
            }
        }

    }

    pdf.finish()
}

fn render_text_box(page: &mut lo_core::PdfPage, slide_h: f32, text_box: &lo_core::TextBox) {
    let x = text_box.frame.origin.x.as_pt();
    let y_top = text_box.frame.origin.y.as_pt();
    let w = text_box.frame.size.width.as_pt().max(16.0);
    let h = text_box.frame.size.height.as_pt().max(12.0);
    let y_bottom = slide_h - y_top - h;
    let style = &text_box.style;
    let padding = style.padding.as_pt().max(4.0);
    if let Some(background) = parse_color(&style.background) {
        page.rect_fill_stroke_rgb(x, y_bottom, w, h, background, (0.75, 0.75, 0.75));
    }
    let (font, size) = pick_text_box_font(text_box.text.lines().next().unwrap_or(""), y_top, h);
    let color = parse_color(&style.foreground).unwrap_or((0.0, 0.0, 0.0));
    let lines = wrap_text_lines(&text_box.text, w - padding * 2.0, size);
    let mut line_y = slide_h - y_top - padding - size;
    for line in lines {
        if line_y < y_bottom + padding {
            break;
        }
        page.text_rgb(x + padding, line_y, size, font, &line, color.0, color.1, color.2);
        line_y -= size * 1.2;
    }
    if style.background.is_empty() {
        page.rect_stroke_rgb(x, y_bottom, w, h, 0.82, 0.82, 0.82);
    }
}

fn pick_text_box_font(first_line: &str, y_top: f32, height: f32) -> (PdfFont, f32) {
    if y_top < 70.0 && height <= 110.0 && first_line.chars().count() <= 80 {
        (PdfFont::HelveticaBold, 26.0)
    } else {
        (PdfFont::Helvetica, 16.0)
    }
}

fn wrap_text_lines(text: &str, width: f32, font_size: f32) -> Vec<String> {
    let max_chars = ((width / (font_size * 0.52)).floor() as usize).max(8);
    let mut out = Vec::new();
    for raw_line in text.split('\n') {
        let mut current = String::new();
        for word in raw_line.split_whitespace() {
            let candidate = if current.is_empty() {
                word.to_string()
            } else {
                format!("{current} {word}")
            };
            if candidate.chars().count() <= max_chars {
                current = candidate;
            } else {
                if !current.is_empty() {
                    out.push(current);
                }
                current = word.to_string();
            }
        }
        if !current.is_empty() {
            out.push(current);
        } else if raw_line.is_empty() {
            out.push(String::new());
        }
    }
    if out.is_empty() {
        out.push(String::new());
    }
    out
}

fn parse_color(input: &str) -> Option<(f32, f32, f32)> {
    let trimmed = input.trim().trim_start_matches('#');
    if trimmed.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&trimmed[0..2], 16).ok()? as f32 / 255.0;
    let g = u8::from_str_radix(&trimmed[2..4], 16).ok()? as f32 / 255.0;
    let b = u8::from_str_radix(&trimmed[4..6], 16).ok()? as f32 / 255.0;
    Some((r, g, b))
}

#[allow(dead_code)]
fn _style(_style: &TextBoxStyle) {}
