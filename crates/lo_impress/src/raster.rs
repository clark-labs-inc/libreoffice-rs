use lo_core::{parse_hex_color, Presentation, RasterImage, Rgba, ShapeKind, SlideElement};

use crate::chart::render_chart_rows_raster;

pub fn render_png_pages(presentation: &Presentation, dpi: u32) -> Vec<Vec<u8>> {
    render_pages(presentation, dpi)
        .into_iter()
        .map(|page| page.encode_png())
        .collect()
}

pub fn render_jpeg_pages(presentation: &Presentation, dpi: u32, quality: u8) -> Vec<Vec<u8>> {
    render_pages(presentation, dpi)
        .into_iter()
        .map(|page| page.encode_jpeg(quality))
        .collect()
}

pub fn render_pages(presentation: &Presentation, dpi: u32) -> Vec<RasterImage> {
    let width = mm_to_px(presentation.page_size.width.as_mm(), dpi).max(320) as u32;
    let height = mm_to_px(presentation.page_size.height.as_mm(), dpi).max(180) as u32;
    let mut pages = Vec::new();
    for slide in &presentation.slides {
        let mut page = RasterImage::new(width, height, Rgba::WHITE);
        page.fill_rect(0, 0, width as i32, 10, Rgba::rgba(220, 230, 242, 255));
        page.draw_text(10, 12, 16, Rgba::rgba(60, 60, 60, 255), &slide.name, true);
        for element in &slide.elements {
            match element {
                SlideElement::TextBox(text_box) => {
                    let x = mm_to_px(text_box.frame.origin.x.as_mm(), dpi);
                    let y = mm_to_px(text_box.frame.origin.y.as_mm(), dpi);
                    let w = mm_to_px(text_box.frame.size.width.as_mm(), dpi).max(20);
                    let h = mm_to_px(text_box.frame.size.height.as_mm(), dpi).max(18);
                    let bg = if text_box.style.background.trim().is_empty() {
                        Rgba::rgba(255, 255, 255, 0)
                    } else {
                        parse_hex_color(&text_box.style.background, Rgba::rgba(245, 245, 245, 255))
                    };
                    if bg.a > 0 {
                        page.fill_rect(x, y, w, h, bg);
                    }
                    page.stroke_rect(x, y, w, h, 1, Rgba::rgba(190, 190, 190, 255));
                    let fg = if text_box.style.foreground.trim().is_empty() {
                        Rgba::BLACK
                    } else {
                        parse_hex_color(&text_box.style.foreground, Rgba::BLACK)
                    };
                    let pad = mm_to_px(text_box.style.padding.as_mm(), dpi).max(4);
                    let font_px = if h > 60 { 16 } else if h > 28 { 14 } else { 12 };
                    let lines = wrap_text(&page, &text_box.text, font_px, w - pad * 2);
                    let mut ty = y + pad;
                    for line in lines.into_iter().take(((h - pad * 2) / (font_px + 3)).max(1) as usize) {
                        page.draw_text(x + pad, ty, font_px, fg, &line, false);
                        ty += font_px + 3;
                    }
                }
                SlideElement::Shape(shape) => {
                    let x = mm_to_px(shape.frame.origin.x.as_mm(), dpi);
                    let y = mm_to_px(shape.frame.origin.y.as_mm(), dpi);
                    let w = mm_to_px(shape.frame.size.width.as_mm(), dpi).max(2);
                    let h = mm_to_px(shape.frame.size.height.as_mm(), dpi).max(2);
                    let fill = parse_hex_color(&shape.style.fill, Rgba::rgba(240, 240, 240, 255));
                    let stroke = if shape.style.stroke.trim().is_empty() {
                        Rgba::rgba(100, 100, 100, 255)
                    } else {
                        parse_hex_color(&shape.style.stroke, Rgba::rgba(100, 100, 100, 255))
                    };
                    let line_w = mm_to_px(shape.style.stroke_width_mm.max(1) as f32, dpi).max(1);
                    match shape.kind {
                        ShapeKind::Rectangle => {
                            if !shape.style.fill.trim().is_empty() {
                                page.fill_rect(x, y, w, h, fill);
                            }
                            page.stroke_rect(x, y, w, h, line_w, stroke);
                        }
                        ShapeKind::Ellipse => {
                            if !shape.style.fill.trim().is_empty() {
                                page.fill_ellipse(x + w / 2, y + h / 2, w / 2, h / 2, fill);
                            }
                            page.stroke_ellipse(x + w / 2, y + h / 2, w / 2, h / 2, line_w, stroke);
                        }
                        ShapeKind::Line => {
                            page.draw_line(x, y, x + w, y + h, line_w, stroke);
                        }
                    }
                }
                SlideElement::Image(image) => {
                    let x = mm_to_px(image.frame.origin.x.as_mm(), dpi);
                    let y = mm_to_px(image.frame.origin.y.as_mm(), dpi);
                    let w = mm_to_px(image.frame.size.width.as_mm(), dpi).max(40);
                    let h = mm_to_px(image.frame.size.height.as_mm(), dpi).max(30);
                    page.fill_rect(x, y, w, h, Rgba::rgba(246, 246, 246, 255));
                    page.stroke_rect(x, y, w, h, 2, Rgba::rgba(150, 150, 150, 255));
                    page.draw_line(x, y, x + w, y + h, 1, Rgba::rgba(180, 180, 180, 255));
                    page.draw_line(x + w, y, x, y + h, 1, Rgba::rgba(180, 180, 180, 255));
                    page.draw_text(x + 8, y + h / 2, 12, Rgba::rgba(80, 80, 80, 255), &format!("image: {}", image.alt), false);
                }
            }
        }
        if !slide.chart_tokens.is_empty() {
            render_chart_rows_raster(&mut page, dpi, &slide.chart_tokens);
        }
        if !slide.notes.is_empty() {
            let band_h = 44;
            let y = height as i32 - band_h;
            page.fill_rect(0, y, width as i32, band_h, Rgba::rgba(252, 249, 235, 255));
            page.stroke_rect(0, y, width as i32, band_h, 1, Rgba::rgba(196, 186, 120, 255));
            page.draw_text(10, y + 8, 12, Rgba::rgba(80, 80, 80, 255), "Notes:", true);
            let joined = slide.notes.join(" • ");
            let lines = wrap_text(&page, &joined, 11, width as i32 - 80);
            let mut ty = y + 22;
            for line in lines.into_iter().take(2) {
                page.draw_text(56, ty, 11, Rgba::rgba(80, 80, 80, 255), &line, false);
                ty += 13;
            }
        }
        pages.push(page);
    }
    pages
}

fn wrap_text(canvas: &RasterImage, text: &str, font_px: i32, max_width: i32) -> Vec<String> {
    let mut lines = Vec::new();
    for paragraph in text.lines() {
        let mut current = String::new();
        for word in paragraph.split_whitespace() {
            let candidate = if current.is_empty() { word.to_string() } else { format!("{} {}", current, word) };
            if !current.is_empty() && canvas.measure_text(&candidate, font_px) > max_width {
                lines.push(current);
                current = word.to_string();
            } else {
                current = candidate;
            }
        }
        if !current.is_empty() {
            lines.push(current);
        }
    }
    if lines.is_empty() {
        lines.push(String::new());
    }
    lines
}

fn mm_to_px(mm: f32, dpi: u32) -> i32 {
    ((mm / 25.4) * dpi as f32).round() as i32
}
