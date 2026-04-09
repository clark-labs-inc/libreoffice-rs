//! SVG render of a `TextDocument`. This is a deliberately simple "preview"
//! renderer: it lays out the plain-text version of the document as wrapped
//! lines on a single page, drawn inside a bordered rectangle.

use lo_core::{
    geometry::{Point, Rect},
    svg_footer, svg_header, svg_rect, svg_text,
    units::Length,
    Size, TextDocument,
};

const LINE_HEIGHT_PT: f32 = 18.0;
const FONT_SIZE_PT: u32 = 14;
const MARGIN_PT: f32 = 24.0;

/// Render the document as a single-page SVG. The size is interpreted in
/// PDF/SVG points (1pt = 1/72 in).
pub fn render_svg(document: &TextDocument, size: Size) -> String {
    let mut svg = String::new();
    svg.push_str(&svg_header(size.width, size.height));

    svg.push_str(&svg_rect(
        Rect {
            origin: Point::new(Length::pt(0.5), Length::pt(0.5)),
            size: Size::new(
                Length::pt(size.width.as_pt() - 1.0),
                Length::pt(size.height.as_pt() - 1.0),
            ),
        },
        "#888888",
        Some("#ffffff"),
    ));

    let max_y = size.height.as_pt() - MARGIN_PT;
    let mut y = MARGIN_PT + FONT_SIZE_PT as f32;
    for line in document.plain_text().lines() {
        if line.is_empty() {
            y += LINE_HEIGHT_PT * 0.5;
            continue;
        }
        if y > max_y {
            break;
        }
        svg.push_str(&svg_text(
            Length::pt(MARGIN_PT),
            Length::pt(y),
            line,
            FONT_SIZE_PT,
            "#000000",
            "normal",
        ));
        y += LINE_HEIGHT_PT;
    }

    svg.push_str(svg_footer());
    svg
}
