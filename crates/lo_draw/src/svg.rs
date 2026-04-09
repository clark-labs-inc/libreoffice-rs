//! SVG render of a `Drawing`. The first page is rendered to scale; subsequent
//! pages are stacked vertically with a small caption.

use lo_core::{
    geometry::Point, svg_ellipse, svg_footer, svg_header, svg_line, svg_rect, svg_text,
    units::Length, DrawElement, Drawing, Rect, ShapeKind, Size,
};

const PADDING_PT: f32 = 12.0;
const CAPTION_HEIGHT_PT: f32 = 16.0;

pub fn render_svg(drawing: &Drawing) -> String {
    let page_w = drawing.page_size.width.as_pt();
    let page_h = drawing.page_size.height.as_pt();
    let pages = drawing.pages.len() as f32;

    let total_w = PADDING_PT * 2.0 + page_w;
    let total_h = PADDING_PT * 2.0 + (CAPTION_HEIGHT_PT + page_h + PADDING_PT) * pages;

    let mut svg = String::new();
    svg.push_str(&svg_header(Length::pt(total_w), Length::pt(total_h)));

    let mut y = PADDING_PT;
    for page in &drawing.pages {
        svg.push_str(&svg_text(
            Length::pt(PADDING_PT),
            Length::pt(y + 12.0),
            &page.name,
            12,
            "#222222",
            "bold",
        ));
        y += CAPTION_HEIGHT_PT;

        svg.push_str(&svg_rect(
            Rect {
                origin: Point::new(Length::pt(PADDING_PT), Length::pt(y)),
                size: Size::new(Length::pt(page_w), Length::pt(page_h)),
            },
            "#888888",
            Some("#ffffff"),
        ));

        for element in &page.elements {
            match element {
                DrawElement::TextBox(tb) => {
                    let x = PADDING_PT + tb.frame.origin.x.as_pt();
                    let mut ty = y + tb.frame.origin.y.as_pt() + 14.0;
                    for line in tb.text.lines() {
                        svg.push_str(&svg_text(
                            Length::pt(x),
                            Length::pt(ty),
                            line,
                            12,
                            "#000000",
                            "normal",
                        ));
                        ty += 14.0;
                    }
                }
                DrawElement::Shape(shape) => {
                    let x = PADDING_PT + shape.frame.origin.x.as_pt();
                    let sy = y + shape.frame.origin.y.as_pt();
                    let w = shape.frame.size.width.as_pt();
                    let h = shape.frame.size.height.as_pt();
                    let stroke = if shape.style.stroke.is_empty() {
                        "#000000"
                    } else {
                        &shape.style.stroke
                    };
                    let fill = if shape.style.fill.is_empty() {
                        None
                    } else {
                        Some(shape.style.fill.as_str())
                    };
                    match shape.kind {
                        ShapeKind::Rectangle => {
                            svg.push_str(&svg_rect(
                                Rect {
                                    origin: Point::new(Length::pt(x), Length::pt(sy)),
                                    size: Size::new(Length::pt(w), Length::pt(h)),
                                },
                                stroke,
                                fill,
                            ));
                        }
                        ShapeKind::Ellipse => {
                            svg.push_str(&svg_ellipse(
                                Length::pt(x + w / 2.0),
                                Length::pt(sy + h / 2.0),
                                Length::pt(w / 2.0),
                                Length::pt(h / 2.0),
                                stroke,
                                fill,
                            ));
                        }
                        ShapeKind::Line => {
                            svg.push_str(&svg_line(
                                Length::pt(x),
                                Length::pt(sy),
                                Length::pt(x + w),
                                Length::pt(sy + h),
                                stroke,
                            ));
                        }
                    }
                }
                DrawElement::Image(_) => {}
            }
        }

        y += page_h + PADDING_PT;
    }

    svg.push_str(svg_footer());
    svg
}
