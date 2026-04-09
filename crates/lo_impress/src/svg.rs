//! SVG render of a `Presentation`. Each slide is rendered into a horizontal
//! filmstrip of bordered cards. Shapes and text boxes are projected from the
//! slide's millimeter coordinate system into the SVG point coordinate system.

use lo_core::{
    geometry::Point, svg_ellipse, svg_footer, svg_header, svg_line, svg_rect, svg_text,
    units::Length, Presentation, Rect, ShapeKind, Size, SlideElement,
};

const PADDING_PT: f32 = 16.0;
const SLIDE_GAP_PT: f32 = 24.0;
const TITLE_HEIGHT_PT: f32 = 18.0;

fn mm_to_pt(mm: Length) -> f32 {
    mm.as_pt()
}

pub fn render_svg(presentation: &Presentation) -> String {
    let slide_w_pt = mm_to_pt(presentation.page_size.width);
    let slide_h_pt = mm_to_pt(presentation.page_size.height);

    let total_w = PADDING_PT * 2.0 + slide_w_pt;
    let total_h = PADDING_PT * 2.0
        + (slide_h_pt + TITLE_HEIGHT_PT + SLIDE_GAP_PT) * presentation.slides.len() as f32;

    let mut svg = String::new();
    svg.push_str(&svg_header(Length::pt(total_w), Length::pt(total_h)));

    let mut y_offset = PADDING_PT;
    for (idx, slide) in presentation.slides.iter().enumerate() {
        // Slide title
        svg.push_str(&svg_text(
            Length::pt(PADDING_PT),
            Length::pt(y_offset + 12.0),
            &format!("Slide {} — {}", idx + 1, slide.name),
            12,
            "#222222",
            "bold",
        ));
        y_offset += TITLE_HEIGHT_PT;

        // Slide canvas border
        svg.push_str(&svg_rect(
            Rect {
                origin: Point::new(Length::pt(PADDING_PT), Length::pt(y_offset)),
                size: Size::new(Length::pt(slide_w_pt), Length::pt(slide_h_pt)),
            },
            "#888888",
            Some("#fafafa"),
        ));

        // Slide elements
        for element in &slide.elements {
            match element {
                SlideElement::TextBox(tb) => {
                    let x = PADDING_PT + mm_to_pt(tb.frame.origin.x);
                    let mut y = y_offset + mm_to_pt(tb.frame.origin.y) + 14.0;
                    for line in tb.text.lines() {
                        svg.push_str(&svg_text(
                            Length::pt(x),
                            Length::pt(y),
                            line,
                            12,
                            "#000000",
                            "normal",
                        ));
                        y += 14.0;
                    }
                }
                SlideElement::Shape(shape) => {
                    let x = PADDING_PT + mm_to_pt(shape.frame.origin.x);
                    let y = y_offset + mm_to_pt(shape.frame.origin.y);
                    let w = mm_to_pt(shape.frame.size.width);
                    let h = mm_to_pt(shape.frame.size.height);
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
                                    origin: Point::new(Length::pt(x), Length::pt(y)),
                                    size: Size::new(Length::pt(w), Length::pt(h)),
                                },
                                stroke,
                                fill,
                            ));
                        }
                        ShapeKind::Ellipse => {
                            svg.push_str(&svg_ellipse(
                                Length::pt(x + w / 2.0),
                                Length::pt(y + h / 2.0),
                                Length::pt(w / 2.0),
                                Length::pt(h / 2.0),
                                stroke,
                                fill,
                            ));
                        }
                        ShapeKind::Line => {
                            svg.push_str(&svg_line(
                                Length::pt(x),
                                Length::pt(y),
                                Length::pt(x + w),
                                Length::pt(y + h),
                                stroke,
                            ));
                        }
                    }
                }
                SlideElement::Image(_) => {}
            }
        }

        y_offset += slide_h_pt + SLIDE_GAP_PT;
    }

    svg.push_str(svg_footer());
    svg
}
