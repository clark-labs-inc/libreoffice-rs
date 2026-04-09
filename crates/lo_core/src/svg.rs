//! SVG render helpers shared by every document crate.
//!
//! These produce strings that can be concatenated into a complete SVG document.
//! Colors are CSS color strings (`"#rrggbb"`, `"red"`, …) to match the rest of
//! the lo_core style structs, which already store colors as `String`.

use crate::geometry::Rect;
use crate::units::Length;
use crate::xml::escape_text;

const DEFAULT_FONT_FAMILY: &str = "Arial, Helvetica, sans-serif";

pub fn svg_header(width: Length, height: Length) -> String {
    let w = width.as_pt();
    let h = height.as_pt();
    format!(
        r#"<svg xmlns="http://www.w3.org/2000/svg" width="{w:.2}" height="{h:.2}" viewBox="0 0 {w:.2} {h:.2}">"#
    )
}

pub fn svg_footer() -> &'static str {
    "</svg>"
}

pub fn svg_text(
    x: Length,
    y: Length,
    text: &str,
    size_pt: u32,
    color: &str,
    weight: &str,
) -> String {
    let color = if color.is_empty() { "#000000" } else { color };
    let weight = if weight.is_empty() { "normal" } else { weight };
    format!(
        r#"<text x="{x:.2}" y="{y:.2}" font-family="{font}" font-size="{size_pt}" fill="{color}" font-weight="{weight}">{escaped}</text>"#,
        x = x.as_pt(),
        y = y.as_pt(),
        font = DEFAULT_FONT_FAMILY,
        escaped = escape_text(text),
    )
}

pub fn svg_rect(rect: Rect, stroke: &str, fill: Option<&str>) -> String {
    let stroke = if stroke.is_empty() { "#000000" } else { stroke };
    let fill = fill.unwrap_or("none");
    format!(
        r#"<rect x="{x:.2}" y="{y:.2}" width="{w:.2}" height="{h:.2}" stroke="{stroke}" fill="{fill}"/>"#,
        x = rect.origin.x.as_pt(),
        y = rect.origin.y.as_pt(),
        w = rect.size.width.as_pt(),
        h = rect.size.height.as_pt(),
    )
}

pub fn svg_line(x1: Length, y1: Length, x2: Length, y2: Length, stroke: &str) -> String {
    let stroke = if stroke.is_empty() { "#000000" } else { stroke };
    format!(
        r#"<line x1="{:.2}" y1="{:.2}" x2="{:.2}" y2="{:.2}" stroke="{stroke}"/>"#,
        x1.as_pt(),
        y1.as_pt(),
        x2.as_pt(),
        y2.as_pt(),
    )
}

pub fn svg_ellipse(
    cx: Length,
    cy: Length,
    rx: Length,
    ry: Length,
    stroke: &str,
    fill: Option<&str>,
) -> String {
    let stroke = if stroke.is_empty() { "#000000" } else { stroke };
    let fill = fill.unwrap_or("none");
    format!(
        r#"<ellipse cx="{:.2}" cy="{:.2}" rx="{:.2}" ry="{:.2}" stroke="{stroke}" fill="{fill}"/>"#,
        cx.as_pt(),
        cy.as_pt(),
        rx.as_pt(),
        ry.as_pt(),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn header_uses_pt_dimensions() {
        let h = svg_header(Length::pt(100.0), Length::pt(50.0));
        assert!(h.contains(r#"width="100.00""#));
        assert!(h.contains(r#"viewBox="0 0 100.00 50.00""#));
    }

    #[test]
    fn text_escapes_xml_special_chars() {
        let t = svg_text(
            Length::pt(0.0),
            Length::pt(0.0),
            "a<b&c",
            12,
            "#ff0000",
            "bold",
        );
        assert!(t.contains("a&lt;b&amp;c"));
        assert!(t.contains("fill=\"#ff0000\""));
        assert!(t.contains("font-weight=\"bold\""));
    }
}
