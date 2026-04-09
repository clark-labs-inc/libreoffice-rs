//! Binary importers for `Drawing`.
//!
//! Supports `odg` and `svg`.

use lo_core::{
    parse_xml_document, DrawElement, DrawPage, Drawing, Length, LoError, Rect, Result, Shape,
    ShapeKind, ShapeStyle, TextBox, TextBoxStyle, XmlNode,
};
use lo_zip::ZipArchive;

pub fn load_bytes(title: impl Into<String>, bytes: &[u8], format: &str) -> Result<Drawing> {
    let title = title.into();
    match format.to_ascii_lowercase().as_str() {
        "odg" => from_odg_bytes(title, bytes),
        "svg" => from_svg(title, &String::from_utf8_lossy(bytes)),
        other => Err(LoError::Unsupported(format!("draw import format {other}"))),
    }
}

pub fn from_odg_bytes(title: impl Into<String>, bytes: &[u8]) -> Result<Drawing> {
    let title = title.into();
    let zip = ZipArchive::new(bytes)?;
    let content = parse_xml_document(&zip.read_string("content.xml")?)?;
    let body = content
        .child("body")
        .ok_or_else(|| LoError::Parse("content.xml missing office:body".to_string()))?;
    let drawing_node = body
        .child("drawing")
        .ok_or_else(|| LoError::Parse("content.xml missing office:drawing".to_string()))?;
    let mut drawing = Drawing::new(title);
    drawing.pages.clear();
    for page_node in drawing_node.children_named("page") {
        drawing.pages.push(parse_odg_page(page_node));
    }
    if drawing.pages.is_empty() {
        drawing.pages.push(DrawPage::default());
    }
    Ok(drawing)
}

pub fn from_svg(title: impl Into<String>, svg: &str) -> Result<Drawing> {
    let title = title.into();
    let root = parse_xml_document(svg)?;
    if root.local_name() != "svg" {
        return Err(LoError::Parse("expected svg root element".to_string()));
    }
    let mut drawing = Drawing::new(title);
    drawing.pages.clear();
    let mut page = DrawPage {
        name: "Page1".to_string(),
        elements: Vec::new(),
    };
    for child in &root.children {
        match child.local_name() {
            "rect" => page.elements.push(DrawElement::Shape(Shape {
                kind: ShapeKind::Rectangle,
                frame: Rect::new(
                    Length::mm(parse_f32(child.attr("x")).unwrap_or(0.0)),
                    Length::mm(parse_f32(child.attr("y")).unwrap_or(0.0)),
                    Length::mm(parse_f32(child.attr("width")).unwrap_or(0.0)),
                    Length::mm(parse_f32(child.attr("height")).unwrap_or(0.0)),
                ),
                style: ShapeStyle::default(),
            })),
            "ellipse" => {
                let cx = parse_f32(child.attr("cx")).unwrap_or(0.0);
                let cy = parse_f32(child.attr("cy")).unwrap_or(0.0);
                let rx = parse_f32(child.attr("rx")).unwrap_or(0.0);
                let ry = parse_f32(child.attr("ry")).unwrap_or(0.0);
                page.elements.push(DrawElement::Shape(Shape {
                    kind: ShapeKind::Ellipse,
                    frame: Rect::new(
                        Length::mm(cx - rx),
                        Length::mm(cy - ry),
                        Length::mm(rx * 2.0),
                        Length::mm(ry * 2.0),
                    ),
                    style: ShapeStyle::default(),
                }));
            }
            "circle" => {
                let cx = parse_f32(child.attr("cx")).unwrap_or(0.0);
                let cy = parse_f32(child.attr("cy")).unwrap_or(0.0);
                let r = parse_f32(child.attr("r")).unwrap_or(0.0);
                page.elements.push(DrawElement::Shape(Shape {
                    kind: ShapeKind::Ellipse,
                    frame: Rect::new(
                        Length::mm(cx - r),
                        Length::mm(cy - r),
                        Length::mm(r * 2.0),
                        Length::mm(r * 2.0),
                    ),
                    style: ShapeStyle::default(),
                }));
            }
            "line" => {
                let x1 = parse_f32(child.attr("x1")).unwrap_or(0.0);
                let y1 = parse_f32(child.attr("y1")).unwrap_or(0.0);
                let x2 = parse_f32(child.attr("x2")).unwrap_or(0.0);
                let y2 = parse_f32(child.attr("y2")).unwrap_or(0.0);
                page.elements.push(DrawElement::Shape(Shape {
                    kind: ShapeKind::Line,
                    frame: Rect::new(
                        Length::mm(x1),
                        Length::mm(y1),
                        Length::mm(x2 - x1),
                        Length::mm(y2 - y1),
                    ),
                    style: ShapeStyle::default(),
                }));
            }
            "text" => page.elements.push(DrawElement::TextBox(TextBox {
                frame: Rect::new(
                    Length::mm(parse_f32(child.attr("x")).unwrap_or(0.0)),
                    Length::mm(parse_f32(child.attr("y")).unwrap_or(0.0)),
                    Length::mm(60.0),
                    Length::mm(10.0),
                ),
                text: child.text_content(),
                style: TextBoxStyle::default(),
            })),
            _ => {}
        }
    }
    drawing.pages.push(page);
    Ok(drawing)
}

fn parse_odg_page(node: &XmlNode) -> DrawPage {
    let mut page = DrawPage {
        name: node.attr("name").unwrap_or("Page1").to_string(),
        elements: Vec::new(),
    };
    for child in &node.children {
        match child.local_name() {
            "frame" => {
                let x = parse_odf_len_mm(child.attr("x")).unwrap_or(0.0);
                let y = parse_odf_len_mm(child.attr("y")).unwrap_or(0.0);
                let width = parse_odf_len_mm(child.attr("width")).unwrap_or(60.0);
                let height = parse_odf_len_mm(child.attr("height")).unwrap_or(20.0);
                let text = child
                    .child("text-box")
                    .map(|node| node.text_content())
                    .unwrap_or_default();
                if text.trim().is_empty() {
                    page.elements.push(DrawElement::Shape(Shape {
                        kind: ShapeKind::Rectangle,
                        frame: Rect::new(
                            Length::mm(x),
                            Length::mm(y),
                            Length::mm(width),
                            Length::mm(height),
                        ),
                        style: ShapeStyle::default(),
                    }));
                } else {
                    page.elements.push(DrawElement::TextBox(TextBox {
                        frame: Rect::new(
                            Length::mm(x),
                            Length::mm(y),
                            Length::mm(width),
                            Length::mm(height),
                        ),
                        text,
                        style: TextBoxStyle::default(),
                    }));
                }
            }
            "ellipse" => page.elements.push(DrawElement::Shape(Shape {
                kind: ShapeKind::Ellipse,
                frame: Rect::new(
                    Length::mm(parse_odf_len_mm(child.attr("x")).unwrap_or(0.0)),
                    Length::mm(parse_odf_len_mm(child.attr("y")).unwrap_or(0.0)),
                    Length::mm(parse_odf_len_mm(child.attr("width")).unwrap_or(40.0)),
                    Length::mm(parse_odf_len_mm(child.attr("height")).unwrap_or(20.0)),
                ),
                style: ShapeStyle::default(),
            })),
            "line" => {
                let x1 = parse_odf_len_mm(child.attr("x1")).unwrap_or(0.0);
                let y1 = parse_odf_len_mm(child.attr("y1")).unwrap_or(0.0);
                let x2 = parse_odf_len_mm(child.attr("x2")).unwrap_or(0.0);
                let y2 = parse_odf_len_mm(child.attr("y2")).unwrap_or(0.0);
                page.elements.push(DrawElement::Shape(Shape {
                    kind: ShapeKind::Line,
                    frame: Rect::new(
                        Length::mm(x1),
                        Length::mm(y1),
                        Length::mm(x2 - x1),
                        Length::mm(y2 - y1),
                    ),
                    style: ShapeStyle::default(),
                }));
            }
            "custom-shape" | "rect" => page.elements.push(DrawElement::Shape(Shape {
                kind: ShapeKind::Rectangle,
                frame: Rect::new(
                    Length::mm(parse_odf_len_mm(child.attr("x")).unwrap_or(0.0)),
                    Length::mm(parse_odf_len_mm(child.attr("y")).unwrap_or(0.0)),
                    Length::mm(parse_odf_len_mm(child.attr("width")).unwrap_or(40.0)),
                    Length::mm(parse_odf_len_mm(child.attr("height")).unwrap_or(20.0)),
                ),
                style: ShapeStyle::default(),
            })),
            _ => {}
        }
    }
    page
}

fn parse_f32(value: Option<&str>) -> Option<f32> {
    let trimmed = value?.trim();
    let trimmed = trimmed
        .strip_suffix("px")
        .or_else(|| trimmed.strip_suffix("pt"))
        .unwrap_or(trimmed);
    trimmed.parse::<f32>().ok()
}

fn parse_odf_len_mm(value: Option<&str>) -> Option<f32> {
    let value = value?;
    let trimmed = value.trim();
    if let Some(number) = trimmed.strip_suffix("cm") {
        return number.parse::<f32>().ok().map(|cm| cm * 10.0);
    }
    if let Some(number) = trimmed.strip_suffix("mm") {
        return number.parse::<f32>().ok();
    }
    if let Some(number) = trimmed.strip_suffix("in") {
        return number.parse::<f32>().ok().map(|inch| inch * 25.4);
    }
    if let Some(number) = trimmed.strip_suffix("pt") {
        return number.parse::<f32>().ok().map(|pt| pt * 25.4 / 72.0);
    }
    trimmed.parse::<f32>().ok()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::demo_drawing;

    #[test]
    fn odg_round_trip_imports_shapes() {
        let drawing = demo_drawing("Demo");
        let tmp = std::env::temp_dir().join("lo_draw_import_test.odg");
        lo_odf::save_drawing_document(&tmp, &drawing).expect("save odg");
        let bytes = std::fs::read(&tmp).expect("read");
        let _ = std::fs::remove_file(&tmp);
        let loaded = from_odg_bytes("Demo", &bytes).expect("import odg");
        assert!(!loaded.pages.is_empty());
        assert!(!loaded.pages[0].elements.is_empty());
    }

    #[test]
    fn svg_import_reads_basic_shapes() {
        let svg = r#"<svg xmlns="http://www.w3.org/2000/svg"><rect x="10" y="20" width="30" height="40"/><text x="5" y="6">hi</text></svg>"#;
        let drawing = from_svg("Demo", svg).expect("import svg");
        assert_eq!(drawing.pages[0].elements.len(), 2);
    }
}
