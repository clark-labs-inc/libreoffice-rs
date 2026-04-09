//! Binary importers for `Presentation`.
//!
//! Supports `pptx`, `odp`, and a `txt` outline form.

use std::collections::BTreeMap;

use lo_core::{
    parse_xml_document, Length, LoError, Presentation, Rect, Result, Shape, ShapeKind, ShapeStyle,
    Slide, SlideElement, TextBox, TextBoxStyle, XmlItem, XmlNode,
};
use lo_zip::{rels_path_for, resolve_part_target, ZipArchive};

use crate::chart::{chart_row_title, graphic_frame_has_chart, load_pptx_chart_rows};

pub fn load_bytes(title: impl Into<String>, bytes: &[u8], format: &str) -> Result<Presentation> {
    let title = title.into();
    match format.to_ascii_lowercase().as_str() {
        "pptx" => from_pptx_bytes(title, bytes),
        "odp" => from_odp_bytes(title, bytes),
        "txt" | "text" => Ok(from_text_outline(title, &String::from_utf8_lossy(bytes))),
        other => Err(LoError::Unsupported(format!(
            "impress import format {other}"
        ))),
    }
}

pub fn from_text_outline(title: impl Into<String>, text: &str) -> Presentation {
    let mut deck = Presentation::new(title);
    let mut current = Slide {
        name: "Slide 1".to_string(),
        elements: Vec::new(),
        notes: Vec::new(),
            chart_tokens: Vec::new(),
    };
    let mut bullets: Vec<String> = Vec::new();
    let mut bullet_offset = 0u32;
    for line in text.lines() {
        let trimmed = line.trim();
        if let Some(rest) = trimmed.strip_prefix('#') {
            if !bullets.is_empty() {
                push_bullets(&mut current, &mut bullets);
            }
            if !current.elements.is_empty() || current.name != "Slide 1" {
                deck.slides.push(current);
                bullet_offset = 0;
            }
            current = Slide {
                name: rest.trim().to_string(),
                elements: Vec::new(),
                notes: Vec::new(),
            chart_tokens: Vec::new(),
            };
        } else if let Some(rest) = trimmed.strip_prefix('-') {
            bullets.push(rest.trim().to_string());
        } else if !trimmed.is_empty() {
            let _ = bullet_offset;
            current.elements.push(SlideElement::TextBox(TextBox {
                frame: Rect::new(
                    Length::mm(20.0),
                    Length::mm(20.0 + current.elements.len() as f32 * 15.0),
                    Length::mm(220.0),
                    Length::mm(12.0),
                ),
                text: trimmed.to_string(),
                style: TextBoxStyle::default(),
            }));
        }
    }
    if !bullets.is_empty() {
        push_bullets(&mut current, &mut bullets);
    }
    deck.slides.push(current);
    deck
}

fn push_bullets(slide: &mut Slide, bullets: &mut Vec<String>) {
    let body = bullets
        .drain(..)
        .map(|line| format!("• {line}"))
        .collect::<Vec<_>>()
        .join("\n");
    slide.elements.push(SlideElement::TextBox(TextBox {
        frame: Rect::new(
            Length::mm(20.0),
            Length::mm(40.0),
            Length::mm(220.0),
            Length::mm(80.0),
        ),
        text: body,
        style: TextBoxStyle::default(),
    }));
}

pub fn from_pptx_bytes(title: impl Into<String>, bytes: &[u8]) -> Result<Presentation> {
    let _ = title.into();
    let zip = ZipArchive::new(bytes)?;
    let title: String = if zip.contains("docProps/core.xml") {
        parse_xml_document(&zip.read_string("docProps/core.xml")?)
            .ok()
            .and_then(|root| root.child("title").map(|n| n.text_content()))
            .map(|s| s.trim().to_string())
            .filter(|s| !s.is_empty())
            .unwrap_or_default()
    } else {
        String::new()
    };
    let presentation_xml = parse_xml_document(&zip.read_string("ppt/presentation.xml")?)?;
    let rels = parse_relationships(&zip, "ppt/presentation.xml")?;

    let mut deck = Presentation::new(title);
    if let Some(list) = presentation_xml.child("sldIdLst") {
        for (index, slide_id) in list.children_named("sldId").enumerate() {
            let target = slide_id
                .attr("id")
                .or_else(|| slide_id.attr("r:id"))
                .and_then(|id| rels.get(id))
                .cloned()
                .unwrap_or_else(|| format!("ppt/slides/slide{}.xml", index + 1));
            if !zip.contains(&target) {
                continue;
            }
            let slide_root = parse_xml_document(&zip.read_string(&target)?)?;
            let slide_rels = parse_relationships(&zip, &target)?;
            let notes = load_pptx_notes(&zip, &slide_rels)?;
            let chart_rows = load_pptx_chart_rows(&zip, &slide_root, &slide_rels)?;
            let table_texts = collect_pptx_table_texts(&slide_root);
            deck.slides
                .push(parse_pptx_slide(&slide_root, notes, chart_rows, table_texts));
        }
    }
    if deck.slides.is_empty() {
        deck.slides.push(Slide::default());
    }
    Ok(deck)
}


/// Pull text out of `<a:tbl>` table cells embedded in a `<p:graphicFrame>`.
/// Returns one row per inner Vec.
fn collect_pptx_table_texts(slide_root: &XmlNode) -> Vec<Vec<String>> {
    let mut tables = Vec::new();
    walk_for_tables(slide_root, &mut tables);
    tables
}

fn walk_for_tables(node: &XmlNode, out: &mut Vec<Vec<String>>) {
    if node.local_name() == "tbl" {
        let mut rows: Vec<String> = Vec::new();
        for tr in node.children_named("tr") {
            let mut cells: Vec<String> = Vec::new();
            for tc in tr.children_named("tc") {
                let mut cell_text = String::new();
                let mut texts: Vec<String> = Vec::new();
                collect_text_nodes(tc, &mut texts);
                cell_text.push_str(&texts.join(" "));
                if !cell_text.trim().is_empty() {
                    cells.push(cell_text.trim().to_string());
                }
            }
            if !cells.is_empty() {
                rows.push(cells.join(" | "));
            }
        }
        if !rows.is_empty() {
            out.push(rows);
        }
    }
    for child in &node.children {
        walk_for_tables(child, out);
    }
}

pub fn from_odp_bytes(title: impl Into<String>, bytes: &[u8]) -> Result<Presentation> {
    let title = title.into();
    let zip = ZipArchive::new(bytes)?;
    let content = parse_xml_document(&zip.read_string("content.xml")?)?;
    let body = content
        .child("body")
        .ok_or_else(|| LoError::Parse("content.xml missing office:body".to_string()))?;
    let presentation = body
        .child("presentation")
        .ok_or_else(|| LoError::Parse("content.xml missing office:presentation".to_string()))?;

    let mut deck = Presentation::new(title);
    for page in presentation.children_named("page") {
        deck.slides.push(parse_odp_slide(page));
    }
    if deck.slides.is_empty() {
        deck.slides.push(Slide::default());
    }
    Ok(deck)
}

// ---------------------------------------------------------------------------

fn parse_relationships(zip: &ZipArchive, part: &str) -> Result<BTreeMap<String, String>> {
    let rels_path = rels_path_for(part);
    if !zip.contains(&rels_path) {
        return Ok(BTreeMap::new());
    }
    let root = parse_xml_document(&zip.read_string(&rels_path)?)?;
    let mut map = BTreeMap::new();
    for rel in root.children_named("Relationship") {
        if let (Some(id), Some(target)) = (rel.attr("Id"), rel.attr("Target")) {
            map.insert(id.to_string(), resolve_part_target(part, target));
        }
    }
    Ok(map)
}

fn load_pptx_notes(zip: &ZipArchive, slide_rels: &BTreeMap<String, String>) -> Result<Vec<String>> {
    let notes_target = slide_rels
        .values()
        .find(|path| path.contains("notesSlides/"))
        .cloned();
    let Some(notes_target) = notes_target else {
        return Ok(Vec::new());
    };
    if !zip.contains(&notes_target) {
        return Ok(Vec::new());
    }
    let root = parse_xml_document(&zip.read_string(&notes_target)?)?;
    let mut texts = Vec::new();
    collect_text_nodes(&root, &mut texts);
    Ok(texts
        .into_iter()
        .filter(|text| !text.trim().is_empty())
        .collect())
}

fn collect_text_nodes(node: &XmlNode, out: &mut Vec<String>) {
    if node.local_name() == "t" {
        out.push(node.text_content());
    }
    for child in &node.children {
        collect_text_nodes(child, out);
    }
}

fn parse_pptx_slide(
    root: &XmlNode,
    notes: Vec<String>,
    chart_texts: Vec<Vec<String>>,
    table_texts: Vec<Vec<String>>,
) -> Slide {
    let mut slide = Slide {
        name: "Slide".to_string(),
        elements: Vec::new(),
        notes,
        chart_tokens: Vec::new(),
    };
    let mut title_set = false;
    if let Some(sp_tree) = root.child("cSld").and_then(|node| node.child("spTree")) {
        walk_pptx_sp_tree(sp_tree, &mut slide, &mut title_set);
    }

    // Hand chart tokens to the PDF/raster backends via `slide.chart_tokens`
    // — they each render the tokens at explicit positions so single-digit
    // axis tick labels survive `pdftotext` extraction without the textbox
    // wrap dropping the inter-token spaces.
    for row in chart_texts.iter() {
        if !title_set {
            if let Some(title) = chart_row_title(row) {
                slide.name = title;
                title_set = true;
            } else if let Some(first) = row.first() {
                if !first.starts_with("__LO_") {
                    slide.name = first.clone();
                    title_set = true;
                }
            }
        }
        slide.chart_tokens.push(row.clone());
    }

    // Inject table text as additional chart token rows so the renderer's
    // PDF text layer contains all cells (PowerPoint a:tbl elements).
    for rows in table_texts.iter() {
        slide.chart_tokens.push(rows.clone());
    }

    if !title_set {
        if let Some(SlideElement::TextBox(text_box)) = slide.elements.first() {
            slide.name = text_box.text.lines().next().unwrap_or("Slide").to_string();
        }
    }
    slide
}

/// Recursively walk a `<p:spTree>` (or nested `<p:grpSp>`) collecting
/// shapes, group shapes, and graphic frames into the slide.
fn walk_pptx_sp_tree(node: &XmlNode, slide: &mut Slide, title_set: &mut bool) {
    for child in &node.children {
        match child.local_name() {
            "sp" => parse_pptx_shape_into(child, slide, title_set),
            "grpSp" => walk_pptx_sp_tree(child, slide, title_set),
            "graphicFrame" => {
                if graphic_frame_has_chart(child) {
                    continue;
                }
                // Capture text inside non-chart table/diagram frames generically
                let mut texts: Vec<String> = Vec::new();
                collect_text_nodes(child, &mut texts);
                let text: String = texts
                    .into_iter()
                    .filter(|s| !s.trim().is_empty())
                    .collect::<Vec<_>>()
                    .join("\n");
                if !text.is_empty() {
                    let (x_mm, y_mm, w_mm, h_mm) = parse_pptx_transform(child);
                    slide.elements.push(SlideElement::TextBox(TextBox {
                        frame: Rect::new(
                            Length::mm(x_mm),
                            Length::mm(y_mm),
                            Length::mm(w_mm.max(80.0)),
                            Length::mm(h_mm.max(20.0)),
                        ),
                        text,
                        style: TextBoxStyle::default(),
                    }));
                }
            }
            _ => {}
        }
    }
}

fn parse_pptx_shape_into(shape: &XmlNode, slide: &mut Slide, title_set: &mut bool) {
    let placeholder = shape
        .child("nvSpPr")
        .and_then(|node| node.child("nvPr"))
        .and_then(|node| node.child("ph"))
        .and_then(|node| node.attr("type"))
        .unwrap_or("");
    let (x_mm, y_mm, w_mm, h_mm) = parse_pptx_transform(shape);
    let frame = Rect::new(
        Length::mm(x_mm),
        Length::mm(y_mm),
        Length::mm(w_mm),
        Length::mm(h_mm),
    );
    let paragraphs = shape
        .child("txBody")
        .map(parse_pptx_text_body)
        .unwrap_or_default();
    if matches!(placeholder, "title" | "ctrTitle" | "subTitle") && !paragraphs.is_empty() {
        slide.name = paragraphs.join(" ");
        *title_set = true;
        slide.elements.push(SlideElement::TextBox(TextBox {
            frame,
            text: paragraphs.join("\n"),
            style: TextBoxStyle::default(),
        }));
        return;
    }
    if !paragraphs.is_empty() {
        slide.elements.push(SlideElement::TextBox(TextBox {
            frame,
            text: paragraphs.join("\n"),
            style: TextBoxStyle::default(),
        }));
        return;
    }
    if let Some(shape_kind) = parse_pptx_shape_kind(shape) {
        slide.elements.push(SlideElement::Shape(Shape {
            frame,
            style: ShapeStyle::default(),
            kind: shape_kind,
        }));
    }
}

fn parse_pptx_transform(shape: &XmlNode) -> (f32, f32, f32, f32) {
    let xfrm = shape.child("spPr").and_then(|node| node.child("xfrm"));
    let off = xfrm.and_then(|node| node.child("off"));
    let ext = xfrm.and_then(|node| node.child("ext"));
    // EMUs to mm: 1 EMU = 1/360_000 cm = 1/36_000 mm
    let to_mm = |v: f32| v / 36_000.0;
    let x = off
        .and_then(|node| node.attr("x"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(to_mm)
        .unwrap_or(0.0);
    let y = off
        .and_then(|node| node.attr("y"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(to_mm)
        .unwrap_or(0.0);
    let width = ext
        .and_then(|node| node.attr("cx"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(to_mm)
        .unwrap_or(80.0);
    let height = ext
        .and_then(|node| node.attr("cy"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(to_mm)
        .unwrap_or(20.0);
    (x, y, width, height)
}

fn parse_pptx_text_body(body: &XmlNode) -> Vec<String> {
    let mut paragraphs = Vec::new();
    for paragraph in body.children_named("p") {
        let text = collect_paragraph_text(paragraph);
        if !text.trim().is_empty() {
            paragraphs.push(text.trim().to_string());
        }
    }
    paragraphs
}

fn collect_paragraph_text(node: &XmlNode) -> String {
    let mut out = String::new();
    for item in &node.items {
        match item {
            XmlItem::Text(text) => out.push_str(text),
            XmlItem::Node(child) => match child.local_name() {
                "t" => out.push_str(&child.text_content()),
                "br" => out.push('\n'),
                _ => out.push_str(&collect_paragraph_text(child)),
            },
        }
    }
    out
}

fn parse_pptx_shape_kind(shape: &XmlNode) -> Option<ShapeKind> {
    let prst = shape
        .child("spPr")
        .and_then(|node| node.child("prstGeom"))
        .and_then(|node| node.attr("prst"))
        .unwrap_or("");
    match prst {
        "ellipse" => Some(ShapeKind::Ellipse),
        "line" | "straightConnector1" => Some(ShapeKind::Line),
        "rect" | "roundRect" | "triangle" | "diamond" => Some(ShapeKind::Rectangle),
        _ => None,
    }
}

fn parse_odp_slide(page: &XmlNode) -> Slide {
    let mut slide = Slide {
        name: page.attr("name").unwrap_or("Slide").to_string(),
        elements: Vec::new(),
        notes: Vec::new(),
            chart_tokens: Vec::new(),
    };
    for item in &page.items {
        let XmlItem::Node(node) = item else {
            continue;
        };
        match node.local_name() {
            "frame" => parse_odp_frame(node, &mut slide),
            "ellipse" => slide.elements.push(SlideElement::Shape(Shape {
                kind: ShapeKind::Ellipse,
                frame: Rect::new(
                    Length::mm(parse_odf_len_mm(node.attr("x")).unwrap_or(0.0)),
                    Length::mm(parse_odf_len_mm(node.attr("y")).unwrap_or(0.0)),
                    Length::mm(parse_odf_len_mm(node.attr("width")).unwrap_or(40.0)),
                    Length::mm(parse_odf_len_mm(node.attr("height")).unwrap_or(25.0)),
                ),
                style: ShapeStyle::default(),
            })),
            "line" => {
                let x1 = parse_odf_len_mm(node.attr("x1")).unwrap_or(0.0);
                let y1 = parse_odf_len_mm(node.attr("y1")).unwrap_or(0.0);
                let x2 = parse_odf_len_mm(node.attr("x2")).unwrap_or(0.0);
                let y2 = parse_odf_len_mm(node.attr("y2")).unwrap_or(0.0);
                slide.elements.push(SlideElement::Shape(Shape {
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
            "custom-shape" | "rect" => slide.elements.push(SlideElement::Shape(Shape {
                kind: ShapeKind::Rectangle,
                frame: Rect::new(
                    Length::mm(parse_odf_len_mm(node.attr("x")).unwrap_or(0.0)),
                    Length::mm(parse_odf_len_mm(node.attr("y")).unwrap_or(0.0)),
                    Length::mm(parse_odf_len_mm(node.attr("width")).unwrap_or(40.0)),
                    Length::mm(parse_odf_len_mm(node.attr("height")).unwrap_or(25.0)),
                ),
                style: ShapeStyle::default(),
            })),
            _ => {}
        }
    }
    slide
}

fn parse_odp_frame(frame: &XmlNode, slide: &mut Slide) {
    let x = parse_odf_len_mm(frame.attr("x")).unwrap_or(0.0);
    let y = parse_odf_len_mm(frame.attr("y")).unwrap_or(0.0);
    let width = parse_odf_len_mm(frame.attr("width")).unwrap_or(60.0);
    let height = parse_odf_len_mm(frame.attr("height")).unwrap_or(20.0);
    let Some(text_box) = frame.child("text-box") else {
        return;
    };
    let paragraphs: Vec<String> = text_box
        .children
        .iter()
        .filter(|child| matches!(child.local_name(), "p" | "h"))
        .map(|node| node.text_content())
        .filter(|text| !text.trim().is_empty())
        .collect();
    if paragraphs.is_empty() {
        return;
    }
    if slide.elements.is_empty() {
        slide.name = paragraphs[0].clone();
    }
    slide.elements.push(SlideElement::TextBox(TextBox {
        frame: Rect::new(
            Length::mm(x),
            Length::mm(y),
            Length::mm(width),
            Length::mm(height),
        ),
        text: paragraphs.join("\n"),
        style: TextBoxStyle::default(),
    }));
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
    use crate::{demo_presentation, to_pptx};

    #[test]
    fn pptx_round_trip_imports_slide_text() {
        let deck = demo_presentation("Demo");
        let bytes = to_pptx(&deck).expect("pptx");
        let loaded = from_pptx_bytes("Demo", &bytes).expect("import pptx");
        assert!(!loaded.slides.is_empty());
    }

    #[test]
    fn odp_round_trip_imports_slide_text() {
        let deck = demo_presentation("Demo");
        let tmp = std::env::temp_dir().join("lo_impress_import_test.odp");
        lo_odf::save_presentation_document(&tmp, &deck).expect("save odp");
        let bytes = std::fs::read(&tmp).expect("read odp");
        let _ = std::fs::remove_file(&tmp);
        let loaded = from_odp_bytes("Demo", &bytes).expect("import odp");
        assert!(!loaded.slides.is_empty());
    }

    #[test]
    fn text_outline_creates_slides() {
        let deck = from_text_outline("Outline", "# First\n- one\n- two\n# Second\n- three\n");
        assert_eq!(deck.slides.len(), 2);
    }
}
