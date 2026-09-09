//! Binary importers for `TextDocument`.
//!
//! Supported formats:
//! - `txt`/`text`, `md`/`markdown`, `html`/`htm` (string-based, no zip)
//! - `pdf` (native PDF text import)
//! - `docx` (Office Open XML)
//! - `odt` (OpenDocument Text)

mod docx;
pub use docx::from_docx_bytes;

use std::collections::BTreeMap;

use lo_core::{
    parse_xml_document, Block, Heading, Inline, ListBlock, ListItem, LoError, Paragraph, Result,
    Table, TableCell, TableRow, TextDocument, XmlItem, XmlNode,
};
use lo_zip::{rels_path_for, resolve_part_target, ZipArchive};

use crate::{from_markdown, from_plain_text};

/// Style hints attached to a span while we walk the source document.
#[derive(Clone, Debug, Default, PartialEq)]
struct SpanStyle {
    bold: bool,
    italic: bool,
    code: bool,
    link: Option<String>,
    font_size_pt: Option<u16>,
}

#[derive(Clone, Debug, Default)]
struct StyleProps {
    bold: bool,
    italic: bool,
    code: bool,
    font_size_pt: Option<u16>,
    heading_level: Option<u8>,
    page_break_before: bool,
}

#[derive(Clone, Debug, Default)]
struct StyledSpan {
    text: String,
    style: SpanStyle,
}

#[derive(Clone, Debug, Default)]
struct ParagraphInfo {
    spans: Vec<StyledSpan>,
    heading_level: Option<u8>,
    list_key: Option<String>,
    page_break: bool,
}

/// Dispatch to the appropriate importer based on a format hint.
pub fn load_bytes(title: impl Into<String>, bytes: &[u8], format: &str) -> Result<TextDocument> {
    let title = title.into();
    match format.to_ascii_lowercase().as_str() {
        "txt" | "text" => Ok(from_plain_text(title, &bytes_to_utf8(bytes)?)),
        "md" | "markdown" => Ok(from_markdown(title, &bytes_to_utf8(bytes)?)),
        "html" | "htm" => Ok(from_html(title, &bytes_to_utf8(bytes)?)),
        "pdf" => from_pdf_bytes(title, bytes),
        "docx" => from_docx_bytes(title, bytes),
        "doc" => from_doc_bytes(title, bytes),
        "odt" => from_odt_bytes(title, bytes),
        other => Err(LoError::Unsupported(format!(
            "writer import format {other}"
        ))),
    }
}

/// Read text from a legacy binary `.doc` (Word 97-2003) file. The CFB
/// stream is parsed natively; the resulting plain text is wrapped in a
/// `TextDocument` via `from_plain_text`.
pub fn from_doc_bytes(title: impl Into<String>, bytes: &[u8]) -> Result<TextDocument> {
    let text = crate::legacy_doc::extract_text_from_doc(bytes)?;
    Ok(from_plain_text(title, &text))
}

/// Read text from a PDF byte stream using the shared native PDF parser.
/// Each extracted page is mapped to paragraphs, with explicit page-break
/// blocks inserted between pages.
pub fn from_pdf_bytes(title: impl Into<String>, bytes: &[u8]) -> Result<TextDocument> {
    let title = title.into();
    let pages = lo_core::extract_pages_from_pdf(bytes)?;
    let mut doc = TextDocument::new(title);
    let mut emitted_any = false;
    for (page_index, page_text) in pages.iter().enumerate() {
        if page_index > 0 && emitted_any {
            doc.body.push(Block::PageBreak);
        }
        let normalized = page_text.replace("\u{000C}", "\n");
        let mut page_emitted = false;
        for paragraph in normalized.split("\n\n") {
            let trimmed = paragraph.trim();
            if trimmed.is_empty() {
                continue;
            }
            let joined = trimmed
                .lines()
                .map(str::trim)
                .filter(|line| !line.is_empty())
                .collect::<Vec<_>>()
                .join(" ");
            if !joined.is_empty() {
                doc.body.push(Block::Paragraph(Paragraph::plain(joined)));
                emitted_any = true;
                page_emitted = true;
            }
        }
        if !page_emitted && !normalized.trim().is_empty() {
            doc.body
                .push(Block::Paragraph(Paragraph::plain(normalized.trim().to_string())));
            emitted_any = true;
        }
    }
    if !emitted_any {
        let text = lo_core::extract_text_from_pdf(bytes)?;
        let trimmed = text.trim();
        if !trimmed.is_empty() {
            doc.body
                .push(Block::Paragraph(Paragraph::plain(trimmed.to_string())));
            emitted_any = true;
        }
    }
    if !emitted_any {
        doc.body.push(Block::Paragraph(Paragraph::default()));
    }
    Ok(doc)
}

fn bytes_to_utf8(bytes: &[u8]) -> Result<String> {
    String::from_utf8(bytes.to_vec())
        .map_err(|err| LoError::Parse(format!("invalid utf-8 input: {err}")))
}

/// Lossy HTML to text conversion. Falls back through `from_plain_text` so
/// downstream callers still get a structured document.
pub fn from_html(title: impl Into<String>, html: &str) -> TextDocument {
    let normalized = html
        .replace("<br>", "\n")
        .replace("<br/>", "\n")
        .replace("<br />", "\n")
        .replace("</p>", "\n\n")
        .replace("</div>", "\n")
        .replace("</li>", "\n")
        .replace("<li>", "- ")
        .replace("</h1>", "\n\n")
        .replace("</h2>", "\n\n")
        .replace("</h3>", "\n\n");
    let mut plain = String::new();
    let mut in_tag = false;
    for ch in normalized.chars() {
        match ch {
            '<' => in_tag = true,
            '>' => in_tag = false,
            _ if !in_tag => plain.push(ch),
            _ => {}
        }
    }
    from_plain_text(title, &lo_core::decode_entities(&plain))
}

fn renumber_footnote_markers(doc: &mut TextDocument) {
    let mut foot = 0u32;
    let mut endn = 0u32;
    fn replace_in_text(s: &str, foot: &mut u32, endn: &mut u32) -> String {
        let mut out = String::with_capacity(s.len());
        for ch in s.chars() {
            match ch {
                '\u{f001}' => {
                    *foot += 1;
                    out.push_str(&foot.to_string());
                }
                '\u{f002}' => {
                    *endn += 1;
                    out.push_str(&to_lower_roman(*endn));
                }
                other => out.push(other),
            }
        }
        out
    }
    fn walk_inline(inline: &mut Inline, foot: &mut u32, endn: &mut u32) {
        match inline {
            Inline::Text(t)
            | Inline::Bold(t)
            | Inline::Italic(t)
            | Inline::Code(t) => *t = replace_in_text(t, foot, endn),
            Inline::Link { label, .. } => *label = replace_in_text(label, foot, endn),
            Inline::Styled { text, .. } => *text = replace_in_text(text, foot, endn),
            Inline::LineBreak => {}
        }
    }
    fn walk_paragraph(p: &mut Paragraph, foot: &mut u32, endn: &mut u32) {
        for span in &mut p.spans {
            walk_inline(span, foot, endn);
        }
    }
    fn walk_block(block: &mut Block, foot: &mut u32, endn: &mut u32) {
        match block {
            Block::Paragraph(p) => walk_paragraph(p, foot, endn),
            Block::Heading(h) => walk_paragraph(&mut h.content, foot, endn),
            Block::List(list) => {
                for item in &mut list.items {
                    for sub in &mut item.blocks {
                        walk_block(sub, foot, endn);
                    }
                }
            }
            Block::Table(table) => {
                for row in &mut table.rows {
                    for cell in &mut row.cells {
                        for p in &mut cell.paragraphs {
                            walk_paragraph(p, foot, endn);
                        }
                    }
                }
            }
            _ => {}
        }
    }
    for block in &mut doc.body {
        walk_block(block, &mut foot, &mut endn);
    }
}

fn to_lower_roman(mut value: u32) -> String {
    if value == 0 {
        return "0".to_string();
    }
    let table = [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ];
    let mut out = String::new();
    for (n, s) in table {
        while value >= n {
            out.push_str(s);
            value -= n;
        }
    }
    out
}

fn collect_w_text_nodes(node: &XmlNode, out: &mut Vec<String>) {
    if node.local_name() == "t" {
        out.push(node.text_content());
    }
    for child in &node.children {
        collect_w_text_nodes(child, out);
    }
}

pub fn from_odt_bytes(title: impl Into<String>, bytes: &[u8]) -> Result<TextDocument> {
    let title = title.into();
    let zip = ZipArchive::new(bytes)?;
    let content = parse_xml_document(&zip.read_string("content.xml")?)?;
    let mut styles = if zip.contains("styles.xml") {
        parse_odt_styles(&parse_xml_document(&zip.read_string("styles.xml")?)?)
    } else {
        BTreeMap::new()
    };
    merge_styles(&mut styles, &parse_odt_styles(&content));
    let body = content
        .child("body")
        .ok_or_else(|| LoError::Parse("content.xml missing office:body".to_string()))?;
    let text = body
        .child("text")
        .ok_or_else(|| LoError::Parse("content.xml missing office:text".to_string()))?;

    let mut doc = TextDocument::new(title);
    let mut pending_list = Vec::new();
    parse_odt_text_children(text, &styles, &mut doc, &mut pending_list);
    flush_list(&mut doc, &mut pending_list);
    if doc.body.is_empty() {
        doc.body.push(Block::Paragraph(Paragraph::default()));
    }
    Ok(doc)
}

// ---------------------------------------------------------------------------
// Span/paragraph helpers
// ---------------------------------------------------------------------------

fn flush_list(doc: &mut TextDocument, pending: &mut Vec<ListItem>) {
    flush_list_with(doc, pending, false);
}

fn flush_list_with(doc: &mut TextDocument, pending: &mut Vec<ListItem>, ordered: bool) {
    if !pending.is_empty() {
        doc.body.push(Block::List(ListBlock {
            ordered,
            items: std::mem::take(pending),
        }));
    }
}

/// Build a `numId -> ordered?` map by walking `word/numbering.xml` and
/// resolving every `<w:num w:numId>` to its `<w:abstractNum>` definition,
/// then peeking at level 0's `<w:numFmt w:val>` to decide whether the
/// list should be rendered as bullets or with numeric markers.
fn parse_docx_numbering(zip: &ZipArchive) -> BTreeMap<String, bool> {
    let mut out: BTreeMap<String, bool> = BTreeMap::new();
    if !zip.contains("word/numbering.xml") {
        return out;
    }
    let xml = match zip.read_string("word/numbering.xml") {
        Ok(x) => x,
        Err(_) => return out,
    };
    let root = match parse_xml_document(&xml) {
        Ok(r) => r,
        Err(_) => return out,
    };
    // Map abstractNumId -> ordered?
    let mut abstracts: BTreeMap<String, bool> = BTreeMap::new();
    for an in root.children_named("abstractNum") {
        let Some(id) = an.attr("abstractNumId") else { continue };
        let mut ordered = false;
        for lvl in an.children_named("lvl") {
            if lvl.attr("ilvl").unwrap_or("") != "0" {
                continue;
            }
            if let Some(fmt) = lvl.child("numFmt").and_then(|n| n.attr("val")) {
                let lower = fmt.to_ascii_lowercase();
                ordered = !lower.contains("bullet") && !lower.is_empty();
            }
        }
        abstracts.insert(id.to_string(), ordered);
    }
    for num in root.children_named("num") {
        let Some(num_id) = num.attr("numId") else { continue };
        let abstract_id = num
            .child("abstractNumId")
            .and_then(|n| n.attr("val"))
            .unwrap_or_default();
        let ordered = abstracts.get(abstract_id).copied().unwrap_or(false);
        out.insert(num_id.to_string(), ordered);
    }
    out
}

fn normalize_spans(spans: Vec<StyledSpan>) -> Vec<StyledSpan> {
    let mut out: Vec<StyledSpan> = Vec::new();
    for span in spans {
        if span.text.is_empty() {
            continue;
        }
        if let Some(last) = out.last_mut() {
            if last.style == span.style {
                last.text.push_str(&span.text);
                continue;
            }
        }
        out.push(span);
    }
    out
}

fn build_paragraph(spans: Vec<StyledSpan>) -> Paragraph {
    let normalized = normalize_spans(spans);
    let mut inlines: Vec<Inline> = Vec::new();
    for span in normalized {
        inlines.push(span_to_inline(span));
    }
    Paragraph {
        spans: inlines,
        ..Paragraph::default()
    }
}

fn span_to_inline(span: StyledSpan) -> Inline {
    if let Some(font_size_pt) = span.style.font_size_pt {
        return Inline::Styled {
            text: span.text,
            style: lo_core::TextStyle {
                font_size_pt, bold: span.style.bold, italic: span.style.italic,
                font_family: if span.style.code { "monospace".into() } else { String::new() },
                ..lo_core::TextStyle::default()
            },
            url: span.style.link,
        };
    }
    if let Some(url) = span.style.link {
        return Inline::Link {
            label: span.text,
            url,
        };
    }
    if span.style.code {
        return Inline::Code(span.text);
    }
    if span.style.bold {
        return Inline::Bold(span.text);
    }
    if span.style.italic {
        return Inline::Italic(span.text);
    }
    Inline::Text(span.text)
}

fn inline_text(inline: &Inline) -> &str {
    match inline {
        Inline::Text(text) | Inline::Bold(text) | Inline::Italic(text) | Inline::Code(text) => text,
        Inline::Link { label, .. } => label,
        Inline::Styled { text, .. } => text,
        Inline::LineBreak => "\n",
    }
}

// ---------------------------------------------------------------------------
// DOCX parsing
// ---------------------------------------------------------------------------

fn parse_relationships(zip: &ZipArchive, part: &str) -> Result<BTreeMap<String, String>> {
    let rels_path = rels_path_for(part);
    if !zip.contains(&rels_path) {
        return Ok(BTreeMap::new());
    }
    let rels = parse_xml_document(&zip.read_string(&rels_path)?)?;
    let mut map = BTreeMap::new();
    for rel in rels.children_named("Relationship") {
        if let Some(id) = rel.attr("Id") {
            if let Some(target) = rel.attr("Target") {
                let resolved = if rel.attr("TargetMode") == Some("External") {
                    target.to_string()
                } else {
                    resolve_part_target(part, target)
                };
                map.insert(id.to_string(), resolved);
            }
        }
    }
    Ok(map)
}

fn parse_odt_styles(root: &XmlNode) -> BTreeMap<String, StyleProps> {
    let mut nodes = Vec::new();
    root.descendants_named("style", &mut nodes);
    let mut styles = BTreeMap::new();
    for style in nodes {
        let Some(name) = style.attr("name").or_else(|| style.attr("style:name")) else {
            continue;
        };
        let mut props = StyleProps::default();
        if let Some(text_props) = style.child("text-properties") {
            if matches!(text_props.attr("font-weight"), Some("bold")) {
                props.bold = true;
            }
            if matches!(text_props.attr("font-style"), Some("italic")) {
                props.italic = true;
            }
            if let Some(font_name) = text_props.attr("font-name") {
                let lower = font_name.to_ascii_lowercase();
                if lower.contains("mono") || lower.contains("courier") {
                    props.code = true;
                }
            }
        }
        if let Some(paragraph_props) = style.child("paragraph-properties") {
            if paragraph_props.attr("break-before") == Some("page") {
                props.page_break_before = true;
            }
        }
        styles.insert(name.to_string(), props);
    }
    styles
}

fn merge_styles(target: &mut BTreeMap<String, StyleProps>, source: &BTreeMap<String, StyleProps>) {
    for (name, props) in source {
        target.insert(name.clone(), props.clone());
    }
}

fn parse_odt_text_children(
    root: &XmlNode,
    styles: &BTreeMap<String, StyleProps>,
    doc: &mut TextDocument,
    pending_list: &mut Vec<ListItem>,
) {
    for item in &root.items {
        let XmlItem::Node(node) = item else {
            continue;
        };
        match node.local_name() {
            "p" => {
                let props = node
                    .attr("style-name")
                    .and_then(|name| styles.get(name))
                    .cloned()
                    .unwrap_or_default();
                if props.page_break_before {
                    flush_list(doc, pending_list);
                    doc.body.push(Block::PageBreak);
                }
                flush_list(doc, pending_list);
                doc.body
                    .push(Block::Paragraph(build_paragraph(parse_odt_inline(
                        node,
                        styles,
                        &SpanStyle::default(),
                    ))));
            }
            "h" => {
                flush_list(doc, pending_list);
                let level = node
                    .attr("outline-level")
                    .and_then(|value| value.parse::<u8>().ok())
                    .unwrap_or(1)
                    .clamp(1, 6);
                doc.body.push(Block::Heading(Heading {
                    level,
                    content: build_paragraph(parse_odt_inline(node, styles, &SpanStyle::default())),
                }));
            }
            "list" => {
                flush_list(doc, pending_list);
                let mut items: Vec<ListItem> = Vec::new();
                for list_item in node.children_named("list-item") {
                    let paragraph_node = list_item
                        .children
                        .iter()
                        .find(|child| matches!(child.local_name(), "p" | "h"));
                    if let Some(paragraph_node) = paragraph_node {
                        items.push(ListItem {
                            blocks: vec![Block::Paragraph(build_paragraph(parse_odt_inline(
                                paragraph_node,
                                styles,
                                &SpanStyle::default(),
                            )))],
                        });
                    }
                }
                if !items.is_empty() {
                    doc.body.push(Block::List(ListBlock {
                        ordered: false,
                        items,
                    }));
                }
            }
            "table" => {
                flush_list(doc, pending_list);
                doc.body.push(Block::Table(parse_odt_table(node, styles)));
            }
            "section" => parse_odt_text_children(node, styles, doc, pending_list),
            "frame" => {
                for text_box in node.children_named("text-box") {
                    parse_odt_text_children(text_box, styles, doc, pending_list);
                }
            }
            _ => {}
        }
    }
}

fn parse_odt_inline(
    node: &XmlNode,
    styles: &BTreeMap<String, StyleProps>,
    inherited: &SpanStyle,
) -> Vec<StyledSpan> {
    let mut spans = Vec::new();
    parse_odt_items(&node.items, styles, inherited, &mut spans);
    spans
}

fn parse_odt_items(
    items: &[XmlItem],
    styles: &BTreeMap<String, StyleProps>,
    inherited: &SpanStyle,
    spans: &mut Vec<StyledSpan>,
) {
    for item in items {
        match item {
            XmlItem::Text(text) => {
                if !text.is_empty() {
                    spans.push(StyledSpan {
                        text: text.clone(),
                        style: inherited.clone(),
                    });
                }
            }
            XmlItem::Node(node) => match node.local_name() {
                "span" => {
                    let mut style = inherited.clone();
                    if let Some(style_name) = node.attr("style-name") {
                        if let Some(props) = styles.get(style_name) {
                            style.bold |= props.bold;
                            style.italic |= props.italic;
                            style.code |= props.code;
                        }
                    }
                    parse_odt_items(&node.items, styles, &style, spans);
                }
                "a" => {
                    let mut style = inherited.clone();
                    style.link = node
                        .attr("href")
                        .or_else(|| node.attr("xlink:href"))
                        .map(str::to_string);
                    parse_odt_items(&node.items, styles, &style, spans);
                }
                "s" => {
                    let count = node
                        .attr("c")
                        .and_then(|value| value.parse::<usize>().ok())
                        .unwrap_or(1);
                    spans.push(StyledSpan {
                        text: " ".repeat(count),
                        style: inherited.clone(),
                    });
                }
                "tab" => spans.push(StyledSpan {
                    text: "\t".to_string(),
                    style: inherited.clone(),
                }),
                "line-break" => spans.push(StyledSpan {
                    text: "\n".to_string(),
                    style: inherited.clone(),
                }),
                _ => parse_odt_items(&node.items, styles, inherited, spans),
            },
        }
    }
}

fn parse_odt_table(table: &XmlNode, styles: &BTreeMap<String, StyleProps>) -> Table {
    let mut rows: Vec<TableRow> = Vec::new();
    for row in table.children_named("table-row") {
        let repeat_rows = row
            .attr("number-rows-repeated")
            .and_then(|value| value.parse::<usize>().ok())
            .unwrap_or(1);
        let mut cells: Vec<TableCell> = Vec::new();
        for cell in row
            .children
            .iter()
            .filter(|child| matches!(child.local_name(), "table-cell" | "covered-table-cell"))
        {
            let repeat_cols = cell
                .attr("number-columns-repeated")
                .and_then(|value| value.parse::<usize>().ok())
                .unwrap_or(1);
            let mut paragraphs: Vec<Paragraph> = Vec::new();
            for paragraph in cell
                .children
                .iter()
                .filter(|child| matches!(child.local_name(), "p" | "h"))
            {
                let para =
                    build_paragraph(parse_odt_inline(paragraph, styles, &SpanStyle::default()));
                paragraphs.push(para);
            }
            if paragraphs.is_empty() {
                paragraphs.push(Paragraph::default());
            }
            for _ in 0..repeat_cols {
                cells.push(TableCell {
                    paragraphs: paragraphs.clone(),
                });
            }
        }
        for _ in 0..repeat_rows {
            rows.push(TableRow {
                cells: cells.clone(),
            });
        }
    }
    Table {
        name: "Table1".to_string(),
        rows,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{from_markdown, to_docx};
    use lo_odf::save_text_document;

    #[test]
    fn docx_round_trip_imports_basic_structure() {
        let doc = from_markdown("demo", "# Title\n\nHello **world**\n\n- one\n- two");
        let bytes = to_docx(&doc).expect("docx");
        let loaded = from_docx_bytes("demo", &bytes).expect("import docx");
        let text = loaded.plain_text();
        assert!(text.contains("Title"));
        assert!(text.contains("world"));
    }

    #[test]
    fn odt_round_trip_imports_basic_structure() {
        let doc = from_markdown("demo", "# Title\n\nHello *world*");
        let tmp = std::env::temp_dir().join("lo_writer_import_test.odt");
        save_text_document(&tmp, &doc).expect("save odt");
        let bytes = std::fs::read(&tmp).expect("read odt");
        let _ = std::fs::remove_file(&tmp);
        let loaded = from_odt_bytes("demo", &bytes).expect("import odt");
        let text = loaded.plain_text();
        assert!(text.contains("Title"));
        assert!(text.contains("Hello"));
    }

    #[test]
    fn html_to_text_roundtrip() {
        let doc = from_html("h", "<h1>hi</h1><p>one<br/>two</p>");
        let text = doc.plain_text();
        assert!(text.contains("hi"));
        assert!(text.contains("one"));
        assert!(text.contains("two"));
    }

    #[test]
    fn load_bytes_dispatches() {
        let html = b"<p>hello</p>";
        let doc = load_bytes("h", html, "html").unwrap();
        assert!(doc.plain_text().contains("hello"));
    }
}
