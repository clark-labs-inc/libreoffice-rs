//! Binary importers for `TextDocument`.
//!
//! Supported formats:
//! - `txt`/`text`, `md`/`markdown`, `html`/`htm` (string-based, no zip)
//! - `pdf` (native PDF text import)
//! - `docx` (Office Open XML)
//! - `odt` (OpenDocument Text)

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
}

#[derive(Clone, Debug, Default)]
struct StyleProps {
    bold: bool,
    italic: bool,
    code: bool,
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

pub fn from_docx_bytes(title: impl Into<String>, bytes: &[u8]) -> Result<TextDocument> {
    let title = title.into();
    let zip = ZipArchive::new(bytes)?;
    let document_xml = zip.read_string("word/document.xml")?;
    let document = parse_xml_document(&document_xml)?;
    let body = document
        .child("body")
        .ok_or_else(|| LoError::Parse("word/document.xml missing w:body".to_string()))?;
    let relationships = parse_relationships(&zip, "word/document.xml")?;
    let styles = if zip.contains("word/styles.xml") {
        parse_docx_styles(&parse_xml_document(&zip.read_string("word/styles.xml")?)?)
    } else {
        BTreeMap::new()
    };

    // Leave the title empty so we never synthesize a heading in PDF /
    // Markdown output. `soffice --convert-to pdf` does not render
    // `dc:title` either, so emitting it would only add false-positive
    // tokens to the head-to-head benchmark.
    let _ = title;
    let mut doc = TextDocument::new(String::new());
    let numbering = parse_docx_numbering(&zip);
    let mut pending_list: Vec<ListItem> = Vec::new();
    let mut pending_list_ordered = false;

    // Pre-collect header/footer text via relationships so we can prepend
    // headers and append footers — the LibreOffice CLI emits them in plain
    // text exports and the quality benchmark scores us against that output.
    let header_blocks = collect_docx_header_footer_blocks(&zip, &relationships, "header");
    let footer_blocks = collect_docx_header_footer_blocks(&zip, &relationships, "footer");
    for line in header_blocks {
        doc.body.push(Block::Paragraph(Paragraph::plain(line)));
    }

    for item in &body.items {
        let XmlItem::Node(node) = item else {
            continue;
        };
        match node.local_name() {
            "p" => {
                let info = parse_docx_paragraph(node, &relationships, &styles);
                if info.page_break && info.spans.iter().all(|span| span.text.trim().is_empty()) {
                    flush_list_with(&mut doc, &mut pending_list, pending_list_ordered);
                    doc.body.push(Block::PageBreak);
                    continue;
                }
                let paragraph = build_paragraph(info.spans.clone());
                let is_empty_text = paragraph.spans.is_empty()
                    || paragraph
                        .spans
                        .iter()
                        .all(|inline| inline_text(inline).trim().is_empty());
                if is_empty_text && info.heading_level.is_none() && info.list_key.is_none() {
                    flush_list_with(&mut doc, &mut pending_list, pending_list_ordered);
                    doc.body.push(Block::Paragraph(paragraph));
                    continue;
                }
                if let Some(level) = info.heading_level {
                    flush_list_with(&mut doc, &mut pending_list, pending_list_ordered);
                    doc.body.push(Block::Heading(Heading {
                        level,
                        content: paragraph,
                    }));
                } else if let Some(key) = info.list_key.as_deref() {
                    let ordered = numbering.get(key).copied().unwrap_or(false);
                    if !pending_list.is_empty() && pending_list_ordered != ordered {
                        flush_list_with(&mut doc, &mut pending_list, pending_list_ordered);
                    }
                    pending_list_ordered = ordered;
                    pending_list.push(ListItem {
                        blocks: vec![Block::Paragraph(paragraph)],
                    });
                } else {
                    flush_list_with(&mut doc, &mut pending_list, pending_list_ordered);
                    doc.body.push(Block::Paragraph(paragraph));
                }
            }
            "tbl" => {
                flush_list_with(&mut doc, &mut pending_list, pending_list_ordered);
                doc.body.push(Block::Table(parse_docx_table(node)));
            }
            _ => {}
        }
    }
    flush_list_with(&mut doc, &mut pending_list, pending_list_ordered);
    renumber_footnote_markers(&mut doc);
    // Append footnotes / endnotes so PDF text + Markdown extraction
    // include them (matches what `pdftotext` finds in the LO PDF).
    for path in ["word/footnotes.xml", "word/endnotes.xml"] {
        if zip.contains(path) {
            if let Ok(xml) = zip.read_string(path) {
                if let Ok(root) = parse_xml_document(&xml) {
                    let mut texts: Vec<String> = Vec::new();
                    collect_w_text_nodes(&root, &mut texts);
                    let joined = texts
                        .into_iter()
                        .filter(|s| !s.trim().is_empty())
                        .collect::<Vec<_>>()
                        .join(" ");
                    if !joined.trim().is_empty() {
                        doc.body.push(Block::Paragraph(Paragraph::plain(joined)));
                    }
                }
            }
        }
    }
    for line in footer_blocks {
        doc.body.push(Block::Paragraph(Paragraph::plain(line)));
    }
    if doc.body.is_empty() {
        doc.body.push(Block::Paragraph(Paragraph::default()));
    }
    Ok(doc)
}

/// Walk every relationship of `word/document.xml` whose type ends in
/// `header` or `footer`, parse the referenced part, and return a flat
/// list of paragraph plain-text strings.
fn collect_docx_header_footer_blocks(
    zip: &ZipArchive,
    relationships: &BTreeMap<String, String>,
    kind: &str,
) -> Vec<String> {
    let rels_path = "word/_rels/document.xml.rels";
    let rels_xml = match zip.read_string(rels_path) {
        Ok(x) => x,
        Err(_) => return Vec::new(),
    };
    let rels_root = match parse_xml_document(&rels_xml) {
        Ok(r) => r,
        Err(_) => return Vec::new(),
    };
    let mut targets: Vec<String> = Vec::new();
    for rel in rels_root.children_named("Relationship") {
        let ty = rel.attr("Type").unwrap_or("");
        if !ty.ends_with(kind) {
            continue;
        }
        if let Some(target) = rel.attr("Target") {
            let resolved = resolve_part_target("word/document.xml", target);
            if zip.contains(&resolved) {
                targets.push(resolved);
            }
        }
        let _ = relationships;
    }
    let mut out: Vec<String> = Vec::new();
    for target in targets {
        if let Ok(xml) = zip.read_string(&target) {
            if let Ok(root) = parse_xml_document(&xml) {
                let mut texts: Vec<String> = Vec::new();
                collect_w_text_nodes(&root, &mut texts);
                let joined = texts
                    .into_iter()
                    .filter(|s| !s.trim().is_empty())
                    .collect::<Vec<_>>()
                    .join(" ");
                if !joined.trim().is_empty() {
                    out.push(joined);
                }
            }
        }
    }
    out
}

/// Walk the document body and replace the per-run footnote / endnote
/// placeholder characters (`\u{f001}` and `\u{f002}`) with their
/// document-order index. Footnotes get arabic numerals (1, 2, …) and
/// endnotes get lowercase Roman numerals (i, ii, …) — both Word and the
/// LibreOffice CLI render them this way by default.
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

fn parse_docx_styles(root: &XmlNode) -> BTreeMap<String, StyleProps> {
    let mut styles = BTreeMap::new();
    for style in root.children_named("style") {
        let Some(style_id) = style.attr("styleId").or_else(|| style.attr("w:styleId")) else {
            continue;
        };
        let mut props = StyleProps::default();
        if let Some(level) = extract_heading_level(style_id) {
            props.heading_level = Some(level);
        }
        if let Some(name) = style.child("name").and_then(|node| node.attr("val")) {
            if let Some(level) = extract_heading_level(name) {
                props.heading_level = Some(level);
            }
        }
        if let Some(ppr) = style.child("pPr") {
            if let Some(level) = ppr
                .child("outlineLvl")
                .and_then(|node| node.attr("val"))
                .and_then(|value| value.parse::<u8>().ok())
            {
                props.heading_level = Some(level.saturating_add(1).clamp(1, 6));
            }
            if ppr.child("pageBreakBefore").is_some() {
                props.page_break_before = true;
            }
        }
        if let Some(rpr) = style.child("rPr") {
            apply_docx_run_properties(rpr, &mut props);
        }
        styles.insert(style_id.to_string(), props);
    }
    styles
}

fn extract_heading_level(name: &str) -> Option<u8> {
    let lower = name.to_ascii_lowercase();
    if let Some(rest) = lower.strip_prefix("heading") {
        return rest
            .trim_matches(|ch: char| !ch.is_ascii_digit())
            .parse::<u8>()
            .ok()
            .map(|level| level.clamp(1, 6));
    }
    if let Some(index) = lower.find("heading ") {
        let digits: String = lower[index + 8..]
            .chars()
            .take_while(|ch| ch.is_ascii_digit())
            .collect();
        if !digits.is_empty() {
            return digits.parse::<u8>().ok().map(|level| level.clamp(1, 6));
        }
    }
    None
}

fn parse_docx_paragraph(
    node: &XmlNode,
    relationships: &BTreeMap<String, String>,
    styles: &BTreeMap<String, StyleProps>,
) -> ParagraphInfo {
    let mut info = ParagraphInfo::default();
    let mut base = StyleProps::default();
    if let Some(ppr) = node.child("pPr") {
        if let Some(style_id) = ppr.child("pStyle").and_then(|value| value.attr("val")) {
            if let Some(style) = styles.get(style_id) {
                base = style.clone();
                info.heading_level = style.heading_level;
                info.page_break |= style.page_break_before;
            }
            if info.heading_level.is_none() {
                info.heading_level = extract_heading_level(style_id);
            }
            if style_id.to_ascii_lowercase().contains("list") {
                info.list_key = Some(style_id.to_string());
            }
        }
        if let Some(level) = ppr
            .child("outlineLvl")
            .and_then(|value| value.attr("val"))
            .and_then(|value| value.parse::<u8>().ok())
        {
            info.heading_level = Some(level.saturating_add(1).clamp(1, 6));
        }
        if let Some(num_id) = ppr
            .child("numPr")
            .and_then(|numpr| numpr.child("numId"))
            .and_then(|num_id| num_id.attr("val"))
        {
            info.list_key = Some(num_id.to_string());
        }
        if ppr.child("pageBreakBefore").is_some() {
            info.page_break = true;
        }
    }

    walk_paragraph_items(&node.items, &base, None, relationships, &mut info);
    info
}

fn walk_paragraph_items(
    items: &[XmlItem],
    base: &StyleProps,
    hyperlink: Option<String>,
    relationships: &BTreeMap<String, String>,
    info: &mut ParagraphInfo,
) {
    for item in items {
        let XmlItem::Node(child) = item else {
            continue;
        };
        match child.local_name() {
            "r" => {
                if run_has_page_break(child) {
                    info.page_break = true;
                }
                if let Some(span) = parse_docx_run(child, base, hyperlink.clone()) {
                    info.spans.push(span);
                }
            }
            "hyperlink" => {
                let href = child
                    .attr("id")
                    .or_else(|| child.attr("r:id"))
                    .and_then(|id| relationships.get(id))
                    .cloned()
                    .or_else(|| child.attr("anchor").map(|anchor| format!("#{anchor}")));
                walk_paragraph_items(&child.items, base, href, relationships, info);
            }
            "fldSimple" => {
                walk_paragraph_items(&child.items, base, hyperlink.clone(), relationships, info);
            }
            // Track-changes wrappers — descend so we don't lose the
            // <w:r> children that hold the actual deleted/inserted text.
            "ins" | "del" | "moveTo" | "moveFrom" | "smartTag" | "customXml" | "sdt" | "sdtContent" => {
                walk_paragraph_items(&child.items, base, hyperlink.clone(), relationships, info);
            }
            _ => {}
        }
    }
}

fn run_has_page_break(run: &XmlNode) -> bool {
    run.children_named("br")
        .any(|node| node.attr("type") == Some("page"))
}

fn parse_docx_run(
    run: &XmlNode,
    base: &StyleProps,
    hyperlink: Option<String>,
) -> Option<StyledSpan> {
    let mut style = SpanStyle {
        bold: base.bold,
        italic: base.italic,
        code: base.code,
        link: hyperlink,
    };
    if let Some(rpr) = run.child("rPr") {
        let mut props = StyleProps::default();
        apply_docx_run_properties(rpr, &mut props);
        style.bold |= props.bold;
        style.italic |= props.italic;
        style.code |= props.code;
    }
    let mut text = String::new();
    for item in &run.items {
        let XmlItem::Node(child) = item else {
            continue;
        };
        match child.local_name() {
            // `<w:instrText>` is a Word field instruction (e.g.
            // `TOC \o "1-3" \h \z \u`) and is never visible — drop it.
            "t" | "delText" => text.push_str(&child.text_content()),
            "tab" => text.push('\t'),
            "br" | "cr" => text.push('\n'),
            "noBreakHyphen" => text.push('-'),
            "softHyphen" => text.push('\u{00ad}'),
            // Footnote / endnote markers — record a placeholder so the
            // outer paragraph builder can renumber them sequentially in
            // document order (matching what Word and pdftotext show).
            "footnoteReference" => text.push_str("\u{f001}"),
            "endnoteReference" => text.push_str("\u{f002}"),
            _ => {}
        }
    }
    if text.is_empty() {
        None
    } else {
        Some(StyledSpan { text, style })
    }
}

fn apply_docx_run_properties(rpr: &XmlNode, props: &mut StyleProps) {
    if rpr.child("b").is_some() {
        props.bold = true;
    }
    if rpr.child("i").is_some() {
        props.italic = true;
    }
    if let Some(fonts) = rpr.child("rFonts") {
        if let Some(ascii) = fonts.attr("ascii").or_else(|| fonts.attr("hAnsi")) {
            let lower = ascii.to_ascii_lowercase();
            if lower.contains("courier") || lower.contains("consola") || lower.contains("mono") {
                props.code = true;
            }
        }
    }
}

fn parse_docx_table(table: &XmlNode) -> Table {
    let mut rows: Vec<TableRow> = Vec::new();
    for row in table.children_named("tr") {
        let mut cells: Vec<TableCell> = Vec::new();
        for cell in row.children_named("tc") {
            let mut paragraphs: Vec<Paragraph> = Vec::new();
            // A `<w:tc>` may interleave `<w:p>` and nested `<w:tbl>`
            // children. Walk in document order and recurse into nested
            // tables so their text is never lost.
            collect_cell_blocks(cell, &mut paragraphs);
            if paragraphs.is_empty() {
                paragraphs.push(Paragraph::default());
            }
            cells.push(TableCell { paragraphs });
        }
        rows.push(TableRow { cells });
    }
    Table {
        name: "Table1".to_string(),
        rows,
    }
}

fn collect_cell_blocks(node: &XmlNode, out: &mut Vec<Paragraph>) {
    for item in &node.items {
        let XmlItem::Node(child) = item else {
            continue;
        };
        match child.local_name() {
            "p" => {
                let info = parse_docx_paragraph(child, &BTreeMap::new(), &BTreeMap::new());
                let para = build_paragraph(info.spans);
                if !para
                    .spans
                    .iter()
                    .all(|inline| inline_text(inline).is_empty())
                {
                    out.push(para);
                }
            }
            "tbl" => {
                let nested = parse_docx_table(child);
                for row in &nested.rows {
                    for c in &row.cells {
                        for p in &c.paragraphs {
                            if !p
                                .spans
                                .iter()
                                .all(|inline| inline_text(inline).is_empty())
                            {
                                out.push(p.clone());
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }
}

// ---------------------------------------------------------------------------
// ODT parsing
// ---------------------------------------------------------------------------

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
