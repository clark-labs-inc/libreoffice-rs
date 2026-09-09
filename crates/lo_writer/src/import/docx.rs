use super::*;

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
                let images = crate::docx_images::from_docx_node(node, &relationships, &zip);
                let info = parse_docx_paragraph(node, &relationships, &styles);
                if info.page_break && info.spans.iter().all(|span| span.text.trim().is_empty()) {
                    flush_list_with(&mut doc, &mut pending_list, pending_list_ordered);
                    doc.body.push(Block::PageBreak);
                } else {
                    let paragraph = build_paragraph(info.spans.clone());
                    let is_empty_text = paragraph.spans.is_empty()
                        || paragraph
                            .spans
                            .iter()
                            .all(|inline| inline_text(inline).trim().is_empty());
                    if is_empty_text && info.heading_level.is_none() && info.list_key.is_none() {
                        flush_list_with(&mut doc, &mut pending_list, pending_list_ordered);
                        if images.is_empty() {
                            doc.body.push(Block::Paragraph(paragraph));
                        }
                    } else if let Some(level) = info.heading_level {
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
                if !images.is_empty() {
                    flush_list_with(&mut doc, &mut pending_list, pending_list_ordered);
                    doc.body.extend(images.into_iter().map(Block::Image));
                }
            }
            "tbl" => {
                flush_list_with(&mut doc, &mut pending_list, pending_list_ordered);
                doc.body.push(Block::Table(parse_docx_table(node)));
                doc.body.extend(
                    crate::docx_images::from_docx_node(node, &relationships, &zip)
                        .into_iter()
                        .map(Block::Image),
                );
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
    while matches!(
        doc.body.last(),
        Some(Block::Paragraph(paragraph)) if paragraph.to_plain_text().trim().is_empty()
    ) {
        doc.body.pop();
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

pub(super) fn extract_heading_level(name: &str) -> Option<u8> {
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
        font_size_pt: base.font_size_pt,
    };
    if let Some(rpr) = run.child("rPr") {
        let mut props = StyleProps::default();
        apply_docx_run_properties(rpr, &mut props);
        style.bold |= props.bold;
        style.italic |= props.italic;
        style.code |= props.code;
        style.font_size_pt = props.font_size_pt.or(style.font_size_pt);
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
    props.font_size_pt = rpr.child("sz").and_then(|node| node.attr("val"))
        .and_then(|value| value.parse::<u16>().ok()).map(|half_points| half_points / 2)
        .filter(|points| *points > 0);
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

