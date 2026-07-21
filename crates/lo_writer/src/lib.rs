pub mod docx;
mod font_metrics;
pub mod html;
pub mod import;
pub mod layout;
pub mod legacy_doc;
pub mod markdown;
pub mod pdf;
pub mod raster;
pub mod svg;

pub use docx::to_docx;
pub use html::to_html;
pub use import::{
    from_doc_bytes, from_docx_bytes, from_html, from_odt_bytes, from_pdf_bytes, load_bytes,
};
pub use legacy_doc::extract_text_from_doc;
pub use markdown::to_markdown;
pub use pdf::{to_pdf, to_pdf_with_size};
pub use raster::{render_jpeg_pages, render_pages, render_png_pages};
pub use svg::render_svg;

use std::fs;
use std::path::{Path, PathBuf};

use lo_core::{
    Block, Heading, ImageBlock, Inline, Length, ListBlock, ListItem, LoError, Paragraph, Result,
    Size, Table, TableCell, TableRow, TextDocument, TextStyle,
};

pub struct WriterEditor {
    pub document: TextDocument,
}

impl WriterEditor {
    pub fn new(title: impl Into<String>) -> Self {
        Self {
            document: TextDocument::new(title),
        }
    }

    pub fn push_paragraph(&mut self, text: impl Into<String>) -> &mut Self {
        self.document.push_paragraph(text);
        self
    }

    pub fn push_heading(&mut self, level: u8, text: impl Into<String>) -> &mut Self {
        self.document.push_heading(level, text);
        self
    }

    pub fn save_odt(&self, path: impl AsRef<Path>) -> Result<()> {
        lo_odf::save_text_document(path, &self.document)
    }
}

pub fn save_odt(path: impl AsRef<Path>, document: &TextDocument) -> Result<()> {
    lo_odf::save_text_document(path, document)
}

/// Render the document into bytes for the requested format.
///
/// Supported format strings (case-insensitive): `txt`, `md`, `html`, `svg`, `pdf`,
/// `odt`, `docx`. Multi-page raster output is exposed through
/// [`render_png_pages`] and [`render_jpeg_pages`].
pub fn save_as(document: &TextDocument, format: &str) -> Result<Vec<u8>> {
    match format.to_ascii_lowercase().as_str() {
        "txt" => Ok(document.plain_text().into_bytes()),
        "md" | "markdown" => Ok(to_markdown(document).into_bytes()),
        "html" => Ok(to_html(document).into_bytes()),
        "svg" => {
            let size = lo_core::Size::new(
                lo_core::units::Length::pt(595.0),
                lo_core::units::Length::pt(842.0),
            );
            Ok(render_svg(document, size).into_bytes())
        }
        "pdf" => Ok(to_pdf(document)),
        "odt" => {
            // Round-trip through a temp file using lo_odf::save_text_document.
            let tmp = std::env::temp_dir().join(format!(
                "lo_writer_{}.odt",
                std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|d| d.as_nanos())
                    .unwrap_or(0)
            ));
            lo_odf::save_text_document(&tmp, document)?;
            let bytes = std::fs::read(&tmp)?;
            let _ = std::fs::remove_file(&tmp);
            Ok(bytes)
        }
        "docx" => to_docx(document),
        other => Err(LoError::Unsupported(format!(
            "writer format not supported: {other}"
        ))),
    }
}

pub fn from_plain_text(title: impl Into<String>, input: &str) -> TextDocument {
    let mut document = TextDocument::new(title);
    for paragraph in input.split("\n\n") {
        let trimmed = paragraph.trim();
        if !trimmed.is_empty() {
            document.push_paragraph(trimmed);
        }
    }
    document
}

pub fn to_plain_text(document: &TextDocument) -> String {
    document.plain_text()
}

pub fn from_markdown(title: impl Into<String>, markdown: &str) -> TextDocument {
    parse_markdown(title.into(), markdown, |_| Ok(None))
        .expect("Markdown parsing without external assets is infallible")
}

/// Parse Markdown and resolve image references relative to `base_dir`.
/// Missing, remote, or unsupported image sources are reported instead of
/// silently replacing the visual with a placeholder.
pub fn from_markdown_with_base(
    title: impl Into<String>,
    markdown: &str,
    base_dir: impl AsRef<Path>,
) -> Result<TextDocument> {
    let base = base_dir.as_ref().to_path_buf();
    from_markdown_with_resolver(title, markdown, move |source| {
        if has_uri_scheme(source) {
            return Err(LoError::Unsupported(format!(
                "remote Markdown image requires a custom resolver: {source}"
            )));
        }
        let path = normalize_asset_path(&base, source);
        fs::read(&path)
            .map(Some)
            .map_err(|error| LoError::Io(format!("{}: {error}", path.display())))
    })
}

/// Parse Markdown with caller-provided image loading. This is the reusable
/// boundary for network-backed artifacts: callers may fetch remote URLs while
/// the pure-Rust CLI uses [`from_markdown_with_base`] for local assets.
pub fn from_markdown_with_resolver<F>(
    title: impl Into<String>,
    markdown: &str,
    resolver: F,
) -> Result<TextDocument>
where
    F: FnMut(&str) -> Result<Option<Vec<u8>>>,
{
    parse_markdown(title.into(), markdown, resolver)
}

fn parse_markdown<F>(title: String, markdown: &str, mut resolver: F) -> Result<TextDocument>
where
    F: FnMut(&str) -> Result<Option<Vec<u8>>>,
{
    let mut document = TextDocument::new(title);
    let lines: Vec<&str> = markdown.lines().collect();
    let mut index = 0usize;

    while index < lines.len() {
        let line = lines[index].trim_end();
        let trimmed = line.trim();

        if trimmed.is_empty() {
            index += 1;
            continue;
        }

        if let Some((level, text)) = parse_heading(trimmed) {
            document.body.push(Block::Heading(Heading {
                level,
                content: Paragraph {
                    spans: parse_inlines(text),
                    ..Paragraph::default()
                },
            }));
            index += 1;
            continue;
        }

        if let Some(fence) = parse_fence(trimmed) {
            let language = trimmed[fence.len()..].trim().to_ascii_lowercase();
            let mut code = Vec::new();
            index += 1;
            while index < lines.len() && !lines[index].trim_start().starts_with(fence) {
                code.push(lines[index]);
                index += 1;
            }
            if index < lines.len() {
                index += 1;
            }
            if language == "mermaid" {
                document.body.push(Block::Image(rich_markdown_block(
                    "diagram.mermaid",
                    "application/vnd.mermaid+text",
                    "Mermaid flowchart",
                    code.join("\n"),
                    170.0,
                    90.0,
                )));
                continue;
            }
            if matches!(language.as_str(), "math" | "latex" | "tex") {
                let source = code.join("\n");
                document.body.push(Block::Image(rich_markdown_block(
                    "formula.tex",
                    "application/x-latex",
                    &format!("Formula: {}", accessible_formula_text(&source)),
                    source,
                    170.0,
                    36.0,
                )));
                continue;
            }
            let mut paragraph = Paragraph {
                spans: vec![Inline::Code(code.join("\n"))],
                ..Paragraph::default()
            };
            paragraph.text_style.font_family = "monospace".to_string();
            paragraph.text_style.background = "#f1f0ec".to_string();
            paragraph.style.margin_top_mm = 2;
            paragraph.style.margin_bottom_mm = 3;
            document.body.push(Block::Paragraph(paragraph));
            continue;
        }

        if trimmed.starts_with("$$") {
            let mut formula = Vec::new();
            if trimmed.len() > 4 && trimmed.ends_with("$$") {
                formula.push(trimmed[2..trimmed.len() - 2].trim());
                index += 1;
            } else {
                let first = trimmed.trim_start_matches("$$").trim();
                if !first.is_empty() {
                    formula.push(first);
                }
                index += 1;
                while index < lines.len() && !lines[index].trim().ends_with("$$") {
                    formula.push(lines[index].trim());
                    index += 1;
                }
                if index < lines.len() {
                    let last = lines[index].trim().trim_end_matches("$$").trim();
                    if !last.is_empty() {
                        formula.push(last);
                    }
                    index += 1;
                }
            }
            let source = formula.join(" ");
            document.body.push(Block::Image(rich_markdown_block(
                "formula.tex",
                "application/x-latex",
                &format!("Formula: {}", accessible_formula_text(&source)),
                source,
                170.0,
                36.0,
            )));
            continue;
        }

        if let Some((alt, source)) = parse_image(trimmed) {
            let data = resolver(source)?.unwrap_or_default();
            document
                .body
                .push(Block::Image(markdown_image(source, alt, data)?));
            index += 1;
            continue;
        }

        if trimmed.starts_with('>') {
            let mut quote_lines = Vec::new();
            while index < lines.len() && lines[index].trim_start().starts_with('>') {
                quote_lines.push(
                    lines[index]
                        .trim_start()
                        .trim_start_matches('>')
                        .trim_start(),
                );
                index += 1;
            }
            let mut paragraph = Paragraph {
                spans: parse_inlines(&quote_lines.join("\n")),
                ..Paragraph::default()
            };
            paragraph.style.margin_left_mm = 6;
            paragraph.style.margin_right_mm = 3;
            paragraph.text_style.italic = true;
            paragraph.text_style.color = "#544f49".to_string();
            document.body.push(Block::Paragraph(paragraph));
            continue;
        }

        if trimmed == "---" || trimmed == "***" {
            document.body.push(Block::HorizontalRule);
            index += 1;
            continue;
        }

        if let Some((ordered, _)) = parse_list_item(trimmed) {
            let mut list = ListBlock {
                ordered,
                items: Vec::new(),
            };
            while index < lines.len() {
                let Some((item_ordered, item_text)) = parse_list_item(lines[index].trim()) else {
                    break;
                };
                if item_ordered != ordered {
                    break;
                }
                list.items.push(ListItem {
                    blocks: vec![Block::Paragraph(Paragraph {
                        spans: parse_inlines(item_text),
                        ..Paragraph::default()
                    })],
                });
                index += 1;
            }
            document.body.push(Block::List(list));
            continue;
        }

        if is_table_row(trimmed) {
            let mut rows = Vec::new();
            while index < lines.len() && is_table_row(lines[index].trim()) {
                let current = lines[index].trim();
                if is_table_separator(current) {
                    index += 1;
                    continue;
                }
                let header = rows.is_empty();
                let cells = split_table_row(current)
                    .into_iter()
                    .map(|cell| {
                        let mut paragraph = Paragraph {
                            spans: parse_inlines(cell.trim()),
                            ..Paragraph::default()
                        };
                        paragraph.text_style.bold = header;
                        TableCell {
                            paragraphs: vec![paragraph],
                        }
                    })
                    .collect();
                rows.push(TableRow { cells });
                index += 1;
            }
            document.body.push(Block::Table(Table {
                name: "Table1".to_string(),
                rows,
            }));
            continue;
        }

        let mut paragraph_lines = vec![trimmed.to_string()];
        index += 1;
        while index < lines.len() {
            let current = lines[index].trim();
            if current.is_empty()
                || parse_heading(current).is_some()
                || parse_list_item(current).is_some()
                || is_table_row(current)
                || parse_image(current).is_some()
                || parse_fence(current).is_some()
                || current.starts_with('>')
                || current == "---"
                || current == "***"
            {
                break;
            }
            paragraph_lines.push(current.to_string());
            index += 1;
        }
        document.body.push(Block::Paragraph(Paragraph {
            spans: parse_inlines(&paragraph_lines.join("\n")),
            ..Paragraph::default()
        }));
    }

    Ok(document)
}

fn parse_heading(line: &str) -> Option<(u8, &str)> {
    let hashes = line.chars().take_while(|ch| *ch == '#').count();
    if hashes == 0 || hashes > 6 {
        return None;
    }
    let rest = line[hashes..].trim();
    if rest.is_empty() {
        return None;
    }
    Some((hashes as u8, rest))
}

fn parse_list_item(line: &str) -> Option<(bool, &str)> {
    if let Some(text) = line
        .strip_prefix("- ")
        .or_else(|| line.strip_prefix("* "))
        .or_else(|| line.strip_prefix("+ "))
    {
        return Some((false, text.trim()));
    }
    let digits = line.chars().take_while(|ch| ch.is_ascii_digit()).count();
    if digits > 0 && line.get(digits..digits + 2) == Some(". ") {
        return Some((true, line[digits + 2..].trim()));
    }
    None
}

fn parse_fence(line: &str) -> Option<&str> {
    if line.starts_with("```") {
        Some("```")
    } else if line.starts_with("~~~") {
        Some("~~~")
    } else {
        None
    }
}

fn parse_image(line: &str) -> Option<(&str, &str)> {
    let rest = line.strip_prefix("![")?;
    let label_end = rest.find("](")?;
    if !rest.ends_with(')') {
        return None;
    }
    let source = rest[label_end + 2..rest.len() - 1].trim();
    let source = source.split_whitespace().next()?.trim_matches(['<', '>']);
    Some((&rest[..label_end], source))
}

fn has_uri_scheme(source: &str) -> bool {
    let bytes = source.as_bytes();
    if bytes.len() >= 3
        && bytes[0].is_ascii_alphabetic()
        && bytes[1] == b':'
        && matches!(bytes[2], b'/' | b'\\')
    {
        return false;
    }
    source.find(':').is_some_and(|index| {
        index > 0
            && source[..index]
                .chars()
                .all(|ch| ch.is_ascii_alphanumeric() || matches!(ch, '+' | '-' | '.'))
    })
}

fn normalize_asset_path(base: &Path, source: &str) -> PathBuf {
    let decoded = percent_decode(source.split(['?', '#']).next().unwrap_or(source));
    let path = Path::new(&decoded);
    if path.is_absolute() {
        path.to_path_buf()
    } else {
        base.join(path)
    }
}

fn percent_decode(source: &str) -> String {
    let bytes = source.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'%' && index + 2 < bytes.len() {
            if let (Some(high), Some(low)) = (hex(bytes[index + 1]), hex(bytes[index + 2])) {
                out.push(high * 16 + low);
                index += 3;
                continue;
            }
        }
        out.push(bytes[index]);
        index += 1;
    }
    String::from_utf8_lossy(&out).to_string()
}

fn hex(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn markdown_image(source: &str, alt: &str, data: Vec<u8>) -> Result<ImageBlock> {
    let (mime_type, pixel_size) = if data.starts_with(b"\x89PNG\r\n\x1a\n") {
        let decoded = lo_core::decode_png(&data)?;
        ("image/png", Some((decoded.width, decoded.height)))
    } else if data.starts_with(&[0xff, 0xd8]) {
        ("image/jpeg", jpeg_dimensions(&data))
    } else if data.starts_with(b"RIFF") && data.get(8..12) == Some(b"WEBP") {
        let decoded = lo_core::decode_webp(&data)?;
        ("image/webp", Some((decoded.width, decoded.height)))
    } else if String::from_utf8_lossy(&data)
        .trim_start()
        .starts_with("<svg")
    {
        ("image/svg+xml", svg_dimensions(&data))
    } else if data.is_empty() {
        ("application/octet-stream", None)
    } else {
        return Err(LoError::Unsupported(format!(
            "Markdown image format is not supported: {source}"
        )));
    };
    let size = pixel_size
        .map(markdown_image_size)
        .unwrap_or_else(|| Size::new(Length::mm(120.0), Length::mm(72.0)));
    Ok(ImageBlock {
        name: source.to_string(),
        mime_type: mime_type.to_string(),
        data,
        alt: alt.to_string(),
        size,
    })
}

fn svg_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    let text = std::str::from_utf8(bytes).ok()?;
    let root = lo_core::parse_xml_document(text).ok()?;
    let parse_dimension = |name: &str| {
        root.attr(name).and_then(|value| {
            value
                .trim_end_matches(|ch: char| ch.is_ascii_alphabetic() || ch == '%')
                .parse::<f32>()
                .ok()
        })
    };
    if let (Some(width), Some(height)) = (parse_dimension("width"), parse_dimension("height")) {
        return Some((
            width.max(1.0).round() as u32,
            height.max(1.0).round() as u32,
        ));
    }
    let view_box = root.attr("viewBox").or_else(|| root.attr("viewbox"))?;
    let values = view_box
        .split(|ch: char| ch.is_whitespace() || ch == ',')
        .filter_map(|value| value.parse::<f32>().ok())
        .collect::<Vec<_>>();
    (values.len() == 4).then(|| {
        (
            values[2].max(1.0).round() as u32,
            values[3].max(1.0).round() as u32,
        )
    })
}

fn markdown_image_size((width, height): (u32, u32)) -> Size {
    let natural_width = width as f32 * 25.4 / 96.0;
    let display_width = natural_width.clamp(25.0, 170.0);
    let display_height = (display_width * height as f32 / width.max(1) as f32).min(220.0);
    Size::new(Length::mm(display_width), Length::mm(display_height))
}

fn jpeg_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    let mut index = 2usize;
    while index + 4 <= bytes.len() {
        if bytes[index] != 0xff {
            index += 1;
            continue;
        }
        let marker = bytes[index + 1];
        index += 2;
        if matches!(marker, 0xd8 | 0xd9) || (0xd0..=0xd7).contains(&marker) {
            continue;
        }
        let len = u16::from_be_bytes(bytes.get(index..index + 2)?.try_into().ok()?) as usize;
        if len < 2 || index + len > bytes.len() {
            return None;
        }
        if matches!(
            marker,
            0xc0 | 0xc1
                | 0xc2
                | 0xc3
                | 0xc5
                | 0xc6
                | 0xc7
                | 0xc9
                | 0xca
                | 0xcb
                | 0xcd
                | 0xce
                | 0xcf
        ) {
            let height =
                u16::from_be_bytes(bytes.get(index + 3..index + 5)?.try_into().ok()?) as u32;
            let width =
                u16::from_be_bytes(bytes.get(index + 5..index + 7)?.try_into().ok()?) as u32;
            return Some((width, height));
        }
        index += len;
    }
    None
}

fn is_table_row(line: &str) -> bool {
    line.contains('|') && line.trim_matches('|').contains('|')
}

fn is_table_separator(line: &str) -> bool {
    line.chars()
        .all(|ch| ch == '|' || ch == '-' || ch == ':' || ch.is_whitespace())
}

fn split_table_row(line: &str) -> Vec<&str> {
    line.trim_matches('|').split('|').collect()
}

fn parse_inlines(input: &str) -> Vec<Inline> {
    let chars: Vec<char> = input.chars().collect();
    let mut index = 0usize;
    let mut spans = Vec::new();
    let mut text_buffer = String::new();

    while index < chars.len() {
        if chars[index] == '\n' {
            flush_text(&mut spans, &mut text_buffer);
            spans.push(Inline::LineBreak);
            index += 1;
            continue;
        }
        if chars[index] == '<' {
            let remaining = chars[index..].iter().collect::<String>();
            if let Some((consumed, inline)) = parse_inline_html(&remaining) {
                flush_text(&mut spans, &mut text_buffer);
                spans.push(inline);
                index += consumed.chars().count();
                continue;
            }
        }
        if chars[index] == '$' && chars.get(index + 1) != Some(&'$') {
            if let Some(end) = find_single_marker(&chars, index + 1, '$') {
                flush_text(&mut spans, &mut text_buffer);
                let source = chars[index + 1..end].iter().collect::<String>();
                let mut style = TextStyle::default();
                style.font_family = "Times".to_string();
                style.italic = true;
                style.color = "#3d315f".to_string();
                spans.push(Inline::Styled {
                    text: flatten_inline_math(&source),
                    style,
                    url: None,
                });
                index = end + 1;
                continue;
            }
        }
        if index + 1 < chars.len() && chars[index] == '*' && chars[index + 1] == '*' {
            if let Some(end) = find_double_marker(&chars, index + 2, '*') {
                flush_text(&mut spans, &mut text_buffer);
                let content: String = chars[index + 2..end].iter().collect();
                spans.push(Inline::Bold(content));
                index = end + 2;
                continue;
            }
        }

        if chars[index] == '*' {
            if let Some(end) = find_single_marker(&chars, index + 1, '*') {
                flush_text(&mut spans, &mut text_buffer);
                let content: String = chars[index + 1..end].iter().collect();
                spans.push(Inline::Italic(content));
                index = end + 1;
                continue;
            }
        }

        if chars[index] == '`' {
            if let Some(end) = find_single_marker(&chars, index + 1, '`') {
                flush_text(&mut spans, &mut text_buffer);
                let content: String = chars[index + 1..end].iter().collect();
                spans.push(Inline::Code(content));
                index = end + 1;
                continue;
            }
        }

        if chars[index] == '[' {
            if let Some(label_end) = find_single_marker(&chars, index + 1, ']') {
                if chars.get(label_end + 1) == Some(&'(') {
                    if let Some(url_end) = find_single_marker(&chars, label_end + 2, ')') {
                        flush_text(&mut spans, &mut text_buffer);
                        let label: String = chars[index + 1..label_end].iter().collect();
                        let url: String = chars[label_end + 2..url_end].iter().collect();
                        spans.push(Inline::Link { label, url });
                        index = url_end + 1;
                        continue;
                    }
                }
            }
        }

        text_buffer.push(chars[index]);
        index += 1;
    }

    flush_text(&mut spans, &mut text_buffer);
    spans
}

fn parse_inline_html(input: &str) -> Option<(&str, Inline)> {
    let tag_end = input.find('>')?;
    let opening = &input[1..tag_end];
    let opening_trimmed = opening.trim();
    if matches!(opening_trimmed.to_ascii_lowercase().as_str(), "br" | "br/") {
        return Some((&input[..=tag_end], Inline::LineBreak));
    }
    if opening_trimmed.starts_with('/') {
        return None;
    }
    let name_end = opening_trimmed
        .find(char::is_whitespace)
        .unwrap_or(opening_trimmed.len());
    let name = opening_trimmed[..name_end].to_ascii_lowercase();
    if !matches!(
        name.as_str(),
        "span" | "strong" | "b" | "em" | "i" | "u" | "mark" | "code" | "a" | "div"
    ) {
        return None;
    }
    let closing = format!("</{name}>");
    let lower = input.to_ascii_lowercase();
    let close_start = lower[tag_end + 1..].find(&closing)? + tag_end + 1;
    let consumed_end = close_start + closing.len();
    let content = strip_html_tags(&input[tag_end + 1..close_start]);
    let mut style = TextStyle::default();
    match name.as_str() {
        "strong" | "b" => style.bold = true,
        "em" | "i" => style.italic = true,
        "u" => style.underline = true,
        "mark" => style.background = "#fff1a8".to_string(),
        "code" => {
            style.font_family = "monospace".to_string();
            style.background = "#f1f0ec".to_string();
        }
        _ => {}
    }
    if let Some(value) = html_attribute(opening_trimmed, "style") {
        apply_inline_css(&mut style, value);
    }
    let url = if name == "a" {
        html_attribute(opening_trimmed, "href").map(str::to_string)
    } else {
        None
    };
    if url.is_some() {
        style.underline = true;
        if style.color.is_empty() {
            style.color = "#1e56b3".to_string();
        }
    }
    Some((
        &input[..consumed_end],
        Inline::Styled {
            text: lo_core::decode_entities(&content),
            style,
            url,
        },
    ))
}

fn html_attribute<'a>(tag: &'a str, name: &str) -> Option<&'a str> {
    let lower = tag.to_ascii_lowercase();
    let needle = format!("{name}=");
    let start = lower.find(&needle)? + needle.len();
    let quote = tag.as_bytes().get(start).copied()?;
    if matches!(quote, b'\'' | b'\"') {
        let end = tag[start + 1..].find(quote as char)? + start + 1;
        Some(&tag[start + 1..end])
    } else {
        let end = tag[start..]
            .find(char::is_whitespace)
            .map(|offset| start + offset)
            .unwrap_or(tag.len());
        Some(&tag[start..end])
    }
}

fn apply_inline_css(style: &mut TextStyle, css: &str) {
    for declaration in css.split(';') {
        let Some((name, value)) = declaration.split_once(':') else {
            continue;
        };
        let name = name.trim().to_ascii_lowercase();
        let value = value.trim();
        match name.as_str() {
            "color" => style.color = normalize_css_color(value),
            "background" | "background-color" => style.background = normalize_css_color(value),
            "font-weight" if value.eq_ignore_ascii_case("bold") || value == "700" => {
                style.bold = true
            }
            "font-style" if value.eq_ignore_ascii_case("italic") => style.italic = true,
            "text-decoration" if value.to_ascii_lowercase().contains("underline") => {
                style.underline = true
            }
            "font-size" => {
                let numeric = value
                    .trim_end_matches("pt")
                    .trim_end_matches("px")
                    .parse::<f32>()
                    .ok();
                if let Some(size) = numeric {
                    style.font_size_pt = size.clamp(6.0, 72.0).round() as u16;
                }
            }
            "font-family" => style.font_family = value.trim_matches(['\'', '\"']).to_string(),
            _ => {}
        }
    }
}

fn normalize_css_color(value: &str) -> String {
    let trimmed = value.trim();
    if trimmed.starts_with('#') {
        return trimmed.to_string();
    }
    if let Some(inner) = trimmed
        .strip_prefix("rgb(")
        .and_then(|value| value.strip_suffix(')'))
    {
        let channels = inner
            .split(',')
            .filter_map(|part| part.trim().parse::<u8>().ok())
            .collect::<Vec<_>>();
        if channels.len() == 3 {
            return format!("#{:02x}{:02x}{:02x}", channels[0], channels[1], channels[2]);
        }
    }
    match trimmed.to_ascii_lowercase().as_str() {
        "red" => "#d32f2f",
        "green" => "#2e7d32",
        "blue" => "#1565c0",
        "purple" => "#6d4cc7",
        "orange" => "#ef8c32",
        "black" => "#000000",
        "white" => "#ffffff",
        "gray" | "grey" => "#777777",
        _ => trimmed,
    }
    .to_string()
}

fn strip_html_tags(input: &str) -> String {
    let mut out = String::new();
    let mut in_tag = false;
    for ch in input.chars() {
        match ch {
            '<' => in_tag = true,
            '>' => in_tag = false,
            _ if !in_tag => out.push(ch),
            _ => {}
        }
    }
    out
}

fn flatten_inline_math(source: &str) -> String {
    source
        .replace("\\times", "x")
        .replace("\\cdot", "·")
        .replace("\\pi", "pi")
        .replace("\\alpha", "alpha")
        .replace("\\beta", "beta")
        .replace("^2", "²")
        .replace("^3", "³")
        .replace(['{', '}'], "")
}

fn accessible_formula_text(input: &str) -> String {
    input
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
        .replace("\\frac{", "(")
        .replace("}{", ") / (")
        .replace('}', ")")
        .replace('{', "")
        .replace("^2", " squared")
        .replace("^3", " cubed")
}

fn rich_markdown_block(
    name: &str,
    mime_type: &str,
    alt: &str,
    source: String,
    width_mm: f32,
    height_mm: f32,
) -> ImageBlock {
    ImageBlock {
        name: name.to_string(),
        mime_type: mime_type.to_string(),
        data: source.into_bytes(),
        alt: alt.to_string(),
        size: Size::new(Length::mm(width_mm), Length::mm(height_mm)),
    }
}

fn flush_text(spans: &mut Vec<Inline>, text_buffer: &mut String) {
    if !text_buffer.is_empty() {
        spans.push(Inline::Text(std::mem::take(text_buffer)));
    }
}

fn find_single_marker(chars: &[char], start: usize, marker: char) -> Option<usize> {
    (start..chars.len()).find(|&index| chars[index] == marker)
}

fn find_double_marker(chars: &[char], start: usize, marker: char) -> Option<usize> {
    (start..chars.len().saturating_sub(1))
        .find(|&index| chars[index] == marker && chars[index + 1] == marker)
}

#[cfg(test)]
mod tests {
    use super::{from_markdown, to_plain_text};
    use lo_core::{Block, Inline};

    #[test]
    fn markdown_headings_and_lists_parse() {
        let doc = from_markdown("Test", "# Title\n\nA **bold** word.\n\n- one\n- two\n");
        assert!(matches!(doc.body[0], Block::Heading(_)));
        assert!(matches!(doc.body[1], Block::Paragraph(_)));
        assert!(matches!(doc.body[2], Block::List(_)));
    }

    #[test]
    fn windows_drive_image_paths_are_local_assets() {
        assert!(!super::has_uri_scheme(r"C:\reports\chart.png"));
        assert!(super::has_uri_scheme("https://example.com/chart.png"));
    }

    #[test]
    fn rich_markdown_parses_html_css_math_and_mermaid() {
        let doc = from_markdown(
            "Rich",
            "<span style=\"color:#6d5bd0;background-color:#fff1a8;font-weight:bold\">styled</span> and $x^2$\n\n```mermaid\nflowchart LR\nA[Markdown] --> B{PDF}\n```\n\n$$E = \\frac{mc^2}{1 + alpha}$$",
        );
        let Block::Paragraph(paragraph) = &doc.body[0] else {
            panic!("expected paragraph");
        };
        assert!(paragraph.spans.iter().any(|span| matches!(
            span,
            Inline::Styled { text, style, .. }
                if text == "styled" && style.color == "#6d5bd0" && style.bold
        )));
        assert!(paragraph
            .spans
            .iter()
            .any(|span| matches!(span, Inline::Styled { text, .. } if text == "x²")));
        assert!(doc.body.iter().any(|block| matches!(
            block,
            Block::Image(image) if image.mime_type == "application/vnd.mermaid+text"
        )));
        assert!(doc.body.iter().any(|block| matches!(
            block,
            Block::Image(image) if image.mime_type == "application/x-latex"
        )));
    }

    #[test]
    fn plain_text_export_contains_content() {
        let doc = from_markdown("Test", "# Title\n\nhello [site](https://example.com)");
        let text = to_plain_text(&doc);
        assert!(text.contains("Title"));
        assert!(text.contains("hello site"));
    }

    #[test]
    fn html_export_includes_strong_and_links() {
        let doc = from_markdown(
            "HTML Test",
            "# Hi\n\nA **bold** [link](https://example.com).",
        );
        let html = super::to_html(&doc);
        assert!(html.contains("<title>HTML Test</title>"));
        assert!(html.contains("<h1>Hi</h1>"));
        assert!(html.contains("<strong>bold</strong>"));
        assert!(html.contains("href=\"https://example.com\""));
    }

    #[test]
    fn pdf_export_starts_with_pdf_header() {
        let doc = from_markdown("PDF", "Hello PDF");
        let pdf = super::to_pdf(&doc);
        assert!(pdf.starts_with(b"%PDF-1.4"));
        assert!(pdf.ends_with(b"%%EOF\n"));
    }

    #[test]
    fn docx_export_is_a_zip_archive() {
        let doc = from_markdown("DOCX", "# Title\n\nA paragraph.");
        let bytes = super::to_docx(&doc).expect("docx");
        assert!(bytes.starts_with(b"PK"));
    }

    #[test]
    fn save_as_dispatches_by_format() {
        let doc = from_markdown("Demo", "# Hi\n\nbody");
        for fmt in ["txt", "html", "svg", "pdf", "odt", "docx"] {
            let bytes = super::save_as(&doc, fmt).unwrap_or_else(|e| panic!("{fmt}: {e}"));
            assert!(!bytes.is_empty(), "{fmt} produced empty output");
        }
        assert!(super::save_as(&doc, "xyz").is_err());
    }
}
