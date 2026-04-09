pub mod docx;
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
pub use import::{from_doc_bytes, from_docx_bytes, from_html, from_odt_bytes, from_pdf_bytes, load_bytes};
pub use legacy_doc::extract_text_from_doc;
pub use markdown::to_markdown;
pub use pdf::{to_pdf, to_pdf_with_size};
pub use raster::{render_jpeg_pages, render_pages, render_png_pages};
pub use svg::render_svg;

use std::path::Path;

use lo_core::{
    Block, Heading, Inline, ListBlock, ListItem, LoError, Paragraph, Result, Table, TableCell,
    TableRow, TextDocument,
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

        if trimmed == "---" || trimmed == "***" {
            document.body.push(Block::HorizontalRule);
            index += 1;
            continue;
        }

        if is_list_item(trimmed) {
            let mut list = ListBlock {
                ordered: false,
                items: Vec::new(),
            };
            while index < lines.len() && is_list_item(lines[index].trim()) {
                let item_text = lines[index].trim()[2..].trim();
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
                let cells = split_table_row(current)
                    .into_iter()
                    .map(|cell| TableCell {
                        paragraphs: vec![Paragraph {
                            spans: parse_inlines(cell.trim()),
                            ..Paragraph::default()
                        }],
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
                || is_list_item(current)
                || is_table_row(current)
                || current == "---"
                || current == "***"
            {
                break;
            }
            paragraph_lines.push(current.to_string());
            index += 1;
        }
        document.body.push(Block::Paragraph(Paragraph {
            spans: parse_inlines(&paragraph_lines.join(" ")),
            ..Paragraph::default()
        }));
    }

    document
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

fn is_list_item(line: &str) -> bool {
    line.starts_with("- ") || line.starts_with("* ")
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
    use lo_core::Block;

    #[test]
    fn markdown_headings_and_lists_parse() {
        let doc = from_markdown("Test", "# Title\n\nA **bold** word.\n\n- one\n- two\n");
        assert!(matches!(doc.body[0], Block::Heading(_)));
        assert!(matches!(doc.body[1], Block::Paragraph(_)));
        assert!(matches!(doc.body[2], Block::List(_)));
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
