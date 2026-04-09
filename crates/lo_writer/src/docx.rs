//! Minimal DOCX (WordprocessingML) export.
//!
//! Generates a self-contained `.docx` file with `[Content_Types].xml`,
//! `_rels/.rels`, `word/document.xml`, `word/styles.xml`, and document
//! relationships. The supported model covers paragraphs, headings, bullet/
//! ordered lists, tables, hyperlinks, bold/italic/underline/code spans,
//! horizontal rules and page breaks — i.e. the same blocks `lo_core::writer`
//! already represents.

use lo_core::Result;
use lo_core::{
    escape_attr, escape_text, Block, Heading, Inline, ListBlock, Paragraph, Table, TextDocument,
};
use lo_zip::{ooxml_package, ZipEntry};

const XML_DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

fn render_run(span: &Inline) -> String {
    match span {
        Inline::Text(text) => format!(
            "<w:r><w:t xml:space=\"preserve\">{}</w:t></w:r>",
            escape_text(text)
        ),
        Inline::Bold(text) => format!(
            "<w:r><w:rPr><w:b/></w:rPr><w:t xml:space=\"preserve\">{}</w:t></w:r>",
            escape_text(text)
        ),
        Inline::Italic(text) => format!(
            "<w:r><w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\">{}</w:t></w:r>",
            escape_text(text)
        ),
        Inline::Code(text) => format!(
            "<w:r><w:rPr><w:rFonts w:ascii=\"Courier New\" w:hAnsi=\"Courier New\"/></w:rPr><w:t xml:space=\"preserve\">{}</w:t></w:r>",
            escape_text(text)
        ),
        // We deliberately do NOT emit a `<w:hyperlink>` element here:
        // emitting one with an empty `r:id` is invalid OOXML and causes
        // LibreOffice's importer to hang. Instead, render the link as an
        // underlined blue run followed by a parenthesized URL — readers
        // still get the URL, no validity warnings, no infinite loops.
        Inline::Link { label, url } => format!(
            "<w:r><w:rPr><w:color w:val=\"0000FF\"/><w:u w:val=\"single\"/></w:rPr><w:t xml:space=\"preserve\">{}</w:t></w:r><w:r><w:t xml:space=\"preserve\"> ({})</w:t></w:r>",
            escape_text(label),
            escape_text(url)
        ),
        Inline::LineBreak => "<w:r><w:br/></w:r>".to_string(),
    }
}

fn render_paragraph(p: &Paragraph, style: Option<&str>) -> String {
    let mut runs = String::new();
    for span in &p.spans {
        runs.push_str(&render_run(span));
    }
    if let Some(style_id) = style {
        format!(
            "<w:p><w:pPr><w:pStyle w:val=\"{}\"/></w:pPr>{}</w:p>",
            escape_attr(style_id),
            runs
        )
    } else {
        format!("<w:p>{}</w:p>", runs)
    }
}

fn render_heading(h: &Heading) -> String {
    let level = h.level.clamp(1, 6);
    render_paragraph(&h.content, Some(&format!("Heading{level}")))
}

fn render_list(list: &ListBlock) -> String {
    let mut out = String::new();
    let bullet = if list.ordered { "1." } else { "•" };
    for item in &list.items {
        for nested in &item.blocks {
            if let Block::Paragraph(p) = nested {
                let mut runs = format!("<w:r><w:t xml:space=\"preserve\">{} </w:t></w:r>", bullet);
                for span in &p.spans {
                    runs.push_str(&render_run(span));
                }
                out.push_str(&format!("<w:p>{}</w:p>", runs));
            }
        }
    }
    out
}

fn render_table(table: &Table) -> String {
    let mut out = String::from("<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>");
    for row in &table.rows {
        out.push_str("<w:tr>");
        for cell in &row.cells {
            out.push_str("<w:tc><w:tcPr/>");
            if cell.paragraphs.is_empty() {
                out.push_str("<w:p/>");
            }
            for p in &cell.paragraphs {
                out.push_str(&render_paragraph(p, None));
            }
            out.push_str("</w:tc>");
        }
        out.push_str("</w:tr>");
    }
    out.push_str("</w:tbl>");
    out
}

fn render_block(block: &Block) -> String {
    match block {
        Block::Heading(h) => render_heading(h),
        Block::Paragraph(p) => render_paragraph(p, None),
        Block::List(list) => render_list(list),
        Block::Table(t) => render_table(t),
        Block::Image(_) => render_paragraph(&Paragraph::plain("[image]"), None),
        Block::Section(section) => {
            let mut out = render_paragraph(&Paragraph::plain(&section.name), Some("Heading2"));
            for nested in &section.blocks {
                out.push_str(&render_block(nested));
            }
            out
        }
        Block::PageBreak => "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>".to_string(),
        Block::HorizontalRule => "<w:p><w:pPr><w:pBdr><w:bottom w:val=\"single\" w:sz=\"6\" w:space=\"1\" w:color=\"auto\"/></w:pBdr></w:pPr></w:p>".to_string(),
    }
}

/// Serialize the document into the bytes of a `.docx` file.
pub fn to_docx(document: &TextDocument) -> Result<Vec<u8>> {
    let mut body = String::new();
    for block in &document.body {
        body.push_str(&render_block(block));
    }
    // OOXML requires a paragraph immediately before <w:sectPr>. If the body
    // ended with a table (or with no blocks at all) the importer in Word /
    // LibreOffice can hang trying to anchor the table to a paragraph that
    // doesn't exist. Always emit a trailing empty paragraph to anchor the
    // section properties.
    body.push_str("<w:p/>");
    body.push_str("<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr>");

    let document_xml = format!(
        "{XML_DECL}<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>{body}</w:body></w:document>"
    );

    let styles_xml = format!(
        "{XML_DECL}<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\
<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>\
<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:b/><w:sz w:val=\"36\"/></w:rPr></w:style>\
<w:style w:type=\"paragraph\" w:styleId=\"Heading2\"><w:name w:val=\"heading 2\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:b/><w:sz w:val=\"30\"/></w:rPr></w:style>\
<w:style w:type=\"paragraph\" w:styleId=\"Heading3\"><w:name w:val=\"heading 3\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:b/><w:sz w:val=\"26\"/></w:rPr></w:style>\
<w:style w:type=\"paragraph\" w:styleId=\"Heading4\"><w:name w:val=\"heading 4\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:b/><w:sz w:val=\"24\"/></w:rPr></w:style>\
<w:style w:type=\"paragraph\" w:styleId=\"Heading5\"><w:name w:val=\"heading 5\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:b/><w:sz w:val=\"22\"/></w:rPr></w:style>\
<w:style w:type=\"paragraph\" w:styleId=\"Heading6\"><w:name w:val=\"heading 6\"/><w:basedOn w:val=\"Normal\"/><w:rPr><w:b/><w:sz w:val=\"20\"/></w:rPr></w:style>\
<w:style w:type=\"character\" w:styleId=\"Hyperlink\"><w:name w:val=\"Hyperlink\"/><w:rPr><w:color w:val=\"0000FF\"/><w:u w:val=\"single\"/></w:rPr></w:style>\
</w:styles>"
    );

    let content_types = format!(
        "{XML_DECL}<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
<Default Extension=\"xml\" ContentType=\"application/xml\"/>\
<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\
<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>\
</Types>"
    );

    let rels = format!(
        "{XML_DECL}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\
</Relationships>"
    );

    let doc_rels = format!(
        "{XML_DECL}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\
</Relationships>"
    );

    ooxml_package(&[
        ZipEntry::new("[Content_Types].xml", content_types.into_bytes()),
        ZipEntry::new("_rels/.rels", rels.into_bytes()),
        ZipEntry::new("word/document.xml", document_xml.into_bytes()),
        ZipEntry::new("word/styles.xml", styles_xml.into_bytes()),
        ZipEntry::new("word/_rels/document.xml.rels", doc_rels.into_bytes()),
    ])
}
