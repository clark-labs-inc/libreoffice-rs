use std::path::Path;

use lo_core::{escape_attr, escape_text, Block, Inline, Paragraph, Result, TextDocument};

use crate::common::{content_root_attrs, image_extras, package_document, MIME_ODT};

fn inline_xml(spans: &[Inline]) -> String {
    let mut out = String::new();
    for span in spans {
        match span {
            Inline::Text(text) => out.push_str(&escape_text(text)),
            Inline::Bold(text) => out.push_str(&format!(
                "<text:span text:style-name=\"Strong\">{}</text:span>",
                escape_text(text)
            )),
            Inline::Italic(text) => out.push_str(&format!(
                "<text:span text:style-name=\"Emphasis\">{}</text:span>",
                escape_text(text)
            )),
            Inline::Code(text) => out.push_str(&format!(
                "<text:span text:style-name=\"Code\">{}</text:span>",
                escape_text(text)
            )),
            Inline::Link { label, url } => out.push_str(&format!(
                "<text:a xlink:href=\"{}\">{}</text:a>",
                escape_attr(url),
                escape_text(label)
            )),
            Inline::LineBreak => out.push_str("<text:line-break/>"),
        }
    }
    out
}

fn paragraph_xml(tag: &str, paragraph: &Paragraph, attrs: &str) -> String {
    format!("<{tag}{attrs}>{}</{tag}>", inline_xml(&paragraph.spans))
}

fn block_xml(block: &Block) -> String {
    match block {
        Block::Heading(heading) => {
            let level = heading.level.max(1).min(6);
            paragraph_xml(
                "text:h",
                &heading.content,
                &format!(
                    " text:style-name=\"Heading_20_{level}\" text:outline-level=\"{level}\""
                ),
            )
        }
        Block::Paragraph(paragraph) => paragraph_xml("text:p", paragraph, ""),
        Block::List(list) => {
            let mut out = String::from("<text:list text:style-name=\"L1\">");
            for item in &list.items {
                out.push_str("<text:list-item>");
                for nested in &item.blocks {
                    out.push_str(&block_xml(nested));
                }
                out.push_str("</text:list-item>");
            }
            out.push_str("</text:list>");
            out
        }
        Block::Table(table) => {
            let col_count = table
                .rows
                .iter()
                .map(|row| row.cells.len())
                .max()
                .unwrap_or(0);
            let mut out = format!(
                "<table:table table:name=\"{}\" table:style-name=\"{}\">",
                escape_attr(&table.name),
                escape_attr(&table.name),
            );
            for _ in 0..col_count {
                out.push_str(&format!(
                    "<table:table-column table:style-name=\"{}.A\"/>",
                    escape_attr(&table.name)
                ));
            }
            for row in &table.rows {
                out.push_str("<table:table-row>");
                for cell in &row.cells {
                    out.push_str("<table:table-cell office:value-type=\"string\">");
                    if cell.paragraphs.is_empty() {
                        out.push_str("<text:p/>");
                    } else {
                        for paragraph in &cell.paragraphs {
                            out.push_str(&paragraph_xml("text:p", paragraph, ""));
                        }
                    }
                    out.push_str("</table:table-cell>");
                }
                out.push_str("</table:table-row>");
            }
            out.push_str("</table:table>");
            out
        }
        Block::Image(image) => format!(
            "<text:p><draw:frame svg:width=\"{}\" svg:height=\"{}\"><draw:image xlink:href=\"Pictures/{}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/></draw:frame></text:p>",
            image.size.width.css(),
            image.size.height.css(),
            escape_attr(&image.name)
        ),
        Block::Section(section) => {
            let mut out = format!(
                "<text:section text:name=\"{}\">",
                escape_attr(&section.name)
            );
            for nested in &section.blocks {
                out.push_str(&block_xml(nested));
            }
            out.push_str("</text:section>");
            out
        }
        Block::PageBreak => "<text:p><text:soft-page-break/></text:p>".to_string(),
        Block::HorizontalRule => "<text:p>----------------</text:p>".to_string(),
    }
}

fn collect_tables<'a>(blocks: &'a [Block], out: &mut Vec<&'a lo_core::Table>) {
    for block in blocks {
        match block {
            Block::Table(table) => out.push(table),
            Block::Section(section) => collect_tables(&section.blocks, out),
            Block::List(list) => {
                for item in &list.items {
                    collect_tables(&item.blocks, out);
                }
            }
            _ => {}
        }
    }
}

pub fn serialize_text_document(doc: &TextDocument) -> String {
    let mut xml = lo_core::XmlBuilder::new();
    xml.declaration();
    xml.open("office:document-content", &content_root_attrs());
    xml.empty("office:scripts", &[]);
    xml.open("office:automatic-styles", &[]);
    let mut tables: Vec<&lo_core::Table> = Vec::new();
    collect_tables(&doc.body, &mut tables);
    for table in &tables {
        xml.raw(&format!(
            "<style:style style:name=\"{name}\" style:family=\"table\">\
<style:table-properties style:width=\"160mm\" table:align=\"margins\"/>\
</style:style>\
<style:style style:name=\"{name}.A\" style:family=\"table-column\">\
<style:table-column-properties style:column-width=\"40mm\"/>\
</style:style>",
            name = escape_attr(&table.name)
        ));
    }
    xml.close();
    xml.open("office:body", &[]);
    xml.open("office:text", &[]);
    for block in &doc.body {
        xml.raw(&block_xml(block));
    }
    xml.close();
    xml.close();
    xml.close();
    xml.finish()
}

pub fn save_text_document(path: impl AsRef<Path>, doc: &TextDocument) -> Result<()> {
    let content = serialize_text_document(doc);
    let extras = image_extras(doc.embedded_images());
    package_document(path, MIME_ODT, content, &doc.meta, extras)
}
