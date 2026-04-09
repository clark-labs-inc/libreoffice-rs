//! HTML export for `TextDocument`.

use lo_core::{
    escape_text, html_escape, Block, Heading, Inline, ListBlock, Paragraph, Table, TextDocument,
};

fn render_inlines(spans: &[Inline]) -> String {
    let mut out = String::new();
    for span in spans {
        match span {
            Inline::Text(text) => out.push_str(&html_escape(text)),
            Inline::Bold(text) => {
                out.push_str("<strong>");
                out.push_str(&html_escape(text));
                out.push_str("</strong>");
            }
            Inline::Italic(text) => {
                out.push_str("<em>");
                out.push_str(&html_escape(text));
                out.push_str("</em>");
            }
            Inline::Code(text) => {
                out.push_str("<code>");
                out.push_str(&html_escape(text));
                out.push_str("</code>");
            }
            Inline::Link { label, url } => {
                out.push_str(&format!(
                    r#"<a href="{}">{}</a>"#,
                    html_escape(url),
                    html_escape(label)
                ));
            }
            Inline::LineBreak => out.push_str("<br/>"),
        }
    }
    out
}

fn render_paragraph(p: &Paragraph) -> String {
    render_inlines(&p.spans)
}

fn render_heading(h: &Heading) -> String {
    let level = h.level.clamp(1, 6);
    format!("<h{level}>{}</h{level}>\n", render_paragraph(&h.content))
}

fn render_list(list: &ListBlock) -> String {
    let tag = if list.ordered { "ol" } else { "ul" };
    let mut out = format!("<{tag}>\n");
    for item in &list.items {
        out.push_str("<li>");
        for nested in &item.blocks {
            match nested {
                Block::Paragraph(p) => out.push_str(&render_paragraph(p)),
                Block::List(sub) => out.push_str(&render_list(sub)),
                other => out.push_str(&render_block(other)),
            }
        }
        out.push_str("</li>\n");
    }
    out.push_str(&format!("</{tag}>\n"));
    out
}

fn render_table(table: &Table) -> String {
    let mut out = String::from("<table border=\"1\">\n");
    for row in &table.rows {
        out.push_str("<tr>");
        for cell in &row.cells {
            out.push_str("<td>");
            for p in &cell.paragraphs {
                out.push_str(&render_paragraph(p));
            }
            out.push_str("</td>");
        }
        out.push_str("</tr>\n");
    }
    out.push_str("</table>\n");
    out
}

fn render_block(block: &Block) -> String {
    match block {
        Block::Heading(h) => render_heading(h),
        Block::Paragraph(p) => format!("<p>{}</p>\n", render_paragraph(p)),
        Block::List(list) => render_list(list),
        Block::Table(t) => render_table(t),
        Block::Image(image) => format!(
            "<figure><figcaption>{}</figcaption></figure>\n",
            html_escape(&image.alt)
        ),
        Block::Section(section) => {
            let mut out = format!("<section><h2>{}</h2>\n", html_escape(&section.name));
            for nested in &section.blocks {
                out.push_str(&render_block(nested));
            }
            out.push_str("</section>\n");
            out
        }
        Block::PageBreak => "<div style=\"page-break-after: always\"></div>\n".to_string(),
        Block::HorizontalRule => "<hr/>\n".to_string(),
    }
}

/// Render the document as a complete `<!DOCTYPE html>` page.
pub fn to_html(document: &TextDocument) -> String {
    let mut body = String::new();
    for block in &document.body {
        body.push_str(&render_block(block));
    }
    format!(
        "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"utf-8\"/>\n<title>{}</title>\n</head>\n<body>\n{}</body>\n</html>\n",
        escape_text(&document.meta.title),
        body
    )
}
