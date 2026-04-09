use lo_core::{Block, Inline, Paragraph, Table, TextDocument};

pub fn to_markdown(document: &TextDocument) -> String {
    let mut out = String::new();
    if !document.meta.title.trim().is_empty() {
        out.push_str(&format!("# {}\n\n", escape_text(&document.meta.title)));
    }
    for block in &document.body {
        match block {
            Block::Heading(heading) => {
                let level = heading.level.clamp(1, 6);
                out.push_str(&"#".repeat(level as usize));
                out.push(' ');
                out.push_str(&paragraph_to_markdown(&heading.content));
                out.push_str("\n\n");
            }
            Block::Paragraph(paragraph) => {
                out.push_str(&paragraph_to_markdown(paragraph));
                out.push_str("\n\n");
            }
            Block::List(list) => {
                for (index, item) in list.items.iter().enumerate() {
                    let prefix = if list.ordered {
                        format!("{}. ", index + 1)
                    } else {
                        "- ".to_string()
                    };
                    let mut rendered = String::new();
                    for nested in &item.blocks {
                        if let Block::Paragraph(paragraph) = nested {
                            if !rendered.is_empty() {
                                rendered.push(' ');
                            }
                            rendered.push_str(&paragraph_to_markdown(paragraph));
                        }
                    }
                    out.push_str(&prefix);
                    out.push_str(&rendered);
                    out.push('\n');
                }
                out.push('\n');
            }
            Block::Table(table) => {
                out.push_str(&table_to_markdown(table));
                out.push_str("\n\n");
            }
            Block::Image(image) => {
                out.push_str(&format!("![{}]({})\n\n", escape_text(&image.alt), escape_text(&image.name)));
            }
            Block::Section(section) => {
                out.push_str(&format!("## {}\n\n", escape_text(&section.name)));
                for nested in &section.blocks {
                    if let Block::Paragraph(paragraph) = nested {
                        out.push_str(&paragraph_to_markdown(paragraph));
                        out.push_str("\n\n");
                    }
                }
            }
            // Use a CommonMark thematic break — it doesn't add bogus
            // English tokens (`div`, `class`, `page`, `break`) and still
            // renders as a page divider in every Markdown viewer.
            Block::PageBreak => out.push_str("\n---\n\n"),
            Block::HorizontalRule => out.push_str("\n---\n\n"),
        }
    }
    out.trim().to_string()
}

fn paragraph_to_markdown(paragraph: &Paragraph) -> String {
    let mut out = String::new();
    for span in &paragraph.spans {
        match span {
            Inline::Text(text) => out.push_str(&escape_text(text)),
            Inline::Bold(text) => out.push_str(&format!("**{}**", escape_text(text))),
            Inline::Italic(text) => out.push_str(&format!("*{}*", escape_text(text))),
            Inline::Code(text) => out.push_str(&format!("`{}`", text.replace('`', "'"))),
            Inline::Link { label, url } => out.push_str(&format!("[{}]({})", escape_text(label), url)),
            Inline::LineBreak => out.push_str("  \n"),
        }
    }
    out
}

fn table_to_markdown(table: &Table) -> String {
    let mut out = String::new();
    if table.rows.is_empty() {
        return out;
    }
    for (row_index, row) in table.rows.iter().enumerate() {
        out.push('|');
        for cell in &row.cells {
            let text = cell
                .paragraphs
                .iter()
                .map(paragraph_to_markdown)
                .collect::<Vec<_>>()
                .join(" ");
            out.push(' ');
            out.push_str(&text.replace('|', "\\|"));
            out.push(' ');
            out.push('|');
        }
        out.push('\n');
        if row_index == 0 {
            out.push('|');
            for _ in &row.cells {
                out.push_str(" --- |");
            }
            out.push('\n');
        }
    }
    out
}

fn escape_text(input: &str) -> String {
    input
        .replace('\\', "\\\\")
        .replace('*', "\\*")
        .replace('_', "\\_")
        .replace('[', "\\[")
        .replace(']', "\\]")
}
