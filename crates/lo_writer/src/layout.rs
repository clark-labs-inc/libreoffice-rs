//! Block-aware PDF layout for `TextDocument`.
//!
//! This renderer stays intentionally simple, but it is now visual enough
//! for Clark's document QA path: it respects page size/margins,
//! paragraph alignment and spacing, heading sizing, basic run styling,
//! tables with borders, and image placeholders with their real document
//! footprint.

use std::collections::BTreeSet;
use std::fs;
use std::path::PathBuf;

use lo_core::{
    Alignment, Block, Heading, Inline, ListBlock, ListItem, PageStyle, Paragraph, PdfDocument,
    PdfFont, Table, TableCell, TextDocument,
};

#[derive(Clone, Debug)]
struct StyledRun {
    text: String,
    font: PdfFont,
    size: f32,
    color: (f32, f32, f32),
}

#[derive(Clone, Debug, Default)]
struct LineLayout {
    runs: Vec<StyledRun>,
    width: f32,
}

#[derive(Clone, Debug)]
struct LayoutContext {
    page_w: f32,
    page_h: f32,
    margin_l: f32,
    margin_r: f32,
    margin_t: f32,
    margin_b: f32,
    fonts: FontResolver,
}

pub fn render_document_pdf(doc: &TextDocument) -> Vec<u8> {
    let ctx = LayoutContext::from_page_style(&doc.page_style);
    let mut pdf = PdfDocument::new();
    let mut page_index = pdf.add_page(ctx.page_w, ctx.page_h);
    let mut y = ctx.page_h - ctx.margin_t;

    if !doc.meta.title.trim().is_empty() {
        let title_run = StyledRun {
            text: doc.meta.title.clone(),
            font: PdfFont::HelveticaBold,
            size: 22.0,
            color: (0.0, 0.0, 0.0),
        };
        y = ensure_room(&ctx, &mut pdf, &mut page_index, y, 30.0);
        render_line(
            pdf.page_mut(page_index).expect("page"),
            &LineLayout {
                runs: vec![title_run],
                width: measure_text(&doc.meta.title, 22.0),
            },
            ctx.margin_l,
            y,
            Alignment::Start,
            ctx.page_w - ctx.margin_l - ctx.margin_r,
            false,
        );
        y -= 30.0;
    }

    for block in &doc.body {
        match block {
            Block::Heading(heading) => {
                y = render_heading(&ctx, &mut pdf, &mut page_index, y, heading);
            }
            Block::Paragraph(paragraph) => {
                y = render_paragraph(&ctx, &mut pdf, &mut page_index, y, paragraph, 12.0, 0.0);
            }
            Block::List(list) => {
                y = render_list(&ctx, &mut pdf, &mut page_index, y, list);
            }
            Block::Table(table) => {
                y = render_table(&ctx, &mut pdf, &mut page_index, y, table);
            }
            Block::Image(image) => {
                let width = image.size.width.as_pt().max(80.0);
                let height = image.size.height.as_pt().max(48.0);
                y = ensure_room(&ctx, &mut pdf, &mut page_index, y, height + 18.0);
                let bottom = y - height;
                let page = pdf.page_mut(page_index).expect("page");
                page.rect_fill_stroke_rgb(
                    ctx.margin_l,
                    bottom,
                    width.min(ctx.page_w - ctx.margin_l - ctx.margin_r),
                    height,
                    (0.98, 0.98, 0.98),
                    (0.60, 0.60, 0.60),
                );
                page.line_rgb(ctx.margin_l, bottom, ctx.margin_l + width, y, 0.70, 0.70, 0.70);
                page.line_rgb(ctx.margin_l, y, ctx.margin_l + width, bottom, 0.70, 0.70, 0.70);
                page.text_rgb(
                    ctx.margin_l + 6.0,
                    bottom + 10.0,
                    10.0,
                    PdfFont::HelveticaOblique,
                    &format!("[image: {}]", image.alt),
                    0.30,
                    0.30,
                    0.30,
                );
                y = bottom - 12.0;
            }
            Block::Section(section) => {
                y = ensure_room(&ctx, &mut pdf, &mut page_index, y, 20.0);
                pdf.page_mut(page_index)
                    .expect("page")
                    .text_rgb(
                        ctx.margin_l,
                        y,
                        13.0,
                        PdfFont::HelveticaBold,
                        &format!("[{}]", section.name),
                        0.15,
                        0.15,
                        0.15,
                    );
                y -= 18.0;
                for nested in &section.blocks {
                    if let Block::Paragraph(paragraph) = nested {
                        y = render_paragraph(&ctx, &mut pdf, &mut page_index, y, paragraph, 12.0, 0.0);
                    }
                }
            }
            Block::HorizontalRule => {
                y = ensure_room(&ctx, &mut pdf, &mut page_index, y, 14.0);
                pdf.page_mut(page_index)
                    .expect("page")
                    .line_rgb(
                        ctx.margin_l,
                        y,
                        ctx.page_w - ctx.margin_r,
                        y,
                        0.55,
                        0.55,
                        0.55,
                    );
                y -= 10.0;
            }
            Block::PageBreak => {
                page_index = pdf.add_page(ctx.page_w, ctx.page_h);
                y = ctx.page_h - ctx.margin_t;
            }
        }
    }

    pdf.finish()
}

fn render_heading(
    ctx: &LayoutContext,
    pdf: &mut PdfDocument,
    page_index: &mut usize,
    mut y: f32,
    heading: &Heading,
) -> f32 {
    let size: f32 = match heading.level {
        1 => 20.0,
        2 => 18.0,
        3 => 16.0,
        4 => 14.0,
        _ => 13.0,
    };
    let mut paragraph = heading.content.clone();
    paragraph.text_style.font_size_pt = size.round() as u16;
    paragraph.text_style.bold = true;
    paragraph.style.margin_top_mm = paragraph.style.margin_top_mm.max(2);
    paragraph.style.margin_bottom_mm = paragraph.style.margin_bottom_mm.max(2);
    y = render_paragraph(ctx, pdf, page_index, y, &paragraph, size, 0.0);
    y - 2.0
}

fn render_list(
    ctx: &LayoutContext,
    pdf: &mut PdfDocument,
    page_index: &mut usize,
    mut y: f32,
    list: &ListBlock,
) -> f32 {
    for (index, item) in list.items.iter().enumerate() {
        let marker = if list.ordered {
            format!("{}.", index + 1)
        } else {
            "•".to_string()
        };
        y = ensure_room(ctx, pdf, page_index, y, 16.0);
        pdf.page_mut(*page_index)
            .expect("page")
            .text(ctx.margin_l + 2.0, y, 12.0, PdfFont::Helvetica, &marker);
        y = render_list_item(ctx, pdf, page_index, y, item);
        y -= 2.0;
    }
    y - 4.0
}

fn render_list_item(
    ctx: &LayoutContext,
    pdf: &mut PdfDocument,
    page_index: &mut usize,
    mut y: f32,
    item: &ListItem,
) -> f32 {
    for block in &item.blocks {
        if let Block::Paragraph(paragraph) = block {
            y = render_paragraph(ctx, pdf, page_index, y, paragraph, 12.0, 18.0);
        }
    }
    y
}

fn render_paragraph(
    ctx: &LayoutContext,
    pdf: &mut PdfDocument,
    page_index: &mut usize,
    mut y: f32,
    paragraph: &Paragraph,
    default_size: f32,
    extra_indent: f32,
) -> f32 {
    let left = ctx.margin_l + mm_to_pt(paragraph.style.margin_left_mm) + extra_indent;
    let right = ctx.margin_r + mm_to_pt(paragraph.style.margin_right_mm);
    let available_width = (ctx.page_w - left - right).max(48.0);
    let margin_top = mm_to_pt(paragraph.style.margin_top_mm);
    let margin_bottom = mm_to_pt(paragraph.style.margin_bottom_mm.max(1));
    y -= margin_top;

    let runs = paragraph_runs(paragraph, default_size, &ctx.fonts);
    let lines = layout_runs(&runs, available_width);
    for (line_index, line) in lines.iter().enumerate() {
        let line_height = line
            .runs
            .iter()
            .map(|run| run.size * 1.25)
            .fold(default_size * 1.25, f32::max);
        y = ensure_room(ctx, pdf, page_index, y, line_height + 2.0);
        let justify = matches!(paragraph.style.alignment, Alignment::Justify)
            && line_index + 1 < lines.len();
        render_line(
            pdf.page_mut(*page_index).expect("page"),
            line,
            left,
            y,
            paragraph.style.alignment.clone(),
            available_width,
            justify,
        );
        y -= line_height;
    }

    y - margin_bottom
}

fn render_table(
    ctx: &LayoutContext,
    pdf: &mut PdfDocument,
    page_index: &mut usize,
    mut y: f32,
    table: &Table,
) -> f32 {
    let cols = table
        .rows
        .iter()
        .map(|row| row.cells.len())
        .max()
        .unwrap_or(1)
        .max(1);
    let available_width = ctx.page_w - ctx.margin_l - ctx.margin_r;
    let mut col_widths = vec![available_width / cols as f32; cols];
    for row in &table.rows {
        for (index, cell) in row.cells.iter().enumerate() {
            let text = cell_plain_text(cell);
            let suggested = measure_text(&text, 10.0) + 12.0;
            col_widths[index] = col_widths[index].max(suggested.min(available_width * 0.55));
        }
    }
    let total: f32 = col_widths.iter().sum();
    if total > available_width {
        let scale = available_width / total;
        for width in &mut col_widths {
            *width *= scale;
        }
    }

    for row in &table.rows {
        let mut cell_lines = Vec::new();
        let mut row_height: f32 = 18.0;
        for (index, cell) in row.cells.iter().enumerate() {
            let lines = layout_cell(cell, col_widths[index] - 10.0, &ctx.fonts);
            let cell_height: f32 = lines
                .iter()
                .map(|line| {
                    line.runs
                        .iter()
                        .map(|run| run.size * 1.2)
                        .fold(12.0, f32::max)
                })
                .sum::<f32>()
                + 8.0;
            row_height = row_height.max(cell_height);
            cell_lines.push(lines);
        }
        y = ensure_room(ctx, pdf, page_index, y, row_height + 2.0);
        let top = y;
        let bottom = y - row_height;
        let page = pdf.page_mut(*page_index).expect("page");
        let mut x = ctx.margin_l;
        for (index, width) in col_widths.iter().enumerate() {
            page.rect_stroke_rgb(x, bottom, *width, row_height, 0.50, 0.50, 0.50);
            let mut line_y = top - 12.0;
            let Some(lines_for_cell) = cell_lines.get(index) else {
                x += *width;
                continue;
            };
            for line in lines_for_cell {
                render_line(
                    page,
                    line,
                    x + 5.0,
                    line_y,
                    Alignment::Start,
                    *width - 10.0,
                    false,
                );
                line_y -= line
                    .runs
                    .iter()
                    .map(|run| run.size * 1.2)
                    .fold(12.0, f32::max);
            }
            x += *width;
        }
        y = bottom - 6.0;
    }

    y - 2.0
}

fn layout_cell(cell: &TableCell, width: f32, fonts: &FontResolver) -> Vec<LineLayout> {
    let mut out = Vec::new();
    for paragraph in &cell.paragraphs {
        let runs = paragraph_runs(paragraph, 10.0, fonts);
        let lines = layout_runs(&runs, width);
        out.extend(lines);
    }
    if out.is_empty() {
        out.push(LineLayout::default());
    }
    out
}

fn render_line(
    page: &mut lo_core::PdfPage,
    line: &LineLayout,
    x: f32,
    y: f32,
    alignment: Alignment,
    available_width: f32,
    justify: bool,
) {
    let base_x = match alignment {
        Alignment::Center => x + ((available_width - line.width).max(0.0) / 2.0),
        Alignment::End => x + (available_width - line.width).max(0.0),
        _ => x,
    };
    let mut cursor = base_x;
    let space_slots = if justify {
        line.runs
            .iter()
            .map(|run| run.text.matches(' ').count())
            .sum::<usize>()
    } else {
        0
    };
    let extra_per_space = if space_slots > 0 {
        (available_width - line.width).max(0.0) / space_slots as f32
    } else {
        0.0
    };
    // Merge consecutive same-style runs (including the space tokens
    // produced by `tokenize_run`) into a single text-show operator. If we
    // emit each space as its own Tj, `pdftotext -raw` cannot tell that the
    // adjacent words are separated and ends up gluing them together
    // (e.g. "PHPWord search" -> "PHPWordsearch").
    let mut idx = 0usize;
    while idx < line.runs.len() {
        let start = idx;
        let head = &line.runs[idx];
        let mut combined = String::new();
        while idx < line.runs.len() {
            let cur = &line.runs[idx];
            if cur.font != head.font || cur.size != head.size || cur.color != head.color {
                break;
            }
            combined.push_str(&cur.text);
            idx += 1;
        }
        let group_width: f32 = line.runs[start..idx]
            .iter()
            .map(|r| measure_text(&r.text, r.size))
            .sum();
        if !combined.is_empty() {
            page.text_rgb(
                cursor,
                y,
                head.size,
                head.font,
                &combined,
                head.color.0,
                head.color.1,
                head.color.2,
            );
        }
        cursor += group_width;
        if extra_per_space > 0.0 {
            cursor += combined.matches(' ').count() as f32 * extra_per_space;
        }
    }
}

fn paragraph_runs(paragraph: &Paragraph, default_size: f32, fonts: &FontResolver) -> Vec<StyledRun> {
    let family = if paragraph.text_style.font_family.trim().is_empty() {
        None
    } else {
        Some(paragraph.text_style.font_family.as_str())
    };
    let size = if paragraph.text_style.font_size_pt == 0 {
        default_size
    } else {
        paragraph.text_style.font_size_pt as f32
    };
    let base_color = parse_color(&paragraph.text_style.color).unwrap_or((0.0, 0.0, 0.0));
    let mut runs = Vec::new();
    for inline in &paragraph.spans {
        match inline {
            Inline::Text(text) => runs.push(StyledRun {
                text: text.clone(),
                font: fonts.pick(family, paragraph.text_style.bold, paragraph.text_style.italic),
                size,
                color: base_color,
            }),
            Inline::Bold(text) => runs.push(StyledRun {
                text: text.clone(),
                font: fonts.pick(family, true, paragraph.text_style.italic),
                size,
                color: base_color,
            }),
            Inline::Italic(text) => runs.push(StyledRun {
                text: text.clone(),
                font: fonts.pick(family, paragraph.text_style.bold, true),
                size,
                color: base_color,
            }),
            Inline::Code(text) => runs.push(StyledRun {
                text: text.clone(),
                font: PdfFont::Courier,
                size: (size - 1.0).max(9.0),
                color: (0.10, 0.10, 0.10),
            }),
            Inline::Link { label, .. } => runs.push(StyledRun {
                text: label.clone(),
                font: fonts.pick(family, false, false),
                size,
                color: (0.10, 0.25, 0.65),
            }),
            Inline::LineBreak => runs.push(StyledRun {
                text: "\n".to_string(),
                font: fonts.pick(family, false, false),
                size,
                color: base_color,
            }),
        }
    }
    if runs.is_empty() {
        runs.push(StyledRun {
            text: String::new(),
            font: fonts.pick(family, false, false),
            size,
            color: base_color,
        });
    }
    runs
}

fn layout_runs(runs: &[StyledRun], max_width: f32) -> Vec<LineLayout> {
    let mut lines = Vec::new();
    let mut current = LineLayout::default();
    for run in runs {
        for token in tokenize_run(run) {
            if token.text == "\n" {
                lines.push(std::mem::take(&mut current));
                continue;
            }
            let token_width = measure_text(&token.text, token.size);
            let is_space = token.text.chars().all(|ch| ch.is_whitespace());
            if !current.runs.is_empty() && current.width + token_width > max_width && !is_space {
                trim_trailing_spaces(&mut current);
                lines.push(std::mem::take(&mut current));
            }
            if current.runs.is_empty() && is_space {
                continue;
            }
            current.width += token_width;
            current.runs.push(token);
        }
    }
    trim_trailing_spaces(&mut current);
    if !current.runs.is_empty() || lines.is_empty() {
        lines.push(current);
    }
    lines
}

fn tokenize_run(run: &StyledRun) -> Vec<StyledRun> {
    let mut out = Vec::new();
    let mut buffer = String::new();
    let flush = |out: &mut Vec<StyledRun>, buffer: &mut String, template: &StyledRun| {
        if !buffer.is_empty() {
            out.push(StyledRun {
                text: std::mem::take(buffer),
                font: template.font,
                size: template.size,
                color: template.color,
            });
        }
    };
    for ch in run.text.chars() {
        if ch == '\n' {
            flush(&mut out, &mut buffer, run);
            out.push(StyledRun {
                text: "\n".to_string(),
                font: run.font,
                size: run.size,
                color: run.color,
            });
            continue;
        }
        if ch.is_whitespace() {
            flush(&mut out, &mut buffer, run);
            out.push(StyledRun {
                text: ch.to_string(),
                font: run.font,
                size: run.size,
                color: run.color,
            });
            continue;
        }
        buffer.push(ch);
    }
    flush(&mut out, &mut buffer, run);
    out
}

fn trim_trailing_spaces(line: &mut LineLayout) {
    while matches!(line.runs.last(), Some(run) if run.text.chars().all(|ch| ch.is_whitespace())) {
        if let Some(run) = line.runs.pop() {
            line.width -= measure_text(&run.text, run.size);
        }
    }
}

fn cell_plain_text(cell: &TableCell) -> String {
    cell.paragraphs
        .iter()
        .map(|paragraph| paragraph.spans.iter().map(inline_text).collect::<Vec<_>>().join(""))
        .collect::<Vec<_>>()
        .join(" ")
}

fn inline_text(inline: &Inline) -> String {
    match inline {
        Inline::Text(text) | Inline::Bold(text) | Inline::Italic(text) | Inline::Code(text) => {
            text.clone()
        }
        Inline::Link { label, .. } => label.clone(),
        Inline::LineBreak => "\n".to_string(),
    }
}

fn ensure_room(
    ctx: &LayoutContext,
    pdf: &mut PdfDocument,
    page_index: &mut usize,
    y: f32,
    needed: f32,
) -> f32 {
    if y - needed >= ctx.margin_b {
        y
    } else {
        *page_index = pdf.add_page(ctx.page_w, ctx.page_h);
        ctx.page_h - ctx.margin_t
    }
}

fn measure_text(text: &str, font_size: f32) -> f32 {
    let mut width = 0.0;
    for ch in text.chars() {
        let factor = match ch {
            'i' | 'l' | '!' | '.' | ',' | ';' | ':' | '|' => 0.26,
            'm' | 'w' | 'M' | 'W' | '@' | '#' => 0.82,
            ' ' => 0.28,
            '0'..='9' => 0.55,
            _ if ch.is_ascii_uppercase() => 0.62,
            _ => 0.52,
        };
        width += font_size * factor;
    }
    width
}

fn parse_color(input: &str) -> Option<(f32, f32, f32)> {
    let trimmed = input.trim().trim_start_matches('#');
    if trimmed.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&trimmed[0..2], 16).ok()? as f32 / 255.0;
    let g = u8::from_str_radix(&trimmed[2..4], 16).ok()? as f32 / 255.0;
    let b = u8::from_str_radix(&trimmed[4..6], 16).ok()? as f32 / 255.0;
    Some((r, g, b))
}

fn mm_to_pt(mm: u16) -> f32 {
    mm as f32 * 72.0 / 25.4
}

impl LayoutContext {
    fn from_page_style(style: &PageStyle) -> Self {
        let page_w = mm_to_pt(style.width_mm.max(1));
        let page_h = mm_to_pt(style.height_mm.max(1));
        let margin = mm_to_pt(style.margin_mm.max(10));
        Self {
            page_w,
            page_h,
            margin_l: margin,
            margin_r: margin,
            margin_t: margin,
            margin_b: margin,
            fonts: FontResolver::scan(),
        }
    }
}

#[derive(Clone, Debug, Default)]
struct FontResolver {
    installed: BTreeSet<String>,
}

impl FontResolver {
    fn scan() -> Self {
        let mut installed = BTreeSet::new();
        let mut roots = vec![PathBuf::from("/usr/share/fonts"), PathBuf::from("/usr/local/share/fonts")];
        if let Ok(home) = std::env::var("HOME") {
            roots.push(PathBuf::from(home).join(".fonts"));
        }
        for root in roots {
            collect_font_names(&root, &mut installed);
        }
        Self { installed }
    }

    fn pick(&self, requested: Option<&str>, bold: bool, italic: bool) -> PdfFont {
        let request = requested.unwrap_or("").to_ascii_lowercase();
        let serif = request.contains("times")
            || request.contains("serif")
            || request.contains("georgia")
            || request.contains("cambria")
            || request.contains("garamond");
        let mono = request.contains("mono")
            || request.contains("code")
            || request.contains("courier")
            || request.contains("consolas");
        let known = if request.is_empty() {
            false
        } else {
            self.installed.iter().any(|name| name.contains(&request))
        };
        if mono {
            return PdfFont::Courier;
        }
        if serif || known && (request.contains("times") || request.contains("serif")) {
            return match (bold, italic) {
                (true, _) => PdfFont::TimesBold,
                (false, true) => PdfFont::TimesItalic,
                _ => PdfFont::TimesRoman,
            };
        }
        match (bold, italic) {
            (true, _) => PdfFont::HelveticaBold,
            (false, true) => PdfFont::HelveticaOblique,
            _ => PdfFont::Helvetica,
        }
    }
}

fn collect_font_names(root: &PathBuf, out: &mut BTreeSet<String>) {
    let Ok(entries) = fs::read_dir(root) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect_font_names(&path, out);
        } else if let Some(ext) = path.extension().and_then(|value| value.to_str()) {
            if matches!(ext.to_ascii_lowercase().as_str(), "ttf" | "otf" | "ttc") {
                if let Some(stem) = path.file_stem().and_then(|value| value.to_str()) {
                    out.insert(stem.to_ascii_lowercase());
                }
            }
        }
    }
}
