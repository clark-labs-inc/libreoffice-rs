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
    decode_png, decode_webp, parse_xml_document, Alignment, Block, Heading, ImageBlock, Inline,
    ListBlock, ListItem, PageStyle, Paragraph, PdfDocument, PdfFont, Table, TableCell,
    TextDocument, XmlNode,
};

use crate::font_metrics::measure_text;

#[derive(Clone, Debug)]
struct StyledRun {
    text: String,
    font: PdfFont,
    size: f32,
    color: (f32, f32, f32),
    background: Option<(f32, f32, f32)>,
    underline: bool,
    link: Option<String>,
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
            background: None,
            underline: false,
            link: None,
        };
        y = ensure_room(&ctx, &mut pdf, &mut page_index, y, 30.0);
        render_line(
            pdf.page_mut(page_index).expect("page"),
            &LineLayout {
                runs: vec![title_run],
                width: measure_text(&doc.meta.title, 22.0, PdfFont::HelveticaBold),
            },
            ctx.margin_l,
            y,
            Alignment::Start,
            ctx.page_w - ctx.margin_l - ctx.margin_r,
            false,
            "H1",
        );
        y -= 30.0;
    }

    for (block_index, block) in doc.body.iter().enumerate() {
        if matches!(block, Block::Heading(_)) {
            if let Some(Block::Image(image)) = doc.body.get(block_index + 1) {
                if matches!(
                    image.mime_type.as_str(),
                    "application/vnd.mermaid+text" | "application/x-latex" | "image/svg+xml"
                ) {
                    let image_height = image
                        .size
                        .height
                        .as_pt()
                        .min(ctx.page_h - ctx.margin_t - ctx.margin_b - 28.0);
                    y = ensure_room(&ctx, &mut pdf, &mut page_index, y, image_height + 58.0);
                }
            }
        }
        match block {
            Block::Heading(heading) => {
                y = render_heading(&ctx, &mut pdf, &mut page_index, y, heading);
            }
            Block::Paragraph(paragraph) => {
                y = render_paragraph(
                    &ctx,
                    &mut pdf,
                    &mut page_index,
                    y,
                    paragraph,
                    12.0,
                    0.0,
                    "P",
                );
            }
            Block::List(list) => {
                y = render_list(&ctx, &mut pdf, &mut page_index, y, list);
            }
            Block::Table(table) => {
                y = render_table(&ctx, &mut pdf, &mut page_index, y, table);
            }
            Block::Image(image) => {
                let max_width = ctx.page_w - ctx.margin_l - ctx.margin_r;
                let mut width = image.size.width.as_pt().max(80.0).min(max_width);
                let mut height = image.size.height.as_pt().max(48.0);
                let max_height = ctx.page_h - ctx.margin_t - ctx.margin_b - 28.0;
                if height > max_height {
                    width *= max_height / height;
                    height = max_height;
                }
                y = ensure_room(&ctx, &mut pdf, &mut page_index, y, height + 18.0);
                let bottom = y - height;
                let resource = register_pdf_image(&mut pdf, image);
                let page = pdf.page_mut(page_index).expect("page");
                page.begin_tag("Figure", Some(&image.alt));
                let rendered = match image.mime_type.as_str() {
                    "image/svg+xml" => {
                        render_svg_image(page, &image.data, ctx.margin_l, bottom, width, height)
                    }
                    "application/vnd.mermaid+text" => {
                        render_mermaid(page, &image.data, ctx.margin_l, bottom, width, height)
                    }
                    "application/x-latex" => {
                        render_formula(page, &image.data, ctx.margin_l, bottom, width, height)
                    }
                    _ => false,
                };
                if let Some(resource) = resource {
                    page.image(&resource, ctx.margin_l, bottom, width, height);
                } else if !rendered {
                    page.rect_fill_stroke_rgb(
                        ctx.margin_l,
                        bottom,
                        width,
                        height,
                        (0.98, 0.98, 0.98),
                        (0.60, 0.60, 0.60),
                    );
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
                }
                page.end_tag();
                if !image.alt.trim().is_empty() {
                    page.begin_tag("Caption", None);
                    page.text_rgb(
                        ctx.margin_l,
                        bottom - 11.0,
                        9.0,
                        PdfFont::HelveticaOblique,
                        &image.alt,
                        0.35,
                        0.33,
                        0.30,
                    );
                    page.end_tag();
                }
                y = bottom
                    - if image.alt.trim().is_empty() {
                        12.0
                    } else {
                        22.0
                    };
            }
            Block::Section(section) => {
                y = ensure_room(&ctx, &mut pdf, &mut page_index, y, 20.0);
                pdf.page_mut(page_index).expect("page").text_rgb(
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
                        y = render_paragraph(
                            &ctx,
                            &mut pdf,
                            &mut page_index,
                            y,
                            paragraph,
                            12.0,
                            0.0,
                            "P",
                        );
                    }
                }
            }
            Block::HorizontalRule => {
                y = ensure_room(&ctx, &mut pdf, &mut page_index, y, 14.0);
                let page = pdf.page_mut(page_index).expect("page");
                page.begin_artifact();
                page.line_rgb(
                    ctx.margin_l,
                    y,
                    ctx.page_w - ctx.margin_r,
                    y,
                    0.55,
                    0.55,
                    0.55,
                );
                page.end_artifact();
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
    // Keep a heading with at least the first two lines of following content.
    // This is deliberately conservative: the full following paragraph applies
    // its own keep-together rule in `render_paragraph`.
    y = ensure_room(ctx, pdf, page_index, y, size * 3.5);
    let mut paragraph = heading.content.clone();
    paragraph.text_style.font_size_pt = size.round() as u16;
    paragraph.text_style.bold = true;
    paragraph.style.margin_top_mm = paragraph.style.margin_top_mm.max(2);
    paragraph.style.margin_bottom_mm = paragraph.style.margin_bottom_mm.max(2);
    let role = format!("H{}", heading.level.clamp(1, 6));
    y = render_paragraph(ctx, pdf, page_index, y, &paragraph, size, 0.0, &role);
    y - 2.0
}

fn render_list(
    ctx: &LayoutContext,
    pdf: &mut PdfDocument,
    page_index: &mut usize,
    mut y: f32,
    list: &ListBlock,
) -> f32 {
    let available_width = (ctx.page_w - ctx.margin_l - ctx.margin_r - 18.0).max(48.0);
    let list_height = list
        .items
        .iter()
        .flat_map(|item| item.blocks.iter())
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph_height(
                paragraph,
                12.0,
                available_width,
                &ctx.fonts,
            )),
            _ => None,
        })
        .sum::<f32>()
        + list.items.len() as f32 * 4.0
        + 4.0;
    let content_height = ctx.page_h - ctx.margin_t - ctx.margin_b;
    if list_height <= content_height
        && y - list_height < ctx.margin_b
        && y < ctx.page_h - ctx.margin_t
    {
        *page_index = pdf.add_page(ctx.page_w, ctx.page_h);
        y = ctx.page_h - ctx.margin_t;
    }
    for (index, item) in list.items.iter().enumerate() {
        let marker = if list.ordered {
            format!("{}.", index + 1)
        } else {
            "•".to_string()
        };
        y = ensure_room(ctx, pdf, page_index, y, 17.0);
        let marker_baseline = y - 15.0;
        let page = pdf.page_mut(*page_index).expect("page");
        page.begin_artifact();
        page.text(
            ctx.margin_l + 2.0,
            marker_baseline,
            12.0,
            PdfFont::Helvetica,
            &marker,
        );
        page.end_artifact();
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
            y = render_paragraph(ctx, pdf, page_index, y, paragraph, 12.0, 18.0, "LI");
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
    semantic_role: &str,
) -> f32 {
    let left = ctx.margin_l + mm_to_pt(paragraph.style.margin_left_mm) + extra_indent;
    let right = ctx.margin_r + mm_to_pt(paragraph.style.margin_right_mm);
    let available_width = (ctx.page_w - left - right).max(48.0);
    let margin_top = mm_to_pt(paragraph.style.margin_top_mm);
    let margin_bottom = mm_to_pt(paragraph.style.margin_bottom_mm.max(1));
    let runs = paragraph_runs(paragraph, default_size, &ctx.fonts);
    let lines = layout_runs(&runs, available_width);
    let line_heights = lines
        .iter()
        .map(|line| {
            line.runs
                .iter()
                .map(|run| run.size * 1.25)
                .fold(default_size * 1.25, f32::max)
        })
        .collect::<Vec<_>>();
    let total_height = margin_top + line_heights.iter().sum::<f32>() + margin_bottom;
    let content_height = ctx.page_h - ctx.margin_t - ctx.margin_b;
    if total_height <= content_height
        && y - total_height < ctx.margin_b
        && y < ctx.page_h - ctx.margin_t
    {
        *page_index = pdf.add_page(ctx.page_w, ctx.page_h);
        y = ctx.page_h - ctx.margin_t;
    }
    y -= margin_top;

    for (line_index, line) in lines.iter().enumerate() {
        let line_height = line_heights[line_index];
        y = ensure_room(ctx, pdf, page_index, y, line_height + 2.0);
        // `y` is the free block boundary; PDF text APIs take a baseline.
        // Descend by one line box before drawing so glyph ascenders never
        // intrude into the preceding table, paragraph, or page margin.
        y -= line_height;
        if let Some((r, g, b)) = parse_color(&paragraph.text_style.background) {
            let page = pdf.page_mut(*page_index).expect("page");
            page.begin_artifact();
            page.rect_fill_rgb(
                left - 4.0,
                y - 3.0,
                available_width + 8.0,
                line_height + 4.0,
                r,
                g,
                b,
            );
            page.end_artifact();
        }
        let justify =
            matches!(paragraph.style.alignment, Alignment::Justify) && line_index + 1 < lines.len();
        render_line(
            pdf.page_mut(*page_index).expect("page"),
            line,
            left,
            y,
            paragraph.style.alignment.clone(),
            available_width,
            justify,
            semantic_role,
        );
    }

    y - margin_bottom
}

fn paragraph_height(
    paragraph: &Paragraph,
    default_size: f32,
    available_width: f32,
    fonts: &FontResolver,
) -> f32 {
    let lines = layout_runs(
        &paragraph_runs(paragraph, default_size, fonts),
        available_width,
    );
    mm_to_pt(paragraph.style.margin_top_mm)
        + lines
            .iter()
            .map(|line| {
                line.runs
                    .iter()
                    .map(|run| run.size * 1.25)
                    .fold(default_size * 1.25, f32::max)
            })
            .sum::<f32>()
        + mm_to_pt(paragraph.style.margin_bottom_mm.max(1))
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
            let suggested = measure_text(&text, 10.0, PdfFont::Helvetica) + 12.0;
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

    for (row_index, row) in table.rows.iter().enumerate() {
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
            if row_index == 0 {
                page.begin_artifact();
                page.rect_fill_rgb(x, bottom, *width, row_height, 0.94, 0.93, 0.90);
                page.end_artifact();
            }
            page.begin_artifact();
            page.rect_stroke_rgb(x, bottom, *width, row_height, 0.50, 0.50, 0.50);
            page.end_artifact();
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
                    if row_index == 0 { "TH" } else { "TD" },
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
    semantic_role: &str,
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
    page.begin_tag(semantic_role, None);
    while idx < line.runs.len() {
        let start = idx;
        let head = &line.runs[idx];
        let mut combined = String::new();
        while idx < line.runs.len() {
            let cur = &line.runs[idx];
            if cur.font != head.font
                || cur.size != head.size
                || cur.color != head.color
                || cur.background != head.background
                || cur.underline != head.underline
                || cur.link != head.link
            {
                break;
            }
            combined.push_str(&cur.text);
            idx += 1;
        }
        let group_width: f32 = line.runs[start..idx]
            .iter()
            .map(|r| measure_text(&r.text, r.size, r.font))
            .sum();
        if !combined.is_empty() {
            if let Some(background) = head.background {
                page.begin_artifact();
                page.rect_fill_rgb(
                    cursor - 1.0,
                    y - 2.0,
                    group_width + 2.0,
                    head.size + 4.0,
                    background.0,
                    background.1,
                    background.2,
                );
                page.end_artifact();
            }
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
            if let Some(uri) = &head.link {
                page.link(cursor, y - 2.0, group_width, head.size + 4.0, uri.clone());
            }
            if head.underline {
                page.begin_artifact();
                page.line_rgb(
                    cursor,
                    y - 1.5,
                    cursor + group_width,
                    y - 1.5,
                    head.color.0,
                    head.color.1,
                    head.color.2,
                );
                page.end_artifact();
            }
        }
        cursor += group_width;
        if extra_per_space > 0.0 {
            cursor += combined.matches(' ').count() as f32 * extra_per_space;
        }
    }
    page.end_tag();
}

fn paragraph_runs(
    paragraph: &Paragraph,
    default_size: f32,
    fonts: &FontResolver,
) -> Vec<StyledRun> {
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
                font: fonts.pick(
                    family,
                    paragraph.text_style.bold,
                    paragraph.text_style.italic,
                ),
                size,
                color: base_color,
                background: None,
                underline: paragraph.text_style.underline,
                link: None,
            }),
            Inline::Bold(text) => runs.push(StyledRun {
                text: text.clone(),
                font: fonts.pick(family, true, paragraph.text_style.italic),
                size,
                color: base_color,
                background: None,
                underline: paragraph.text_style.underline,
                link: None,
            }),
            Inline::Italic(text) => runs.push(StyledRun {
                text: text.clone(),
                font: fonts.pick(family, paragraph.text_style.bold, true),
                size,
                color: base_color,
                background: None,
                underline: paragraph.text_style.underline,
                link: None,
            }),
            Inline::Code(text) => runs.push(StyledRun {
                text: text.clone(),
                font: PdfFont::Courier,
                size: (size - 1.0).max(9.0),
                color: (0.10, 0.10, 0.10),
                background: Some((0.95, 0.94, 0.91)),
                underline: false,
                link: None,
            }),
            Inline::Link { label, url } => runs.push(StyledRun {
                text: label.clone(),
                font: fonts.pick(family, false, false),
                size,
                color: (0.10, 0.25, 0.65),
                background: None,
                underline: true,
                link: Some(url.clone()),
            }),
            Inline::Styled { text, style, url } => {
                let styled_family = if style.font_family.trim().is_empty() {
                    family
                } else {
                    Some(style.font_family.as_str())
                };
                runs.push(StyledRun {
                    text: text.clone(),
                    font: fonts.pick(
                        styled_family,
                        style.bold || paragraph.text_style.bold,
                        style.italic || paragraph.text_style.italic,
                    ),
                    size: if style.font_size_pt == 0 {
                        size
                    } else {
                        style.font_size_pt as f32
                    },
                    color: parse_color(&style.color).unwrap_or(base_color),
                    background: parse_color(&style.background),
                    underline: style.underline || paragraph.text_style.underline,
                    link: url.clone(),
                });
            }
            Inline::LineBreak => runs.push(StyledRun {
                text: "\n".to_string(),
                font: fonts.pick(family, false, false),
                size,
                color: base_color,
                background: None,
                underline: false,
                link: None,
            }),
        }
    }
    if runs.is_empty() {
        runs.push(StyledRun {
            text: String::new(),
            font: fonts.pick(family, false, false),
            size,
            color: base_color,
            background: None,
            underline: false,
            link: None,
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
            let token_width = measure_text(&token.text, token.size, token.font);
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
                background: template.background,
                underline: template.underline,
                link: template.link.clone(),
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
                background: run.background,
                underline: run.underline,
                link: run.link.clone(),
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
                background: run.background,
                underline: run.underline,
                link: run.link.clone(),
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
            line.width -= measure_text(&run.text, run.size, run.font);
        }
    }
}

fn cell_plain_text(cell: &TableCell) -> String {
    cell.paragraphs
        .iter()
        .map(|paragraph| {
            paragraph
                .spans
                .iter()
                .map(inline_text)
                .collect::<Vec<_>>()
                .join("")
        })
        .collect::<Vec<_>>()
        .join(" ")
}

fn inline_text(inline: &Inline) -> String {
    match inline {
        Inline::Text(text)
        | Inline::Bold(text)
        | Inline::Italic(text)
        | Inline::Code(text)
        | Inline::Styled { text, .. } => text.clone(),
        Inline::Link { label, .. } => label.clone(),
        Inline::LineBreak => "\n".to_string(),
    }
}

fn register_pdf_image(pdf: &mut PdfDocument, image: &ImageBlock) -> Option<String> {
    match image.mime_type.as_str() {
        "image/png" => {
            let decoded = decode_png(&image.data).ok()?;
            Some(register_raster_image(pdf, decoded))
        }
        "image/webp" => {
            let decoded = decode_webp(&image.data).ok()?;
            Some(register_raster_image(pdf, decoded))
        }
        "image/jpeg" => {
            let (width, height) = jpeg_dimensions(&image.data)?;
            Some(pdf.add_jpeg_image(width, height, image.data.clone()))
        }
        _ => None,
    }
}

fn register_raster_image(pdf: &mut PdfDocument, decoded: lo_core::RasterImage) -> String {
    let mut rgb = Vec::with_capacity(decoded.width as usize * decoded.height as usize * 3);
    for pixel in decoded.pixels.chunks_exact(4) {
        let alpha = pixel[3] as u16;
        let inverse = 255 - alpha;
        rgb.push(((pixel[0] as u16 * alpha + 255 * inverse) / 255) as u8);
        rgb.push(((pixel[1] as u16 * alpha + 255 * inverse) / 255) as u8);
        rgb.push(((pixel[2] as u16 * alpha + 255 * inverse) / 255) as u8);
    }
    pdf.add_rgb_image(decoded.width, decoded.height, rgb)
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

fn render_svg_image(
    page: &mut lo_core::PdfPage,
    bytes: &[u8],
    x: f32,
    bottom: f32,
    width: f32,
    height: f32,
) -> bool {
    let Ok(svg) = std::str::from_utf8(bytes) else {
        return false;
    };
    let Ok(root) = parse_xml_document(svg) else {
        return false;
    };
    if root.local_name() != "svg" {
        return false;
    }
    let view_box = root
        .attr("viewBox")
        .or_else(|| root.attr("viewbox"))
        .and_then(parse_number_list)
        .filter(|values| values.len() == 4);
    let (min_x, min_y, svg_width, svg_height) = if let Some(values) = view_box {
        (values[0], values[1], values[2].max(1.0), values[3].max(1.0))
    } else {
        (
            0.0,
            0.0,
            svg_length(root.attr("width")).unwrap_or(100.0).max(1.0),
            svg_length(root.attr("height")).unwrap_or(100.0).max(1.0),
        )
    };
    let transform = SvgTransform {
        x,
        bottom,
        width,
        height,
        min_x,
        min_y,
        svg_width,
        svg_height,
    };
    for child in &root.children {
        render_svg_node(page, child, transform);
    }
    true
}

#[derive(Clone, Copy)]
struct SvgTransform {
    x: f32,
    bottom: f32,
    width: f32,
    height: f32,
    min_x: f32,
    min_y: f32,
    svg_width: f32,
    svg_height: f32,
}

impl SvgTransform {
    fn x(self, value: f32) -> f32 {
        self.x + (value - self.min_x) * self.width / self.svg_width
    }

    fn y(self, value: f32) -> f32 {
        self.bottom + self.height - (value - self.min_y) * self.height / self.svg_height
    }

    fn w(self, value: f32) -> f32 {
        value * self.width / self.svg_width
    }

    fn h(self, value: f32) -> f32 {
        value * self.height / self.svg_height
    }
}

fn render_svg_node(page: &mut lo_core::PdfPage, node: &XmlNode, transform: SvgTransform) {
    if matches!(node.local_name(), "g" | "svg") {
        for child in &node.children {
            render_svg_node(page, child, transform);
        }
        return;
    }
    let fill = svg_paint(node, "fill", Some((0.12, 0.12, 0.12)));
    let stroke = svg_paint(node, "stroke", Some((0.20, 0.20, 0.20)));
    let stroke_width = svg_style_value(node, "stroke-width")
        .and_then(|value| svg_length(Some(value)))
        .unwrap_or(1.0);
    page.line_width(transform.w(stroke_width).max(0.4));
    match node.local_name() {
        "rect" => {
            let sx = transform.x(svg_number(node, "x", 0.0));
            let sy = transform.y(svg_number(node, "y", 0.0) + svg_number(node, "height", 0.0));
            let sw = transform.w(svg_number(node, "width", 0.0));
            let sh = transform.h(svg_number(node, "height", 0.0));
            match (fill, stroke) {
                (Some(fill), Some(stroke)) => {
                    page.rect_fill_stroke_rgb(sx, sy, sw, sh, fill, stroke)
                }
                (Some(fill), None) => page.rect_fill_rgb(sx, sy, sw, sh, fill.0, fill.1, fill.2),
                (None, Some(stroke)) => {
                    page.rect_stroke_rgb(sx, sy, sw, sh, stroke.0, stroke.1, stroke.2)
                }
                _ => {}
            }
        }
        "circle" | "ellipse" => {
            let cx = transform.x(svg_number(node, "cx", 0.0));
            let cy = transform.y(svg_number(node, "cy", 0.0));
            let rx = transform.w(if node.local_name() == "circle" {
                svg_number(node, "r", 0.0)
            } else {
                svg_number(node, "rx", 0.0)
            });
            let ry = transform.h(if node.local_name() == "circle" {
                svg_number(node, "r", 0.0)
            } else {
                svg_number(node, "ry", 0.0)
            });
            match (fill, stroke) {
                (Some(fill), Some(stroke)) => {
                    page.ellipse_fill_stroke_rgb(cx, cy, rx, ry, fill, stroke)
                }
                (_, Some(stroke)) => {
                    page.ellipse_stroke_rgb(cx, cy, rx, ry, stroke.0, stroke.1, stroke.2)
                }
                (Some(fill), None) => page.ellipse_fill_stroke_rgb(cx, cy, rx, ry, fill, fill),
                _ => {}
            }
        }
        "line" => {
            if let Some(stroke) = stroke {
                page.line_rgb(
                    transform.x(svg_number(node, "x1", 0.0)),
                    transform.y(svg_number(node, "y1", 0.0)),
                    transform.x(svg_number(node, "x2", 0.0)),
                    transform.y(svg_number(node, "y2", 0.0)),
                    stroke.0,
                    stroke.1,
                    stroke.2,
                );
            }
        }
        "polygon" | "polyline" => {
            if let Some(values) = node.attr("points").and_then(parse_number_list) {
                let points = values
                    .chunks_exact(2)
                    .map(|point| (transform.x(point[0]), transform.y(point[1])))
                    .collect::<Vec<_>>();
                if node.local_name() == "polygon" && points.len() >= 3 {
                    let fill = fill.unwrap_or((0.42, 0.35, 0.82));
                    page.polygon_fill_stroke_rgb(
                        &points,
                        fill,
                        stroke.unwrap_or((0.20, 0.17, 0.38)),
                    );
                } else if let Some(stroke) = stroke {
                    for pair in points.windows(2) {
                        page.line_rgb(
                            pair[0].0, pair[0].1, pair[1].0, pair[1].1, stroke.0, stroke.1,
                            stroke.2,
                        );
                    }
                }
            }
        }
        "text" => {
            let color = fill.or(stroke).unwrap_or((0.10, 0.10, 0.10));
            let size = transform
                .h(svg_style_value(node, "font-size")
                    .and_then(|v| svg_length(Some(v)))
                    .unwrap_or(14.0))
                .clamp(7.0, 32.0);
            let font = if svg_style_value(node, "font-weight")
                .is_some_and(|value| value.eq_ignore_ascii_case("bold") || value == "700")
            {
                PdfFont::HelveticaBold
            } else {
                PdfFont::Helvetica
            };
            page.text_rgb(
                transform.x(svg_number(node, "x", 0.0)),
                transform.y(svg_number(node, "y", 0.0)),
                size,
                font,
                &node.text_content(),
                color.0,
                color.1,
                color.2,
            );
        }
        _ => {}
    }
    page.line_width(1.0);
}

fn svg_number(node: &XmlNode, name: &str, fallback: f32) -> f32 {
    svg_length(node.attr(name)).unwrap_or(fallback)
}

fn svg_length(value: Option<&str>) -> Option<f32> {
    value?
        .trim()
        .trim_end_matches(|ch: char| ch.is_ascii_alphabetic() || ch == '%')
        .parse()
        .ok()
}

fn parse_number_list(value: &str) -> Option<Vec<f32>> {
    let numbers = value
        .split(|ch: char| ch.is_whitespace() || ch == ',')
        .filter(|part| !part.is_empty())
        .map(str::parse::<f32>)
        .collect::<std::result::Result<Vec<_>, _>>()
        .ok()?;
    (!numbers.is_empty()).then_some(numbers)
}

fn svg_style_value<'a>(node: &'a XmlNode, name: &str) -> Option<&'a str> {
    if let Some(value) = node.attr(name) {
        return Some(value);
    }
    node.attr("style")?.split(';').find_map(|declaration| {
        let (property, value) = declaration.split_once(':')?;
        property
            .trim()
            .eq_ignore_ascii_case(name)
            .then_some(value.trim())
    })
}

fn svg_paint(
    node: &XmlNode,
    name: &str,
    fallback: Option<(f32, f32, f32)>,
) -> Option<(f32, f32, f32)> {
    let Some(value) = svg_style_value(node, name) else {
        return fallback;
    };
    if value.eq_ignore_ascii_case("none") {
        return None;
    }
    parse_color(&normalize_named_color(value)).or(fallback)
}

fn normalize_named_color(value: &str) -> String {
    match value.trim().to_ascii_lowercase().as_str() {
        "red" => "#d32f2f",
        "green" => "#2e7d32",
        "blue" => "#1565c0",
        "purple" => "#6d4cc7",
        "orange" => "#ef8c32",
        "white" => "#ffffff",
        "black" => "#000000",
        "gray" | "grey" => "#777777",
        _ => value,
    }
    .to_string()
}

#[derive(Clone, Debug)]
struct MermaidNode {
    id: String,
    label: String,
    decision: bool,
}

fn render_mermaid(
    page: &mut lo_core::PdfPage,
    bytes: &[u8],
    x: f32,
    bottom: f32,
    width: f32,
    height: f32,
) -> bool {
    let source = String::from_utf8_lossy(bytes);
    let horizontal = source.lines().next().is_some_and(|line| {
        let upper = line.to_ascii_uppercase();
        upper.contains(" LR") || upper.ends_with("LR")
    });
    let mut nodes = Vec::<MermaidNode>::new();
    let mut edges = Vec::<(String, String)>::new();
    for line in source.lines().map(str::trim) {
        if line.is_empty()
            || line.starts_with("flowchart")
            || line.starts_with("graph")
            || line.starts_with("%%")
        {
            continue;
        }
        if let Some((left, right)) = line.split_once("-->") {
            let source_node = parse_mermaid_node(left.trim());
            let right = right.trim().trim_start_matches('|');
            let right = right.split_once('|').map_or(right, |(_, rest)| rest.trim());
            let target_node = parse_mermaid_node(right.trim());
            add_mermaid_node(&mut nodes, source_node.clone());
            add_mermaid_node(&mut nodes, target_node.clone());
            edges.push((source_node.id, target_node.id));
        } else {
            add_mermaid_node(&mut nodes, parse_mermaid_node(line));
        }
    }
    if nodes.is_empty() {
        return false;
    }
    page.rect_fill_stroke_rgb(
        x,
        bottom,
        width,
        height,
        (0.98, 0.975, 0.99),
        (0.78, 0.74, 0.88),
    );
    let count = nodes.len() as f32;
    let node_width = if horizontal {
        (width / count - 18.0).clamp(70.0, 125.0)
    } else {
        (width * 0.46).clamp(110.0, 190.0)
    };
    let node_height = 38.0;
    let positions = nodes
        .iter()
        .enumerate()
        .map(|(index, node)| {
            let (cx, cy) = if horizontal {
                (
                    x + width * (index as f32 + 0.5) / count,
                    bottom + height / 2.0,
                )
            } else {
                (
                    x + width / 2.0,
                    bottom + height - height * (index as f32 + 0.5) / count,
                )
            };
            (node.id.clone(), cx, cy)
        })
        .collect::<Vec<_>>();
    for (source, target) in &edges {
        let Some((_, sx, sy)) = positions.iter().find(|(id, _, _)| id == source) else {
            continue;
        };
        let Some((_, tx, ty)) = positions.iter().find(|(id, _, _)| id == target) else {
            continue;
        };
        let (start_x, start_y, end_x, end_y) = if horizontal {
            (sx + node_width / 2.0, *sy, tx - node_width / 2.0, *ty)
        } else {
            (*sx, sy - node_height / 2.0, *tx, ty + node_height / 2.0)
        };
        page.line_rgb(start_x, start_y, end_x, end_y, 0.35, 0.31, 0.49);
        let angle = (end_y - start_y).atan2(end_x - start_x);
        let arrow = 7.0;
        page.polygon_fill_stroke_rgb(
            &[
                (end_x, end_y),
                (
                    end_x - arrow * (angle - 0.55).cos(),
                    end_y - arrow * (angle - 0.55).sin(),
                ),
                (
                    end_x - arrow * (angle + 0.55).cos(),
                    end_y - arrow * (angle + 0.55).sin(),
                ),
            ],
            (0.35, 0.31, 0.49),
            (0.35, 0.31, 0.49),
        );
    }
    for (node, (_, cx, cy)) in nodes
        .iter()
        .zip(positions.iter().map(|(_, x, y)| ((), *x, *y)))
    {
        let fill = if node.decision {
            (0.98, 0.76, 0.42)
        } else {
            (0.42, 0.35, 0.82)
        };
        if node.decision {
            page.polygon_fill_stroke_rgb(
                &[
                    (cx, cy + node_height / 2.0),
                    (cx + node_width / 2.0, cy),
                    (cx, cy - node_height / 2.0),
                    (cx - node_width / 2.0, cy),
                ],
                fill,
                (0.28, 0.23, 0.43),
            );
        } else {
            page.rect_fill_stroke_rgb(
                cx - node_width / 2.0,
                cy - node_height / 2.0,
                node_width,
                node_height,
                fill,
                (0.28, 0.23, 0.43),
            );
        }
        let text_color = if node.decision {
            (0.18, 0.13, 0.08)
        } else {
            (1.0, 1.0, 1.0)
        };
        let label_width = measure_text(&node.label, 10.0, PdfFont::HelveticaBold);
        page.text_rgb(
            cx - label_width / 2.0,
            cy - 3.5,
            10.0,
            PdfFont::HelveticaBold,
            &node.label,
            text_color.0,
            text_color.1,
            text_color.2,
        );
    }
    true
}

fn add_mermaid_node(nodes: &mut Vec<MermaidNode>, node: MermaidNode) {
    if !node.id.is_empty() && !nodes.iter().any(|existing| existing.id == node.id) {
        nodes.push(node);
    }
}

fn parse_mermaid_node(token: &str) -> MermaidNode {
    let token = token.trim().trim_end_matches(';');
    let boundary = token
        .find(|ch| matches!(ch, '[' | '{' | '('))
        .unwrap_or(token.len());
    let id = token[..boundary].trim().to_string();
    let decision = token[boundary..].starts_with('{');
    let label = if boundary < token.len() {
        token[boundary + 1..]
            .trim_end_matches(|ch| matches!(ch, ']' | '}' | ')'))
            .trim_matches(['\'', '\"'])
            .to_string()
    } else {
        id.clone()
    };
    MermaidNode {
        id,
        label,
        decision,
    }
}

fn render_formula(
    page: &mut lo_core::PdfPage,
    bytes: &[u8],
    x: f32,
    bottom: f32,
    width: f32,
    height: f32,
) -> bool {
    let source = String::from_utf8_lossy(bytes);
    page.rect_fill_stroke_rgb(
        x,
        bottom,
        width,
        height,
        (0.975, 0.97, 0.99),
        (0.75, 0.70, 0.86),
    );
    let formula = source.trim().replace(['\n', '\r'], " ");
    let Some(frac_start) = formula.find("\\frac{") else {
        let text = readable_latex(&formula);
        let size = 18.0;
        let text_width = measure_text(&text, size, PdfFont::TimesItalic);
        page.text_rgb(
            x + (width - text_width) / 2.0,
            bottom + height / 2.0 - 6.0,
            size,
            PdfFont::TimesItalic,
            &text,
            0.20,
            0.16,
            0.33,
        );
        return true;
    };
    let prefix = readable_latex(&formula[..frac_start]);
    let fraction_source = &formula[frac_start + "\\frac".len()..];
    let Some((numerator, after_numerator)) = take_braced(fraction_source) else {
        return false;
    };
    let Some((denominator, suffix)) = take_braced(after_numerator) else {
        return false;
    };
    let numerator = readable_latex(numerator);
    let denominator = readable_latex(denominator);
    let suffix = readable_latex(suffix);
    let text_size = 17.0;
    let fraction_size = 13.0;
    let prefix_width = measure_text(&prefix, text_size, PdfFont::TimesItalic);
    let numerator_width = measure_text(&numerator, fraction_size, PdfFont::TimesItalic);
    let denominator_width = measure_text(&denominator, fraction_size, PdfFont::TimesItalic);
    let fraction_width = numerator_width.max(denominator_width) + 12.0;
    let suffix_width = measure_text(&suffix, text_size, PdfFont::TimesItalic);
    let total_width = prefix_width + fraction_width + suffix_width;
    let start = x + (width - total_width) / 2.0;
    let center_y = bottom + height / 2.0;
    page.text_rgb(
        start,
        center_y - 5.0,
        text_size,
        PdfFont::TimesItalic,
        &prefix,
        0.20,
        0.16,
        0.33,
    );
    let fraction_x = start + prefix_width;
    page.text_rgb(
        fraction_x + (fraction_width - numerator_width) / 2.0,
        center_y + 7.0,
        fraction_size,
        PdfFont::TimesItalic,
        &numerator,
        0.20,
        0.16,
        0.33,
    );
    page.line_rgb(
        fraction_x + 3.0,
        center_y + 3.0,
        fraction_x + fraction_width - 3.0,
        center_y + 3.0,
        0.20,
        0.16,
        0.33,
    );
    page.text_rgb(
        fraction_x + (fraction_width - denominator_width) / 2.0,
        center_y - 12.0,
        fraction_size,
        PdfFont::TimesItalic,
        &denominator,
        0.20,
        0.16,
        0.33,
    );
    page.text_rgb(
        fraction_x + fraction_width,
        center_y - 5.0,
        text_size,
        PdfFont::TimesItalic,
        &suffix,
        0.20,
        0.16,
        0.33,
    );
    true
}

fn take_braced(input: &str) -> Option<(&str, &str)> {
    let input = input.trim_start();
    let inner = input.strip_prefix('{')?;
    let mut depth = 1usize;
    for (index, ch) in inner.char_indices() {
        match ch {
            '{' => depth += 1,
            '}' => {
                depth -= 1;
                if depth == 0 {
                    return Some((&inner[..index], &inner[index + 1..]));
                }
            }
            _ => {}
        }
    }
    None
}

fn readable_latex(source: &str) -> String {
    source
        .replace("\\times", "x")
        .replace("\\cdot", "·")
        .replace("\\pi", "pi")
        .replace("\\alpha", "alpha")
        .replace("\\beta", "beta")
        .replace("\\sqrt", "sqrt")
        .replace("^2", "²")
        .replace("^3", "³")
        .replace(['{', '}'], "")
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
        let mut roots = vec![
            PathBuf::from("/usr/share/fonts"),
            PathBuf::from("/usr/local/share/fonts"),
        ];
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
