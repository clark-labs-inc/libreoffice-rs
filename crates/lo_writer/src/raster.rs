use lo_core::{
    parse_hex_color, Alignment, Block, Heading, Inline, ListBlock, Paragraph, RasterImage, Rgba,
    Table, TextDocument,
};

#[derive(Clone, Debug)]
struct Run {
    text: String,
    size_px: i32,
    bold: bool,
    color: Rgba,
}

#[derive(Clone, Debug, Default)]
struct Line {
    runs: Vec<Run>,
    width: i32,
    height: i32,
}

pub fn render_png_pages(document: &TextDocument, dpi: u32) -> Vec<Vec<u8>> {
    render_pages(document, dpi)
        .into_iter()
        .map(|page| page.encode_png())
        .collect()
}

pub fn render_jpeg_pages(document: &TextDocument, dpi: u32, quality: u8) -> Vec<Vec<u8>> {
    render_pages(document, dpi)
        .into_iter()
        .map(|page| page.encode_jpeg(quality))
        .collect()
}

pub fn render_pages(document: &TextDocument, dpi: u32) -> Vec<RasterImage> {
    let width = mm_to_px(document.page_style.width_mm as f32, dpi).max(256);
    let height = mm_to_px(document.page_style.height_mm as f32, dpi).max(256);
    let margin = mm_to_px(document.page_style.margin_mm as f32, dpi).max(16);
    let mut pages = vec![RasterImage::new(width as u32, height as u32, Rgba::WHITE)];
    let mut page_index = 0usize;
    let mut y = margin;

    if !document.meta.title.trim().is_empty() {
        let title = Paragraph::plain(document.meta.title.clone());
        y = draw_paragraph(
            &mut pages,
            &mut page_index,
            y,
            width,
            height,
            margin,
            &title,
            20,
            0,
            true,
        );
    }

    for block in &document.body {
        match block {
            Block::Heading(heading) => {
                y = draw_heading(&mut pages, &mut page_index, y, width, height, margin, heading, dpi);
            }
            Block::Paragraph(paragraph) => {
                y = draw_paragraph(&mut pages, &mut page_index, y, width, height, margin, paragraph, 12, 0, false);
            }
            Block::List(list) => {
                y = draw_list(&mut pages, &mut page_index, y, width, height, margin, list);
            }
            Block::Table(table) => {
                y = draw_table(&mut pages, &mut page_index, y, width, height, margin, table);
            }
            Block::Image(image) => {
                let box_w = mm_to_px(image.size.width.as_mm(), dpi).max(120).min(width - margin * 2);
                let box_h = mm_to_px(image.size.height.as_mm(), dpi).max(70);
                y = ensure_room(&mut pages, &mut page_index, y, box_h + 24, width, height, margin);
                let page = &mut pages[page_index];
                page.fill_rect(margin, y, box_w, box_h, Rgba::rgba(245, 245, 245, 255));
                page.stroke_rect(margin, y, box_w, box_h, 2, Rgba::rgba(160, 160, 160, 255));
                page.draw_line(margin, y, margin + box_w, y + box_h, 1, Rgba::rgba(180, 180, 180, 255));
                page.draw_line(margin + box_w, y, margin, y + box_h, 1, Rgba::rgba(180, 180, 180, 255));
                let label = format!("image: {}", image.alt);
                page.draw_text(margin + 8, y + box_h / 2, 14, Rgba::rgba(70, 70, 70, 255), &label, false);
                y += box_h + 18;
            }
            Block::Section(section) => {
                let heading = Heading { level: 2, content: Paragraph::plain(section.name.clone()) };
                y = draw_heading(&mut pages, &mut page_index, y, width, height, margin, &heading, dpi);
                for nested in &section.blocks {
                    if let Block::Paragraph(paragraph) = nested {
                        y = draw_paragraph(&mut pages, &mut page_index, y, width, height, margin, paragraph, 12, 0, false);
                    }
                }
            }
            Block::HorizontalRule => {
                y = ensure_room(&mut pages, &mut page_index, y, 18, width, height, margin);
                pages[page_index].draw_line(margin, y + 6, width - margin, y + 6, 2, Rgba::rgba(140, 140, 140, 255));
                y += 18;
            }
            Block::PageBreak => {
                pages.push(RasterImage::new(width as u32, height as u32, Rgba::WHITE));
                page_index = pages.len() - 1;
                y = margin;
            }
        }
    }

    pages
}

fn draw_heading(
    pages: &mut Vec<RasterImage>,
    page_index: &mut usize,
    y: i32,
    width: i32,
    height: i32,
    margin: i32,
    heading: &Heading,
    _dpi: u32,
) -> i32 {
    let size = match heading.level {
        1 => 22,
        2 => 18,
        3 => 16,
        4 => 14,
        _ => 13,
    };
    draw_paragraph(pages, page_index, y, width, height, margin, &heading.content, size, 0, true)
}

fn draw_list(
    pages: &mut Vec<RasterImage>,
    page_index: &mut usize,
    mut y: i32,
    width: i32,
    height: i32,
    margin: i32,
    list: &ListBlock,
) -> i32 {
    for (index, item) in list.items.iter().enumerate() {
        let marker = if list.ordered { format!("{}.", index + 1) } else { "•".to_string() };
        y = ensure_room(pages, page_index, y, 18, width, height, margin);
        pages[*page_index].draw_text(margin + 4, y + 2, 13, Rgba::BLACK, &marker, true);
        for nested in &item.blocks {
            if let Block::Paragraph(paragraph) = nested {
                y = draw_paragraph(pages, page_index, y, width, height, margin, paragraph, 12, 22, false);
            }
        }
    }
    y + 6
}

fn draw_table(
    pages: &mut Vec<RasterImage>,
    page_index: &mut usize,
    mut y: i32,
    width: i32,
    height: i32,
    margin: i32,
    table: &Table,
) -> i32 {
    if table.rows.is_empty() {
        return y;
    }
    let col_count = table.rows.iter().map(|row| row.cells.len()).max().unwrap_or(1).max(1);
    let avail = width - margin * 2;
    let mut widths = vec![avail / col_count as i32; col_count];
    for row in &table.rows {
        for (index, cell) in row.cells.iter().enumerate() {
            let text = cell.paragraphs.iter().map(paragraph_plain).collect::<Vec<_>>().join(" ");
            widths[index] = widths[index].max((text.chars().count() as i32 * 7 + 18).min(avail));
        }
    }
    let total: i32 = widths.iter().sum();
    if total > avail {
        let scale = avail as f32 / total as f32;
        for width in &mut widths {
            *width = (*width as f32 * scale).max(48.0) as i32;
        }
    }
    for (row_index, row) in table.rows.iter().enumerate() {
        let mut row_height = 26;
        for cell in &row.cells {
            let text = cell.paragraphs.iter().map(paragraph_plain).collect::<Vec<_>>().join(" ");
            row_height = row_height.max(18 + ((text.chars().count() as i32 * 7) / 180) * 14);
        }
        y = ensure_room(pages, page_index, y, row_height + 2, width, height, margin);
        let page = &mut pages[*page_index];
        let mut x = margin;
        for col in 0..col_count {
            let cell_text = row
                .cells
                .get(col)
                .map(|cell| cell.paragraphs.iter().map(paragraph_plain).collect::<Vec<_>>().join(" "))
                .unwrap_or_default();
            let fill = if row_index == 0 { Rgba::rgba(235, 241, 247, 255) } else { Rgba::WHITE };
            page.fill_rect(x, y, widths[col], row_height, fill);
            page.stroke_rect(x, y, widths[col], row_height, 1, Rgba::rgba(120, 120, 120, 255));
            let lines = wrap_plain(&page, &cell_text, 11, widths[col] - 10);
            let mut ty = y + 6;
            for line in lines.iter().take(((row_height - 8) / 14).max(1) as usize) {
                page.draw_text(x + 5, ty, 11, Rgba::BLACK, line, row_index == 0);
                ty += 14;
            }
            x += widths[col];
        }
        y += row_height;
    }
    y + 8
}

fn draw_paragraph(
    pages: &mut Vec<RasterImage>,
    page_index: &mut usize,
    mut y: i32,
    width: i32,
    height: i32,
    margin: i32,
    paragraph: &Paragraph,
    default_pt: i32,
    extra_indent: i32,
    force_bold: bool,
) -> i32 {
    let left = margin + paragraph.style.margin_left_mm as i32 * 3 + extra_indent;
    let right = margin + paragraph.style.margin_right_mm as i32 * 3;
    let avail = (width - left - right).max(80);
    y += paragraph.style.margin_top_mm as i32 * 3;
    let base_color = if paragraph.text_style.color.trim().is_empty() {
        Rgba::BLACK
    } else {
        parse_hex_color(&paragraph.text_style.color, Rgba::BLACK)
    };
    let size = if paragraph.text_style.font_size_pt > 0 {
        (paragraph.text_style.font_size_pt as i32 * 4 / 3).max(default_pt)
    } else {
        default_pt * 4 / 3
    };
    let runs = paragraph_runs(paragraph, size, base_color, force_bold || paragraph.text_style.bold);
    let lines = wrap_runs(&pages[*page_index], &runs, avail);
    for line in lines {
        y = ensure_room(pages, page_index, y, line.height + 6, width, height, margin);
        let offset = match paragraph.style.alignment {
            Alignment::Center => ((avail - line.width).max(0)) / 2,
            Alignment::End => (avail - line.width).max(0),
            _ => 0,
        };
        let mut x = left + offset;
        for run in line.runs {
            if !run.text.is_empty() {
                pages[*page_index].draw_text(x, y, run.size_px, run.color, &run.text, run.bold);
                x += pages[*page_index].measure_text(&run.text, run.size_px);
            }
        }
        y += line.height + 2;
    }
    y + (paragraph.style.margin_bottom_mm.max(1) as i32 * 3)
}

fn paragraph_runs(paragraph: &Paragraph, size_px: i32, color: Rgba, force_bold: bool) -> Vec<Run> {
    let mut runs = Vec::new();
    for span in &paragraph.spans {
        match span {
            Inline::Text(text) => runs.push(Run { text: text.clone(), size_px, bold: force_bold, color }),
            Inline::Bold(text) => runs.push(Run { text: text.clone(), size_px, bold: true, color }),
            Inline::Italic(text) => runs.push(Run { text: text.clone(), size_px, bold: force_bold, color }),
            Inline::Code(text) => runs.push(Run { text: text.clone(), size_px: (size_px - 1).max(8), bold: true, color: Rgba::rgba(50, 50, 50, 255) }),
            Inline::Link { label, .. } => runs.push(Run { text: label.clone(), size_px, bold: force_bold, color: Rgba::rgba(0, 80, 180, 255) }),
            Inline::LineBreak => runs.push(Run { text: "\n".to_string(), size_px, bold: force_bold, color }),
        }
    }
    if runs.is_empty() {
        runs.push(Run { text: String::new(), size_px, bold: force_bold, color });
    }
    runs
}

fn wrap_runs(canvas: &RasterImage, runs: &[Run], max_width: i32) -> Vec<Line> {
    let mut lines = Vec::new();
    let mut current = Line::default();
    for run in runs {
        if run.text.contains('\n') {
            for (part_index, part) in run.text.split('\n').enumerate() {
                if !part.is_empty() {
                    push_tokens(canvas, &mut current, &mut lines, Run { text: part.to_string(), ..run.clone() }, max_width);
                }
                if part_index + 1 < run.text.split('\n').count() {
                    finalize_line(&mut current, &mut lines);
                }
            }
            continue;
        }
        push_tokens(canvas, &mut current, &mut lines, run.clone(), max_width);
    }
    finalize_line(&mut current, &mut lines);
    if lines.is_empty() {
        lines.push(Line { runs: Vec::new(), width: 0, height: 16 });
    }
    lines
}

fn push_tokens(canvas: &RasterImage, current: &mut Line, lines: &mut Vec<Line>, run: Run, max_width: i32) {
    for token in tokenize_preserve_spaces(&run.text) {
        let piece = Run { text: token.clone(), ..run.clone() };
        let piece_width = canvas.measure_text(&piece.text, piece.size_px);
        if current.width > 0 && current.width + piece_width > max_width && !piece.text.trim().is_empty() {
            finalize_line(current, lines);
        }
        if current.runs.is_empty() && piece.text.trim().is_empty() {
            continue;
        }
        current.height = current.height.max(piece.size_px + 4);
        current.width += piece_width;
        current.runs.push(piece);
    }
}

fn finalize_line(current: &mut Line, lines: &mut Vec<Line>) {
    while matches!(current.runs.last(), Some(run) if run.text.trim().is_empty()) {
        if let Some(run) = current.runs.pop() {
            current.width -= current.width.min(run.text.chars().count() as i32 * 6 * (run.size_px.max(8) / 8).max(1));
        }
    }
    if !current.runs.is_empty() || current.width > 0 {
        if current.height == 0 {
            current.height = 16;
        }
        lines.push(std::mem::take(current));
    }
}

fn tokenize_preserve_spaces(text: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut buf = String::new();
    let mut last_space = None;
    for ch in text.chars() {
        let is_space = ch.is_whitespace() && ch != '\n';
        match last_space {
            Some(flag) if flag == is_space => buf.push(ch),
            Some(_) => {
                out.push(std::mem::take(&mut buf));
                buf.push(ch);
                last_space = Some(is_space);
            }
            None => {
                buf.push(ch);
                last_space = Some(is_space);
            }
        }
    }
    if !buf.is_empty() {
        out.push(buf);
    }
    out
}

fn wrap_plain(canvas: &RasterImage, text: &str, size_px: i32, max_width: i32) -> Vec<String> {
    let mut lines = Vec::new();
    let mut current = String::new();
    for word in text.split_whitespace() {
        let candidate = if current.is_empty() { word.to_string() } else { format!("{} {}", current, word) };
        if !current.is_empty() && canvas.measure_text(&candidate, size_px) > max_width {
            lines.push(current);
            current = word.to_string();
        } else {
            current = candidate;
        }
    }
    if !current.is_empty() {
        lines.push(current);
    }
    if lines.is_empty() {
        lines.push(String::new());
    }
    lines
}

fn paragraph_plain(paragraph: &Paragraph) -> String {
    paragraph.spans.iter().map(|span| match span {
        Inline::Text(text) | Inline::Bold(text) | Inline::Italic(text) | Inline::Code(text) => text.clone(),
        Inline::Link { label, .. } => label.clone(),
        Inline::LineBreak => " ".to_string(),
    }).collect::<Vec<_>>().join("")
}

fn ensure_room(
    pages: &mut Vec<RasterImage>,
    page_index: &mut usize,
    y: i32,
    needed: i32,
    width: i32,
    height: i32,
    margin: i32,
) -> i32 {
    if y + needed <= height - margin {
        return y;
    }
    pages.push(RasterImage::new(width as u32, height as u32, Rgba::WHITE));
    *page_index = pages.len() - 1;
    margin
}

fn mm_to_px(mm: f32, dpi: u32) -> i32 {
    ((mm / 25.4) * dpi as f32).round() as i32
}
