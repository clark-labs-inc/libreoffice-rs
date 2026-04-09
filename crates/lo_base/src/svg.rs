//! SVG render of the first table in a `DatabaseDocument`.

use lo_core::{
    geometry::Point, svg_footer, svg_header, svg_line, svg_rect, svg_text, units::Length,
    DatabaseDocument, DbValue, Rect, Size,
};

const COL_W_PT: f32 = 120.0;
const ROW_H_PT: f32 = 22.0;
const HEADER_PT: f32 = 24.0;

fn db_value_to_string(value: &DbValue) -> String {
    match value {
        DbValue::Null => String::new(),
        DbValue::Integer(v) => v.to_string(),
        DbValue::Float(v) => v.to_string(),
        DbValue::Bool(v) => v.to_string(),
        DbValue::Text(v) => v.clone(),
    }
}

pub fn render_svg(database: &DatabaseDocument) -> String {
    let table = match database.tables.first() {
        Some(t) => t,
        None => {
            let mut svg = svg_header(Length::pt(200.0), Length::pt(60.0));
            svg.push_str(&svg_text(
                Length::pt(20.0),
                Length::pt(36.0),
                "(empty database)",
                14,
                "#666666",
                "italic",
            ));
            svg.push_str(svg_footer());
            return svg;
        }
    };

    let cols = table.columns.len() as f32;
    let rows = (table.rows.len() + 1) as f32;
    let total_w = cols * COL_W_PT + 32.0;
    let total_h = HEADER_PT + rows * ROW_H_PT + 32.0;

    let mut svg = svg_header(Length::pt(total_w), Length::pt(total_h));
    svg.push_str(&svg_text(
        Length::pt(16.0),
        Length::pt(18.0),
        &table.name,
        14,
        "#222222",
        "bold",
    ));

    svg.push_str(&svg_rect(
        Rect {
            origin: Point::new(Length::pt(16.0), Length::pt(HEADER_PT)),
            size: Size::new(Length::pt(cols * COL_W_PT), Length::pt(rows * ROW_H_PT)),
        },
        "#888888",
        Some("#ffffff"),
    ));

    // Vertical/horizontal grid
    for c in 0..=table.columns.len() {
        let x = 16.0 + c as f32 * COL_W_PT;
        svg.push_str(&svg_line(
            Length::pt(x),
            Length::pt(HEADER_PT),
            Length::pt(x),
            Length::pt(HEADER_PT + rows * ROW_H_PT),
            "#cccccc",
        ));
    }
    for r in 0..=(rows as u32) {
        let y = HEADER_PT + r as f32 * ROW_H_PT;
        svg.push_str(&svg_line(
            Length::pt(16.0),
            Length::pt(y),
            Length::pt(16.0 + cols * COL_W_PT),
            Length::pt(y),
            "#cccccc",
        ));
    }

    // Header row
    for (idx, col) in table.columns.iter().enumerate() {
        svg.push_str(&svg_text(
            Length::pt(20.0 + idx as f32 * COL_W_PT),
            Length::pt(HEADER_PT + 16.0),
            &col.name,
            12,
            "#1f4e79",
            "bold",
        ));
    }

    // Body rows
    for (row_idx, row) in table.rows.iter().enumerate() {
        for (col_idx, value) in row.iter().enumerate() {
            svg.push_str(&svg_text(
                Length::pt(20.0 + col_idx as f32 * COL_W_PT),
                Length::pt(HEADER_PT + (row_idx as f32 + 2.0) * ROW_H_PT - 4.0),
                &db_value_to_string(value),
                11,
                "#000000",
                "normal",
            ));
        }
    }

    svg.push_str(svg_footer());
    svg
}
