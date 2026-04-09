use std::collections::BTreeMap;
use std::f32::consts::PI;

use lo_core::{
    parse_hex_color, parse_xml_document, PdfFont, PdfPage, RasterImage, Rgba, XmlNode,
};
use lo_zip::ZipArchive;

const CHART_SPEC_ROW: &str = "__LO_CHART_SPEC__";
const PALETTE: [&str; 8] = [
    "#4472C4",
    "#ED7D31",
    "#A5A5A5",
    "#FFC000",
    "#5B9BD5",
    "#70AD47",
    "#264478",
    "#9E480E",
];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SeriesKind {
    Column,
    Bar,
    Line,
    Area,
    Scatter,
    Pie,
    Doughnut,
}

impl SeriesKind {
    fn as_str(self) -> &'static str {
        match self {
            Self::Column => "column",
            Self::Bar => "bar",
            Self::Line => "line",
            Self::Area => "area",
            Self::Scatter => "scatter",
            Self::Pie => "pie",
            Self::Doughnut => "doughnut",
        }
    }

    fn from_str(value: &str) -> Option<Self> {
        match value {
            "column" => Some(Self::Column),
            "bar" => Some(Self::Bar),
            "line" => Some(Self::Line),
            "area" => Some(Self::Area),
            "scatter" => Some(Self::Scatter),
            "pie" => Some(Self::Pie),
            "doughnut" => Some(Self::Doughnut),
            _ => None,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
struct ChartSeries {
    name: String,
    categories: Vec<String>,
    values: Vec<f64>,
    x_values: Vec<f64>,
    color: String,
    kind: SeriesKind,
}

#[derive(Clone, Debug, PartialEq)]
struct ChartSpec {
    frame_mm: (f32, f32, f32, f32),
    title: String,
    category_axis_title: String,
    value_axis_title: String,
    categories: Vec<String>,
    series: Vec<ChartSeries>,
    show_value_labels: bool,
    show_category_labels: bool,
}

impl ChartSpec {
    fn new(frame_mm: (f32, f32, f32, f32)) -> Self {
        Self {
            frame_mm,
            title: String::new(),
            category_axis_title: String::new(),
            value_axis_title: String::new(),
            categories: Vec::new(),
            series: Vec::new(),
            show_value_labels: false,
            show_category_labels: false,
        }
    }

    fn dominant_kind(&self) -> SeriesKind {
        self.series
            .first()
            .map(|series| series.kind)
            .unwrap_or(SeriesKind::Column)
    }

    fn type_name(&self) -> &'static str {
        let first = self.dominant_kind();
        if self.series.iter().skip(1).any(|series| series.kind != first) {
            return "combo";
        }
        first.as_str()
    }

    fn axis_categories(&self) -> Vec<String> {
        if !self.categories.is_empty() {
            return self.categories.clone();
        }
        self.series
            .iter()
            .find(|series| !series.categories.is_empty())
            .map(|series| series.categories.clone())
            .unwrap_or_default()
    }
}

fn hex_encode(input: &str) -> String {
    let mut out = String::with_capacity(input.len() * 2);
    for byte in input.as_bytes() {
        out.push_str(&format!("{byte:02x}"));
    }
    out
}

fn hex_decode(input: &str) -> Option<String> {
    if input.len() % 2 != 0 {
        return None;
    }
    let mut bytes = Vec::with_capacity(input.len() / 2);
    let mut index = 0;
    while index < input.len() {
        let byte = u8::from_str_radix(&input[index..index + 2], 16).ok()?;
        bytes.push(byte);
        index += 2;
    }
    String::from_utf8(bytes).ok()
}

fn encode_strings(values: &[String]) -> String {
    values
        .iter()
        .map(|value| hex_encode(value))
        .collect::<Vec<_>>()
        .join(",")
}

fn decode_strings(input: &str) -> Option<Vec<String>> {
    if input.trim().is_empty() {
        return Some(Vec::new());
    }
    input
        .split(',')
        .map(hex_decode)
        .collect::<Option<Vec<_>>>()
}

fn encode_floats(values: &[f64]) -> String {
    values
        .iter()
        .map(|value| {
            if (value - value.round()).abs() < 1e-9 {
                format!("{}", value.round() as i64)
            } else {
                let text = format!("{value:.6}");
                text.trim_end_matches('0').trim_end_matches('.').to_string()
            }
        })
        .collect::<Vec<_>>()
        .join(",")
}

fn decode_floats(input: &str) -> Option<Vec<f64>> {
    if input.trim().is_empty() {
        return Some(Vec::new());
    }
    input
        .split(',')
        .map(|value| value.parse::<f64>().ok())
        .collect::<Option<Vec<_>>>()
}

fn encode_chart_spec(spec: &ChartSpec) -> String {
    let mut lines = vec![
        format!(
            "frame={:.3},{:.3},{:.3},{:.3}",
            spec.frame_mm.0, spec.frame_mm.1, spec.frame_mm.2, spec.frame_mm.3
        ),
        format!("title={}", hex_encode(&spec.title)),
        format!("cat_title={}", hex_encode(&spec.category_axis_title)),
        format!("val_title={}", hex_encode(&spec.value_axis_title)),
        format!("show_values={}", if spec.show_value_labels { 1 } else { 0 }),
        format!("show_categories={}", if spec.show_category_labels { 1 } else { 0 }),
        format!("categories={}", encode_strings(&spec.categories)),
    ];
    for series in &spec.series {
        lines.push(format!(
            "series={}|{}|{}|{}|{}|{}",
            series.kind.as_str(),
            hex_encode(&series.name),
            series.color.trim().trim_start_matches('#'),
            encode_floats(&series.values),
            encode_strings(&series.categories),
            encode_floats(&series.x_values),
        ));
    }
    lines.join("\n")
}

fn decode_chart_spec(input: &str) -> Option<ChartSpec> {
    let mut spec = ChartSpec::new((0.0, 0.0, 120.0, 80.0));
    for line in input.lines() {
        let Some((key, value)) = line.split_once('=') else {
            continue;
        };
        match key {
            "frame" => {
                let parts = value
                    .split(',')
                    .map(|part| part.parse::<f32>().ok())
                    .collect::<Option<Vec<_>>>()?;
                if parts.len() == 4 {
                    spec.frame_mm = (parts[0], parts[1], parts[2], parts[3]);
                }
            }
            "title" => spec.title = hex_decode(value).unwrap_or_default(),
            "cat_title" => spec.category_axis_title = hex_decode(value).unwrap_or_default(),
            "val_title" => spec.value_axis_title = hex_decode(value).unwrap_or_default(),
            "show_values" => spec.show_value_labels = value == "1",
            "show_categories" => spec.show_category_labels = value == "1",
            "categories" => spec.categories = decode_strings(value).unwrap_or_default(),
            "series" => {
                let parts = value.split('|').collect::<Vec<_>>();
                if parts.len() != 6 {
                    continue;
                }
                let kind = SeriesKind::from_str(parts[0])?;
                spec.series.push(ChartSeries {
                    kind,
                    name: hex_decode(parts[1]).unwrap_or_default(),
                    color: if parts[2].trim().is_empty() {
                        String::new()
                    } else {
                        format!("#{}", parts[2].trim().trim_start_matches('#'))
                    },
                    values: decode_floats(parts[3]).unwrap_or_default(),
                    categories: decode_strings(parts[4]).unwrap_or_default(),
                    x_values: decode_floats(parts[5]).unwrap_or_default(),
                });
            }
            _ => {}
        }
    }
    Some(spec)
}

fn spec_to_row(spec: &ChartSpec) -> Vec<String> {
    vec![CHART_SPEC_ROW.to_string(), encode_chart_spec(spec)]
}

fn row_to_spec(row: &[String]) -> Option<ChartSpec> {
    if row.first()?.as_str() != CHART_SPEC_ROW {
        return None;
    }
    decode_chart_spec(row.get(1)?)
}

pub(crate) fn chart_row_title(row: &[String]) -> Option<String> {
    let spec = row_to_spec(row)?;
    if spec.title.trim().is_empty() {
        None
    } else {
        Some(spec.title)
    }
}

pub(crate) fn graphic_frame_has_chart(node: &XmlNode) -> bool {
    extract_chart_rel_id(node).is_some()
}

pub(crate) fn load_pptx_chart_rows(
    zip: &ZipArchive,
    slide_root: &XmlNode,
    slide_rels: &BTreeMap<String, String>,
) -> lo_core::Result<Vec<Vec<String>>> {
    let mut refs = Vec::new();
    collect_chart_refs(slide_root, slide_rels, &mut refs);
    let mut out = Vec::new();
    for (chart_index, (target, frame_mm)) in refs.into_iter().enumerate() {
        if !zip.contains(&target) {
            continue;
        }
        let chart = parse_xml_document(&zip.read_string(&target)?)?;
        let spec = parse_chart_spec(&chart, frame_mm, chart_index);
        if !spec.series.is_empty()
            || !spec.title.trim().is_empty()
            || !spec.category_axis_title.trim().is_empty()
            || !spec.value_axis_title.trim().is_empty()
        {
            out.push(spec_to_row(&spec));
        }
    }
    Ok(out)
}

fn collect_chart_refs(
    node: &XmlNode,
    slide_rels: &BTreeMap<String, String>,
    out: &mut Vec<(String, (f32, f32, f32, f32))>,
) {
    if node.local_name() == "graphicFrame" {
        if let Some(rel_id) = extract_chart_rel_id(node) {
            if let Some(target) = slide_rels.get(&rel_id) {
                out.push((target.clone(), parse_graphic_frame_transform(node)));
            }
        }
    }
    for child in &node.children {
        collect_chart_refs(child, slide_rels, out);
    }
}

fn extract_chart_rel_id(node: &XmlNode) -> Option<String> {
    if node.local_name() == "chart" {
        if let Some(rel_id) = node.attr("r:id").or_else(|| node.attr("id")) {
            return Some(rel_id.to_string());
        }
    }
    for child in &node.children {
        if let Some(rel_id) = extract_chart_rel_id(child) {
            return Some(rel_id);
        }
    }
    None
}

fn parse_graphic_frame_transform(node: &XmlNode) -> (f32, f32, f32, f32) {
    let xfrm = node.child("xfrm");
    let off = xfrm.and_then(|n| n.child("off"));
    let ext = xfrm.and_then(|n| n.child("ext"));
    let to_mm = |value: f32| value / 36_000.0;
    let x = off
        .and_then(|n| n.attr("x"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(to_mm)
        .unwrap_or(20.0);
    let y = off
        .and_then(|n| n.attr("y"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(to_mm)
        .unwrap_or(20.0);
    let w = ext
        .and_then(|n| n.attr("cx"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(to_mm)
        .unwrap_or(160.0);
    let h = ext
        .and_then(|n| n.attr("cy"))
        .and_then(|v| v.parse::<f32>().ok())
        .map(to_mm)
        .unwrap_or(95.0);
    (x, y, w, h)
}

fn parse_chart_spec(chart_root: &XmlNode, frame_mm: (f32, f32, f32, f32), palette_offset: usize) -> ChartSpec {
    let chart = chart_root.child("chart").unwrap_or(chart_root);
    let plot_area = chart.child("plotArea").unwrap_or(chart);

    let mut spec = ChartSpec::new(frame_mm);
    spec.title = extract_text(chart.child("title"));
    spec.category_axis_title = first_non_empty(&[
        axis_title(plot_area, &["catAx"]),
        axis_title(plot_area, &["dateAx"]),
        axis_title(plot_area, &["serAx"]),
    ]);
    spec.value_axis_title = first_non_empty(&[axis_title(plot_area, &["valAx"])]);
    spec.show_value_labels = contains_flag(chart, "showVal");
    spec.show_category_labels = contains_flag(chart, "showCatName") || contains_flag(chart, "showLegendKey");

    for child in &plot_area.children {
        match child.local_name() {
            "barChart" => parse_bar_chart(child, &mut spec, palette_offset),
            "lineChart" => parse_simple_series(child, &mut spec, palette_offset, SeriesKind::Line),
            "areaChart" => parse_simple_series(child, &mut spec, palette_offset, SeriesKind::Area),
            "scatterChart" => parse_scatter_chart(child, &mut spec, palette_offset),
            "pieChart" | "ofPieChart" => parse_simple_series(child, &mut spec, palette_offset, SeriesKind::Pie),
            "doughnutChart" => parse_simple_series(child, &mut spec, palette_offset, SeriesKind::Doughnut),
            "radarChart" => parse_simple_series(child, &mut spec, palette_offset, SeriesKind::Line),
            "bubbleChart" => parse_scatter_chart(child, &mut spec, palette_offset),
            "stockChart" => parse_simple_series(child, &mut spec, palette_offset, SeriesKind::Line),
            _ => {}
        }
    }

    if spec.categories.is_empty() {
        spec.categories = spec.axis_categories();
    }

    if spec.title.trim().is_empty() {
        if let Some(name) = spec.series.iter().map(|series| series.name.trim()).find(|name| !name.is_empty()) {
            spec.title = name.to_string();
        }
    }

    spec
}

fn parse_bar_chart(node: &XmlNode, spec: &mut ChartSpec, palette_offset: usize) {
    let bar_dir = node
        .child("barDir")
        .and_then(|bar_dir| bar_dir.attr("val"))
        .unwrap_or("col");
    let kind = if bar_dir.eq_ignore_ascii_case("bar") {
        SeriesKind::Bar
    } else {
        SeriesKind::Column
    };
    parse_simple_series(node, spec, palette_offset, kind);
}

fn parse_simple_series(node: &XmlNode, spec: &mut ChartSpec, palette_offset: usize, kind: SeriesKind) {
    let chart_show_values = contains_flag(node, "showVal");
    let chart_show_categories = contains_flag(node, "showCatName") || contains_flag(node, "showLegendKey");
    if chart_show_values {
        spec.show_value_labels = true;
    }
    if chart_show_categories {
        spec.show_category_labels = true;
    }
    for (series_index, ser) in node.children_named("ser").enumerate() {
        let name = series_name(ser);
        let categories = parse_category_values(ser.child("cat"));
        let values = parse_numeric_branch(ser.child("val"));
        let color = series_color(ser, palette_offset + spec.series.len() + series_index);
        if spec.categories.is_empty() && !categories.is_empty() {
            spec.categories = categories.clone();
        }
        spec.series.push(ChartSeries {
            name,
            categories,
            values,
            x_values: Vec::new(),
            color,
            kind,
        });
        if contains_flag(ser, "showVal") {
            spec.show_value_labels = true;
        }
        if contains_flag(ser, "showCatName") || contains_flag(ser, "showLegendKey") {
            spec.show_category_labels = true;
        }
    }
}

fn parse_scatter_chart(node: &XmlNode, spec: &mut ChartSpec, palette_offset: usize) {
    if contains_flag(node, "showVal") {
        spec.show_value_labels = true;
    }
    for (series_index, ser) in node.children_named("ser").enumerate() {
        let name = series_name(ser);
        let x_values = parse_numeric_branch(ser.child("xVal"));
        let values = parse_numeric_branch(ser.child("yVal"));
        let color = series_color(ser, palette_offset + spec.series.len() + series_index);
        spec.series.push(ChartSeries {
            name,
            categories: x_values.iter().map(|value| format_tick(*value)).collect(),
            values,
            x_values,
            color,
            kind: SeriesKind::Scatter,
        });
        if contains_flag(ser, "showVal") {
            spec.show_value_labels = true;
        }
    }
}

fn axis_title(plot_area: &XmlNode, kinds: &[&str]) -> String {
    for kind in kinds {
        for axis in plot_area.children_named(kind) {
            let title = extract_text(axis.child("title"));
            if !title.trim().is_empty() {
                return title;
            }
        }
    }
    String::new()
}

fn contains_flag(node: &XmlNode, local_name: &str) -> bool {
    if node.local_name() == local_name {
        if let Some(value) = node.attr("val") {
            return value == "1" || value.eq_ignore_ascii_case("true");
        }
        return true;
    }
    for child in &node.children {
        if contains_flag(child, local_name) {
            return true;
        }
    }
    false
}

fn first_non_empty(values: &[String]) -> String {
    values
        .iter()
        .find(|value| !value.trim().is_empty())
        .cloned()
        .unwrap_or_default()
}

fn extract_text(node: Option<&XmlNode>) -> String {
    fn walk(node: &XmlNode, out: &mut Vec<String>) {
        let local = node.local_name();
        if local == "t" || local == "v" {
            let text = node.text_content().trim().to_string();
            if !text.is_empty() {
                out.push(text);
            }
        }
        for child in &node.children {
            walk(child, out);
        }
    }
    let Some(node) = node else {
        return String::new();
    };
    let mut parts = Vec::new();
    walk(node, &mut parts);
    normalize_whitespace(&parts.join(" "))
}

fn normalize_whitespace(input: &str) -> String {
    input.split_whitespace().collect::<Vec<_>>().join(" ")
}

fn series_name(ser: &XmlNode) -> String {
    let text = extract_text(ser.child("tx"));
    if text.is_empty() {
        extract_text(ser.child("title"))
    } else {
        text
    }
}

fn parse_category_values(branch: Option<&XmlNode>) -> Vec<String> {
    let Some(branch) = branch else {
        return Vec::new();
    };
    if let Some(cache) = find_cache_node(branch, &["strCache", "strLit", "multiLvlStrCache"]) {
        let mut values = cache_strings(cache);
        if !values.is_empty() {
            return values.drain(..).map(|value| normalize_whitespace(&value)).collect();
        }
    }
    if let Some(cache) = find_cache_node(branch, &["numCache", "numLit"]) {
        if is_date_format_cache(cache) {
            return Vec::new();
        }
        return cache_strings(cache)
            .into_iter()
            .filter_map(|value| value.parse::<f64>().ok())
            .map(format_tick)
            .collect();
    }
    Vec::new()
}

fn parse_numeric_branch(branch: Option<&XmlNode>) -> Vec<f64> {
    let Some(branch) = branch else {
        return Vec::new();
    };
    if let Some(cache) = find_cache_node(branch, &["numCache", "numLit"]) {
        if is_date_format_cache(cache) {
            // Date-formatted serial numbers (Excel epoch) should not be
            // rendered as raw integers — LO renders them as date strings,
            // which we cannot reproduce, so skip the series entirely.
            return Vec::new();
        }
        return cache_strings(cache)
            .into_iter()
            .filter_map(|value| value.parse::<f64>().ok())
            .collect();
    }
    Vec::new()
}

fn is_date_format_cache(node: &XmlNode) -> bool {
    if let Some(fc) = node.child("formatCode") {
        let code = fc.text_content().to_ascii_lowercase();
        if code.contains("yy") || code.contains("mmm") || code.contains("dd")
            || code.contains("/yy") || code.contains("h:mm") || code.contains("d-m")
        {
            return true;
        }
    }
    false
}

fn find_cache_node<'a>(node: &'a XmlNode, names: &[&str]) -> Option<&'a XmlNode> {
    if names.iter().any(|name| *name == node.local_name()) {
        return Some(node);
    }
    for child in &node.children {
        if let Some(found) = find_cache_node(child, names) {
            return Some(found);
        }
    }
    None
}

fn cache_strings(node: &XmlNode) -> Vec<String> {
    let mut points = Vec::new();
    collect_pt_values(node, &mut points);
    points.sort_by_key(|(index, _)| *index);
    points.into_iter().map(|(_, value)| value).collect()
}

fn collect_pt_values(node: &XmlNode, out: &mut Vec<(usize, String)>) {
    if node.local_name() == "pt" {
        let index = node
            .attr("idx")
            .and_then(|value| value.parse::<usize>().ok())
            .unwrap_or(out.len());
        let value = extract_text(Some(node));
        if !value.trim().is_empty() {
            out.push((index, value));
        }
        return;
    }
    for child in &node.children {
        collect_pt_values(child, out);
    }
}

fn series_color(ser: &XmlNode, palette_index: usize) -> String {
    fn walk(node: &XmlNode) -> Option<String> {
        if node.local_name() == "srgbClr" {
            if let Some(value) = node.attr("val") {
                return Some(format!("#{}", value.trim().trim_start_matches('#')));
            }
        }
        for child in &node.children {
            if let Some(value) = walk(child) {
                return Some(value);
            }
        }
        None
    }
    walk(ser).unwrap_or_else(|| PALETTE[palette_index % PALETTE.len()].to_string())
}

fn value_range(spec: &ChartSpec) -> Option<(f64, f64)> {
    let mut min = f64::INFINITY;
    let mut max = f64::NEG_INFINITY;
    for series in &spec.series {
        for value in &series.values {
            if *value < min {
                min = *value;
            }
            if *value > max {
                max = *value;
            }
        }
    }
    if !min.is_finite() || !max.is_finite() {
        return None;
    }
    if min > 0.0 {
        min = 0.0;
    }
    if (max - min).abs() < 1e-9 {
        max = min + 1.0;
    }
    Some((min, max))
}

fn x_numeric_range(spec: &ChartSpec) -> Option<(f64, f64)> {
    let mut min = f64::INFINITY;
    let mut max = f64::NEG_INFINITY;
    let mut found = false;
    for series in &spec.series {
        for value in &series.x_values {
            found = true;
            if *value < min {
                min = *value;
            }
            if *value > max {
                max = *value;
            }
        }
    }
    if !found {
        return None;
    }
    if (max - min).abs() < 1e-9 {
        max = min + 1.0;
    }
    Some((min, max))
}

fn axis_ticks(min: f64, max: f64) -> Vec<f64> {
    let span = (max - min).max(1.0);
    let raw_step = span / 6.0;
    let step = nice_step(raw_step);
    let start = (min / step).floor() * step;
    // LibreOffice's chart engine extends the value axis one step past
    // the rounded-up max so the last data point doesn't sit on the top
    // gridline; mirror that so synthetic ticks line up with LO's output.
    let end = (max / step).ceil() * step + step;
    let mut out = Vec::new();
    let mut tick = start;
    let mut guard = 0;
    while tick <= end + step * 0.5 && guard < 32 {
        out.push(tick);
        tick += step;
        guard += 1;
    }
    out
}

fn nice_step(raw: f64) -> f64 {
    if raw <= 0.0 {
        return 1.0;
    }
    let exp = raw.log10().floor();
    let base = 10f64.powf(exp);
    let frac = raw / base;
    let nice = if frac < 1.5 {
        1.0
    } else if frac < 3.5 {
        2.0
    } else if frac < 7.5 {
        5.0
    } else {
        10.0
    };
    nice * base
}

fn format_tick(value: f64) -> String {
    if (value - value.round()).abs() < 1e-9 {
        format!("{}", value.round() as i64)
    } else {
        let text = format!("{value:.4}");
        text.trim_end_matches('0').trim_end_matches('.').to_string()
    }
}

fn mm_to_pt(mm: f32) -> f32 {
    mm * 72.0 / 25.4
}

fn mm_to_px(mm: f32, dpi: u32) -> i32 {
    ((mm / 25.4) * dpi as f32).round() as i32
}

fn pdf_text_width(text: &str, size: f32) -> f32 {
    text.chars().count() as f32 * size * 0.53
}

fn draw_pdf_centered(page: &mut PdfPage, x_center: f32, y: f32, size: f32, font: PdfFont, text: &str, color: (f32, f32, f32)) {
    let width = pdf_text_width(text, size);
    page.text_rgb(
        x_center - width / 2.0,
        y,
        size,
        font,
        text,
        color.0,
        color.1,
        color.2,
    );
}

fn draw_pdf_right(page: &mut PdfPage, x_right: f32, y: f32, size: f32, font: PdfFont, text: &str, color: (f32, f32, f32)) {
    let width = pdf_text_width(text, size);
    page.text_rgb(
        x_right - width,
        y,
        size,
        font,
        text,
        color.0,
        color.1,
        color.2,
    );
}

fn render_pdf_polygon(page: &mut PdfPage, points: &[(f32, f32)], fill: (f32, f32, f32), stroke: (f32, f32, f32)) {
    if points.len() < 3 {
        return;
    }
    let mut command = format!(
        "{:.3} {:.3} {:.3} rg\n{:.3} {:.3} {:.3} RG\n{:.2} {:.2} m\n",
        fill.0, fill.1, fill.2, stroke.0, stroke.1, stroke.2, points[0].0, points[0].1
    );
    for point in &points[1..] {
        command.push_str(&format!("{:.2} {:.2} l\n", point.0, point.1));
    }
    command.push_str("h B\n0 0 0 rg\n0 0 0 RG\n");
    page.raw(&command);
}

fn render_pdf_polyline(page: &mut PdfPage, points: &[(f32, f32)], stroke: (f32, f32, f32)) {
    if points.len() < 2 {
        return;
    }
    let mut command = format!(
        "{:.3} {:.3} {:.3} RG\n{:.2} {:.2} m\n",
        stroke.0, stroke.1, stroke.2, points[0].0, points[0].1
    );
    for point in &points[1..] {
        command.push_str(&format!("{:.2} {:.2} l\n", point.0, point.1));
    }
    command.push_str("S\n0 0 0 RG\n");
    page.raw(&command);
}

fn wedge_points(center: (f32, f32), radius: f32, inner_radius: f32, start: f32, end: f32) -> Vec<(f32, f32)> {
    let steps = (((end - start).abs() / (PI / 18.0)).ceil() as usize).max(8);
    let mut points = Vec::with_capacity(steps * 2 + 4);
    if inner_radius <= 0.0 {
        points.push(center);
    }
    for index in 0..=steps {
        let angle = start + (end - start) * (index as f32 / steps as f32);
        points.push((center.0 + radius * angle.cos(), center.1 + radius * angle.sin()));
    }
    if inner_radius > 0.0 {
        for index in (0..=steps).rev() {
            let angle = start + (end - start) * (index as f32 / steps as f32);
            points.push((
                center.0 + inner_radius * angle.cos(),
                center.1 + inner_radius * angle.sin(),
            ));
        }
    }
    points
}

fn fill_raster_polygon(image: &mut RasterImage, points: &[(i32, i32)], color: Rgba) {
    if points.len() < 3 {
        return;
    }
    let min_y = points.iter().map(|(_, y)| *y).min().unwrap_or(0).max(0);
    let max_y = points
        .iter()
        .map(|(_, y)| *y)
        .max()
        .unwrap_or(0)
        .min(image.height as i32 - 1);
    for y in min_y..=max_y {
        let mut xs = Vec::new();
        for index in 0..points.len() {
            let (x1, y1) = points[index];
            let (x2, y2) = points[(index + 1) % points.len()];
            if y1 == y2 {
                continue;
            }
            let (ax, ay, bx, by) = if y1 < y2 {
                (x1, y1, x2, y2)
            } else {
                (x2, y2, x1, y1)
            };
            if y < ay || y >= by {
                continue;
            }
            let t = (y - ay) as f32 / (by - ay) as f32;
            xs.push(ax as f32 + (bx - ax) as f32 * t);
        }
        xs.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
        for pair in xs.chunks(2) {
            if let [left, right] = pair {
                let x0 = left.floor() as i32;
                let x1 = right.ceil() as i32;
                for x in x0..x1 {
                    image.blend_pixel(x, y, color);
                }
            }
        }
    }
}

fn stroke_raster_polygon(image: &mut RasterImage, points: &[(i32, i32)], thickness: i32, color: Rgba) {
    if points.len() < 2 {
        return;
    }
    for index in 0..points.len() {
        let (x1, y1) = points[index];
        let (x2, y2) = points[(index + 1) % points.len()];
        image.draw_line(x1, y1, x2, y2, thickness, color);
    }
}

fn chart_mode(spec: &ChartSpec) -> SeriesKind {
    if spec.series.iter().any(|series| series.kind == SeriesKind::Doughnut) {
        return SeriesKind::Doughnut;
    }
    if spec.series.iter().any(|series| series.kind == SeriesKind::Pie) {
        return SeriesKind::Pie;
    }
    if spec.series.iter().any(|series| series.kind == SeriesKind::Bar) {
        return SeriesKind::Bar;
    }
    if spec.series.iter().any(|series| series.kind == SeriesKind::Scatter) {
        return SeriesKind::Scatter;
    }
    SeriesKind::Column
}

pub(crate) fn render_chart_rows_pdf(page: &mut PdfPage, slide_h: f32, slide_w: f32, rows: &[Vec<String>]) {
    let mut fallback = Vec::new();
    for row in rows {
        if let Some(spec) = row_to_spec(row) {
            render_chart_pdf(page, slide_h, &spec);
        } else {
            fallback.push(row.clone());
        }
    }
    if !fallback.is_empty() {
        render_plain_rows_pdf(page, slide_w, slide_h, &fallback);
    }
}

pub(crate) fn render_chart_rows_raster(image: &mut RasterImage, dpi: u32, rows: &[Vec<String>]) {
    let mut fallback = Vec::new();
    for row in rows {
        if let Some(spec) = row_to_spec(row) {
            render_chart_raster(image, dpi, &spec);
        } else {
            fallback.push(row.clone());
        }
    }
    if !fallback.is_empty() {
        render_plain_rows_raster(image, &fallback);
    }
}

pub(crate) fn append_chart_rows_markdown(out: &mut String, rows: &[Vec<String>]) {
    for row in rows {
        if let Some(spec) = row_to_spec(row) {
            out.push_str("- ");
            out.push_str(&format!("[chart: {}]", spec.type_name()));
            if !spec.title.trim().is_empty() {
                out.push(' ');
                out.push_str(&spec.title);
            }
            out.push('\n');
            if !spec.category_axis_title.trim().is_empty() {
                out.push_str("  - X axis: ");
                out.push_str(&spec.category_axis_title);
                out.push('\n');
            }
            if !spec.value_axis_title.trim().is_empty() {
                out.push_str("  - Y axis: ");
                out.push_str(&spec.value_axis_title);
                out.push('\n');
            }
            let categories = spec.axis_categories();
            if !categories.is_empty() {
                out.push_str("  - Categories: ");
                out.push_str(&categories.join(", "));
                out.push('\n');
            }
            for series in &spec.series {
                out.push_str("  - Series ");
                if series.name.trim().is_empty() {
                    out.push_str(series.kind.as_str());
                } else {
                    out.push_str(&series.name);
                }
                out.push_str(": ");
                if series.kind == SeriesKind::Scatter && !series.x_values.is_empty() {
                    let pairs = series
                        .x_values
                        .iter()
                        .zip(series.values.iter())
                        .map(|(x, y)| format!("({}, {})", format_tick(*x), format_tick(*y)))
                        .collect::<Vec<_>>();
                    out.push_str(&pairs.join(", "));
                } else {
                    out.push_str(
                        &series
                            .values
                            .iter()
                            .map(|value| format_tick(*value))
                            .collect::<Vec<_>>()
                            .join(", "),
                    );
                }
                out.push('\n');
            }
            continue;
        }
        if row.is_empty() {
            continue;
        }
        out.push_str("- ");
        out.push_str(&row.join(" "));
        out.push('\n');
    }
}

fn render_plain_rows_pdf(page: &mut PdfPage, slide_w: f32, slide_h: f32, rows: &[Vec<String>]) {
    let margin = 24.0;
    let mut y = slide_h - 36.0;
    for row in rows {
        let mut cursor = margin;
        for token in row {
            if cursor > slide_w - margin {
                cursor = margin;
                y -= 14.0;
            }
            if y < 24.0 {
                return;
            }
            page.text_rgb(cursor, y, 11.0, PdfFont::Helvetica, token, 0.25, 0.25, 0.25);
            cursor += pdf_text_width(token, 11.0) + 18.0;
        }
        y -= 14.0;
    }
}

fn render_plain_rows_raster(image: &mut RasterImage, rows: &[Vec<String>]) {
    let mut y = 20;
    for row in rows {
        let mut x = 10;
        for token in row {
            image.draw_text(x, y, 12, Rgba::rgba(60, 60, 60, 255), token, false);
            x += image.measure_text(token, 12) + 14;
        }
        y += 16;
        if y > image.height as i32 - 18 {
            return;
        }
    }
}

fn render_chart_pdf(page: &mut PdfPage, slide_h: f32, spec: &ChartSpec) {
    let x = mm_to_pt(spec.frame_mm.0);
    let y_top = mm_to_pt(spec.frame_mm.1);
    let w = mm_to_pt(spec.frame_mm.2.max(40.0));
    let h = mm_to_pt(spec.frame_mm.3.max(30.0));
    let y = slide_h - y_top - h;
    page.rect_fill_stroke_rgb(x, y, w, h, (1.0, 1.0, 1.0), (0.72, 0.72, 0.72));

    let title_h = if spec.title.trim().is_empty() { 0.0 } else { 16.0 };
    let legend_w = legend_width_pt(spec);
    let mode = chart_mode(spec);
    match mode {
        SeriesKind::Pie | SeriesKind::Doughnut => render_pdf_pie_like(page, x, y, w, h, title_h, legend_w, spec, mode == SeriesKind::Doughnut),
        SeriesKind::Bar => render_pdf_bar(page, x, y, w, h, title_h, legend_w, spec),
        SeriesKind::Scatter => render_pdf_scatter(page, x, y, w, h, title_h, legend_w, spec),
        _ => render_pdf_cartesian(page, x, y, w, h, title_h, legend_w, spec),
    }
}

fn render_chart_raster(image: &mut RasterImage, dpi: u32, spec: &ChartSpec) {
    let x = mm_to_px(spec.frame_mm.0, dpi);
    let y = mm_to_px(spec.frame_mm.1, dpi);
    let w = mm_to_px(spec.frame_mm.2.max(40.0), dpi).max(80);
    let h = mm_to_px(spec.frame_mm.3.max(30.0), dpi).max(60);
    image.fill_rect(x, y, w, h, Rgba::WHITE);
    image.stroke_rect(x, y, w, h, 1, Rgba::rgba(180, 180, 180, 255));

    let title_h = if spec.title.trim().is_empty() { 0 } else { 18 };
    let legend_w = legend_width_px(spec, dpi);
    let mode = chart_mode(spec);
    match mode {
        SeriesKind::Pie | SeriesKind::Doughnut => render_raster_pie_like(image, x, y, w, h, title_h, legend_w, spec, mode == SeriesKind::Doughnut),
        SeriesKind::Bar => render_raster_bar(image, x, y, w, h, title_h, legend_w, spec),
        SeriesKind::Scatter => render_raster_scatter(image, x, y, w, h, title_h, legend_w, spec),
        _ => render_raster_cartesian(image, x, y, w, h, title_h, legend_w, spec),
    }
}

fn legend_width_pt(spec: &ChartSpec) -> f32 {
    let entries = legend_entries(spec);
    if entries.is_empty() {
        return 0.0;
    }
    (spec.frame_mm.2 * 72.0 / 25.4 * 0.24).clamp(64.0, 120.0)
}

fn legend_width_px(spec: &ChartSpec, dpi: u32) -> i32 {
    let entries = legend_entries(spec);
    if entries.is_empty() {
        return 0;
    }
    mm_to_px(spec.frame_mm.2 * 0.24, dpi).clamp(70, 140)
}

fn legend_entries(spec: &ChartSpec) -> Vec<(String, String)> {
    let mode = chart_mode(spec);
    if matches!(mode, SeriesKind::Pie | SeriesKind::Doughnut) {
        let categories = spec.axis_categories();
        let series = spec.series.first();
        return categories
            .into_iter()
            .enumerate()
            .map(|(index, category)| {
                let color = series
                    .map(|series| slice_color(index, &series.color))
                    .unwrap_or_else(|| PALETTE[index % PALETTE.len()].to_string());
                (category, color)
            })
            .collect();
    }
    spec.series
        .iter()
        .enumerate()
        .map(|(index, series)| {
            let label = if series.name.trim().is_empty() {
                format!("{} {}", series.kind.as_str(), index + 1)
            } else {
                series.name.clone()
            };
            (label, series.color.clone())
        })
        .collect()
}

fn render_pdf_legend(page: &mut PdfPage, x: f32, y: f32, _w: f32, h: f32, entries: &[(String, String)]) {
    let mut cursor_y = y + h - 16.0;
    for (label, color) in entries {
        if cursor_y < y + 10.0 {
            break;
        }
        let rgb = pdf_color(color);
        page.rect_fill_stroke_rgb(x, cursor_y - 3.0, 8.0, 8.0, rgb, (0.3, 0.3, 0.3));
        page.text_rgb(x + 12.0, cursor_y, 9.5, PdfFont::Helvetica, label, 0.2, 0.2, 0.2);
        cursor_y -= 12.0;
    }
}

fn render_raster_legend(image: &mut RasterImage, x: i32, y: i32, entries: &[(String, String)]) {
    let mut cursor_y = y;
    for (label, color) in entries {
        if cursor_y > image.height as i32 - 12 {
            break;
        }
        let swatch = parse_hex_color(color, Rgba::rgba(80, 80, 80, 255));
        image.fill_rect(x, cursor_y - 8, 8, 8, swatch);
        image.stroke_rect(x, cursor_y - 8, 8, 8, 1, Rgba::rgba(70, 70, 70, 255));
        image.draw_text(x + 12, cursor_y - 8, 10, Rgba::rgba(60, 60, 60, 255), label, false);
        cursor_y += 12;
    }
}

fn render_pdf_cartesian(page: &mut PdfPage, x: f32, y: f32, w: f32, h: f32, title_h: f32, legend_w: f32, spec: &ChartSpec) {
    let top_pad = if title_h > 0.0 { title_h + 6.0 } else { 8.0 };
    let left_pad = 36.0;
    let bottom_pad = if spec.category_axis_title.trim().is_empty() { 28.0 } else { 40.0 };
    let right_pad = if legend_w > 0.0 { legend_w + 20.0 } else { 12.0 };
    let plot_x = x + left_pad;
    let plot_y = y + bottom_pad;
    let plot_w = (w - left_pad - right_pad).max(50.0);
    let plot_h = (h - top_pad - bottom_pad).max(30.0);
    page.rect_fill_rgb(plot_x, plot_y, plot_w, plot_h, 0.985, 0.985, 0.985);
    page.rect_stroke_rgb(plot_x, plot_y, plot_w, plot_h, 0.80, 0.80, 0.80);
    if title_h > 0.0 {
        let baseline = y + h - title_h + 2.0;
        draw_pdf_centered(page, x + w / 2.0, baseline, 11.5, PdfFont::HelveticaBold, &spec.title, (0.1, 0.1, 0.1));
    }
    let Some((min_val, max_val)) = value_range(spec) else {
        return;
    };
    let ticks = axis_ticks(min_val, max_val);
    let baseline = if min_val <= 0.0 && max_val >= 0.0 {
        0.0
    } else {
        min_val
    };
    for tick in &ticks {
        let ratio = (*tick - min_val) / (max_val - min_val);
        let yy = plot_y + plot_h * ratio as f32;
        page.line_rgb(plot_x, yy, plot_x + plot_w, yy, 0.88, 0.88, 0.88);
        draw_pdf_right(page, plot_x - 6.0, yy - 3.5, 8.5, PdfFont::Helvetica, &format_tick(*tick), (0.32, 0.32, 0.32));
    }
    page.line_rgb(plot_x, plot_y, plot_x, plot_y + plot_h, 0.35, 0.35, 0.35);
    let zero_ratio = (baseline - min_val) / (max_val - min_val);
    let zero_y = plot_y + plot_h * zero_ratio as f32;
    page.line_rgb(plot_x, zero_y, plot_x + plot_w, zero_y, 0.35, 0.35, 0.35);

    let categories = spec.axis_categories();
    let category_count = categories.len().max(spec.series.iter().map(|series| series.values.len()).max().unwrap_or(0));
    if category_count == 0 {
        return;
    }
    let cluster_w = plot_w / category_count as f32;
    for (index, category) in categories.iter().enumerate() {
        let center_x = plot_x + cluster_w * (index as f32 + 0.5);
        draw_pdf_centered(page, center_x, plot_y - 14.0, 8.5, PdfFont::Helvetica, category, (0.25, 0.25, 0.25));
    }
    if !spec.category_axis_title.trim().is_empty() {
        draw_pdf_centered(page, plot_x + plot_w / 2.0, y + 8.0, 9.5, PdfFont::HelveticaOblique, &spec.category_axis_title, (0.2, 0.2, 0.2));
    }
    if !spec.value_axis_title.trim().is_empty() {
        page.text_rgb(plot_x, y + h - top_pad + 1.0, 9.0, PdfFont::HelveticaOblique, &spec.value_axis_title, 0.2, 0.2, 0.2);
    }

    let bar_series = spec
        .series
        .iter()
        .filter(|series| matches!(series.kind, SeriesKind::Column))
        .count()
        .max(1);
    let mut bar_slot = 0usize;
    for series in &spec.series {
        match series.kind {
            SeriesKind::Column => {
                let bar_area = cluster_w * 0.76;
                let bar_w = (bar_area / bar_series as f32).max(4.0);
                let inset = (cluster_w - bar_area) / 2.0;
                let rgb = pdf_color(&series.color);
                for (index, value) in series.values.iter().enumerate() {
                    let x0 = plot_x + cluster_w * index as f32 + inset + bar_w * bar_slot as f32;
                    let y0 = plot_y + plot_h * ((*value - min_val) / (max_val - min_val)) as f32;
                    let base_y = zero_y;
                    let (rect_y, rect_h) = if y0 >= base_y {
                        (base_y, (y0 - base_y).max(1.5))
                    } else {
                        (y0, (base_y - y0).max(1.5))
                    };
                    page.rect_fill_stroke_rgb(x0, rect_y, bar_w - 1.0, rect_h, rgb, (0.35, 0.35, 0.35));
                    if spec.show_value_labels {
                        let label_y = if y0 >= base_y { rect_y + rect_h + 4.0 } else { rect_y - 11.0 };
                        draw_pdf_centered(page, x0 + (bar_w - 1.0) / 2.0, label_y, 8.5, PdfFont::Helvetica, &format_tick(*value), (0.15, 0.15, 0.15));
                    }
                }
                bar_slot += 1;
            }
            SeriesKind::Line | SeriesKind::Area => {
                let rgb = pdf_color(&series.color);
                let mut points = Vec::new();
                for (index, value) in series.values.iter().enumerate() {
                    let cx = plot_x + cluster_w * (index as f32 + 0.5);
                    let cy = plot_y + plot_h * ((*value - min_val) / (max_val - min_val)) as f32;
                    points.push((cx, cy));
                }
                if series.kind == SeriesKind::Area && points.len() >= 2 {
                    let mut area = Vec::new();
                    area.push((points[0].0, zero_y));
                    area.extend(points.iter().copied());
                    if let Some(last) = points.last() {
                        area.push((last.0, zero_y));
                    }
                    render_pdf_polygon(page, &area, pdf_color_alpha(&series.color, 0.20), rgb);
                }
                render_pdf_polyline(page, &points, rgb);
                for (point, value) in points.iter().zip(series.values.iter()) {
                    page.ellipse_fill_stroke_rgb(point.0, point.1, 2.6, 2.6, rgb, rgb);
                    if spec.show_value_labels {
                        draw_pdf_centered(page, point.0, point.1 + 6.0, 8.2, PdfFont::Helvetica, &format_tick(*value), (0.15, 0.15, 0.15));
                    }
                }
            }
            _ => {}
        }
    }
    let entries = legend_entries(spec);
    if legend_w > 0.0 {
        render_pdf_legend(page, plot_x + plot_w + 10.0, plot_y + 14.0, legend_w - 14.0, plot_h - 8.0, &entries);
    }
}

fn render_pdf_bar(page: &mut PdfPage, x: f32, y: f32, w: f32, h: f32, title_h: f32, legend_w: f32, spec: &ChartSpec) {
    let top_pad = if title_h > 0.0 { title_h + 6.0 } else { 8.0 };
    let left_pad = 56.0;
    let bottom_pad = if spec.value_axis_title.trim().is_empty() { 28.0 } else { 40.0 };
    let right_pad = if legend_w > 0.0 { legend_w + 22.0 } else { 16.0 };
    let plot_x = x + left_pad;
    let plot_y = y + bottom_pad;
    let plot_w = (w - left_pad - right_pad).max(50.0);
    let plot_h = (h - top_pad - bottom_pad).max(30.0);
    page.rect_fill_rgb(plot_x, plot_y, plot_w, plot_h, 0.985, 0.985, 0.985);
    page.rect_stroke_rgb(plot_x, plot_y, plot_w, plot_h, 0.80, 0.80, 0.80);
    if title_h > 0.0 {
        let baseline = y + h - title_h + 2.0;
        draw_pdf_centered(page, x + w / 2.0, baseline, 11.5, PdfFont::HelveticaBold, &spec.title, (0.1, 0.1, 0.1));
    }
    let Some((min_val, max_val)) = value_range(spec) else {
        return;
    };
    let ticks = axis_ticks(min_val, max_val);
    let baseline = if min_val <= 0.0 && max_val >= 0.0 { 0.0 } else { min_val };
    for tick in &ticks {
        let ratio = (*tick - min_val) / (max_val - min_val);
        let xx = plot_x + plot_w * ratio as f32;
        page.line_rgb(xx, plot_y, xx, plot_y + plot_h, 0.88, 0.88, 0.88);
        draw_pdf_centered(page, xx, plot_y - 14.0, 8.5, PdfFont::Helvetica, &format_tick(*tick), (0.32, 0.32, 0.32));
    }
    page.line_rgb(plot_x, plot_y, plot_x + plot_w, plot_y, 0.35, 0.35, 0.35);
    let zero_ratio = (baseline - min_val) / (max_val - min_val);
    let zero_x = plot_x + plot_w * zero_ratio as f32;
    page.line_rgb(zero_x, plot_y, zero_x, plot_y + plot_h, 0.35, 0.35, 0.35);

    let categories = spec.axis_categories();
    let category_count = categories.len().max(spec.series.iter().map(|series| series.values.len()).max().unwrap_or(0));
    if category_count == 0 {
        return;
    }
    let cluster_h = plot_h / category_count as f32;
    for (index, category) in categories.iter().enumerate() {
        let center_y = plot_y + plot_h - cluster_h * (index as f32 + 0.5);
        draw_pdf_right(page, plot_x - 8.0, center_y - 3.0, 8.5, PdfFont::Helvetica, category, (0.25, 0.25, 0.25));
    }
    if !spec.category_axis_title.trim().is_empty() {
        page.text_rgb(plot_x, y + h - top_pad + 1.0, 9.0, PdfFont::HelveticaOblique, &spec.category_axis_title, 0.2, 0.2, 0.2);
    }
    if !spec.value_axis_title.trim().is_empty() {
        draw_pdf_centered(page, plot_x + plot_w / 2.0, y + 8.0, 9.5, PdfFont::HelveticaOblique, &spec.value_axis_title, (0.2, 0.2, 0.2));
    }

    let bar_series = spec.series.iter().filter(|series| series.kind == SeriesKind::Bar).count().max(1);
    let mut bar_slot = 0usize;
    for series in &spec.series {
        if series.kind != SeriesKind::Bar {
            continue;
        }
        let bar_area = cluster_h * 0.76;
        let bar_h = (bar_area / bar_series as f32).max(4.0);
        let inset = (cluster_h - bar_area) / 2.0;
        let rgb = pdf_color(&series.color);
        for (index, value) in series.values.iter().enumerate() {
            let y0 = plot_y + plot_h - cluster_h * (index as f32 + 1.0) + inset + bar_h * bar_slot as f32;
            let x1 = plot_x + plot_w * ((*value - min_val) / (max_val - min_val)) as f32;
            let (rect_x, rect_w) = if x1 >= zero_x {
                (zero_x, (x1 - zero_x).max(1.5))
            } else {
                (x1, (zero_x - x1).max(1.5))
            };
            page.rect_fill_stroke_rgb(rect_x, y0, rect_w, bar_h - 1.0, rgb, (0.35, 0.35, 0.35));
            if spec.show_value_labels {
                page.text_rgb(rect_x + rect_w + 6.0, y0 + 1.0, 8.5, PdfFont::Helvetica, &format_tick(*value), 0.15, 0.15, 0.15);
            }
        }
        bar_slot += 1;
    }
    let entries = legend_entries(spec);
    if legend_w > 0.0 {
        render_pdf_legend(page, plot_x + plot_w + 10.0, plot_y + 14.0, legend_w - 14.0, plot_h - 8.0, &entries);
    }
}

fn render_pdf_scatter(page: &mut PdfPage, x: f32, y: f32, w: f32, h: f32, title_h: f32, legend_w: f32, spec: &ChartSpec) {
    let top_pad = if title_h > 0.0 { title_h + 6.0 } else { 8.0 };
    let left_pad = 36.0;
    let bottom_pad = if spec.category_axis_title.trim().is_empty() { 28.0 } else { 40.0 };
    let right_pad = if legend_w > 0.0 { legend_w + 20.0 } else { 12.0 };
    let plot_x = x + left_pad;
    let plot_y = y + bottom_pad;
    let plot_w = (w - left_pad - right_pad).max(50.0);
    let plot_h = (h - top_pad - bottom_pad).max(30.0);
    page.rect_fill_rgb(plot_x, plot_y, plot_w, plot_h, 0.985, 0.985, 0.985);
    page.rect_stroke_rgb(plot_x, plot_y, plot_w, plot_h, 0.80, 0.80, 0.80);
    if title_h > 0.0 {
        let baseline = y + h - title_h + 2.0;
        draw_pdf_centered(page, x + w / 2.0, baseline, 11.5, PdfFont::HelveticaBold, &spec.title, (0.1, 0.1, 0.1));
    }
    let Some((min_y, max_y)) = value_range(spec) else { return; };
    let Some((min_x, max_x)) = x_numeric_range(spec) else { return; };
    let y_ticks = axis_ticks(min_y, max_y);
    let x_ticks = axis_ticks(min_x, max_x);
    for tick in &y_ticks {
        let ratio = (*tick - min_y) / (max_y - min_y);
        let yy = plot_y + plot_h * ratio as f32;
        page.line_rgb(plot_x, yy, plot_x + plot_w, yy, 0.88, 0.88, 0.88);
        draw_pdf_right(page, plot_x - 6.0, yy - 3.5, 8.5, PdfFont::Helvetica, &format_tick(*tick), (0.32, 0.32, 0.32));
    }
    for tick in &x_ticks {
        let ratio = (*tick - min_x) / (max_x - min_x);
        let xx = plot_x + plot_w * ratio as f32;
        page.line_rgb(xx, plot_y, xx, plot_y + plot_h, 0.92, 0.92, 0.92);
        draw_pdf_centered(page, xx, plot_y - 14.0, 8.5, PdfFont::Helvetica, &format_tick(*tick), (0.32, 0.32, 0.32));
    }
    page.line_rgb(plot_x, plot_y, plot_x, plot_y + plot_h, 0.35, 0.35, 0.35);
    page.line_rgb(plot_x, plot_y, plot_x + plot_w, plot_y, 0.35, 0.35, 0.35);
    if !spec.category_axis_title.trim().is_empty() {
        draw_pdf_centered(page, plot_x + plot_w / 2.0, y + 8.0, 9.5, PdfFont::HelveticaOblique, &spec.category_axis_title, (0.2, 0.2, 0.2));
    }
    if !spec.value_axis_title.trim().is_empty() {
        page.text_rgb(plot_x, y + h - top_pad + 1.0, 9.0, PdfFont::HelveticaOblique, &spec.value_axis_title, 0.2, 0.2, 0.2);
    }
    for series in &spec.series {
        let rgb = pdf_color(&series.color);
        for (x_value, y_value) in series.x_values.iter().zip(series.values.iter()) {
            let px = plot_x + plot_w * ((*x_value - min_x) / (max_x - min_x)) as f32;
            let py = plot_y + plot_h * ((*y_value - min_y) / (max_y - min_y)) as f32;
            page.ellipse_fill_stroke_rgb(px, py, 3.2, 3.2, rgb, rgb);
            if spec.show_value_labels {
                page.text_rgb(px + 5.0, py + 4.0, 8.2, PdfFont::Helvetica, &format_tick(*y_value), 0.15, 0.15, 0.15);
            }
        }
    }
    let entries = legend_entries(spec);
    if legend_w > 0.0 {
        render_pdf_legend(page, plot_x + plot_w + 10.0, plot_y + 14.0, legend_w - 14.0, plot_h - 8.0, &entries);
    }
}

fn render_pdf_pie_like(page: &mut PdfPage, x: f32, y: f32, w: f32, h: f32, title_h: f32, legend_w: f32, spec: &ChartSpec, doughnut: bool) {
    let top_pad = if title_h > 0.0 { title_h + 6.0 } else { 8.0 };
    let right_pad = if legend_w > 0.0 { legend_w + 18.0 } else { 12.0 };
    let left_pad = 12.0;
    let bottom_pad = 12.0;
    let plot_x = x + left_pad;
    let plot_y = y + bottom_pad;
    let plot_w = (w - left_pad - right_pad).max(40.0);
    let plot_h = (h - top_pad - bottom_pad).max(40.0);
    if title_h > 0.0 {
        let baseline = y + h - title_h + 2.0;
        draw_pdf_centered(page, x + w / 2.0, baseline, 11.5, PdfFont::HelveticaBold, &spec.title, (0.1, 0.1, 0.1));
    }
    let Some(series) = spec.series.first() else { return; };
    let categories = if spec.categories.is_empty() { series.categories.clone() } else { spec.categories.clone() };
    let values = &series.values;
    let total: f64 = values.iter().copied().filter(|value| *value > 0.0).sum();
    if total <= 0.0 { return; }
    let cx = plot_x + plot_w * 0.46;
    let cy = plot_y + plot_h * 0.50;
    let radius = plot_w.min(plot_h) * 0.34;
    let inner = if doughnut { radius * 0.48 } else { 0.0 };
    let mut start = -PI / 2.0;
    for (index, value) in values.iter().enumerate() {
        if *value <= 0.0 { continue; }
        let sweep = (*value as f32 / total as f32) * PI * 2.0;
        let end = start + sweep;
        let color = pdf_color(&slice_color(index, &series.color));
        let points = wedge_points((cx, cy), radius, inner, start, end);
        render_pdf_polygon(page, &points, color, (0.35, 0.35, 0.35));
        if spec.show_value_labels || spec.show_category_labels {
            let label_angle = start + sweep / 2.0;
            let label_radius = if doughnut { radius * 1.10 } else { radius * 0.72 };
            let lx = cx + label_radius * label_angle.cos();
            let ly = cy + label_radius * label_angle.sin();
            let mut label_parts = Vec::new();
            if spec.show_category_labels {
                if let Some(category) = categories.get(index) {
                    label_parts.push(category.clone());
                }
            }
            if spec.show_value_labels {
                label_parts.push(format_tick(*value));
            }
            if label_parts.is_empty() {
                label_parts.push(format_tick(*value));
            }
            page.text_rgb(lx, ly, 8.5, PdfFont::Helvetica, &label_parts.join(" "), 0.15, 0.15, 0.15);
        }
        start = end;
    }
    let entries = legend_entries(spec);
    if legend_w > 0.0 {
        render_pdf_legend(page, plot_x + plot_w + 10.0, plot_y + plot_h - 8.0, legend_w - 14.0, plot_h - 8.0, &entries);
    }
}

fn render_raster_cartesian(image: &mut RasterImage, x: i32, y: i32, w: i32, h: i32, title_h: i32, legend_w: i32, spec: &ChartSpec) {
    let top_pad = if title_h > 0 { title_h + 6 } else { 8 };
    let left_pad = 44;
    let bottom_pad = if spec.category_axis_title.trim().is_empty() { 30 } else { 42 };
    let right_pad = if legend_w > 0 { legend_w + 18 } else { 12 };
    let plot_x = x + left_pad;
    let plot_y = y + top_pad;
    let plot_w = (w - left_pad - right_pad).max(40);
    let plot_h = (h - top_pad - bottom_pad).max(28);
    image.fill_rect(plot_x, plot_y, plot_w, plot_h, Rgba::rgba(251, 251, 251, 255));
    image.stroke_rect(plot_x, plot_y, plot_w, plot_h, 1, Rgba::rgba(200, 200, 200, 255));
    if title_h > 0 {
        let tx = x + w / 2 - image.measure_text(&spec.title, 12) / 2;
        image.draw_text(tx, y + 8, 12, Rgba::rgba(30, 30, 30, 255), &spec.title, true);
    }
    let Some((min_val, max_val)) = value_range(spec) else { return; };
    let ticks = axis_ticks(min_val, max_val);
    let baseline = if min_val <= 0.0 && max_val >= 0.0 { 0.0 } else { min_val };
    for tick in &ticks {
        let ratio = (*tick - min_val) / (max_val - min_val);
        let yy = plot_y + plot_h - (plot_h as f64 * ratio) as i32;
        image.draw_line(plot_x, yy, plot_x + plot_w, yy, 1, Rgba::rgba(228, 228, 228, 255));
        let text = format_tick(*tick);
        let tw = image.measure_text(&text, 10);
        image.draw_text(plot_x - tw - 6, yy - 4, 10, Rgba::rgba(80, 80, 80, 255), &text, false);
    }
    image.draw_line(plot_x, plot_y, plot_x, plot_y + plot_h, 1, Rgba::rgba(70, 70, 70, 255));
    let zero_ratio = (baseline - min_val) / (max_val - min_val);
    let zero_y = plot_y + plot_h - (plot_h as f64 * zero_ratio) as i32;
    image.draw_line(plot_x, zero_y, plot_x + plot_w, zero_y, 1, Rgba::rgba(70, 70, 70, 255));
    let categories = spec.axis_categories();
    let category_count = categories.len().max(spec.series.iter().map(|series| series.values.len()).max().unwrap_or(0));
    if category_count == 0 {
        return;
    }
    let cluster_w = plot_w as f32 / category_count as f32;
    for (index, category) in categories.iter().enumerate() {
        let cx = plot_x + (cluster_w * (index as f32 + 0.5)).round() as i32;
        let tw = image.measure_text(category, 10);
        image.draw_text(cx - tw / 2, plot_y + plot_h + 6, 10, Rgba::rgba(70, 70, 70, 255), category, false);
    }
    if !spec.category_axis_title.trim().is_empty() {
        let tw = image.measure_text(&spec.category_axis_title, 10);
        image.draw_text(plot_x + plot_w / 2 - tw / 2, y + h - 14, 10, Rgba::rgba(70, 70, 70, 255), &spec.category_axis_title, false);
    }
    if !spec.value_axis_title.trim().is_empty() {
        image.draw_text(plot_x, y + 12, 10, Rgba::rgba(70, 70, 70, 255), &spec.value_axis_title, false);
    }
    let bar_series = spec.series.iter().filter(|series| series.kind == SeriesKind::Column).count().max(1);
    let mut bar_slot = 0usize;
    for series in &spec.series {
        match series.kind {
            SeriesKind::Column => {
                let bar_area = cluster_w * 0.76;
                let bar_w = (bar_area / bar_series as f32).max(4.0);
                let inset = (cluster_w - bar_area) / 2.0;
                let fill = parse_hex_color(&series.color, Rgba::rgba(80, 120, 180, 255));
                for (index, value) in series.values.iter().enumerate() {
                    let x0 = plot_x + (cluster_w * index as f32 + inset + bar_w * bar_slot as f32).round() as i32;
                    let y1 = plot_y + plot_h - (plot_h as f64 * ((*value - min_val) / (max_val - min_val))) as i32;
                    let (rect_y, rect_h) = if y1 <= zero_y {
                        (y1, (zero_y - y1).max(2))
                    } else {
                        (zero_y, (y1 - zero_y).max(2))
                    };
                    image.fill_rect(x0, rect_y, (bar_w - 1.0).round() as i32, rect_h, fill);
                    image.stroke_rect(x0, rect_y, (bar_w - 1.0).round() as i32, rect_h, 1, Rgba::rgba(60, 60, 60, 255));
                    if spec.show_value_labels {
                        let text = format_tick(*value);
                        let tw = image.measure_text(&text, 10);
                        image.draw_text(x0 + (bar_w as i32 - tw) / 2, rect_y - 12, 10, Rgba::rgba(30, 30, 30, 255), &text, false);
                    }
                }
                bar_slot += 1;
            }
            SeriesKind::Line | SeriesKind::Area => {
                let stroke = parse_hex_color(&series.color, Rgba::rgba(80, 120, 180, 255));
                let mut points = Vec::new();
                for (index, value) in series.values.iter().enumerate() {
                    let px = plot_x + (cluster_w * (index as f32 + 0.5)).round() as i32;
                    let py = plot_y + plot_h - (plot_h as f64 * ((*value - min_val) / (max_val - min_val))) as i32;
                    points.push((px, py));
                }
                if series.kind == SeriesKind::Area && points.len() >= 2 {
                    let mut polygon = Vec::new();
                    polygon.push((points[0].0, zero_y));
                    polygon.extend(points.iter().copied());
                    if let Some(last) = points.last() {
                        polygon.push((last.0, zero_y));
                    }
                    fill_raster_polygon(image, &polygon, stroke.with_alpha(60));
                    stroke_raster_polygon(image, &polygon, 1, stroke);
                }
                for window in points.windows(2) {
                    image.draw_line(window[0].0, window[0].1, window[1].0, window[1].1, 2, stroke);
                }
                for (point, value) in points.iter().zip(series.values.iter()) {
                    image.fill_ellipse(point.0, point.1, 3, 3, stroke);
                    image.stroke_ellipse(point.0, point.1, 3, 3, 1, Rgba::rgba(40, 40, 40, 255));
                    if spec.show_value_labels {
                        image.draw_text(point.0 + 4, point.1 - 12, 10, Rgba::rgba(30, 30, 30, 255), &format_tick(*value), false);
                    }
                }
            }
            _ => {}
        }
    }
    let entries = legend_entries(spec);
    if legend_w > 0 {
        render_raster_legend(image, plot_x + plot_w + 10, plot_y + 18, &entries);
    }
}

fn render_raster_bar(image: &mut RasterImage, x: i32, y: i32, w: i32, h: i32, title_h: i32, legend_w: i32, spec: &ChartSpec) {
    let top_pad = if title_h > 0 { title_h + 6 } else { 8 };
    let left_pad = 62;
    let bottom_pad = if spec.value_axis_title.trim().is_empty() { 30 } else { 42 };
    let right_pad = if legend_w > 0 { legend_w + 18 } else { 12 };
    let plot_x = x + left_pad;
    let plot_y = y + top_pad;
    let plot_w = (w - left_pad - right_pad).max(40);
    let plot_h = (h - top_pad - bottom_pad).max(28);
    image.fill_rect(plot_x, plot_y, plot_w, plot_h, Rgba::rgba(251, 251, 251, 255));
    image.stroke_rect(plot_x, plot_y, plot_w, plot_h, 1, Rgba::rgba(200, 200, 200, 255));
    if title_h > 0 {
        let tx = x + w / 2 - image.measure_text(&spec.title, 12) / 2;
        image.draw_text(tx, y + 8, 12, Rgba::rgba(30, 30, 30, 255), &spec.title, true);
    }
    let Some((min_val, max_val)) = value_range(spec) else { return; };
    let ticks = axis_ticks(min_val, max_val);
    let baseline = if min_val <= 0.0 && max_val >= 0.0 { 0.0 } else { min_val };
    for tick in &ticks {
        let ratio = (*tick - min_val) / (max_val - min_val);
        let xx = plot_x + (plot_w as f64 * ratio) as i32;
        image.draw_line(xx, plot_y, xx, plot_y + plot_h, 1, Rgba::rgba(228, 228, 228, 255));
        let text = format_tick(*tick);
        let tw = image.measure_text(&text, 10);
        image.draw_text(xx - tw / 2, plot_y + plot_h + 6, 10, Rgba::rgba(80, 80, 80, 255), &text, false);
    }
    image.draw_line(plot_x, plot_y + plot_h, plot_x + plot_w, plot_y + plot_h, 1, Rgba::rgba(70, 70, 70, 255));
    let zero_ratio = (baseline - min_val) / (max_val - min_val);
    let zero_x = plot_x + (plot_w as f64 * zero_ratio) as i32;
    image.draw_line(zero_x, plot_y, zero_x, plot_y + plot_h, 1, Rgba::rgba(70, 70, 70, 255));
    let categories = spec.axis_categories();
    let category_count = categories.len().max(spec.series.iter().map(|series| series.values.len()).max().unwrap_or(0));
    if category_count == 0 { return; }
    let cluster_h = plot_h as f32 / category_count as f32;
    for (index, category) in categories.iter().enumerate() {
        let cy = plot_y + plot_h - (cluster_h * (index as f32 + 0.5)).round() as i32;
        let tw = image.measure_text(category, 10);
        image.draw_text(plot_x - tw - 8, cy - 4, 10, Rgba::rgba(70, 70, 70, 255), category, false);
    }
    if !spec.category_axis_title.trim().is_empty() {
        image.draw_text(plot_x, y + 12, 10, Rgba::rgba(70, 70, 70, 255), &spec.category_axis_title, false);
    }
    if !spec.value_axis_title.trim().is_empty() {
        let tw = image.measure_text(&spec.value_axis_title, 10);
        image.draw_text(plot_x + plot_w / 2 - tw / 2, y + h - 14, 10, Rgba::rgba(70, 70, 70, 255), &spec.value_axis_title, false);
    }
    let bar_series = spec.series.iter().filter(|series| series.kind == SeriesKind::Bar).count().max(1);
    let mut bar_slot = 0usize;
    for series in &spec.series {
        if series.kind != SeriesKind::Bar { continue; }
        let bar_area = cluster_h * 0.76;
        let bar_h = (bar_area / bar_series as f32).max(4.0);
        let inset = (cluster_h - bar_area) / 2.0;
        let fill = parse_hex_color(&series.color, Rgba::rgba(80, 120, 180, 255));
        for (index, value) in series.values.iter().enumerate() {
            let y0 = plot_y + plot_h - (cluster_h * (index as f32 + 1.0)).round() as i32 + inset.round() as i32 + (bar_h * bar_slot as f32).round() as i32;
            let x1 = plot_x + (plot_w as f64 * ((*value - min_val) / (max_val - min_val))) as i32;
            let (rect_x, rect_w) = if x1 >= zero_x { (zero_x, (x1 - zero_x).max(2)) } else { (x1, (zero_x - x1).max(2)) };
            image.fill_rect(rect_x, y0, rect_w, (bar_h - 1.0).round() as i32, fill);
            image.stroke_rect(rect_x, y0, rect_w, (bar_h - 1.0).round() as i32, 1, Rgba::rgba(60, 60, 60, 255));
            if spec.show_value_labels {
                image.draw_text(rect_x + rect_w + 5, y0 - 1, 10, Rgba::rgba(30, 30, 30, 255), &format_tick(*value), false);
            }
        }
        bar_slot += 1;
    }
    let entries = legend_entries(spec);
    if legend_w > 0 {
        render_raster_legend(image, plot_x + plot_w + 10, plot_y + 18, &entries);
    }
}

fn render_raster_scatter(image: &mut RasterImage, x: i32, y: i32, w: i32, h: i32, title_h: i32, legend_w: i32, spec: &ChartSpec) {
    let top_pad = if title_h > 0 { title_h + 6 } else { 8 };
    let left_pad = 44;
    let bottom_pad = if spec.category_axis_title.trim().is_empty() { 30 } else { 42 };
    let right_pad = if legend_w > 0 { legend_w + 18 } else { 12 };
    let plot_x = x + left_pad;
    let plot_y = y + top_pad;
    let plot_w = (w - left_pad - right_pad).max(40);
    let plot_h = (h - top_pad - bottom_pad).max(28);
    image.fill_rect(plot_x, plot_y, plot_w, plot_h, Rgba::rgba(251, 251, 251, 255));
    image.stroke_rect(plot_x, plot_y, plot_w, plot_h, 1, Rgba::rgba(200, 200, 200, 255));
    if title_h > 0 {
        let tx = x + w / 2 - image.measure_text(&spec.title, 12) / 2;
        image.draw_text(tx, y + 8, 12, Rgba::rgba(30, 30, 30, 255), &spec.title, true);
    }
    let Some((min_y, max_y)) = value_range(spec) else { return; };
    let Some((min_x, max_x)) = x_numeric_range(spec) else { return; };
    let y_ticks = axis_ticks(min_y, max_y);
    let x_ticks = axis_ticks(min_x, max_x);
    for tick in &y_ticks {
        let ratio = (*tick - min_y) / (max_y - min_y);
        let yy = plot_y + plot_h - (plot_h as f64 * ratio) as i32;
        image.draw_line(plot_x, yy, plot_x + plot_w, yy, 1, Rgba::rgba(228, 228, 228, 255));
        let text = format_tick(*tick);
        let tw = image.measure_text(&text, 10);
        image.draw_text(plot_x - tw - 6, yy - 4, 10, Rgba::rgba(80, 80, 80, 255), &text, false);
    }
    for tick in &x_ticks {
        let ratio = (*tick - min_x) / (max_x - min_x);
        let xx = plot_x + (plot_w as f64 * ratio) as i32;
        image.draw_line(xx, plot_y, xx, plot_y + plot_h, 1, Rgba::rgba(234, 234, 234, 255));
        let text = format_tick(*tick);
        let tw = image.measure_text(&text, 10);
        image.draw_text(xx - tw / 2, plot_y + plot_h + 6, 10, Rgba::rgba(80, 80, 80, 255), &text, false);
    }
    image.draw_line(plot_x, plot_y, plot_x, plot_y + plot_h, 1, Rgba::rgba(70, 70, 70, 255));
    image.draw_line(plot_x, plot_y + plot_h, plot_x + plot_w, plot_y + plot_h, 1, Rgba::rgba(70, 70, 70, 255));
    if !spec.category_axis_title.trim().is_empty() {
        let tw = image.measure_text(&spec.category_axis_title, 10);
        image.draw_text(plot_x + plot_w / 2 - tw / 2, y + h - 14, 10, Rgba::rgba(70, 70, 70, 255), &spec.category_axis_title, false);
    }
    if !spec.value_axis_title.trim().is_empty() {
        image.draw_text(plot_x, y + 12, 10, Rgba::rgba(70, 70, 70, 255), &spec.value_axis_title, false);
    }
    for series in &spec.series {
        let fill = parse_hex_color(&series.color, Rgba::rgba(80, 120, 180, 255));
        for (x_value, y_value) in series.x_values.iter().zip(series.values.iter()) {
            let px = plot_x + (plot_w as f64 * ((*x_value - min_x) / (max_x - min_x))) as i32;
            let py = plot_y + plot_h - (plot_h as f64 * ((*y_value - min_y) / (max_y - min_y))) as i32;
            image.fill_ellipse(px, py, 3, 3, fill);
            image.stroke_ellipse(px, py, 3, 3, 1, Rgba::rgba(40, 40, 40, 255));
            if spec.show_value_labels {
                image.draw_text(px + 4, py - 12, 10, Rgba::rgba(30, 30, 30, 255), &format_tick(*y_value), false);
            }
        }
    }
    let entries = legend_entries(spec);
    if legend_w > 0 {
        render_raster_legend(image, plot_x + plot_w + 10, plot_y + 18, &entries);
    }
}

fn render_raster_pie_like(image: &mut RasterImage, x: i32, y: i32, w: i32, h: i32, title_h: i32, legend_w: i32, spec: &ChartSpec, doughnut: bool) {
    let top_pad = if title_h > 0 { title_h + 6 } else { 8 };
    let right_pad = if legend_w > 0 { legend_w + 18 } else { 12 };
    let left_pad = 12;
    let bottom_pad = 12;
    let plot_x = x + left_pad;
    let plot_y = y + top_pad;
    let plot_w = (w - left_pad - right_pad).max(40);
    let plot_h = (h - top_pad - bottom_pad).max(40);
    if title_h > 0 {
        let tx = x + w / 2 - image.measure_text(&spec.title, 12) / 2;
        image.draw_text(tx, y + 8, 12, Rgba::rgba(30, 30, 30, 255), &spec.title, true);
    }
    let Some(series) = spec.series.first() else { return; };
    let categories = if spec.categories.is_empty() { series.categories.clone() } else { spec.categories.clone() };
    let total: f64 = series.values.iter().copied().filter(|value| *value > 0.0).sum();
    if total <= 0.0 { return; }
    let cx = plot_x + (plot_w as f32 * 0.46).round() as i32;
    let cy = plot_y + (plot_h as f32 * 0.50).round() as i32;
    let radius = ((plot_w.min(plot_h) as f32) * 0.34).round() as i32;
    let inner = if doughnut { (radius as f32 * 0.48).round() as i32 } else { 0 };
    let mut start = -PI / 2.0;
    for (index, value) in series.values.iter().enumerate() {
        if *value <= 0.0 { continue; }
        let sweep = (*value as f32 / total as f32) * PI * 2.0;
        let end = start + sweep;
        let color = parse_hex_color(&slice_color(index, &series.color), Rgba::rgba(80, 120, 180, 255));
        let points = wedge_points((cx as f32, cy as f32), radius as f32, inner as f32, start, end)
            .into_iter()
            .map(|(px, py)| (px.round() as i32, py.round() as i32))
            .collect::<Vec<_>>();
        fill_raster_polygon(image, &points, color);
        stroke_raster_polygon(image, &points, 1, Rgba::rgba(60, 60, 60, 255));
        if spec.show_value_labels || spec.show_category_labels {
            let angle = start + sweep / 2.0;
            let label_radius = if doughnut { radius as f32 * 1.10 } else { radius as f32 * 0.72 };
            let lx = cx + (label_radius * angle.cos()).round() as i32;
            let ly = cy + (label_radius * angle.sin()).round() as i32;
            let mut label_parts = Vec::new();
            if spec.show_category_labels {
                if let Some(category) = categories.get(index) {
                    label_parts.push(category.clone());
                }
            }
            if spec.show_value_labels {
                label_parts.push(format_tick(*value));
            }
            if label_parts.is_empty() {
                label_parts.push(format_tick(*value));
            }
            image.draw_text(lx, ly, 10, Rgba::rgba(30, 30, 30, 255), &label_parts.join(" "), false);
        }
        start = end;
    }
    let entries = legend_entries(spec);
    if legend_w > 0 {
        render_raster_legend(image, plot_x + plot_w + 10, plot_y + 18, &entries);
    }
}

fn pdf_color(input: &str) -> (f32, f32, f32) {
    let color = parse_hex_color(input, Rgba::rgba(80, 120, 180, 255));
    (
        color.r as f32 / 255.0,
        color.g as f32 / 255.0,
        color.b as f32 / 255.0,
    )
}

fn pdf_color_alpha(input: &str, alpha: f32) -> (f32, f32, f32) {
    let base = pdf_color(input);
    let blend = |component: f32| component * alpha + (1.0 - alpha);
    (blend(base.0), blend(base.1), blend(base.2))
}

fn slice_color(index: usize, base: &str) -> String {
    if index == 0 && !base.trim().is_empty() {
        return base.to_string();
    }
    PALETTE[index % PALETTE.len()].to_string()
}

#[cfg(test)]
mod tests {
    use super::*;
    use lo_core::{parse_xml_document, Presentation, Slide};
    use crate::pdf::to_pdf;

    #[test]
    fn chart_spec_round_trip() {
        let spec = ChartSpec {
            frame_mm: (25.0, 30.0, 160.0, 90.0),
            title: "Chart Types".to_string(),
            category_axis_title: "Category".to_string(),
            value_axis_title: "Value".to_string(),
            categories: vec!["A".to_string(), "B".to_string()],
            series: vec![ChartSeries {
                name: "Series 1".to_string(),
                categories: vec!["A".to_string(), "B".to_string()],
                values: vec![4.3, 2.4],
                x_values: Vec::new(),
                color: "#4472C4".to_string(),
                kind: SeriesKind::Column,
            }],
            show_value_labels: true,
            show_category_labels: false,
        };
        let row = spec_to_row(&spec);
        let decoded = row_to_spec(&row).expect("decode chart row");
        assert_eq!(decoded.title, "Chart Types");
        assert_eq!(decoded.series[0].values, vec![4.3, 2.4]);
        assert_eq!(decoded.category_axis_title, "Category");
        assert!(decoded.show_value_labels);
    }

    #[test]
    fn parses_bar_chart_axis_titles_and_values() {
        let xml = r#"
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:chart>
            <c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:plotArea>
              <c:barChart>
                <c:barDir val="col"/>
                <c:ser>
                  <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>North</c:v></c:pt></c:strCache></c:strRef></c:tx>
                  <c:cat><c:strRef><c:strCache>
                    <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                    <c:pt idx="1"><c:v>Q2</c:v></c:pt>
                  </c:strCache></c:strRef></c:cat>
                  <c:val><c:numRef><c:numCache>
                    <c:pt idx="0"><c:v>4.3</c:v></c:pt>
                    <c:pt idx="1"><c:v>2.4</c:v></c:pt>
                  </c:numCache></c:numRef></c:val>
                </c:ser>
                <c:dLbls><c:showVal val="1"/></c:dLbls>
              </c:barChart>
              <c:catAx><c:title><c:tx><c:rich><a:p><a:r><a:t>Quarter</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx>
              <c:valAx><c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
        "#;
        let root = parse_xml_document(xml).expect("chart xml");
        let spec = parse_chart_spec(&root, (20.0, 20.0, 160.0, 90.0), 0);
        assert_eq!(spec.title, "Sales");
        assert_eq!(spec.category_axis_title, "Quarter");
        assert_eq!(spec.value_axis_title, "Revenue");
        assert_eq!(spec.series[0].values, vec![4.3, 2.4]);
        assert_eq!(spec.axis_categories(), vec!["Q1".to_string(), "Q2".to_string()]);
        assert!(spec.show_value_labels);
    }

    #[test]
    fn pdf_keeps_data_labels_separate() {
        let spec = ChartSpec {
            frame_mm: (20.0, 20.0, 160.0, 90.0),
            title: "Labels".to_string(),
            category_axis_title: "Quarter".to_string(),
            value_axis_title: "Revenue".to_string(),
            categories: vec!["Q1".to_string(), "Q2".to_string()],
            series: vec![ChartSeries {
                name: "North".to_string(),
                categories: vec!["Q1".to_string(), "Q2".to_string()],
                values: vec![4.3, 2.4],
                x_values: Vec::new(),
                color: "#4472C4".to_string(),
                kind: SeriesKind::Column,
            }],
            show_value_labels: true,
            show_category_labels: false,
        };
        let mut presentation = Presentation::new("Charts");
        presentation.slides.push(Slide {
            name: "Chart".to_string(),
            elements: Vec::new(),
            notes: Vec::new(),
            chart_tokens: vec![spec_to_row(&spec)],
        });
        let pdf = to_pdf(&presentation);
        let text = String::from_utf8_lossy(&pdf);
        assert!(text.contains("(4.3) Tj"));
        assert!(text.contains("(2.4) Tj"));
        assert!(text.contains("(Quarter) Tj"));
        assert!(text.contains("(Revenue) Tj"));
    }
}
