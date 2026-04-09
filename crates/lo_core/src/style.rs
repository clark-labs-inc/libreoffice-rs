use crate::units::Length;

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Alignment {
    Start,
    Center,
    End,
    Justify,
}

impl Default for Alignment {
    fn default() -> Self {
        Self::Start
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct TextStyle {
    pub font_family: String,
    pub font_size_pt: u16,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub color: String,
    pub background: String,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ParagraphStyle {
    pub alignment: Alignment,
    pub margin_top_mm: u16,
    pub margin_bottom_mm: u16,
    pub margin_left_mm: u16,
    pub margin_right_mm: u16,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct PageStyle {
    pub width_mm: u16,
    pub height_mm: u16,
    pub margin_mm: u16,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CellStyle {
    pub format_code: String,
    pub background: String,
    pub foreground: String,
    pub alignment: Alignment,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ShapeStyle {
    pub fill: String,
    pub stroke: String,
    pub stroke_width_mm: u16,
}

#[derive(Clone, Debug, PartialEq)]
pub struct TextBoxStyle {
    pub padding: Length,
    pub background: String,
    pub foreground: String,
}

impl Default for TextBoxStyle {
    fn default() -> Self {
        Self {
            padding: Length::mm(1.5),
            background: String::new(),
            foreground: String::new(),
        }
    }
}
