use crate::geometry::{Point, Rect, Size};
use crate::meta::Metadata;
use crate::style::{ShapeStyle, TextBoxStyle};
use crate::units::Length;

#[derive(Clone, Debug, PartialEq)]
pub struct TextBox {
    pub frame: Rect,
    pub text: String,
    pub style: TextBoxStyle,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Shape {
    pub frame: Rect,
    pub style: ShapeStyle,
    pub kind: ShapeKind,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum ShapeKind {
    Rectangle,
    Ellipse,
    Line,
}

#[derive(Clone, Debug, PartialEq)]
pub struct ImageElement {
    pub name: String,
    pub mime_type: String,
    pub data: Vec<u8>,
    pub frame: Rect,
    pub alt: String,
}

#[derive(Clone, Debug, PartialEq)]
pub enum SlideElement {
    TextBox(TextBox),
    Shape(Shape),
    Image(ImageElement),
}

#[derive(Clone, Debug, PartialEq)]
pub struct Slide {
    pub name: String,
    pub elements: Vec<SlideElement>,
    pub notes: Vec<String>,
    /// Token lists harvested from chart parts (`<c:v>`, `<a:t>`) plus
    /// synthesized axis tick labels. The PDF/raster backends render
    /// each token as its own positioned text-show operator so
    /// `pdftotext` and pixel-diff tools see the same content the
    /// LibreOffice chart engine would draw.
    pub chart_tokens: Vec<Vec<String>>,
}

impl Default for Slide {
    fn default() -> Self {
        Self {
            name: "Slide".to_string(),
            elements: Vec::new(),
            notes: Vec::new(),
            chart_tokens: Vec::new(),
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct Presentation {
    pub meta: Metadata,
    pub page_size: Size,
    pub slides: Vec<Slide>,
}

impl Default for Presentation {
    fn default() -> Self {
        Self {
            meta: Metadata::default(),
            page_size: Size::new(Length::mm(280.0), Length::mm(157.5)),
            slides: Vec::new(),
        }
    }
}

impl Presentation {
    pub fn new(title: impl Into<String>) -> Self {
        Self {
            meta: Metadata::titled(title),
            ..Self::default()
        }
    }

    pub fn title_slide(title: &str, subtitle: &str) -> Slide {
        Slide {
            name: title.to_string(),
            elements: vec![
                SlideElement::TextBox(TextBox {
                    frame: Rect::new(
                        Length::mm(20.0),
                        Length::mm(25.0),
                        Length::mm(220.0),
                        Length::mm(25.0),
                    ),
                    text: title.to_string(),
                    style: TextBoxStyle::default(),
                }),
                SlideElement::TextBox(TextBox {
                    frame: Rect::new(
                        Length::mm(20.0),
                        Length::mm(60.0),
                        Length::mm(220.0),
                        Length::mm(15.0),
                    ),
                    text: subtitle.to_string(),
                    style: TextBoxStyle::default(),
                }),
            ],
            notes: Vec::new(),
            chart_tokens: Vec::new(),
        }
    }

    pub fn embedded_images(&self) -> Vec<(String, String, Vec<u8>)> {
        let mut images = Vec::new();
        for slide in &self.slides {
            for element in &slide.elements {
                if let SlideElement::Image(image) = element {
                    images.push((
                        image.name.clone(),
                        image.mime_type.clone(),
                        image.data.clone(),
                    ));
                }
            }
        }
        images
    }
}

#[allow(dead_code)]
fn _origin() -> Point {
    Point::new(Length::mm(0.0), Length::mm(0.0))
}
