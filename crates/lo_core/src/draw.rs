use crate::geometry::Rect;
use crate::impress::{ImageElement, Shape, TextBox};
use crate::meta::Metadata;
use crate::units::Length;

#[derive(Clone, Debug, PartialEq)]
pub enum DrawElement {
    TextBox(TextBox),
    Shape(Shape),
    Image(ImageElement),
}

#[derive(Clone, Debug, PartialEq)]
pub struct DrawPage {
    pub name: String,
    pub elements: Vec<DrawElement>,
}

impl Default for DrawPage {
    fn default() -> Self {
        Self {
            name: "Page1".to_string(),
            elements: Vec::new(),
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct Drawing {
    pub meta: Metadata,
    pub page_size: crate::geometry::Size,
    pub pages: Vec<DrawPage>,
}

impl Default for Drawing {
    fn default() -> Self {
        Self {
            meta: Metadata::default(),
            page_size: crate::geometry::Size::new(Length::mm(297.0), Length::mm(210.0)),
            pages: vec![DrawPage::default()],
        }
    }
}

impl Drawing {
    pub fn new(title: impl Into<String>) -> Self {
        Self {
            meta: Metadata::titled(title),
            ..Self::default()
        }
    }

    pub fn embedded_images(&self) -> Vec<(String, String, Vec<u8>)> {
        let mut images = Vec::new();
        for page in &self.pages {
            for element in &page.elements {
                if let DrawElement::Image(image) = element {
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
fn _rect() -> Rect {
    Rect::new(
        Length::mm(0.0),
        Length::mm(0.0),
        Length::mm(10.0),
        Length::mm(10.0),
    )
}
