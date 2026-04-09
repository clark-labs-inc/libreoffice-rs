use std::path::Path;

use lo_core::{DrawElement, Drawing, Result, ShapeKind};

use crate::common::{content_root_attrs, image_extras, package_document, MIME_ODG};

fn shape_xml(shape: &lo_core::Shape) -> String {
    match shape.kind {
        ShapeKind::Rectangle => format!(
            "<draw:rect svg:x=\"{}\" svg:y=\"{}\" svg:width=\"{}\" svg:height=\"{}\" draw:fill-color=\"{}\" draw:stroke-color=\"{}\"/>",
            shape.frame.origin.x.css(),
            shape.frame.origin.y.css(),
            shape.frame.size.width.css(),
            shape.frame.size.height.css(),
            lo_core::escape_attr(&shape.style.fill),
            lo_core::escape_attr(&shape.style.stroke)
        ),
        ShapeKind::Ellipse => format!(
            "<draw:ellipse svg:x=\"{}\" svg:y=\"{}\" svg:width=\"{}\" svg:height=\"{}\" draw:fill-color=\"{}\" draw:stroke-color=\"{}\"/>",
            shape.frame.origin.x.css(),
            shape.frame.origin.y.css(),
            shape.frame.size.width.css(),
            shape.frame.size.height.css(),
            lo_core::escape_attr(&shape.style.fill),
            lo_core::escape_attr(&shape.style.stroke)
        ),
        ShapeKind::Line => format!(
            "<draw:line svg:x1=\"{}\" svg:y1=\"{}\" svg:x2=\"{}\" svg:y2=\"{}\" draw:stroke-color=\"{}\"/>",
            shape.frame.origin.x.css(),
            shape.frame.origin.y.css(),
            (shape.frame.origin.x.as_mm() + shape.frame.size.width.as_mm()).to_string() + "mm",
            (shape.frame.origin.y.as_mm() + shape.frame.size.height.as_mm()).to_string() + "mm",
            lo_core::escape_attr(&shape.style.stroke)
        ),
    }
}

pub fn serialize_drawing_document(drawing: &Drawing) -> String {
    let mut xml = lo_core::XmlBuilder::new();
    xml.declaration();
    xml.open("office:document-content", &content_root_attrs());
    xml.empty("office:scripts", &[]);
    xml.open("office:automatic-styles", &[]);
    xml.raw(
        "<style:style style:name=\"gr1\" style:family=\"graphic\">\
<style:graphic-properties draw:auto-grow-width=\"true\" draw:auto-grow-height=\"true\" \
fo:wrap-option=\"no-wrap\" draw:textarea-horizontal-align=\"center\" \
draw:textarea-vertical-align=\"middle\" fo:padding=\"1mm\"/>\
</style:style>",
    );
    xml.close();
    xml.open("office:body", &[]);
    xml.open("office:drawing", &[]);
    for (page_index, page) in drawing.pages.iter().enumerate() {
        xml.raw(&format!(
            "<draw:page draw:name=\"{}\" draw:id=\"page{}\" draw:master-page-name=\"Default\">",
            lo_core::escape_attr(&page.name),
            page_index + 1
        ));
        for element in &page.elements {
            let fragment = match element {
                DrawElement::TextBox(text_box) => format!(
                    "<draw:frame draw:style-name=\"gr1\" svg:x=\"{}\" svg:y=\"{}\" svg:width=\"{}\" svg:height=\"{}\"><draw:text-box><text:p>{}</text:p></draw:text-box></draw:frame>",
                    text_box.frame.origin.x.css(),
                    text_box.frame.origin.y.css(),
                    text_box.frame.size.width.css(),
                    text_box.frame.size.height.css(),
                    lo_core::escape_text(&text_box.text)
                ),
                DrawElement::Shape(shape) => shape_xml(shape),
                DrawElement::Image(image) => format!(
                    "<draw:frame svg:x=\"{}\" svg:y=\"{}\" svg:width=\"{}\" svg:height=\"{}\"><draw:image xlink:href=\"Pictures/{}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/></draw:frame>",
                    image.frame.origin.x.css(),
                    image.frame.origin.y.css(),
                    image.frame.size.width.css(),
                    image.frame.size.height.css(),
                    lo_core::escape_attr(&image.name)
                ),
            };
            xml.raw(&fragment);
        }
        xml.raw("</draw:page>");
    }
    xml.close();
    xml.close();
    xml.close();
    xml.finish()
}

pub fn save_drawing_document(path: impl AsRef<Path>, drawing: &Drawing) -> Result<()> {
    let content = serialize_drawing_document(drawing);
    let extras = image_extras(drawing.embedded_images());
    package_document(path, MIME_ODG, content, &drawing.meta, extras)
}
