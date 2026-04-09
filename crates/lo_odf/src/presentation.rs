use std::path::Path;

use lo_core::{Presentation, Result, ShapeKind, SlideElement};

use crate::common::{content_root_attrs, image_extras, package_document, MIME_ODP};

fn text_box_body(text: &str) -> String {
    // Split on newlines so each line becomes its own paragraph, otherwise
    // LibreOffice renders the entire block as a single wrapped line. If every
    // non-empty line starts with a bullet glyph, emit a real text:list so the
    // PPTX exporter produces proper list items.
    let lines: Vec<&str> = text.split('\n').collect();
    let all_bullets = !lines.is_empty()
        && lines
            .iter()
            .filter(|l| !l.trim().is_empty())
            .all(|l| l.trim_start().starts_with("• ") || l.trim_start().starts_with("- "));
    if all_bullets && lines.iter().any(|l| !l.trim().is_empty()) {
        let mut out = String::from("<text:list text:style-name=\"L1\">");
        for line in lines {
            if line.trim().is_empty() {
                continue;
            }
            let stripped = line
                .trim_start()
                .trim_start_matches("• ")
                .trim_start_matches("- ");
            out.push_str("<text:list-item><text:p>");
            out.push_str(&lo_core::escape_text(stripped));
            out.push_str("</text:p></text:list-item>");
        }
        out.push_str("</text:list>");
        out
    } else {
        let mut out = String::new();
        for line in lines {
            out.push_str("<text:p>");
            out.push_str(&lo_core::escape_text(line));
            out.push_str("</text:p>");
        }
        out
    }
}

fn text_box_xml(text_box: &lo_core::TextBox) -> String {
    format!(
        "<draw:frame draw:style-name=\"gr1\" svg:x=\"{}\" svg:y=\"{}\" svg:width=\"{}\" svg:height=\"{}\"><draw:text-box>{}</draw:text-box></draw:frame>",
        text_box.frame.origin.x.css(),
        text_box.frame.origin.y.css(),
        text_box.frame.size.width.css(),
        text_box.frame.size.height.css(),
        text_box_body(&text_box.text)
    )
}

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

fn image_xml(image: &lo_core::ImageElement) -> String {
    format!(
        "<draw:frame svg:x=\"{}\" svg:y=\"{}\" svg:width=\"{}\" svg:height=\"{}\"><draw:image xlink:href=\"Pictures/{}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/></draw:frame>",
        image.frame.origin.x.css(),
        image.frame.origin.y.css(),
        image.frame.size.width.css(),
        image.frame.size.height.css(),
        lo_core::escape_attr(&image.name)
    )
}

pub fn serialize_presentation_document(presentation: &Presentation) -> String {
    let mut xml = lo_core::XmlBuilder::new();
    xml.declaration();
    xml.open("office:document-content", &content_root_attrs());
    xml.empty("office:scripts", &[]);
    // Automatic styles: a graphic style that enables text auto-grow + wrap
    // so text boxes in shapes don't clip, and a bullet list style L1 used by
    // text_box_body when it detects bullet lines.
    xml.open("office:automatic-styles", &[]);
    xml.raw(
        "<style:style style:name=\"gr1\" style:family=\"graphic\">\
<style:graphic-properties draw:auto-grow-height=\"true\" draw:auto-grow-width=\"false\" \
fo:wrap-option=\"wrap\" draw:textarea-horizontal-align=\"center\" \
draw:textarea-vertical-align=\"middle\"/>\
</style:style>",
    );
    xml.raw(
        "<text:list-style style:name=\"L1\">\
<text:list-level-style-bullet text:level=\"1\" text:bullet-char=\"•\">\
<style:list-level-properties text:space-before=\"6mm\" text:min-label-width=\"5mm\"/>\
</text:list-level-style-bullet>\
</text:list-style>",
    );
    xml.close();
    xml.open("office:body", &[]);
    xml.open("office:presentation", &[]);
    for (index, slide) in presentation.slides.iter().enumerate() {
        xml.raw(&format!(
            "<draw:page draw:name=\"{}\" draw:id=\"slide{}\" draw:master-page-name=\"Default\">",
            lo_core::escape_attr(&slide.name),
            index + 1
        ));
        for element in &slide.elements {
            let fragment = match element {
                SlideElement::TextBox(text_box) => text_box_xml(text_box),
                SlideElement::Shape(shape) => shape_xml(shape),
                SlideElement::Image(image) => image_xml(image),
            };
            xml.raw(&fragment);
        }
        if !slide.notes.is_empty() {
            xml.raw(&format!(
                "<presentation:notes><text:p>{}</text:p></presentation:notes>",
                lo_core::escape_text(&slide.notes.join("\n"))
            ));
        }
        xml.raw("</draw:page>");
    }
    xml.close();
    xml.close();
    xml.close();
    xml.finish()
}

pub fn save_presentation_document(
    path: impl AsRef<Path>,
    presentation: &Presentation,
) -> Result<()> {
    let content = serialize_presentation_document(presentation);
    let extras = image_extras(presentation.embedded_images());
    package_document(path, MIME_ODP, content, &presentation.meta, extras)
}
