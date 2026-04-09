//! Minimal PPTX (PresentationML) export.
//!
//! Generates a `.pptx` file with `[Content_Types].xml`, `_rels/.rels`,
//! `ppt/presentation.xml`, `ppt/_rels/presentation.xml.rels`, and one
//! `ppt/slides/slideN.xml` per slide. Each slide contains its text boxes as
//! `<p:sp>` shapes whose body holds the text-box text.

use lo_core::{escape_text, Presentation, Result, SlideElement};
use lo_zip::{ooxml_package, ZipEntry};

const XML_DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

fn render_slide(slide: &lo_core::Slide) -> String {
    let mut shapes = String::new();
    let mut shape_id = 2u32;
    for element in &slide.elements {
        if let SlideElement::TextBox(tb) = element {
            let mut paragraphs = String::new();
            for line in tb.text.lines() {
                paragraphs.push_str(&format!(
                    "<a:p><a:r><a:rPr lang=\"en-US\"/><a:t>{}</a:t></a:r></a:p>",
                    escape_text(line)
                ));
            }
            if paragraphs.is_empty() {
                paragraphs = "<a:p/>".to_string();
            }
            shapes.push_str(&format!(
                "<p:sp><p:nvSpPr><p:cNvPr id=\"{shape_id}\" name=\"TextBox{shape_id}\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</p:txBody></p:sp>"
            ));
            shape_id += 1;
        }
    }
    format!(
        "{XML_DECL}<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>"
    )
}

pub fn to_pptx(presentation: &Presentation) -> Result<Vec<u8>> {
    let mut entries: Vec<ZipEntry> = Vec::new();

    // Per-slide parts
    let mut slide_meta: Vec<(usize, String)> = Vec::new();
    for (idx, slide) in presentation.slides.iter().enumerate() {
        let path = format!("ppt/slides/slide{}.xml", idx + 1);
        entries.push(ZipEntry::new(
            path.clone(),
            render_slide(slide).into_bytes(),
        ));
        slide_meta.push((idx + 1, path));
    }

    // Slide ID list inside presentation.xml
    let mut sldid_list = String::new();
    for (idx, _) in &slide_meta {
        sldid_list.push_str(&format!(
            "<p:sldId id=\"{}\" r:id=\"rId{}\"/>",
            255 + idx,
            idx
        ));
    }
    let presentation_xml = format!(
        "{XML_DECL}<p:presentation xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:sldIdLst>{sldid_list}</p:sldIdLst><p:sldSz cx=\"9144000\" cy=\"6858000\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/></p:presentation>"
    );
    entries.push(ZipEntry::new(
        "ppt/presentation.xml",
        presentation_xml.into_bytes(),
    ));

    // Presentation rels
    let mut pres_rels = String::from(
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );
    for (idx, path) in &slide_meta {
        let rel_target = path.trim_start_matches("ppt/");
        pres_rels.push_str(&format!(
            "<Relationship Id=\"rId{idx}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"{rel_target}\"/>"
        ));
    }
    pres_rels.push_str("</Relationships>");
    entries.push(ZipEntry::new(
        "ppt/_rels/presentation.xml.rels",
        format!("{XML_DECL}{pres_rels}").into_bytes(),
    ));

    // Top-level package parts
    let mut content_types = String::from(
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
<Default Extension=\"xml\" ContentType=\"application/xml\"/>\
<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>",
    );
    for (idx, _) in &slide_meta {
        content_types.push_str(&format!(
            "<Override PartName=\"/ppt/slides/slide{idx}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>"
        ));
    }
    content_types.push_str("</Types>");
    entries.insert(
        0,
        ZipEntry::new(
            "[Content_Types].xml",
            format!("{XML_DECL}{content_types}").into_bytes(),
        ),
    );

    let rels = format!(
        "{XML_DECL}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rIdPres\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>\
</Relationships>"
    );
    entries.insert(1, ZipEntry::new("_rels/.rels", rels.into_bytes()));

    ooxml_package(&entries)
}
