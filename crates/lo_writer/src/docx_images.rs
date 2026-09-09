//! DrawingML/VML image import for DOCX packages.

use std::collections::BTreeMap;
use std::path::Path;

use lo_core::{ImageBlock, Length, Size, XmlNode};
use lo_zip::ZipArchive;

const EMUS_PER_MM: f32 = 36_000.0;

pub(crate) fn from_docx_node(
    node: &XmlNode,
    relationships: &BTreeMap<String, String>,
    zip: &ZipArchive,
) -> Vec<ImageBlock> {
    let mut images = Vec::new();
    let mut drawings = Vec::new();
    node.descendants_named("drawing", &mut drawings);
    for drawing in drawings {
        images.extend(images_from_drawing(drawing, relationships, zip));
    }

    let mut picts = Vec::new();
    node.descendants_named("pict", &mut picts);
    for pict in picts {
        let mut image_nodes = Vec::new();
        pict.descendants_named("imagedata", &mut image_nodes);
        for image_node in image_nodes {
            if let Some(image) = image_from_relationship(
                image_node.attr("id"),
                image_node.attr("title").unwrap_or_default(),
                None,
                relationships,
                zip,
            ) {
                images.push(image);
            }
        }
    }
    images
}

fn images_from_drawing(
    drawing: &XmlNode,
    relationships: &BTreeMap<String, String>,
    zip: &ZipArchive,
) -> Vec<ImageBlock> {
    let size = drawing_extent(drawing);
    let alt = drawing_alt_text(drawing);
    let mut blips = Vec::new();
    drawing.descendants_named("blip", &mut blips);
    blips
        .into_iter()
        .filter_map(|blip| {
            image_from_relationship(
                blip.attr("embed").or_else(|| blip.attr("link")),
                &alt,
                size,
                relationships,
                zip,
            )
        })
        .collect()
}

fn image_from_relationship(
    relationship_id: Option<&str>,
    alt: &str,
    size: Option<Size>,
    relationships: &BTreeMap<String, String>,
    zip: &ZipArchive,
) -> Option<ImageBlock> {
    let target = relationships.get(relationship_id?)?;
    if target.contains("://") {
        return None;
    }
    let data = zip.read(target).ok()?;
    let name = Path::new(target)
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("image.bin")
        .to_string();
    Some(ImageBlock {
        mime_type: image_mime_type(&name, &data).to_string(),
        data,
        alt: alt.split_whitespace().collect::<Vec<_>>().join(" "),
        size: size.unwrap_or_else(|| ImageBlock::default().size),
        name,
    })
}

fn drawing_extent(drawing: &XmlNode) -> Option<Size> {
    let mut extents = Vec::new();
    drawing.descendants_named("extent", &mut extents);
    let extent = extents.into_iter().find(|node| {
        node.attr("cx").and_then(|value| value.parse::<f32>().ok()).is_some()
            && node.attr("cy").and_then(|value| value.parse::<f32>().ok()).is_some()
    })?;
    let width = extent.attr("cx")?.parse::<f32>().ok()? / EMUS_PER_MM;
    let height = extent.attr("cy")?.parse::<f32>().ok()? / EMUS_PER_MM;
    if width <= 0.0 || height <= 0.0 {
        return None;
    }
    Some(Size::new(Length::mm(width), Length::mm(height)))
}

fn drawing_alt_text(drawing: &XmlNode) -> String {
    for name in ["docPr", "cNvPr"] {
        let mut nodes = Vec::new();
        drawing.descendants_named(name, &mut nodes);
        for node in nodes {
            for attribute in ["descr", "title"] {
                if let Some(value) = node.attr(attribute).filter(|value| !value.trim().is_empty()) {
                    return value.split_whitespace().collect::<Vec<_>>().join(" ");
                }
            }
        }
    }
    String::new()
}

fn image_mime_type(name: &str, data: &[u8]) -> &'static str {
    if data.starts_with(b"\x89PNG\r\n\x1a\n") {
        return "image/png";
    }
    if data.starts_with(b"\xff\xd8\xff") {
        return "image/jpeg";
    }
    if data.starts_with(b"RIFF") && data.get(8..12) == Some(b"WEBP") {
        return "image/webp";
    }
    match Path::new(name)
        .extension()
        .and_then(|extension| extension.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase()
        .as_str()
    {
        "svg" => "image/svg+xml",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        "emf" => "image/emf",
        "wmf" => "image/wmf",
        _ => "application/octet-stream",
    }
}
