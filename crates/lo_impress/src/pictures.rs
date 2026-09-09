//! Embedded picture bytes shared by PPTX import and rendering.
use std::collections::BTreeMap;
use lo_core::{ImageElement, Length, LoError, RasterImage, Rect, Result, Rgba, XmlNode};
use lo_zip::ZipArchive;

pub(crate) fn decode(bytes: &[u8]) -> Result<image::RgbaImage> {
    image::load_from_memory(bytes).map(|image| image.to_rgba8())
        .map_err(|error| LoError::Parse(format!("could not decode presentation picture: {error}")))
}

pub(crate) fn load(node: &XmlNode, zip: &ZipArchive, relationships: &BTreeMap<String, String>) -> Result<ImageElement> {
    let id = node.child("blipFill").and_then(|fill| fill.child("blip"))
        .and_then(|blip| blip.attr("r:embed"))
        .ok_or_else(|| LoError::Unsupported("PPTX picture requires an embedded image relationship".into()))?;
    let path = relationships.get(id)
        .ok_or_else(|| LoError::Parse(format!("missing PPTX image relationship {id}")))?;
    let mut data = zip.read(path)?;
    // A missing or unsupported image must fail import instead of producing a
    // successful-looking blank preview that sends the author into repair loops.
    let mut pixels = decode(&data)?;
    let properties = node.child("nvPicPr").and_then(|node| node.child("cNvPr"));
    let transform = node.child("spPr").and_then(|node| node.child("xfrm"));
    let offset = transform.and_then(|node| node.child("off"));
    let extent = transform.and_then(|node| node.child("ext"));
    let mm = |node: Option<&XmlNode>, key: &str, default: f32| {
        node.and_then(|node| node.attr(key)).and_then(|value| value.parse::<f32>().ok())
            .map(|value| value / 36_000.0).unwrap_or(default)
    };
    let mut frame = Rect::new(Length::mm(mm(offset, "x", 0.0)), Length::mm(mm(offset, "y", 0.0)),
        Length::mm(mm(extent, "cx", 80.0)), Length::mm(mm(extent, "cy", 20.0)));
    let mut changed = false;
    if let Some(crop) = node.child("blipFill").and_then(|fill| fill.child("srcRect")) {
        let fraction = |key: &str| -> Result<f64> {
            let value = crop.attr(key).unwrap_or("0").parse::<f64>()
                .map_err(|_| LoError::Parse(format!("invalid picture crop {key}")))? / 100_000.0;
            if !(0.0..1.0).contains(&value) {
                return Err(LoError::Unsupported("picture crop outside the source image".into()));
            }
            Ok(value)
        };
        let (left, top, right, bottom) = (fraction("l")?, fraction("t")?, fraction("r")?, fraction("b")?);
        if left + right >= 1.0 || top + bottom >= 1.0 {
            return Err(LoError::Parse("picture crop removes the whole image".into()));
        }
        if [left, top, right, bottom].iter().any(|value| *value > 0.0) {
            let x = (left * pixels.width() as f64).round() as u32;
            let y = (top * pixels.height() as f64).round() as u32;
            let width = ((1.0 - left - right) * pixels.width() as f64).round() as u32;
            let height = ((1.0 - top - bottom) * pixels.height() as f64).round() as u32;
            if width == 0 || height == 0 || x + width > pixels.width() || y + height > pixels.height() {
                return Err(LoError::Parse("picture crop has no valid pixel extent".into()));
            }
            pixels = image::imageops::crop_imm(&pixels, x, y, width, height).to_image();
            changed = true;
        }
    }
    if let Some(transform) = transform {
        for (attribute, horizontal) in [("flipH", true), ("flipV", false)] {
            if matches!(transform.attr(attribute), Some("1" | "true")) {
                pixels = if horizontal { image::imageops::flip_horizontal(&pixels) }
                    else { image::imageops::flip_vertical(&pixels) };
                changed = true;
            }
        }
    }
    let angle = transform.and_then(|node| node.attr("rot")).unwrap_or("0")
        .parse::<f64>().map_err(|_| LoError::Parse("invalid picture rotation".into()))? / 60_000.0;
    if !angle.is_finite() { return Err(LoError::Parse("invalid picture rotation".into())); }
    if angle.rem_euclid(360.0).abs() > f64::EPSILON {
        (pixels, frame) = rotate(pixels, frame, angle)?;
        changed = true;
    }
    if changed {
        let mut encoded = std::io::Cursor::new(Vec::new());
        image::DynamicImage::ImageRgba8(pixels).write_to(&mut encoded, image::ImageFormat::Png)
            .map_err(|error| LoError::Parse(error.to_string()))?;
        data = encoded.into_inner();
    }
    let format = image::guess_format(&data).map_err(|error| LoError::Parse(error.to_string()))?;
    Ok(ImageElement {
        name: properties.and_then(|node| node.attr("name")).unwrap_or(path).into(),
        alt: properties.and_then(|node| node.attr("descr")).unwrap_or("").into(),
        mime_type: format.to_mime_type().into(), data,
        frame,
    })
}

pub(crate) fn paint(page: &mut RasterImage, picture: &ImageElement, dpi: u32) {
    let Ok(image) = decode(&picture.data) else { return };
    let scale = |mm: f32| (mm * dpi as f32 / 25.4).round() as i32;
    let x = scale(picture.frame.origin.x.as_mm());
    let y = scale(picture.frame.origin.y.as_mm());
    let width = scale(picture.frame.size.width.as_mm());
    let height = scale(picture.frame.size.height.as_mm());
    if width <= 0 || height <= 0 { return; }
    // Clip before iterating so partially off-slide images retain their transform.
    for py in y.max(0)..(y.saturating_add(height)).min(page.height as i32) {
        for px in x.max(0)..(x.saturating_add(width)).min(page.width as i32) {
            let sx = ((px - x) as u64 * image.width() as u64 / width as u64) as u32;
            let sy = ((py - y) as u64 * image.height() as u64 / height as u64) as u32;
            let color = image.get_pixel(sx, sy).0;
            page.blend_pixel(px, py, Rgba::rgba(color[0], color[1], color[2], color[3]));
        }
    }
}

// Bake picture rotation into transparent pixels and its bounding frame. PDF and
// raster backends then consume the same geometry, including non-square frames.
fn rotate(source: image::RgbaImage, frame: Rect, angle: f64) -> Result<(image::RgbaImage, Rect)> {
    let width = frame.size.width.as_mm() as f64;
    let height = frame.size.height.as_mm() as f64;
    if !width.is_finite() || !height.is_finite() || width <= 0.0 || height <= 0.0 {
        return Err(LoError::Parse("picture has invalid dimensions".into()));
    }
    let (sin, cos) = angle.rem_euclid(360.0).to_radians().sin_cos();
    let new_width = width * cos.abs() + height * sin.abs();
    let new_height = width * sin.abs() + height * cos.abs();
    let density = (source.width() as f64 / width).max(source.height() as f64 / height);
    let px_width = (new_width * density).ceil().max(1.0) as u32;
    let px_height = (new_height * density).ceil().max(1.0) as u32;
    let mut target = image::RgbaImage::new(px_width, px_height);
    for (x, y, pixel) in target.enumerate_pixels_mut() {
        let tx = (x as f64 + 0.5) / px_width as f64 * new_width - new_width / 2.0;
        let ty = (y as f64 + 0.5) / px_height as f64 * new_height - new_height / 2.0;
        let sx = cos * tx + sin * ty + width / 2.0;
        let sy = -sin * tx + cos * ty + height / 2.0;
        if sx >= 0.0 && sy >= 0.0 && sx < width && sy < height {
            *pixel = *source.get_pixel((sx / width * source.width() as f64) as u32,
                (sy / height * source.height() as f64) as u32);
        }
    }
    Ok((target, Rect::new(
        Length::mm((frame.origin.x.as_mm() as f64 + (width - new_width) / 2.0) as f32),
        Length::mm((frame.origin.y.as_mm() as f64 + (height - new_height) / 2.0) as f32),
        Length::mm(new_width as f32), Length::mm(new_height as f32),
    )))
}

/// Apply the group's child-coordinate mapping after nested elements are parsed.
pub(crate) fn place_group(node: &XmlNode, elements: &mut [lo_core::SlideElement]) -> Result<()> {
    let Some(transform) = node.child("grpSpPr").and_then(|node| node.child("xfrm")) else { return Ok(()) };
    if transform.attr("rot").unwrap_or("0") != "0"
        || matches!(transform.attr("flipH"), Some("1" | "true"))
        || matches!(transform.attr("flipV"), Some("1" | "true")) {
        return Err(LoError::Unsupported("rotated or reflected presentation groups".into()));
    }
    let value = |child: &str, attribute: &str, fallback: f32| {
        transform.child(child).and_then(|node| node.attr(attribute))
            .and_then(|value| value.parse::<f32>().ok()).unwrap_or(fallback)
    };
    let x = value("off", "x", 0.0) / 36_000.0;
    let y = value("off", "y", 0.0) / 36_000.0;
    let child_x = value("chOff", "x", 0.0) / 36_000.0;
    let child_y = value("chOff", "y", 0.0) / 36_000.0;
    let sx = value("ext", "cx", 1.0) / value("chExt", "cx", 1.0);
    let sy = value("ext", "cy", 1.0) / value("chExt", "cy", 1.0);
    if !sx.is_finite() || !sy.is_finite() || sx <= 0.0 || sy <= 0.0 {
        return Err(LoError::Parse("invalid presentation group extent".into()));
    }
    for element in elements {
        let frame = match element {
            lo_core::SlideElement::Image(image) => &mut image.frame,
            lo_core::SlideElement::TextBox(text) => &mut text.frame,
            lo_core::SlideElement::Shape(shape) => &mut shape.frame,
        };
        *frame = Rect::new(Length::mm(x + (frame.origin.x.as_mm() - child_x) * sx),
            Length::mm(y + (frame.origin.y.as_mm() - child_y) * sy),
            Length::mm(frame.size.width.as_mm() * sx), Length::mm(frame.size.height.as_mm() * sy));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn quarter_turn_preserves_center_and_pixel_orientation() {
        let source = image::RgbaImage::from_fn(4, 2, |x, _| {
            if x < 2 { image::Rgba([240, 10, 10, 255]) } else { image::Rgba([10, 10, 240, 255]) }
        });
        let frame = Rect::new(Length::mm(10.0), Length::mm(20.0), Length::mm(4.0), Length::mm(2.0));
        let (turned, frame) = rotate(source, frame, 90.0).unwrap();
        assert!((frame.origin.x.as_mm() - 11.0).abs() < 0.001);
        assert!((frame.origin.y.as_mm() - 19.0).abs() < 0.001);
        assert!((frame.size.width.as_mm() - 2.0).abs() < 0.001);
        assert!((frame.size.height.as_mm() - 4.0).abs() < 0.001);
        assert_eq!(turned.get_pixel(turned.width() / 2, 0).0, [240, 10, 10, 255]);
        assert_eq!(turned.get_pixel(turned.width() / 2, turned.height() - 1).0, [10, 10, 240, 255]);
    }

    #[test]
    fn arbitrary_rotation_has_transparent_corners() {
        let source = image::RgbaImage::from_pixel(20, 10, image::Rgba([240, 10, 10, 255]));
        let frame = Rect::new(Length::mm(0.0), Length::mm(0.0), Length::mm(20.0), Length::mm(10.0));
        let (turned, _) = rotate(source, frame, 45.0).unwrap();
        assert_eq!(turned.get_pixel(0, 0).0[3], 0);
        assert_eq!(turned.get_pixel(turned.width() / 2, turned.height() / 2).0, [240, 10, 10, 255]);
    }

    #[test]
    fn crop_and_flip_use_embedded_pixels() {
        let mut source = RasterImage::new(4, 2, Rgba::rgba(10, 10, 240, 255));
        source.fill_rect(0, 0, 2, 2, Rgba::rgba(240, 10, 10, 255));
        let bytes = lo_zip::ooxml_package(&[lo_zip::ZipEntry::new("ppt/media/p.png", source.encode_png())]).unwrap();
        let zip = ZipArchive::new(&bytes).unwrap();
        let relationships = BTreeMap::from([("rId1".into(), "ppt/media/p.png".into())]);
        let node = lo_core::parse_xml_document(r#"<p:pic xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:blipFill><a:blip r:embed="rId1"/><a:srcRect l="25000" r="0"/></p:blipFill><p:spPr><a:xfrm flipH="1"><a:off x="0" y="0"/><a:ext cx="360000" cy="360000"/></a:xfrm></p:spPr></p:pic>"#).unwrap();
        let picture = load(&node, &zip, &relationships).unwrap();
        let pixels = decode(&picture.data).unwrap();
        assert_eq!(pixels.dimensions(), (3, 2));
        assert_eq!(pixels.get_pixel(0, 0).0, [10, 10, 240, 255]);
        assert_eq!(pixels.get_pixel(2, 0).0, [240, 10, 10, 255]);
    }
    #[test]
    fn group_child_coordinates_map_to_slide_coordinates() {
        let group = lo_core::parse_xml_document(r#"<p:grpSp xmlns:p="p" xmlns:a="a"><p:grpSpPr><a:xfrm><a:off x="3600000" y="7200000"/><a:ext cx="72000" cy="108000"/><a:chOff x="180000" y="360000"/><a:chExt cx="36000" cy="36000"/></a:xfrm></p:grpSpPr></p:grpSp>"#).unwrap();
        let mut elements = vec![lo_core::SlideElement::Image(ImageElement {
            name: "picture".into(), alt: String::new(), mime_type: "image/png".into(), data: Vec::new(),
            frame: Rect::new(Length::mm(10.0), Length::mm(20.0), Length::mm(30.0), Length::mm(40.0)),
        })];
        place_group(&group, &mut elements).unwrap();
        let lo_core::SlideElement::Image(image) = &elements[0] else { unreachable!() };
        assert_eq!(image.frame.origin.x.as_mm(), 110.0);
        assert_eq!(image.frame.origin.y.as_mm(), 230.0);
        assert_eq!(image.frame.size.width.as_mm(), 60.0);
        assert_eq!(image.frame.size.height.as_mm(), 120.0);
    }

}
