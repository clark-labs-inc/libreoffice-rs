use lo_core::{ImageElement, Length, Presentation, RasterImage, Rect, Rgba, Slide, SlideElement};
use lo_zip::{ooxml_package, ZipEntry};

fn picture() -> ImageElement {
    ImageElement {
        name: "image1.png".into(), mime_type: "image/png".into(),
        data: RasterImage::new(8, 8, Rgba::rgba(12, 34, 200, 255)).encode_png(),
        frame: Rect::new(Length::mm(0.0), Length::mm(0.0), Length::mm(280.0), Length::mm(157.5)),
        alt: "Blue control".into(),
    }
}

fn deck() -> Presentation {
    let mut deck = Presentation::new("Image control");
    deck.slides.push(Slide { elements: vec![SlideElement::Image(picture())], ..Slide::default() });
    deck
}

#[test]
fn picture_pixels_survive_raster_export() {
    let pages = lo_impress::render_pages(&deck(), 96);
    let page = &pages[0];
    let center = ((page.height / 2 * page.width + page.width / 2) * 4) as usize;
    assert_eq!(&page.pixels[center..center + 4], &[12, 34, 200, 255]);
}

#[test]
fn picture_pixels_are_embedded_in_pdf() {
    let pdf = lo_impress::to_pdf(&deck());
    assert!(pdf.windows(b"/Subtype /Image".len()).any(|b| b == b"/Subtype /Image"));
    assert!(!pdf.windows(b"[image:".len()).any(|b| b == b"[image:"));
}

#[test]
fn pptx_picture_relationship_loads_bytes_and_slide_size() {
    let entries = vec![
        ZipEntry::new("ppt/presentation.xml", br#"<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#.to_vec()),
        ZipEntry::new("ppt/_rels/presentation.xml.rels", br#"<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/></Relationships>"#.to_vec()),
        ZipEntry::new("ppt/slides/slide1.xml", br#"<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="2" name="image1.png" descr="Blue control"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm></p:spPr></p:pic></p:spTree></p:cSld></p:sld>"#.to_vec()),
        ZipEntry::new("ppt/slides/_rels/slide1.xml.rels", br#"<Relationships><Relationship Id="rId2" Target="../media/image1.png" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/></Relationships>"#.to_vec()),
        ZipEntry::new("ppt/media/image1.png", picture().data),
    ];
    let loaded = lo_impress::from_pptx_bytes("control", &ooxml_package(&entries).unwrap()).unwrap();
    assert_eq!(loaded.slides[0].elements.len(), 1, "p:pic must not be silently discarded");
    let SlideElement::Image(image) = &loaded.slides[0].elements[0] else { panic!("picture element") };
    assert_eq!(image.data, picture().data);
    assert!((loaded.page_size.width.as_mm() - 338.6667).abs() < 0.01);
    assert!((loaded.page_size.height.as_mm() - 190.5).abs() < 0.01);
}

#[test]
fn transparent_picture_keeps_pixels_and_a_pdf_soft_mask() {
    let mut deck = deck();
    let SlideElement::Image(picture) = &mut deck.slides[0].elements[0] else { unreachable!() };
    picture.data = RasterImage::new(8, 8, Rgba::rgba(12, 34, 200, 128)).encode_png();
    let pdf = lo_impress::to_pdf(&deck);
    assert!(pdf.windows(b"/SMask".len()).any(|bytes| bytes == b"/SMask"));
    let pages = lo_impress::render_pages(&deck, 96);
    let page = &pages[0];
    let center = ((page.height / 2 * page.width + page.width / 2) * 4) as usize;
    assert_eq!(&page.pixels[center..center + 4], &[133, 144, 227, 255]);
    if let Ok(directory) = std::env::var("CLARK_RENDERING_EVIDENCE_DIR") {
        std::fs::create_dir_all(&directory).unwrap();
        std::fs::write(std::path::Path::new(&directory).join("transparent-picture.pdf"), &pdf).unwrap();
        std::fs::write(std::path::Path::new(&directory).join("transparent-picture.png"), page.encode_png()).unwrap();
    }
}

#[test]
fn jpeg_is_embedded_without_reencoding() {
    let mut deck = deck();
    let jpeg = RasterImage::new(16, 16, Rgba::rgba(12, 34, 200, 255)).encode_jpeg(90);
    let SlideElement::Image(picture) = &mut deck.slides[0].elements[0] else { unreachable!() };
    picture.data = jpeg.clone();
    picture.mime_type = "image/jpeg".into();
    let pdf = lo_impress::to_pdf(&deck);
    assert!(pdf.windows(jpeg.len()).any(|bytes| bytes == jpeg));
    let pages = lo_impress::render_pages(&deck, 96);
    let page = &pages[0];
    let center = ((page.height / 2 * page.width + page.width / 2) * 4) as usize;
    for (actual, expected) in page.pixels[center..center + 3].iter().zip([12_u8, 34, 200]) {
        assert!((*actual as i16 - expected as i16).abs() <= 5);
    }
}
