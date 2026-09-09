use libreoffice_pure::docx_to_pdf_bytes;
use lo_core::{Block, RasterImage, Rgba};
use lo_zip::{ooxml_package, ZipEntry};

fn drawing(relationship_id: &str, name: &str, description: &str, cx: u32, cy: u32) -> String {
    format!(
        r#"<w:r><w:drawing><wp:inline><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="1" name="{name}" descr="{description}"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="{relationship_id}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
    )
}

fn image_docx() -> Vec<u8> {
    let first = RasterImage::new(8, 6, Rgba::rgba(220, 20, 60, 255)).encode_png();
    let second = RasterImage::new(5, 9, Rgba::rgba(30, 144, 255, 255)).encode_png();
    let first_drawing = drawing("rId1", "Schedule", "Interview rotation map", 1_800_000, 900_000);
    let second_drawing = drawing("rId2", "Logo", "Company logo", 720_000, 360_000);
    let document = format!(
        r#"<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
 <w:body>
  <w:p><w:r><w:t>Before image</w:t></w:r>{first_drawing}</w:p>
  <w:tbl><w:tr><w:tc><w:p>{second_drawing}</w:p></w:tc></w:tr></w:tbl>
 </w:body>
</w:document>"#
    );
    let relationships = r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/schedule.png"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>
</Relationships>"#;
    let package_relationships = r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;
    let content_types = r#"<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Default Extension="png" ContentType="image/png"/>
 <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;
    ooxml_package(&[
        ZipEntry::new("[Content_Types].xml", content_types.as_bytes().to_vec()),
        ZipEntry::new("_rels/.rels", package_relationships.as_bytes().to_vec()),
        ZipEntry::new("word/document.xml", document.into_bytes()),
        ZipEntry::new(
            "word/_rels/document.xml.rels",
            relationships.as_bytes().to_vec(),
        ),
        ZipEntry::new("word/media/schedule.png", first),
        ZipEntry::new("word/media/logo.png", second),
    ])
    .expect("build image DOCX")
}

#[test]
fn docx_drawing_media_reaches_writer_model_and_pdf() {
    let bytes = image_docx();
    let document = lo_writer::from_docx_bytes("images", &bytes).expect("import DOCX");
    let images = document
        .body
        .iter()
        .filter_map(|block| match block {
            Block::Image(image) => Some(image),
            _ => None,
        })
        .collect::<Vec<_>>();
    assert_eq!(images.len(), 2, "paragraph and table drawings must import");
    assert_eq!(images[0].name, "schedule.png");
    assert_eq!(images[0].mime_type, "image/png");
    assert_eq!(images[0].alt, "Interview rotation map");
    assert!((images[0].size.width.as_mm() - 50.0).abs() < 0.01);
    assert!((images[0].size.height.as_mm() - 25.0).abs() < 0.01);
    assert_eq!(images[1].name, "logo.png");

    let pdf = docx_to_pdf_bytes(&bytes).expect("render DOCX to PDF");
    assert!(pdf.starts_with(b"%PDF"));
    assert!(
        pdf.windows(b"/Subtype /Image".len())
            .filter(|window| *window == b"/Subtype /Image")
            .count()
            >= 2,
        "both imported images must be embedded in the PDF"
    );
}
