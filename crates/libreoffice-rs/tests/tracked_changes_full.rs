use libreoffice_pure::accept_all_tracked_changes_docx_bytes;
use lo_zip::{ooxml_package, ZipArchive, ZipEntry};

fn package(entries: Vec<ZipEntry>) -> Vec<u8> {
    ooxml_package(&entries).expect("zip")
}

fn content_types() -> ZipEntry {
    ZipEntry::new(
        "[Content_Types].xml",
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#
            .to_vec(),
    )
}

#[test]
fn accepts_common_insertions_and_drops_common_deletions() {
    let bytes = package(vec![
        content_types(),
        ZipEntry::new(
            "word/document.xml",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>before </w:t></w:r>
      <w:ins><w:r><w:t>kept</w:t></w:r></w:ins>
      <w:del><w:r><w:delText>gone</w:delText></w:r></w:del>
      <w:r><w:t> after</w:t></w:r>
    </w:p>
  </w:body>
</w:document>"#
                .to_vec(),
        ),
        ZipEntry::new(
            "word/settings.xml",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:trackRevisions/>
</w:settings>"#
                .to_vec(),
        ),
    ]);

    let out = accept_all_tracked_changes_docx_bytes(&bytes).expect("accept changes");
    let zip = ZipArchive::new(&out).expect("out zip");
    let document = zip.read_string("word/document.xml").expect("document.xml");
    let settings = zip.read_string("word/settings.xml").expect("settings.xml");

    assert!(document.contains("before"));
    assert!(document.contains("kept"));
    assert!(document.contains("after"));
    assert!(!document.contains("<w:ins"));
    assert!(!document.contains("<w:del"));
    assert!(!document.contains("gone"));
    assert!(!settings.contains("trackRevisions"));
}

#[test]
fn preserves_live_comment_anchors_and_prunes_unreferenced_comments() {
    let bytes = package(vec![
        content_types(),
        ZipEntry::new(
            "word/document.xml",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>note</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:r><w:commentReference w:id="0"/></w:r>
    </w:p>
  </w:body>
</w:document>"#
                .to_vec(),
        ),
        ZipEntry::new(
            "word/comments.xml",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0"><w:p><w:r><w:t>keep me</w:t></w:r></w:p></w:comment>
  <w:comment w:id="1"><w:p><w:r><w:t>drop me</w:t></w:r></w:p></w:comment>
</w:comments>"#
                .to_vec(),
        ),
    ]);

    let out = accept_all_tracked_changes_docx_bytes(&bytes).expect("accept changes");
    let zip = ZipArchive::new(&out).expect("out zip");
    let document = zip.read_string("word/document.xml").expect("document.xml");
    let comments = zip.read_string("word/comments.xml").expect("comments.xml");

    assert!(document.contains("commentRangeStart"));
    assert!(document.contains("commentRangeEnd"));
    assert!(document.contains("commentReference"));
    assert!(comments.contains("keep me"));
    assert!(!comments.contains("drop me"));
    assert!(comments.contains("w:id=\"0\""));
    assert!(!comments.contains("w:id=\"1\""));
}

#[test]
fn drops_formatting_change_history_but_keeps_current_formatting_and_text() {
    let bytes = package(vec![
        content_types(),
        ZipEntry::new(
            "word/document.xml",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:rPrChange><w:rPr><w:i/></w:rPr></w:rPrChange>
        <w:t>bold</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#
                .to_vec(),
        ),
    ]);

    let out = accept_all_tracked_changes_docx_bytes(&bytes).expect("accept changes");
    let zip = ZipArchive::new(&out).expect("out zip");
    let document = zip.read_string("word/document.xml").expect("document.xml");

    assert!(document.contains("<w:b"));
    assert!(document.contains("bold"));
    assert!(!document.contains("rPrChange"));
}

#[test]
fn drops_deleted_rows_cells_and_move_from_but_keeps_move_to() {
    let bytes = package(vec![
        content_types(),
        ZipEntry::new(
            "word/document.xml",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:moveFrom><w:r><w:t>old</w:t></w:r></w:moveFrom>
      <w:moveTo><w:r><w:t>new</w:t></w:r></w:moveTo>
    </w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>keep</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr>
        <w:trPr><w:del/></w:trPr>
        <w:tc><w:p><w:r><w:t>drop-row</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:tcPr><w:cellDel/></w:tcPr>
          <w:p><w:r><w:t>drop-cell</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#
                .to_vec(),
        ),
    ]);

    let out = accept_all_tracked_changes_docx_bytes(&bytes).expect("accept changes");
    let zip = ZipArchive::new(&out).expect("out zip");
    let document = zip.read_string("word/document.xml").expect("document.xml");

    assert!(document.contains("new"));
    assert!(document.contains("keep"));
    assert!(!document.contains("old"));
    assert!(!document.contains("drop-row"));
    assert!(!document.contains("drop-cell"));
    assert!(!document.contains("moveFrom"));
}
