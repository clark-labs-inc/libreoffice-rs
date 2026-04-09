use libreoffice_pure::accept_all_tracked_changes_docx_bytes;
use lo_zip::{ooxml_package, ZipArchive, ZipEntry};

fn package(entries: Vec<ZipEntry>) -> Vec<u8> {
    ooxml_package(&entries).expect("zip")
}

#[test]
fn accepts_insertions_and_drops_comment_markers() {
    let bytes = package(vec![
        ZipEntry::new(
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>"#.to_vec(),
        ),
        ZipEntry::new(
            "word/document.xml",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>before </w:t></w:r>
      <w:ins><w:r><w:t>kept</w:t></w:r></w:ins>
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t> note </w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:del><w:r><w:delText>gone</w:delText></w:r></w:del>
    </w:p>
  </w:body>
</w:document>"#.to_vec(),
        ),
        ZipEntry::new(
            "word/settings.xml",
            br#"<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:trackRevisions/></w:settings>"#.to_vec(),
        ),
    ]);

    let out = accept_all_tracked_changes_docx_bytes(&bytes).expect("accept changes");
    let zip = ZipArchive::new(&out).expect("out zip");
    let document = zip.read_string("word/document.xml").expect("document.xml");
    let settings = zip.read_string("word/settings.xml").expect("settings.xml");
    assert!(document.contains(">before <"));
    assert!(document.contains(">kept<"));
    assert!(document.contains("> note <"));
    // 0.4.0+: live comment anchors are preserved (orphans are pruned
    // from `word/comments.xml`), so the range markers can stay.
    assert!(!document.contains("delText"));
    assert!(!settings.contains("trackRevisions"));
}

#[test]
fn drops_formatting_change_history_but_keeps_current_run() {
    let bytes = package(vec![
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
</w:document>"#.to_vec(),
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
fn removes_table_rows_marked_deleted() {
    let bytes = package(vec![
        ZipEntry::new(
            "word/document.xml",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>keep</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr>
        <w:trPr><w:del/></w:trPr>
        <w:tc><w:p><w:r><w:t>drop</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>"#.to_vec(),
        ),
    ]);

    let out = accept_all_tracked_changes_docx_bytes(&bytes).expect("accept changes");
    let zip = ZipArchive::new(&out).expect("out zip");
    let document = zip.read_string("word/document.xml").expect("document.xml");
    assert!(document.contains("keep"));
    assert!(!document.contains("drop"));
    assert!(!document.contains("<w:del"));
}
