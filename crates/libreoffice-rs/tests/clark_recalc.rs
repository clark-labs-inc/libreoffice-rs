use libreoffice_pure::{xlsx_recalc_bytes, xlsx_recalc_check_json};
use lo_zip::{ooxml_package, ZipArchive, ZipEntry};

fn minimal_xlsx(sheet_xml: &str) -> Vec<u8> {
    ooxml_package(&vec![
        ZipEntry::new(
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#.to_vec(),
        ),
        ZipEntry::new(
            "_rels/.rels",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#.to_vec(),
        ),
        ZipEntry::new(
            "xl/workbook.xml",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#.to_vec(),
        ),
        ZipEntry::new(
            "xl/_rels/workbook.xml.rels",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#.to_vec(),
        ),
        ZipEntry::new("xl/worksheets/sheet1.xml", sheet_xml.as_bytes().to_vec()),
    ])
    .expect("xlsx zip")
}

#[test]
fn recalc_rewrites_formula_cache_values() {
    let xlsx = minimal_xlsx(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1"><f>SUM(A1:B1)</f><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#,
    );
    let out = xlsx_recalc_bytes(&xlsx).expect("recalc");
    let zip = ZipArchive::new(&out).expect("zip");
    let sheet = zip.read_string("xl/worksheets/sheet1.xml").expect("sheet1.xml");
    assert!(sheet.contains("<f>SUM(A1:B1)</f>"));
    assert!(sheet.contains("<v>3</v>"));
    assert!(!zip.contains("xl/calcChain.xml"));
}

#[test]
fn recalc_check_json_reports_formula_errors() {
    let xlsx = minimal_xlsx(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>1/0</f><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#,
    );
    let json = xlsx_recalc_check_json(&xlsx).expect("json");
    assert!(json.contains("\"status\":\"error\""));
    assert!(json.contains("#DIV/0!"));
    assert!(json.contains("Sheet1!A1"));
}
