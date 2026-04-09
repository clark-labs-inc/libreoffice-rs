//! Sanity-check that the lo_zip deflate decoder reads a method=8 archive
//! produced by Python's zipfile module (matches what real DOCX/XLSX use).

use std::path::PathBuf;

use lo_zip::ZipArchive;

#[test]
fn reads_real_deflated_archive() {
    let path = PathBuf::from("/tmp/lo_cli_demo/deflated.zip");
    if !path.exists() {
        eprintln!("skipping: {} not present", path.display());
        return;
    }
    let bytes = std::fs::read(&path).expect("read deflated.zip");
    let zip = ZipArchive::new(&bytes).expect("parse zip");
    let hello = zip.read_string("hello.txt").expect("read hello.txt");
    assert_eq!(hello, "Hello world! ".repeat(200));
    let xml = zip.read_string("content.xml").expect("read content.xml");
    assert!(xml.starts_with("<?xml"));
    assert_eq!(xml.matches("<a>x</a>").count(), 500);
}
