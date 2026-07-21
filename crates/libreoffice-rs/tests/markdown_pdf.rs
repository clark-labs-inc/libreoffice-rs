use std::fs;
use std::path::PathBuf;
use std::time::{SystemTime, UNIX_EPOCH};

use libreoffice_pure::{convert_path_bytes, markdown_to_pdf_bytes};
use lo_core::{parse_pdf, RasterImage, Rgba};

struct Fixture {
    dir: PathBuf,
    path: PathBuf,
    markdown: Vec<u8>,
}

impl Drop for Fixture {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.dir);
    }
}

fn fixture() -> Fixture {
    let nonce = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("time")
        .as_nanos();
    let dir = std::env::temp_dir().join(format!("libreoffice-markdown-pdf-{nonce}"));
    fs::create_dir_all(&dir).expect("create fixture dir");
    let mut image = RasterImage::new(24, 12, Rgba::rgba(102, 87, 217, 255));
    image.fill_rect(12, 0, 12, 12, Rgba::rgba(240, 164, 68, 255));
    fs::write(dir.join("chart.png"), image.encode_png()).expect("write image");
    fs::write(
        dir.join("badge.webp"),
        image.encode_webp().expect("encode WebP"),
    )
    .expect("write WebP");
    fs::write(
        dir.join("vector.svg"),
        br##"<svg xmlns="http://www.w3.org/2000/svg" width="640" height="180" viewBox="0 0 640 180">
<rect x="4" y="4" width="632" height="172" fill="#f7f4ff" stroke="#6d5bd0" stroke-width="4"/>
<circle cx="90" cy="90" r="44" fill="#f0a444" stroke="#4a356f" stroke-width="3"/>
<polygon points="180,130 240,50 300,130" fill="#6d5bd0" stroke="#342752"/>
<text x="325" y="102" font-size="26" font-weight="bold" fill="#342752">Native SVG vectors</text>
</svg>"##,
    )
    .expect("write SVG");
    let path = dir.join("report.md");
    let markdown = br#"# Quarterly Review

This keeps **strong text**, *emphasis*, and [a link](https://example.com).

Inline HTML keeps <span style="color:#6d5bd0;background-color:#fff1a8;font-weight:bold;text-decoration:underline">violet CSS styling</span>, while inline math keeps $x^2 + y^2$ readable.

> A structured callout with deliberate color.

1. First ordered item
2. Second ordered item

| Metric | Result |
| --- | ---: |
| Quality | 99% |

```rust
fn main() {
    println!("preserved");
}
```

![Violet and amber chart](chart.png)

![Lossless WebP badge](badge.webp)

![Native SVG illustration](vector.svg)

```mermaid
flowchart LR
A[Markdown] --> B{Rich parser}
B --> C[Tagged PDF]
```

$$E = \frac{mc^2}{1 + alpha}$$
"#
    .to_vec();
    fs::write(&path, &markdown).expect("write markdown");
    Fixture {
        dir,
        path,
        markdown,
    }
}

#[test]
fn markdown_pdf_preserves_structure_text_and_image_xobject() {
    let fixture = fixture();
    let pdf = markdown_to_pdf_bytes(&fixture.path, &fixture.markdown).expect("Markdown PDF");
    assert!(pdf.starts_with(b"%PDF-1.4"));
    assert!(pdf
        .windows(b"/Subtype /Image".len())
        .any(|window| window == b"/Subtype /Image"));
    assert!(pdf
        .windows(b"/Im1 Do".len())
        .any(|window| window == b"/Im1 Do"));
    assert!(pdf.windows(3).any(|window| window == [102, 87, 217]));
    assert!(pdf.windows(3).any(|window| window == [240, 164, 68]));
    assert!(pdf
        .windows(b"/Subtype /Link".len())
        .any(|window| window == b"/Subtype /Link"));
    assert!(pdf
        .windows(b"https://example.com".len())
        .any(|window| window == b"https://example.com"));
    for marker in [
        b"/StructTreeRoot".as_slice(),
        b"/MarkInfo".as_slice(),
        b"/Marked true".as_slice(),
        b"/S /Figure".as_slice(),
        b"/S /H1".as_slice(),
        b"/Alt (Native SVG illustration)".as_slice(),
    ] {
        assert!(pdf.windows(marker.len()).any(|window| window == marker));
    }
    assert!(pdf
        .windows(b"0.427 0.357 0.816 rg".len())
        .any(|window| window == b"0.427 0.357 0.816 rg"));
    assert!(
        pdf.windows(b"/Subtype /Image".len())
            .filter(|window| *window == b"/Subtype /Image")
            .count()
            >= 2
    );
    let parsed = parse_pdf(&pdf).expect("parse generated PDF");
    let text = parsed.extract_text();
    assert!(text.contains("Quarterly Review"), "{text}");
    assert!(text.contains("First ordered item"), "{text}");
    assert!(text.contains("Metric"), "{text}");
    assert!(text.contains("println"), "{text}");
    assert!(text.contains("violet CSS styling"), "{text}");
    assert!(text.contains("Native SVG vectors"), "{text}");
    assert!(text.contains("Markdown"), "{text}");
    assert!(text.contains("Tagged PDF"), "{text}");
    assert!(text.contains("mc²"), "{text}");
}

#[test]
fn generic_path_conversion_uses_the_asset_aware_markdown_pipeline() {
    let fixture = fixture();
    let pdf = convert_path_bytes(fixture.path.to_str().unwrap(), &fixture.markdown, "pdf")
        .expect("generic convert");
    assert!(pdf
        .windows(b"/Subtype /Image".len())
        .any(|window| window == b"/Subtype /Image"));
}
