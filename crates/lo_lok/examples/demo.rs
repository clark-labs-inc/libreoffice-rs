//! Tiny lo_lok demo: open a Writer doc, insert text, render an SVG tile and
//! save it as PDF + DOCX. Mirrors the parity attempt's `lo_lok/examples/demo.rs`.
//!
//! Run with: `cargo run -p lo_lok --example demo`

use std::fs;
use std::sync::Arc;

use lo_lok::{DocumentKind, KitEvent, Office, TileRequest};
use lo_uno::UnoValue;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let office = Office::new();
    office.register_callback(Arc::new(|event: &KitEvent| {
        println!("event: {event:?}");
    }));

    let doc = office.open_empty(DocumentKind::Writer, "lo_lok demo")?;

    doc.execute_command(
        ".uno:AppendHeading",
        &[
            ("level", UnoValue::Int(1)),
            ("text", UnoValue::String("Hello from lo_lok".into())),
        ]
        .into_iter()
        .map(|(k, v)| (k.to_string(), v))
        .collect(),
    )?;
    doc.execute_command(
        ".uno:InsertText",
        &[(
            "text",
            UnoValue::String("This document was created via the lo_lok runtime.".into()),
        )]
        .into_iter()
        .map(|(k, v)| (k.to_string(), v))
        .collect(),
    )?;

    let pdf = doc.save_as("pdf")?;
    let docx = doc.save_as("docx")?;
    let tile = doc.render_tile(TileRequest {
        width: 800,
        height: 600,
    })?;

    fs::write("/tmp/lo_lok_demo.pdf", &pdf)?;
    fs::write("/tmp/lo_lok_demo.docx", &docx)?;
    fs::write("/tmp/lo_lok_demo.svg", &tile.bytes)?;

    println!("wrote /tmp/lo_lok_demo.pdf ({} bytes)", pdf.len());
    println!("wrote /tmp/lo_lok_demo.docx ({} bytes)", docx.len());
    println!("wrote /tmp/lo_lok_demo.svg ({} bytes)", tile.bytes.len());
    Ok(())
}
