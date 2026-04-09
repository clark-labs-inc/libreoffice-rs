//! Quick microbenchmark for the import paths.
//!
//! Runs each importer N times against fixtures under `/tmp/lo_cli_demo/`
//! and reports total + per-iteration time. Not a substitute for criterion;
//! intended for smoke-checking that the round trips stay in the
//! expected order of magnitude.

use std::time::Instant;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let dir = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "/tmp/lo_cli_demo".to_string());

    bench("zip parse + inflate big.docx", 200, || {
        let bytes = std::fs::read(format!("{dir}/big.docx")).unwrap();
        let _ = lo_zip::ZipArchive::new(&bytes).unwrap();
    });

    bench("DOCX -> TextDocument (big.docx, 2k paras)", 50, || {
        let bytes = std::fs::read(format!("{dir}/big.docx")).unwrap();
        let _ = lo_writer::from_docx_bytes("big", &bytes).unwrap();
    });

    bench("DOCX -> TextDocument (real.docx)", 5_000, || {
        let bytes = std::fs::read(format!("{dir}/real.docx")).unwrap();
        let _ = lo_writer::from_docx_bytes("r", &bytes).unwrap();
    });

    bench("XLSX -> Workbook (sheet.xlsx)", 5_000, || {
        let bytes = std::fs::read(format!("{dir}/sheet.xlsx")).unwrap();
        let _ = lo_calc::from_xlsx_bytes("s", &bytes).unwrap();
    });

    bench("ODS -> Workbook (sheet.ods)", 5_000, || {
        let bytes = std::fs::read(format!("{dir}/sheet.ods")).unwrap();
        let _ = lo_calc::from_ods_bytes("s", &bytes).unwrap();
    });

    bench("PPTX -> Presentation (deck.pptx)", 5_000, || {
        let bytes = std::fs::read(format!("{dir}/deck.pptx")).unwrap();
        let _ = lo_impress::from_pptx_bytes("d", &bytes).unwrap();
    });

    bench("ODP -> Presentation (deck.odp)", 5_000, || {
        let bytes = std::fs::read(format!("{dir}/deck.odp")).unwrap();
        let _ = lo_impress::from_odp_bytes("d", &bytes).unwrap();
    });

    bench("ODG -> Drawing (draw.odg)", 5_000, || {
        let bytes = std::fs::read(format!("{dir}/draw.odg")).unwrap();
        let _ = lo_draw::from_odg_bytes("d", &bytes).unwrap();
    });

    bench("MathML parser (formula.mathml)", 20_000, || {
        let bytes = std::fs::read(format!("{dir}/formula.mathml")).unwrap();
        let _ = lo_math::load_bytes("f", &bytes, "mathml").unwrap();
    });

    Ok(())
}

fn bench(label: &str, iters: u32, mut f: impl FnMut()) {
    // warmup
    f();
    let start = Instant::now();
    for _ in 0..iters {
        f();
    }
    let elapsed = start.elapsed();
    let per = elapsed / iters;
    println!(
        "{label:<48} {iters:>6} iters  total {:>10.3?}  per {:>10.3?}",
        elapsed, per
    );
}
