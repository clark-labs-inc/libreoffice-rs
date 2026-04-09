//! Command-line entry point for the workspace.
//!
//! Subcommands wrap the high-level `save_as` and `load_bytes` helpers
//! exposed by each crate so that text and binary documents can be
//! converted from the shell.

use std::collections::BTreeMap;
use std::env;
use std::fs;

use lo_lok::{DocumentKind, Office};
use lo_uno::UnoValue;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        print_usage();
        return Ok(());
    }

    match args[1].as_str() {
        "writer-md" if args.len() >= 4 => {
            let input = fs::read_to_string(&args[2])?;
            let doc = lo_writer::from_markdown("document", &input);
            fs::write(&args[3], lo_writer::save_as(&doc, extension(&args[3]))?)?;
        }
        "writer-txt" if args.len() >= 4 => {
            let input = fs::read_to_string(&args[2])?;
            let doc = lo_writer::from_plain_text("document", &input);
            fs::write(&args[3], lo_writer::save_as(&doc, extension(&args[3]))?)?;
        }
        "writer-import" if args.len() >= 4 => {
            let input = fs::read(&args[2])?;
            let doc = lo_writer::load_bytes(&args[2], &input, extension(&args[2]))?;
            fs::write(&args[3], lo_writer::save_as(&doc, extension(&args[3]))?)?;
        }
        "calc-csv" if args.len() >= 4 => {
            let input = fs::read_to_string(&args[2])?;
            let wb = lo_calc::workbook_from_csv("workbook", "Sheet1", &input)?;
            fs::write(&args[3], lo_calc::save_as(&wb, extension(&args[3]))?)?;
        }
        "calc-import" if args.len() >= 4 => {
            let input = fs::read(&args[2])?;
            let wb = lo_calc::load_bytes(&args[2], &input, extension(&args[2]))?;
            fs::write(&args[3], lo_calc::save_as(&wb, extension(&args[3]))?)?;
        }
        "impress-demo" if args.len() >= 3 => {
            let deck = lo_impress::demo_presentation("Presentation");
            fs::write(&args[2], lo_impress::save_as(&deck, extension(&args[2]))?)?;
        }
        "impress-import" if args.len() >= 4 => {
            let input = fs::read(&args[2])?;
            let deck = lo_impress::load_bytes(&args[2], &input, extension(&args[2]))?;
            fs::write(&args[3], lo_impress::save_as(&deck, extension(&args[3]))?)?;
        }
        "draw-demo" if args.len() >= 3 => {
            let drawing = lo_draw::demo_drawing("Drawing");
            fs::write(&args[2], lo_draw::save_as(&drawing, extension(&args[2]))?)?;
        }
        "draw-import" if args.len() >= 4 => {
            let input = fs::read(&args[2])?;
            let drawing = lo_draw::load_bytes(&args[2], &input, extension(&args[2]))?;
            fs::write(&args[3], lo_draw::save_as(&drawing, extension(&args[3]))?)?;
        }
        "math" if args.len() >= 4 => {
            let doc = lo_math::from_latex("formula", &args[2])?;
            fs::write(&args[3], lo_math::save_as(&doc, extension(&args[3]))?)?;
        }
        "math-import" if args.len() >= 4 => {
            let input = fs::read(&args[2])?;
            let doc = lo_math::load_bytes("formula", &input, extension(&args[2]))?;
            fs::write(&args[3], lo_math::save_as(&doc, extension(&args[3]))?)?;
        }
        "base-csv" if args.len() >= 5 => {
            let input = fs::read_to_string(&args[2])?;
            let db = lo_base::database_from_csv("db", "data", &input)?;
            let result = lo_base::execute_select(&db, &args[3])?;
            fs::write(&args[4], result_to_csv(&result))?;
        }
        "base-import" if args.len() >= 5 => {
            let input = fs::read(&args[2])?;
            let db = lo_base::load_bytes("db", &input, extension(&args[2]), Some("data"))?;
            let result = lo_base::execute_select(&db, &args[3])?;
            fs::write(&args[4], result_to_csv(&result))?;
        }
        "office-demo" if args.len() >= 3 => {
            let out_dir = &args[2];
            fs::create_dir_all(out_dir)?;
            let office = Office::new();
            let writer = office.open_empty(DocumentKind::Writer, "Writer")?;
            writer.execute_command(
                ".uno:InsertText",
                &BTreeMap::from([(
                    "text".to_string(),
                    UnoValue::String("Hello from Office runtime".to_string()),
                )]),
            )?;
            fs::write(format!("{out_dir}/writer.odt"), writer.save_as("odt")?)?;

            let calc = office.open_empty(DocumentKind::Calc, "Calc")?;
            calc.execute_command(
                ".uno:SetCell",
                &BTreeMap::from([
                    ("row".to_string(), UnoValue::Int(0)),
                    ("col".to_string(), UnoValue::Int(0)),
                    ("value".to_string(), UnoValue::String("41".to_string())),
                ]),
            )?;
            calc.execute_command(
                ".uno:SetCell",
                &BTreeMap::from([
                    ("row".to_string(), UnoValue::Int(0)),
                    ("col".to_string(), UnoValue::Int(1)),
                    ("value".to_string(), UnoValue::String("=A1+1".to_string())),
                ]),
            )?;
            fs::write(format!("{out_dir}/calc.ods"), calc.save_as("ods")?)?;

            let impress = office.open_empty(DocumentKind::Impress, "Impress")?;
            impress.execute_command(
                ".uno:InsertSlide",
                &BTreeMap::from([("title".to_string(), UnoValue::String("Slide 1".to_string()))]),
            )?;
            impress.execute_command(
                ".uno:InsertBullets",
                &BTreeMap::from([
                    ("slide".to_string(), UnoValue::Int(1)),
                    (
                        "items".to_string(),
                        UnoValue::String("Pure Rust|Native runtime|ODP export".to_string()),
                    ),
                ]),
            )?;
            fs::write(format!("{out_dir}/slides.odp"), impress.save_as("odp")?)?;
        }
        _ => print_usage(),
    }

    Ok(())
}

fn extension(path: &str) -> &str {
    path.rsplit('.').next().unwrap_or(path)
}

fn result_to_csv(result: &lo_base::QueryResult) -> Vec<u8> {
    let mut out = String::new();
    out.push_str(&result.columns.join(","));
    out.push('\n');
    for row in &result.rows {
        let cells: Vec<String> = row
            .iter()
            .map(|value| match value {
                lo_core::DbValue::Null => String::new(),
                lo_core::DbValue::Integer(v) => v.to_string(),
                lo_core::DbValue::Float(v) => v.to_string(),
                lo_core::DbValue::Text(v) => csv_escape(v),
                lo_core::DbValue::Bool(v) => v.to_string(),
            })
            .collect();
        out.push_str(&cells.join(","));
        out.push('\n');
    }
    out.into_bytes()
}

fn csv_escape(value: &str) -> String {
    if value.contains(',') || value.contains('"') || value.contains('\n') {
        let escaped = value.replace('"', "\"\"");
        format!("\"{escaped}\"")
    } else {
        value.to_string()
    }
}

const USAGE: &str = "usage:
  lo_cli writer-md input.md output.{odt|docx|html|pdf|svg|txt}
  lo_cli writer-txt input.txt output.{odt|docx|html|pdf|svg|txt}
  lo_cli writer-import input.{docx|odt|html|txt} output.{...}
  lo_cli calc-csv input.csv output.{ods|xlsx|html|pdf|svg|csv}
  lo_cli calc-import input.{xlsx|ods|csv} output.{...}
  lo_cli impress-demo output.{odp|pptx|html|pdf|svg}
  lo_cli impress-import input.{pptx|odp|txt} output.{...}
  lo_cli draw-demo output.{odg|svg|pdf}
  lo_cli draw-import input.{odg|svg} output.{...}
  lo_cli math '\\frac{x^2}{y}' output.{mathml|svg|pdf}
  lo_cli math-import input.{mathml|mml|odf|txt} output.{...}
  lo_cli base-csv input.csv \"SELECT * FROM data\" output.csv
  lo_cli base-import input.{odb|csv} \"SELECT * FROM data\" output.csv
  lo_cli office-demo out_dir";

fn print_usage() {
    eprintln!("{USAGE}");
}
