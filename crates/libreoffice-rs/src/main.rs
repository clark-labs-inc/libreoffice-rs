use std::env;
use std::fs;
use std::path::Path;

use std::collections::BTreeMap;

use lo_app::DesktopApp;
use lo_base::{database_from_csv, execute_select};
use lo_calc::{evaluate_formula, save_ods, workbook_from_csv_opts};
use lo_core::{CellValue, Result, Sheet, TextDocument};
use lo_draw::demo_drawing;
use lo_impress::demo_presentation;
use lo_math::{from_latex, to_mathml_string};
use lo_uno::{ServiceRegistry, UnoValue};
use lo_writer::{from_markdown, save_odt};
use lo_zip::list_entries;

fn main() {
    if let Err(error) = run() {
        eprintln!("error: {error}");
        std::process::exit(1);
    }
}

fn run() -> Result<()> {
    let args: Vec<String> = env::args().collect();
    if args.len() < 2 {
        print_usage();
        return Ok(());
    }
    match args[1].as_str() {
        "writer" => writer_command(&args[2..]),
        "calc" => calc_command(&args[2..]),
        "impress" => impress_command(&args[2..]),
        "draw" => draw_command(&args[2..]),
        "math" => math_command(&args[2..]),
        "base" => base_command(&args[2..]),
        "package" => package_command(&args[2..]),
        "uno" => uno_command(&args[2..]),
        "office-demo" => office_demo_command(&args[2..]),
        "desktop-demo" => desktop_demo_command(&args[2..]),
        "help" | "--help" | "-h" => {
            print_usage();
            Ok(())
        }
        other => Err(lo_core::LoError::InvalidInput(format!(
            "unknown command: {other}"
        ))),
    }
}

/// Generate one document of each kind into `outdir` in every supported format.
/// Mirrors the parity attempt's `lo_cli office-demo` command and is the same
/// thing the integration tests will exercise against real LibreOffice.
fn office_demo_command(args: &[String]) -> Result<()> {
    let outdir = required_positional(args, 0, "output directory")?;
    fs::create_dir_all(outdir)?;

    // Writer
    let writer_doc = lo_writer::from_markdown(
        "Office Demo",
        "# Office Demo\n\nThis is a *demo* document with **bold**, `code`, and a [link](https://example.com).\n\n- one\n- two\n- three\n\n| col a | col b |\n| --- | --- |\n| 1 | 2 |\n",
    );
    for fmt in ["txt", "html", "svg", "pdf", "odt", "docx"] {
        let bytes = lo_writer::save_as(&writer_doc, fmt)?;
        fs::write(format!("{outdir}/writer.{fmt}"), bytes)?;
    }

    // Calc
    let workbook = lo_calc::workbook_from_csv_opts(
        "Office Demo",
        "Sales",
        "region,units,price\nNorth,12,9.5\nSouth,8,11.0\nEast,15,7.25",
        true,
    )?;
    for fmt in ["csv", "html", "svg", "pdf", "ods", "xlsx"] {
        let bytes = lo_calc::save_as(&workbook, fmt)?;
        fs::write(format!("{outdir}/calc.{fmt}"), bytes)?;
    }

    // Impress
    let presentation = demo_presentation("Office Demo");
    for fmt in ["html", "svg", "pdf", "odp", "pptx"] {
        let bytes = lo_impress::save_as(&presentation, fmt)?;
        fs::write(format!("{outdir}/impress.{fmt}"), bytes)?;
    }

    // Draw
    let drawing = demo_drawing("Office Demo");
    for fmt in ["svg", "pdf", "odg"] {
        let bytes = lo_draw::save_as(&drawing, fmt)?;
        fs::write(format!("{outdir}/draw.{fmt}"), bytes)?;
    }

    // Math
    let formula = lo_math::from_latex("Quadratic", r"\frac{-b \pm \sqrt{b^2 - 4ac}}{2a}")?;
    for fmt in ["mathml", "svg", "pdf"] {
        let bytes = lo_math::save_as(&formula, fmt)?;
        fs::write(format!("{outdir}/math.{fmt}"), bytes)?;
    }
    // ODF formula goes through lo_odf since that's the canonical writer.
    lo_odf::save_formula_document(format!("{outdir}/math.odf"), &formula)?;

    // Base
    let database = database_from_csv(
        "Office Demo",
        "people",
        "name,age,active\nAlice,30,true\nBob,22,false\nCara,41,true",
    )?;
    for fmt in ["html", "svg", "pdf", "odb"] {
        let bytes = lo_base::save_as(&database, fmt)?;
        fs::write(format!("{outdir}/base.{fmt}"), bytes)?;
    }

    println!("wrote office demo files to {outdir}");
    Ok(())
}

/// Drive the lo_app desktop surface end-to-end: open one document of every
/// kind from a template, exercise a representative command on each, save +
/// shell-render each window, run a recorded macro on Writer, autosave
/// everything to recovery snapshots, persist preferences and the start
/// center, and finally write a workspace report. Mirrors the parity
/// attempt's `lo_cli desktop-demo` end-to-end exercise.
fn desktop_demo_command(args: &[String]) -> Result<()> {
    let profile_dir = required_positional(args, 0, "profile directory")?;
    fs::create_dir_all(profile_dir)?;
    let app = DesktopApp::new(profile_dir)?;

    // ---- Writer ----
    let writer = app.open_template("writer:report")?;
    let mut text_args = BTreeMap::new();
    text_args.insert(
        "text".to_string(),
        UnoValue::String("This desktop shell runs fully in Rust.".to_string()),
    );
    app.execute_window_command(writer, ".uno:InsertText", &text_args)?;
    app.execute_window_command(writer, ".uno:SelectAll", &BTreeMap::new())?;
    app.execute_window_command(writer, ".uno:Bold", &BTreeMap::new())?;
    app.save_window(writer)?;
    app.save_window_shell(writer, format!("{profile_dir}/shells/writer.html"))?;

    // ---- Calc ----
    let calc = app.open_template("calc:budget")?;
    let mut cell_label_args = BTreeMap::new();
    cell_label_args.insert("row".to_string(), UnoValue::Int(5));
    cell_label_args.insert("col".to_string(), UnoValue::Int(0));
    cell_label_args.insert(
        "value".to_string(),
        UnoValue::String("Forecast".to_string()),
    );
    app.execute_window_command(calc, ".uno:SetCell", &cell_label_args)?;
    let mut cell_formula_args = BTreeMap::new();
    cell_formula_args.insert("row".to_string(), UnoValue::Int(5));
    cell_formula_args.insert("col".to_string(), UnoValue::Int(1));
    cell_formula_args.insert("value".to_string(), UnoValue::String("=B2-B3".to_string()));
    app.execute_window_command(calc, ".uno:SetCell", &cell_formula_args)?;
    app.save_window(calc)?;
    app.save_window_shell(calc, format!("{profile_dir}/shells/calc.html"))?;

    // ---- Impress ----
    let impress = app.open_template("impress:blank")?;
    let mut slide_args = BTreeMap::new();
    slide_args.insert(
        "title".to_string(),
        UnoValue::String("Pure Rust Desktop".to_string()),
    );
    app.execute_window_command(impress, ".uno:InsertSlide", &slide_args)?;
    let mut bullets_args = BTreeMap::new();
    bullets_args.insert("slide".to_string(), UnoValue::Int(1));
    bullets_args.insert(
        "items".to_string(),
        UnoValue::String("Writer|Calc|Impress|Draw|Math|Base".to_string()),
    );
    app.execute_window_command(impress, ".uno:InsertBullets", &bullets_args)?;
    app.save_window(impress)?;
    app.save_window_shell(impress, format!("{profile_dir}/shells/impress.html"))?;

    // ---- Draw ----
    let draw = app.open_template("draw:blank")?;
    let mut shape_args = BTreeMap::new();
    shape_args.insert("kind".to_string(), UnoValue::String("rect".to_string()));
    shape_args.insert("text".to_string(), UnoValue::String("Diagram".to_string()));
    app.execute_window_command(draw, ".uno:InsertShape", &shape_args)?;
    app.save_window(draw)?;
    app.save_window_shell(draw, format!("{profile_dir}/shells/draw.html"))?;

    // ---- Math ----
    let math = app.open_template("math:blank")?;
    let mut formula_args = BTreeMap::new();
    formula_args.insert(
        "formula".to_string(),
        UnoValue::String(r"\frac{x^2}{y}".to_string()),
    );
    app.execute_window_command(math, ".uno:SetFormula", &formula_args)?;
    app.export_window(math, format!("{profile_dir}/exports/formula.mathml"))?;
    app.save_window_shell(math, format!("{profile_dir}/shells/math.html"))?;

    // ---- Base ----
    let base = app.open_template("base:blank")?;
    let mut create_table_args = BTreeMap::new();
    create_table_args.insert(
        "name".to_string(),
        UnoValue::String("customers".to_string()),
    );
    create_table_args.insert(
        "columns".to_string(),
        UnoValue::String("id,name,tier".to_string()),
    );
    app.execute_window_command(base, ".uno:CreateTable", &create_table_args)?;
    let mut insert_row_args = BTreeMap::new();
    insert_row_args.insert(
        "table".to_string(),
        UnoValue::String("customers".to_string()),
    );
    insert_row_args.insert(
        "row".to_string(),
        UnoValue::String("1,Ada,Gold".to_string()),
    );
    app.execute_window_command(base, ".uno:InsertRow", &insert_row_args)?;
    app.save_window(base)?;
    app.save_window_shell(base, format!("{profile_dir}/shells/base.html"))?;

    // ---- Macro recording ----
    app.start_macro_recording("writer-uppercase");
    app.execute_window_command(writer, ".uno:Uppercase", &BTreeMap::new())?;
    app.stop_macro_recording();
    app.play_macro("writer-uppercase", writer)?;

    // ---- Persist desktop state ----
    app.autosave_all()?;
    app.save_preferences()?;
    app.save_start_center(format!("{profile_dir}/start-center.html"))?;
    fs::write(
        format!("{profile_dir}/workspace.txt"),
        app.workspace_report(),
    )?;

    println!("wrote desktop demo profile to {profile_dir}");
    Ok(())
}

fn writer_command(args: &[String]) -> Result<()> {
    match args.first().map(String::as_str) {
        Some("new") => {
            let output = required_positional(args, 1, "output path")?;
            let title = flag_value(args, "--title").unwrap_or_else(|| "Document".to_string());
            let text = flag_value(args, "--text")
                .unwrap_or_else(|| "Hello from libreoffice-rs".to_string());
            let mut doc = TextDocument::new(title);
            doc.push_paragraph(text);
            save_odt(output, &doc)
        }
        Some("markdown-to-odt") => {
            let input = required_positional(args, 1, "input markdown path")?;
            let output = required_positional(args, 2, "output odt path")?;
            let title = flag_value(args, "--title")
                .unwrap_or_else(|| file_stem_or_default(input, "Document"));
            let markdown = fs::read_to_string(input)?;
            let doc = from_markdown(title, &markdown);
            save_odt(output, &doc)
        }
        Some("convert") => {
            // Generic markdown → any-supported-format converter that uses
            // lo_writer::save_as. Format is inferred from the output
            // extension or can be overridden with --format.
            let input = required_positional(args, 1, "input markdown path")?;
            let output = required_positional(args, 2, "output path")?;
            let title = flag_value(args, "--title")
                .unwrap_or_else(|| file_stem_or_default(input, "Document"));
            let format = flag_value(args, "--format")
                .unwrap_or_else(|| extension_of(output).unwrap_or_else(|| "odt".to_string()));
            let markdown = fs::read_to_string(input)?;
            let doc = from_markdown(title, &markdown);
            let bytes = lo_writer::save_as(&doc, &format)?;
            fs::write(output, bytes)?;
            Ok(())
        }
        _ => Err(lo_core::LoError::InvalidInput(
            "writer commands: new, markdown-to-odt, convert".to_string(),
        )),
    }
}

fn calc_command(args: &[String]) -> Result<()> {
    match args.first().map(String::as_str) {
        Some("csv-to-ods") => {
            let input = required_positional(args, 1, "input csv path")?;
            let output = required_positional(args, 2, "output ods path")?;
            let sheet_name = flag_value(args, "--sheet").unwrap_or_else(|| "Sheet1".to_string());
            let title = flag_value(args, "--title")
                .unwrap_or_else(|| file_stem_or_default(output, "Workbook"));
            let has_header = has_flag(args, "--has-header");
            let csv = fs::read_to_string(input)?;
            let workbook = workbook_from_csv_opts(title, &sheet_name, &csv, has_header)?;
            save_ods(output, &workbook)
        }
        Some("convert") => {
            let input = required_positional(args, 1, "input csv path")?;
            let output = required_positional(args, 2, "output path")?;
            let sheet_name = flag_value(args, "--sheet").unwrap_or_else(|| "Sheet1".to_string());
            let title = flag_value(args, "--title")
                .unwrap_or_else(|| file_stem_or_default(output, "Workbook"));
            let has_header = has_flag(args, "--has-header");
            let format = flag_value(args, "--format")
                .unwrap_or_else(|| extension_of(output).unwrap_or_else(|| "ods".to_string()));
            let csv = fs::read_to_string(input)?;
            let workbook = workbook_from_csv_opts(title, &sheet_name, &csv, has_header)?;
            let bytes = lo_calc::save_as(&workbook, &format)?;
            fs::write(output, bytes)?;
            Ok(())
        }
        Some("eval") => {
            let formula = required_positional(args, 1, "formula")?;
            let has_header = has_flag(args, "--has-header");
            let sheet = if let Some(csv_path) = flag_value(args, "--csv") {
                let csv = fs::read_to_string(csv_path)?;
                let workbook = workbook_from_csv_opts("Eval", "Sheet1", &csv, has_header)?;
                workbook.sheets[0].clone()
            } else {
                let mut sheet = Sheet::new("Sheet1");
                sheet.set(lo_core::CellAddr::new(0, 0), CellValue::Number(1.0));
                sheet.set(lo_core::CellAddr::new(1, 0), CellValue::Number(2.0));
                sheet.set(lo_core::CellAddr::new(2, 0), CellValue::Number(3.0));
                sheet
            };
            let value = evaluate_formula(formula, &sheet)?;
            println!("{value:?}");
            Ok(())
        }
        _ => Err(lo_core::LoError::InvalidInput(
            "calc commands: csv-to-ods, convert, eval".to_string(),
        )),
    }
}

fn impress_command(args: &[String]) -> Result<()> {
    match args.first().map(String::as_str) {
        Some("demo") => {
            let output = required_positional(args, 1, "output odp path")?;
            let title = flag_value(args, "--title").unwrap_or_else(|| "Demo Deck".to_string());
            let presentation = demo_presentation(&title);
            lo_impress::save_odp(output, &presentation)
        }
        _ => Err(lo_core::LoError::InvalidInput(
            "impress commands: demo".to_string(),
        )),
    }
}

fn draw_command(args: &[String]) -> Result<()> {
    match args.first().map(String::as_str) {
        Some("demo") => {
            let output = required_positional(args, 1, "output odg path")?;
            let title = flag_value(args, "--title").unwrap_or_else(|| "Diagram".to_string());
            let drawing = demo_drawing(&title);
            lo_draw::save_odg(output, &drawing)
        }
        _ => Err(lo_core::LoError::InvalidInput(
            "draw commands: demo".to_string(),
        )),
    }
}

fn math_command(args: &[String]) -> Result<()> {
    match args.first().map(String::as_str) {
        Some("latex-to-mathml") => {
            let input = required_positional(args, 1, "input formula path")?;
            let latex = fs::read_to_string(input)?;
            let doc = from_latex("Formula", &latex)?;
            println!("{}", to_mathml_string(&doc.root));
            Ok(())
        }
        Some("latex-to-odf") => {
            let input = required_positional(args, 1, "input formula path")?;
            let output = required_positional(args, 2, "output odf path")?;
            let title = flag_value(args, "--title").unwrap_or_else(|| "Formula".to_string());
            let latex = fs::read_to_string(input)?;
            let doc = from_latex(title, &latex)?;
            lo_odf::save_formula_document(output, &doc)
        }
        _ => Err(lo_core::LoError::InvalidInput(
            "math commands: latex-to-mathml, latex-to-odf".to_string(),
        )),
    }
}

fn base_command(args: &[String]) -> Result<()> {
    match args.first().map(String::as_str) {
        Some("csv-to-odb") => {
            let input = required_positional(args, 1, "input csv path")?;
            let table = required_positional(args, 2, "table name")?;
            let output = required_positional(args, 3, "output odb path")?;
            let title = flag_value(args, "--title")
                .unwrap_or_else(|| file_stem_or_default(output, "Database"));
            let csv = fs::read_to_string(input)?;
            let database = database_from_csv(title, table, &csv)?;
            lo_base::save_odb(output, &database)
        }
        Some("query") => {
            let input = required_positional(args, 1, "input csv path")?;
            let table = required_positional(args, 2, "table name")?;
            let sql = args.iter().skip(3).cloned().collect::<Vec<_>>().join(" ");
            if sql.trim().is_empty() {
                return Err(lo_core::LoError::InvalidInput(
                    "missing SQL query".to_string(),
                ));
            }
            let csv = fs::read_to_string(input)?;
            let database = database_from_csv("Query", table, &csv)?;
            let result = execute_select(&database, &sql)?;
            println!("{}", result.columns.join("\t"));
            for row in result.rows {
                let cells = row
                    .into_iter()
                    .map(|value| match value {
                        lo_core::DbValue::Null => String::new(),
                        lo_core::DbValue::Integer(v) => v.to_string(),
                        lo_core::DbValue::Float(v) => v.to_string(),
                        lo_core::DbValue::Text(v) => v,
                        lo_core::DbValue::Bool(v) => v.to_string(),
                    })
                    .collect::<Vec<_>>();
                println!("{}", cells.join("\t"));
            }
            Ok(())
        }
        _ => Err(lo_core::LoError::InvalidInput(
            "base commands: csv-to-odb, query".to_string(),
        )),
    }
}

fn package_command(args: &[String]) -> Result<()> {
    match args.first().map(String::as_str) {
        Some("inspect") => {
            let path = required_positional(args, 1, "package path")?;
            for entry in list_entries(path)? {
                println!("{}\t{} bytes", entry.name, entry.uncompressed_size);
            }
            Ok(())
        }
        _ => Err(lo_core::LoError::InvalidInput(
            "package commands: inspect".to_string(),
        )),
    }
}

fn uno_command(args: &[String]) -> Result<()> {
    // ComponentContext::default() pre-registers Echo, TextTransformations
    // and Info as singletons, so the CLI gets them for free.
    let registry = ServiceRegistry::default();
    match args.first().map(String::as_str) {
        Some("list") => {
            for service in registry.list_services() {
                println!("{service}");
            }
            Ok(())
        }
        Some("demo") => {
            let response = registry.invoke(
                "com.libreoffice_rs.Echo",
                "echo",
                &[UnoValue::string("hello from libreoffice-rs")],
            )?;
            println!("{response:?}");
            Ok(())
        }
        _ => Err(lo_core::LoError::InvalidInput(
            "uno commands: list, demo".to_string(),
        )),
    }
}

fn required_positional<'a>(args: &'a [String], index: usize, name: &str) -> Result<&'a str> {
    args.get(index)
        .map(String::as_str)
        .ok_or_else(|| lo_core::LoError::InvalidInput(format!("missing {name}")))
}

fn flag_value(args: &[String], flag: &str) -> Option<String> {
    args.iter()
        .position(|arg| arg == flag)
        .and_then(|index| args.get(index + 1).cloned())
}

fn has_flag(args: &[String], flag: &str) -> bool {
    args.iter().any(|arg| arg == flag)
}

fn file_stem_or_default(path: &str, fallback: &str) -> String {
    Path::new(path)
        .file_stem()
        .and_then(|value| value.to_str())
        .map(|value| value.to_string())
        .unwrap_or_else(|| fallback.to_string())
}

fn extension_of(path: &str) -> Option<String> {
    Path::new(path)
        .extension()
        .and_then(|value| value.to_str())
        .map(|value| value.to_ascii_lowercase())
}

fn print_usage() {
    println!(
        "Usage:\n  libreoffice-rs writer new <out.odt> [--title TITLE] [--text TEXT]\n  libreoffice-rs writer markdown-to-odt <in.md> <out.odt> [--title TITLE]\n  libreoffice-rs writer convert <in.md> <out.{{txt|html|svg|pdf|odt|docx}}> [--title TITLE] [--format FMT]\n  libreoffice-rs calc csv-to-ods <in.csv> <out.ods> [--sheet NAME] [--title TITLE] [--has-header]\n  libreoffice-rs calc convert <in.csv> <out.{{csv|html|svg|pdf|ods|xlsx}}> [--sheet NAME] [--title TITLE] [--has-header] [--format FMT]\n  libreoffice-rs calc eval <formula> [--csv input.csv] [--has-header]\n  libreoffice-rs impress demo <out.odp> [--title TITLE]\n  libreoffice-rs draw demo <out.odg> [--title TITLE]\n  libreoffice-rs math latex-to-mathml <formula.txt>\n  libreoffice-rs math latex-to-odf <formula.txt> <out.odf> [--title TITLE]\n  libreoffice-rs base csv-to-odb <in.csv> <table> <out.odb> [--title TITLE]\n  libreoffice-rs base query <in.csv> <table> <SQL...>\n  libreoffice-rs office-demo <outdir>\n  libreoffice-rs desktop-demo <profile_dir>\n  libreoffice-rs package inspect <file>\n  libreoffice-rs uno list\n  libreoffice-rs uno demo"
    );
}
