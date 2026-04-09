//! `libreoffice-pure` CLI: thin wrapper over the high-level helpers in
//! the `libreoffice_pure` library target.

use std::env;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::ExitCode;

use libreoffice_pure::{
    accept_all_tracked_changes_docx_bytes, convert_bytes, convert_path_bytes, doc_to_docx_bytes,
    docx_to_jpeg_pages, docx_to_md_bytes, docx_to_pdf_bytes, docx_to_png_pages,
    pdf_to_html_bytes, pdf_to_md_bytes, pdf_to_txt_bytes, pptx_to_jpeg_pages,
    pptx_to_md_bytes, pptx_to_pdf_bytes, pptx_to_png_pages, xlsx_recalc_bytes,
    xlsx_recalc_check_json, xlsx_to_md_bytes,
};

fn main() -> ExitCode {
    let args: Vec<String> = env::args().collect();
    match run(&args) {
        Ok(()) => ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("error: {error}");
            ExitCode::FAILURE
        }
    }
}

fn run(args: &[String]) -> Result<(), Box<dyn std::error::Error>> {
    let normalized = normalize_global_flags(args);
    if normalized.len() < 2 {
        print_usage();
        return Err("missing command".into());
    }
    match normalized[1].as_str() {
        "--convert-to" => soffice_convert(&normalized[1..]),
        "convert" => convert_command(&normalized[2..]),
        "docx-to-pdf" => convert_legacy(&normalized[1..], docx_to_pdf_bytes),
        "doc-to-docx" => convert_legacy(&normalized[1..], doc_to_docx_bytes),
        "pptx-to-pdf" => convert_legacy(&normalized[1..], pptx_to_pdf_bytes),
        "xlsx-recalc" => convert_legacy(&normalized[1..], xlsx_recalc_bytes),
        "xlsx-recalc-check" => recalc_check_command(&normalized[2..]),
        "accept-changes" | "accept-tracked-changes" => {
            convert_legacy(&normalized[1..], accept_all_tracked_changes_docx_bytes)
        }
        "docx-to-md" => convert_legacy(&normalized[1..], docx_to_md_bytes),
        "pdf-to-txt" => convert_legacy(&normalized[1..], pdf_to_txt_bytes),
        "pdf-to-md" => convert_legacy(&normalized[1..], pdf_to_md_bytes),
        "pdf-to-html" => convert_legacy(&normalized[1..], pdf_to_html_bytes),
        "pptx-to-md" => convert_legacy(&normalized[1..], pptx_to_md_bytes),
        "xlsx-to-md" => convert_legacy(&normalized[1..], xlsx_to_md_bytes),
        "docx-to-pngs" => raster_command(&normalized[2..], docx_to_png_pages, OutputFormat::Png),
        "docx-to-jpegs" => raster_jpeg_command(&normalized[2..], docx_to_jpeg_pages),
        "pptx-to-pngs" => raster_command(&normalized[2..], pptx_to_png_pages, OutputFormat::Png),
        "pptx-to-jpegs" => raster_jpeg_command(&normalized[2..], pptx_to_jpeg_pages),
        "help" | "--help" | "-h" => {
            print_usage();
            Ok(())
        }
        other => {
            print_usage();
            Err(format!("unknown command: {other}").into())
        }
    }
}

fn normalize_global_flags(args: &[String]) -> Vec<String> {
    let mut out = Vec::with_capacity(args.len());
    for (index, arg) in args.iter().enumerate() {
        if index != 0 && arg == "--headless" {
            continue;
        }
        out.push(arg.clone());
    }
    out
}

fn convert_legacy<F>(args: &[String], handler: F) -> Result<(), Box<dyn std::error::Error>>
where
    F: Fn(&[u8]) -> lo_core::Result<Vec<u8>>,
{
    if args.len() < 3 {
        return Err("expected: <input> <output>".into());
    }
    let input = &args[1];
    let output = &args[2];
    let bytes = fs::read(input)?;
    let out = handler(&bytes)?;
    fs::write(output, out)?;
    Ok(())
}

fn recalc_check_command(args: &[String]) -> Result<(), Box<dyn std::error::Error>> {
    match args {
        [input] => {
            let bytes = fs::read(input)?;
            println!("{}", xlsx_recalc_check_json(&bytes)?);
            Ok(())
        }
        [input, output] => {
            let bytes = fs::read(input)?;
            fs::write(output, xlsx_recalc_check_json(&bytes)?)?;
            Ok(())
        }
        _ => Err("expected: xlsx-recalc-check <input.xlsx> [output.json]".into()),
    }
}

#[derive(Clone, Copy)]
enum OutputFormat {
    Png,
    Jpeg,
}

fn raster_command<F>(
    args: &[String],
    handler: F,
    format: OutputFormat,
) -> Result<(), Box<dyn std::error::Error>>
where
    F: Fn(&[u8], u32) -> lo_core::Result<Vec<Vec<u8>>>,
{
    let mut dpi = 150u32;
    let mut positionals = Vec::new();
    let mut index = 0usize;
    while index < args.len() {
        match args[index].as_str() {
            "--dpi" => {
                index += 1;
                dpi = args.get(index).ok_or("missing value after --dpi")?.parse()?;
            }
            other => positionals.push(other.to_string()),
        }
        index += 1;
    }
    if positionals.len() != 2 {
        return Err("expected: <input> <outdir> [--dpi 150]".into());
    }
    let bytes = fs::read(&positionals[0])?;
    let pages = handler(&bytes, dpi)?;
    write_pages(&positionals[0], &positionals[1], pages, format)
}

fn raster_jpeg_command<F>(args: &[String], handler: F) -> Result<(), Box<dyn std::error::Error>>
where
    F: Fn(&[u8], u32, u8) -> lo_core::Result<Vec<Vec<u8>>>,
{
    let mut dpi = 150u32;
    let mut quality = 85u8;
    let mut positionals = Vec::new();
    let mut index = 0usize;
    while index < args.len() {
        match args[index].as_str() {
            "--dpi" => {
                index += 1;
                dpi = args.get(index).ok_or("missing value after --dpi")?.parse()?;
            }
            "--quality" => {
                index += 1;
                quality = args.get(index).ok_or("missing value after --quality")?.parse()?;
            }
            other => positionals.push(other.to_string()),
        }
        index += 1;
    }
    if positionals.len() != 2 {
        return Err("expected: <input> <outdir> [--dpi 150] [--quality 85]".into());
    }
    let bytes = fs::read(&positionals[0])?;
    let pages = handler(&bytes, dpi, quality)?;
    write_pages(&positionals[0], &positionals[1], pages, OutputFormat::Jpeg)
}

fn write_pages(
    input_path: &str,
    outdir: &str,
    pages: Vec<Vec<u8>>,
    format: OutputFormat,
) -> Result<(), Box<dyn std::error::Error>> {
    fs::create_dir_all(outdir)?;
    let stem = Path::new(input_path)
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("page");
    let ext = match format {
        OutputFormat::Png => "png",
        OutputFormat::Jpeg => "jpg",
    };
    for (index, bytes) in pages.into_iter().enumerate() {
        let path = Path::new(outdir).join(format!("{}_{}.{}", stem, index + 1, ext));
        fs::write(path, bytes)?;
    }
    Ok(())
}

fn convert_command(args: &[String]) -> Result<(), Box<dyn std::error::Error>> {
    let mut from: Option<String> = None;
    let mut to: Option<String> = None;
    let mut outdir: Option<PathBuf> = None;
    let mut positionals: Vec<String> = Vec::new();
    let mut index = 0usize;
    while index < args.len() {
        match args[index].as_str() {
            "--from" => {
                index += 1;
                from = Some(args.get(index).ok_or("missing value after --from")?.clone());
            }
            "--to" => {
                index += 1;
                to = Some(args.get(index).ok_or("missing value after --to")?.clone());
            }
            "--outdir" => {
                index += 1;
                outdir = Some(PathBuf::from(args.get(index).ok_or("missing value after --outdir")?));
            }
            "-h" | "--help" => {
                print_usage();
                return Ok(());
            }
            other => positionals.push(other.to_string()),
        }
        index += 1;
    }
    let to = to.ok_or("convert requires --to <format>")?;
    if positionals.is_empty() {
        return Err("convert requires at least one input path".into());
    }
    if outdir.is_some() && positionals.len() == 2 {
        return Err("use either --outdir with input paths or an explicit output path, not both".into());
    }
    if let Some(dir) = outdir {
        fs::create_dir_all(&dir)?;
        for input in &positionals {
            let output = derived_output_path(input, &to, Some(&dir))?;
            convert_one(input, &output, from.as_deref(), &to)?;
        }
        return Ok(());
    }
    match positionals.as_slice() {
        [input] => {
            let output = derived_output_path(input, &to, None)?;
            convert_one(input, &output, from.as_deref(), &to)
        }
        [input, output] => convert_one(input, output, from.as_deref(), &to),
        _ => Err("multiple inputs require --outdir <dir> or soffice-style --convert-to".into()),
    }
}

fn soffice_convert(args: &[String]) -> Result<(), Box<dyn std::error::Error>> {
    if args.len() < 3 {
        return Err("expected: --convert-to <format> <input...> [--outdir dir]".into());
    }
    let to = args[1].clone();
    let mut outdir: Option<PathBuf> = None;
    let mut inputs: Vec<String> = Vec::new();
    let mut index = 2usize;
    while index < args.len() {
        match args[index].as_str() {
            "--outdir" => {
                index += 1;
                outdir = Some(PathBuf::from(args.get(index).ok_or("missing value after --outdir")?));
            }
            other => inputs.push(other.to_string()),
        }
        index += 1;
    }
    if inputs.is_empty() {
        return Err("--convert-to requires at least one input file".into());
    }
    if let Some(dir) = &outdir {
        fs::create_dir_all(dir)?;
    }
    for input in &inputs {
        let output = derived_output_path(input, &to, outdir.as_deref())?;
        convert_one(input, &output, None, &to)?;
    }
    Ok(())
}

fn convert_one(
    input: &str,
    output: impl AsRef<Path>,
    from_hint: Option<&str>,
    to_hint: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = fs::read(input)?;
    let out = match from_hint {
        Some(from) => convert_bytes(&bytes, from, to_hint)?,
        None => convert_path_bytes(input, &bytes, to_hint)?,
    };
    fs::write(output, out)?;
    Ok(())
}

fn derived_output_path(
    input: &str,
    to_hint: &str,
    outdir: Option<&Path>,
) -> Result<PathBuf, Box<dyn std::error::Error>> {
    let input_path = Path::new(input);
    let stem = input_path
        .file_stem()
        .and_then(|value| value.to_str())
        .ok_or("input path has no valid file stem")?;
    let target_ext = canonical_target_extension(to_hint);
    let parent = match outdir {
        Some(dir) => dir.to_path_buf(),
        None => input_path
            .parent()
            .map(Path::to_path_buf)
            .unwrap_or_else(|| PathBuf::from(".")),
    };
    Ok(parent.join(format!("{stem}.{target_ext}")))
}

fn canonical_target_extension(to_hint: &str) -> String {
    let head = to_hint
        .trim()
        .trim_start_matches('.')
        .split(':')
        .next()
        .unwrap_or(to_hint)
        .to_ascii_lowercase();
    match head.as_str() {
        "text" => "txt".to_string(),
        "htm" => "html".to_string(),
        "mml" => "mathml".to_string(),
        "markdown" => "md".to_string(),
        other => other.to_string(),
    }
}

fn print_usage() {
    eprintln!(
        "usage:
  libreoffice-pure docx-to-pdf <input.docx> <output.pdf>
  libreoffice-pure doc-to-docx <input.doc> <output.docx>
  libreoffice-pure pptx-to-pdf <input.pptx> <output.pdf>
  libreoffice-pure xlsx-recalc <input.xlsx> <output.xlsx>
  libreoffice-pure xlsx-recalc-check <input.xlsx> [output.json]
  libreoffice-pure accept-changes <input.docx> <output.docx>
  libreoffice-pure docx-to-md <input.docx> <output.md>
  libreoffice-pure pdf-to-txt <input.pdf> <output.txt>
  libreoffice-pure pdf-to-md <input.pdf> <output.md>
  libreoffice-pure pdf-to-html <input.pdf> <output.html>
  libreoffice-pure pptx-to-md <input.pptx> <output.md>
  libreoffice-pure xlsx-to-md <input.xlsx> <output.md>
  libreoffice-pure docx-to-pngs <input.docx> <outdir> [--dpi 150]
  libreoffice-pure docx-to-jpegs <input.docx> <outdir> [--dpi 150] [--quality 85]
  libreoffice-pure pptx-to-pngs <input.pptx> <outdir> [--dpi 150]
  libreoffice-pure pptx-to-jpegs <input.pptx> <outdir> [--dpi 150] [--quality 85]
  libreoffice-pure convert --to <format> [--from <format>] <input> [output]
  libreoffice-pure convert --to <format> --outdir <dir> <input>...
  libreoffice-pure --headless --convert-to <format> <input>... [--outdir <dir>]

examples:
  libreoffice-pure convert --to pdf report.docx
  libreoffice-pure convert --to md deck.pptx
  libreoffice-pure convert --to txt paper.pdf
  libreoffice-pure docx-to-pngs report.docx pages --dpi 144
  libreoffice-pure pptx-to-jpegs slides.pptx slides_jpg --dpi 150 --quality 88
  libreoffice-pure --headless --convert-to pdf slide.pptx --outdir out
  libreoffice-pure --convert-to pdf:writer_pdf_Export notes.odt
  libreoffice-pure xlsx-recalc-check model.xlsx"
    );
}
