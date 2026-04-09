# lo_cli

Command-line driver for the [libreoffice-rs](https://github.com/clark-labs-inc/libreoffice-rs)
workspace. Wraps the high-level `save_as` and `load_bytes` helpers in
each crate so you can convert documents from the shell.

## Subcommands

```text
lo_cli writer-md input.md output.{odt|docx|html|pdf|svg|txt}
lo_cli writer-txt input.txt output.{odt|docx|html|pdf|svg|txt}
lo_cli writer-import input.{docx|odt|html|txt} output.{...}
lo_cli calc-csv input.csv output.{ods|xlsx|html|pdf|svg|csv}
lo_cli calc-import input.{xlsx|ods|csv} output.{...}
lo_cli impress-demo output.{odp|pptx|html|pdf|svg}
lo_cli impress-import input.{pptx|odp|txt} output.{...}
lo_cli draw-demo output.{odg|svg|pdf}
lo_cli draw-import input.{odg|svg} output.{...}
lo_cli math '\frac{x^2}{y}' output.{mathml|svg|pdf}
lo_cli math-import input.{mathml|mml|odf|txt} output.{...}
lo_cli base-csv input.csv "SELECT * FROM data" output.csv
lo_cli base-import input.{odb|csv} "SELECT * FROM data" output.csv
lo_cli office-demo out_dir
```

## Install

```sh
cargo install lo_cli
```

License: MIT
