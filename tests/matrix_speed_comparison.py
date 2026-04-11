#!/usr/bin/env python3
"""Full N×M matrix speed comparison: libreoffice-rs vs real LibreOffice.

For each (family, from, to) format pair this script runs both
libreoffice-rs (`libreoffice-pure convert`) and real LibreOffice
(`soffice --headless --convert-to`) on the same input file, records
wall-clock time, and writes a TSV + Markdown summary.

Outputs:
  benchmark_evidence/matrix_speed_comparison.tsv
  benchmark_evidence/matrix_speed_comparison.md
  benchmark_evidence/18_speed_comparison/{rs,lo}/...
"""
import os
import statistics
import subprocess
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
CORPUS = REPO / "benchmark_evidence" / "00_corpus"
OUTDIR = REPO / "benchmark_evidence" / "18_speed_comparison"
TSV = REPO / "benchmark_evidence" / "matrix_speed_comparison.tsv"
SUMMARY_MD = REPO / "benchmark_evidence" / "matrix_speed_comparison.md"
BIN_PURE = REPO / "target" / "release" / "libreoffice-pure"

LO_SKIP_FROM = {"md", "latex", "mathml", "odb", "odf"}
LO_SKIP_TO = {"md", "mathml", "odf", "odb"}

PAIRS = [
    ("writer", "txt",  "synthetic.txt",  ["txt", "md", "html", "svg", "pdf", "odt", "docx"]),
    ("writer", "md",   "synthetic.md",   ["txt", "md", "html", "svg", "pdf", "odt", "docx"]),
    ("writer", "html", "synthetic.html", ["txt", "md", "html", "svg", "pdf", "odt", "docx"]),
    ("writer", "odt",  "synthetic.odt",  ["txt", "md", "html", "svg", "pdf", "odt", "docx"]),
    ("writer", "pdf",  "synthetic.pdf",  ["txt", "md", "html", "svg", "pdf", "odt", "docx"]),
    ("writer", "docx", "synthetic.docx", ["txt", "md", "html", "svg", "pdf", "odt", "docx"]),
    ("writer", "docx", "fixture-calibre-demo.docx", ["txt", "md", "html", "svg", "pdf", "odt", "docx"]),
    ("calc",   "csv",  "synthetic.csv",  ["csv", "md", "html", "svg", "pdf", "ods", "xlsx"]),
    ("calc",   "ods",  "synthetic.ods",  ["csv", "md", "html", "svg", "pdf", "ods", "xlsx"]),
    ("calc",   "xlsx", "synthetic.xlsx", ["csv", "md", "html", "svg", "pdf", "ods", "xlsx"]),
    ("calc",   "xlsx", "gov-census-state-pop.xlsx", ["csv", "md", "html", "svg", "pdf", "ods", "xlsx"]),
    ("impress","odp",  "synthetic.odp",  ["md", "html", "svg", "pdf", "odp", "pptx"]),
    ("impress","pptx", "fixture-python-pptx-datalabels.pptx", ["md", "html", "svg", "pdf", "odp", "pptx"]),
    ("draw",   "svg",  "synthetic.svg",  ["svg", "pdf", "odg"]),
    ("draw",   "odg",  "synthetic.odg",  ["svg", "pdf", "odg"]),
    ("math",   "latex",  "synthetic.latex",  ["mathml", "svg", "pdf", "odf"]),
    ("math",   "mathml", "synthetic.mathml", ["mathml", "svg", "pdf", "odf"]),
    ("math",   "odf",    "synthetic.odf",    ["mathml", "svg", "pdf", "odf"]),
    ("base",   "odb",  "synthetic.odb",  ["html", "svg", "pdf", "odb"]),
]


def now():
    return time.monotonic()


def ensure_binary():
    if BIN_PURE.exists():
        return
    print("Building libreoffice-pure release binary...", flush=True)
    subprocess.run(
        ["cargo", "build", "--release", "--quiet", "-p", "libreoffice-pure"],
        cwd=REPO,
        check=True,
    )


def run_timed(cmd, timeout=180):
    start = now()
    try:
        r = subprocess.run(
            cmd,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=timeout,
        )
        elapsed_ms = (now() - start) * 1000.0
        return r.returncode, elapsed_ms
    except subprocess.TimeoutExpired:
        return -1, float(timeout * 1000)


def measure_rs(from_fmt, to_fmt, input_path, out_root):
    out_dir = out_root / f"{input_path.stem}_{from_fmt}"
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"out.{to_fmt}"
    rc, ms = run_timed(
        [
            str(BIN_PURE),
            "convert",
            "--from",
            from_fmt,
            "--to",
            to_fmt,
            str(input_path),
            str(out),
        ]
    )
    ok = rc == 0 and out.exists() and out.stat().st_size > 0
    return ok, ms


def measure_lo(from_fmt, to_fmt, input_path, out_root):
    if from_fmt in LO_SKIP_FROM or to_fmt in LO_SKIP_TO:
        return None, None
    od = out_root / f"{input_path.stem}_{from_fmt}_to_{to_fmt}"
    od.mkdir(parents=True, exist_ok=True)
    filter_str = "txt:Text (encoded):UTF8" if to_fmt == "txt" else to_fmt
    rc, ms = run_timed(
        [
            "soffice",
            "--headless",
            "--convert-to",
            filter_str,
            "--outdir",
            str(od),
            str(input_path),
        ]
    )
    produced = [p for p in od.iterdir() if p.is_file()] if od.exists() else []
    ok = rc == 0 and len(produced) > 0
    return ok, ms


def soffice_version():
    try:
        r = subprocess.run(
            ["soffice", "--version"],
            capture_output=True,
            text=True,
            timeout=30,
        )
        return r.stdout.strip().splitlines()[0]
    except Exception:
        return "soffice unavailable"


def main():
    ensure_binary()
    OUTDIR.mkdir(parents=True, exist_ok=True)
    rs_root = OUTDIR / "rs"
    lo_root = OUTDIR / "lo"
    rs_root.mkdir(parents=True, exist_ok=True)
    lo_root.mkdir(parents=True, exist_ok=True)

    print("Warming up soffice (first cold start)...", flush=True)
    warmup = OUTDIR / "warmup"
    warmup.mkdir(exist_ok=True)
    subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(warmup),
            str(CORPUS / "synthetic.odt"),
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    rows = []
    with TSV.open("w") as f:
        f.write(
            "family\tfrom\tto\tinput\trs_ok\trs_ms\tlo_supported\tlo_ok\tlo_ms\tspeedup\n"
        )
        for family, frm, input_name, to_list in PAIRS:
            input_path = CORPUS / input_name
            if not input_path.exists():
                print(f"  SKIP missing input {input_name}", flush=True)
                continue
            for to in to_list:
                rs_ok, rs_ms = measure_rs(frm, to, input_path, rs_root)
                lo_ok_raw, lo_ms_raw = measure_lo(frm, to, input_path, lo_root)
                if lo_ok_raw is None:
                    lo_supported = "no"
                    lo_ok = "-"
                    lo_ms_val = None
                else:
                    lo_supported = "yes"
                    lo_ok = "OK" if lo_ok_raw else "ERR"
                    lo_ms_val = lo_ms_raw
                speedup = "-"
                if rs_ok and lo_ok == "OK" and rs_ms > 0 and lo_ms_val:
                    speedup = f"{lo_ms_val / rs_ms:.1f}x"
                rs_ms_str = f"{rs_ms:.0f}"
                lo_ms_str = f"{lo_ms_val:.0f}" if lo_ms_val is not None else "-"
                rs_ok_str = "OK" if rs_ok else "ERR"
                row = {
                    "family": family,
                    "from": frm,
                    "to": to,
                    "input": input_name,
                    "rs_ok": rs_ok_str,
                    "rs_ms": rs_ms,
                    "lo_supported": lo_supported,
                    "lo_ok": lo_ok,
                    "lo_ms": lo_ms_val,
                    "speedup": speedup,
                }
                rows.append(row)
                f.write(
                    f"{family}\t{frm}\t{to}\t{input_name}\t{rs_ok_str}\t{rs_ms_str}\t{lo_supported}\t{lo_ok}\t{lo_ms_str}\t{speedup}\n"
                )
                print(
                    f"  {family:<8} {frm:>5}→{to:<5} {input_name:<44} rs={rs_ms_str:>5}ms "
                    f"lo={lo_ms_str:>6}ms speedup={speedup:<6} [{rs_ok_str}/{lo_ok}]",
                    flush=True,
                )

    write_summary(rows)
    print(f"\nTSV:     {TSV}")
    print(f"Summary: {SUMMARY_MD}")


def write_summary(rows):
    lo_ver = soffice_version()
    total = len(rows)
    rs_ok_rows = [r for r in rows if r["rs_ok"] == "OK"]
    lo_ok_rows = [r for r in rows if r["lo_ok"] == "OK"]
    both_ok = [r for r in rows if r["rs_ok"] == "OK" and r["lo_ok"] == "OK"]
    unsupported_by_lo = [r for r in rows if r["lo_supported"] == "no"]

    def fam_stats(family):
        xs = [r for r in both_ok if r["family"] == family]
        if not xs:
            return None
        rs_mean = statistics.mean(r["rs_ms"] for r in xs)
        lo_mean = statistics.mean(r["lo_ms"] for r in xs)
        speedups = [r["lo_ms"] / r["rs_ms"] for r in xs if r["rs_ms"] > 0]
        return {
            "n": len(xs),
            "rs_mean": rs_mean,
            "lo_mean": lo_mean,
            "speedup_mean": statistics.mean(speedups) if speedups else 0.0,
            "speedup_min": min(speedups) if speedups else 0.0,
            "speedup_max": max(speedups) if speedups else 0.0,
        }

    families = ["writer", "calc", "impress", "draw", "math", "base"]
    fs = {f: fam_stats(f) for f in families}

    all_speedups = [
        r["lo_ms"] / r["rs_ms"] for r in both_ok if r["rs_ms"] > 0
    ]
    overall_rs_mean = statistics.mean(r["rs_ms"] for r in both_ok) if both_ok else 0
    overall_lo_mean = statistics.mean(r["lo_ms"] for r in both_ok) if both_ok else 0
    overall_speedup_mean = statistics.mean(all_speedups) if all_speedups else 0
    overall_speedup_min = min(all_speedups) if all_speedups else 0
    overall_speedup_max = max(all_speedups) if all_speedups else 0

    with SUMMARY_MD.open("w") as f:
        f.write("# Matrix Speed Comparison: libreoffice-rs vs real LibreOffice\n\n")
        f.write(f"- **Reference:** `{lo_ver}`\n")
        f.write(f"- **Total pairs measured:** {total}\n")
        f.write(f"- **libreoffice-rs OK:** {len(rs_ok_rows)}/{total}\n")
        f.write(
            f"- **LibreOffice OK (on pairs it supports):** "
            f"{len(lo_ok_rows)}/{total - len(unsupported_by_lo)}\n"
        )
        f.write(
            f"- **Pairs unsupported by LibreOffice CLI:** {len(unsupported_by_lo)} "
            f"(md/mathml/odf/odb output, or md/latex/mathml input)\n\n"
        )
        f.write(
            f"**Overall (both engines succeeded, n={len(both_ok)}):**  "
            f"libreoffice-rs **{overall_rs_mean:.0f}ms** vs "
            f"LibreOffice **{overall_lo_mean:.0f}ms**  "
            f"(mean speedup **{overall_speedup_mean:.1f}×**, "
            f"range {overall_speedup_min:.1f}×–{overall_speedup_max:.0f}×)\n\n"
        )
        f.write("## Per-family summary\n\n")
        f.write(
            "| Family | Pairs | libreoffice-rs mean | LibreOffice mean | Mean speedup | Speedup range |\n"
        )
        f.write("|---|---:|---:|---:|---:|---|\n")
        for fam in families:
            s = fs[fam]
            if not s:
                f.write(f"| {fam} | 0 | — | — | — | (no head-to-head) |\n")
                continue
            f.write(
                f"| {fam} | {s['n']} | {s['rs_mean']:.0f}ms | {s['lo_mean']:.0f}ms | "
                f"**{s['speedup_mean']:.1f}×** | {s['speedup_min']:.1f}×–{s['speedup_max']:.0f}× |\n"
            )
        f.write("\n## Full matrix\n\n")
        f.write(
            "| family | from → to | input | rs ms | LO ms | speedup | LO support |\n"
        )
        f.write("|---|---|---|---:|---:|---:|---|\n")
        for r in rows:
            rs_ms = f"{r['rs_ms']:.0f}"
            lo_ms = f"{r['lo_ms']:.0f}" if r["lo_ms"] is not None else "—"
            f.write(
                f"| {r['family']} | `{r['from']}`→`{r['to']}` | {r['input']} | "
                f"{rs_ms} | {lo_ms} | {r['speedup']} | "
                f"{'yes' if r['lo_supported']=='yes' else 'no'} |\n"
            )


if __name__ == "__main__":
    main()
