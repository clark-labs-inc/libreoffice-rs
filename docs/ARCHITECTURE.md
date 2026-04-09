# Architecture

## Workspace boundaries

`lo_core` owns the canonical document data structures. Every higher-level crate depends on those shared models.

`lo_zip` is a tiny std-only ZIP implementation that writes stored entries and parses the central directory. `lo_odf` builds on that to emit ODF-like archives.

`lo_writer`, `lo_calc`, `lo_impress`, `lo_draw`, `lo_math`, and `lo_base` add editing, parsing, and domain-specific logic. The CLI crate ties them together.

## Why std-only?

The repository intentionally avoids third-party crates so it can remain easy to audit and portable in constrained environments.

## Known gaps

This architecture is a foundation, not a replacement for LibreOffice. Layout, rendering, import/export compatibility, GUI work, and large parts of the office-suite surface area are still out of scope.
