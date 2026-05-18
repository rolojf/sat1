# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Single-purpose tool that classifies Mexican CFDI 4.0 invoices (XML + PDF pairs) into SAT-aligned categories, refiles them into a year/month/category folder layout under `facturas/`, and produces `clasificacion.xlsx` + two CSV diagnostics (`skipped_xml.csv`, `pdf_sin_xml.csv`).

Python 3.12+, managed with `uv`. Only runtime deps are `openpyxl` and `pytest`.

## Commands

```bash
# Dry-run: parse everything, write the Excel report, but DO NOT move files.
uv run python clasifica.py

# Actually refile XML/PDF pairs into facturas/<YEAR>/<Mes><YEAR>/<YYMM>-<Cat>/
uv run python clasifica.py --apply

# Restrict the Excel report to one year (default: current year)
uv run python clasifica.py --year 2025 --apply

# Conservation tests — must be green before AND after --apply.
uv run python tests/snapshot.py     # record baseline first (writes tests/baseline.json)
uv run pytest                       # run conservation suite
uv run pytest tests/test_conservation.py::test_every_baseline_file_still_present  # one test
```

The conservation tests skip if `tests/baseline.json` is missing, so always run `snapshot.py` immediately before `clasifica.py --apply` and re-run pytest after.

## Architecture

`clasifica.py` is the whole product; everything else is supporting data or legacy.

Pipeline (top-to-bottom in `clasifica.py`):

1. **Pairing** (`discover_pairs`, `stem_key`) — walks `facturas/` recursively, normalizes filenames (strips `(1)` suffixes, `.xml.pdf` doubled extensions, casefolds) so each XML pairs with its companion PDF(s) regardless of how the SAT portal named the download.
2. **Parsing** (`parse_factura`) — XML namespace-aware; supports CFDI v4.0 and v3.3. Pulls emisor/receptor RFC, UsoCFDI, ClaveProdServ list, totals, IVA/ISR retentions, and the TFD UUID. Parse failures land in `skipped_xml.csv` and the companion PDF gets routed to `sin_xml/` rather than left scattered.
3. **Classification** (`classify`) — precedence order matters:
   - `TipoDeComprobante=P` → **Pagos**, `=N` → **Nomina**
   - `UsoCFDI` in D01..D10 → **DedPers** (personal deduction; flagged for review since it should not appear on a business RFC)
   - Auto override: emisor RFC in `AUTO_RFCS` (currently just Pemex) OR any `ClaveProdServ` matching `AUTO_PREFIXES` (1510 fuels, 2511/2512 vehicles, 2517 parts, 7811 transport, 7818 mainto., 8413 vehicle insurance) → **Auto**
   - `UsoCFDI=G01` → **AdqMerca**
   - everything else → **Gastos** (confidence "media" for G03, "baja" otherwise)
4. **Filing** (`file_pair`, `safe_move`, `file_orphan_pdf`) — destination is `facturas/<YEAR>/<MesAbr><YEAR>/<YYMM>-<Categoria>/`. Orphan PDFs (no XML) go to `facturas/<YEAR>/sin_xml/`. `safe_move` suffixes `__dupN` rather than overwriting on collision. `file_orphan_pdf` respects pre-existing manual classification: if a PDF is already in a `YYMM-Cat` or `sin_xml/` folder, it is left alone.
5. **Excel** (`write_excel`) — sheet `Facturas` (one row per CFDI, columns defined by `COLUMNS`) plus `Resumen` pivot of categoría × año-mes.

`tests/test_conservation.py` is a no-loss invariant: SHA1-hashes every `.xml`/`.pdf` under `facturas/` against `tests/baseline.json`. Catches drops, duplications, and unintended copies during refiling.

`cfdv40.xsd` is the official SAT CFDI v4.0 schema, kept for reference / future validation.

`fact.nu` and `xmls.nu` are predecessor nushell scripts that produced HTML summaries from XMLs. Superseded by `clasifica.py`; kept around but not part of the active workflow.

## Notes on the data layout

`facturas/<YEAR>/<MesAbr><YEAR>/<YYMM>-<Cat>/` is the canonical layout written by `clasifica.py --apply`. Anything outside that pattern (loose PDFs, manually-named folders) will be picked up and refiled on the next `--apply` unless it is already inside a `YYMM-Cat` or `sin_xml/` folder.

`MES_ABR` uses Spanish 3-letter month abbreviations: Ene, Feb, Mar, Abr, May, Jun, Jul, Ago, Sep, Oct, Nov, Dic.
