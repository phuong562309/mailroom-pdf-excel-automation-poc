# Mailroom PDF to Excel automation

A standalone Python proof of concept that splits a combined **text-layer PDF** into individual correspondence PDFs, extracts record metadata, and creates a two-sheet Excel report. The included immutable synthetic example has **15 documents, 24 pages and all 11 configured document types**, including one-, two- and three-page documents, horizontal-rule tables, Unicode names, and one person with two different statuses.

This is a local workflow. It does not perform OCR, classify unlabelled scans, upload documents, or connect to live SharePoint. `SharePoint_Synced_Outbox` is a local output folder that can later be placed inside an authorized sync location.

## Quick start

Use Python 3.10 or later. From the cloned or extracted repository:

```bash
python -m venv .venv
```

Activate it on Windows with `.venv\Scripts\activate`, or on macOS/Linux with `source .venv/bin/activate`. Install the declared public Python packages, then run the batch:

```bash
python -m pip install -r requirements.txt
python -X utf8 -B scripts/mailroom_splitter_maker.py --input examples/Combined_Mailroom_Scan_Sample.pdf --output-dir sample-output
```

`--output` is an alias of `--output-dir`. `--config config.json` optionally selects another configuration. The output directory must not exist. On success, exit code is 0; rejected input returns 2 with an actionable `MAILROOM_REJECTED` message. A batch is staged privately, verified for complete pages/records, and published by directory rename. Existing batches are never intentionally replaced. Cooperating processes serialize publication using an exclusive lock file; after an interrupted process, inspect the state before removing a stale `.publish.lock` file.

Run the independent Checker after Maker succeeds:

```bash
python -X utf8 -B scripts/mailroom_audit_checker.py --source examples/Combined_Mailroom_Scan_Sample.pdf --output sample-output --audit AUDIT.json
```

The delivery may already include `sample-output/` and `AUDIT.json` as proof artifacts. Use a new output directory and audit filename for another run; do not overwrite the included evidence. Checker uses `pdfplumber/pdfminer` and read-only `openpyxl`; Maker uses `pypdf`, so the two parsers derive their results separately. A Maker build marked `complete_unverified` still requires acceptance of that independent audit.

No private application libraries, JavaScript runtime, external PDF-rendering executable, environment-specific path or parent repository is required. Tested package versions are `pypdf 6.10.0`, `pdfplumber 0.11.9`, `openpyxl 3.1.5`, `reportlab 4.4.9` and `Pillow 12.3.0`; compatible version ranges are declared in `requirements.txt`. Installation needs access to your chosen Python package index; processing runs locally afterwards.

## Input protocol: mailroom-text-v1

The reusable algorithm derives N documents across M pages from labels and document-local page markers. It never uses the sample's document count or physical boundary positions.

Every page contains exactly one standalone line for each of:

```text
Document Type: Type_01
Document ID: MR-2026-0001
Page 1 of 3
```

Each document's first page also contains exactly one line for each of:

```text
Person Name: Nguyễn Thị Hà
Status: Received
Date: 2026-11-01
```

Values must be nonempty NFC Unicode; `Date` is a valid `YYYY-MM-DD` calendar date. The eleven allowed `Type_XX` labels are configured in `config.json`. Local page numbers start at 1, advance by one, and end at the declared total. Every continuation repeats the same ID, type and total. If a continuation repeats name/status/date, those values must agree. The next document starts only after the previous sequence ends. Document IDs are unique case-insensitively.

Empty/image-only pages, missing/duplicate labels, unknown types, conflicting metadata, orphan continuations and incomplete sequences are rejected before publication. Status categories differing only by case are rejected because Excel conditional counts compare text case-insensitively. Each source page must appear exactly once and in source order. A real unlabelled scan requires an additional OCR/template strategy and validation on representative client documents.

## Outputs

The batch directory contains:

- `SharePoint_Synced_Outbox/<Document_Type>/<Person_Name>_<Status>_<Document_ID>.pdf`
- `11_Types_Summary_Breakdown.xlsx`
- `maker_manifest.json` with complete records, relative source/output references, page ownership and SHA-256 hashes
- `maker_build.json` with preservation-guard results and explicit workbook-cache limitations

The workbook has exactly `11_Type_Summary` and `Extracted_Records`, with table headers at row 5. Detail columns A:J are `Doc_Index`, `Document_Type`, `Person_Name`, `Status`, `Document_ID`, `Date`, `Page_Range`, `Page_Count`, `Output_Filename`, `SharePoint_Relative_Path`. Dates are typed Excel values. `Page_Range` identifies 1-based physical source pages, distinct from the local page markers. Detail headers and the first three identity columns remain visible while scrolling (`D6` freeze panes). Both sheets use filterable tables, readable fonts and explicit number/date formats.

Summary rows cover all eleven configured types, even when a batch has no record of one type. Observed statuses become dynamic columns. `COUNTIFS`, `SUMIFS` and `SUM` formulas remain live in the workbook; nested `SUBSTITUTE` escapes literal status characters `~`, `*` and `?`.

**Formula caches:** `openpyxl` writes formulas but does not evaluate them. The pipeline independently computes initial numeric caches from all admitted records and inserts OOXML `<v>` values while preserving each `<f>` formula. It reopens the file in formula and cached-value modes to verify export. These caches refresh on every pipeline build. The workbook requests automatic/full recalculation when opened in a spreadsheet application; such an application must recalculate after edits. Reading with `data_only=True` alone does not recalculate. Native Microsoft Excel edit/recalculation behavior is not claimed as tested.

## Safe filenames and publication

Each filename component preserves NFC Unicode. Literal `%`, `_`, Windows forbidden characters (`<>:"/\|?*`), control characters, and trailing spaces/dots become uppercase percent escapes of UTF-8 bytes. Separating the encoded components with `_` is reversible: split into three components, then percent-decode. For example, `Alice / Bob` becomes `Alice %2F Bob`, and `A_B % Partner` becomes `A%5FB %25 Partner`. The composed basename contains separators and cannot equal a Windows device name.

The pipeline rejects case-insensitive relative-path collisions, duplicate record IDs, filenames beyond 255 UTF-16 units, and full paths beyond the extended Windows budget. Select a shallow output location for compatibility with older tools using the classic 260-character path limit. Source bytes are never changed and the source hash is checked before publish. File paths saved in metadata are repository-relative; an external input is represented by its basename and SHA-256, so hashes identify it without exposing its host directory.

## Standalone preservation guard

Maker and Checker first try importing `assert_no_filtered_view_overwrite` from an available `guardrail_checks` module. If unavailable, they use `scripts/standalone_guard.py`, without changing `sys.path` or searching ancestor directories. The packaged guard rejects projections, shrinking source stores, invalid counts, incomplete/duplicate identity lists, and loss of complete-source IDs. Maker exposes `GUARD_BACKEND` as `shared` or `standalone` for inspection. The guard supports the complete-source page and record checks; it does not replace independent content auditing.

## Optional sample regeneration

Keep the original example immutable. Generate a separate synthetic example with:

```bash
python -X utf8 -B scripts/generate_mailroom_sample.py --output new_sample.pdf
```

The generator refuses to overwrite existing files. It uses the unmodified bundled DejaVu Sans fonts in `assets/fonts/`, checks required Unicode glyph coverage, and never depends on installed system fonts. The exact embedded font redistribution license is included in `assets/fonts/LICENSE-DejaVu.txt`; the font copyrights remain with their owners. Regenerated PDF bytes differ from the original example because its original font is preserved inside that immutable fixture.

The synthetic sample and audit demonstrate this explicit text-layer protocol. They do not establish extraction accuracy on arbitrary client scans or prove live SharePoint synchronization.

## Delivery verification

`PORTABILITY_REPORT.json` records the independent relocation test and public package versions; `AUDIT.json` covers every page, document and saved workbook record. The test script refuses to overwrite existing evidence. To repeat the complete acceptance suite from the delivered package, copy the source files into a fresh directory while omitting `sample-output/`, `AUDIT.json` and `PORTABILITY_REPORT.json`, then run `python -X utf8 -B tests/test_portability_delta.py`.
