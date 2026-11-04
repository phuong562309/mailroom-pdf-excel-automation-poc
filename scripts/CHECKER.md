# Independent portable audit

Run from the repository root after processing a labeled PDF:

```text
python scripts/mailroom_audit_checker.py --source examples/Combined_Mailroom_Scan_Sample.pdf --output sample-output --audit AUDIT.json
```

The checker derives document boundaries and all expected metadata directly from the source through pdfplumber/pdfminer. It never imports the Maker, workbook writer or fixture manifests. It checks every source and child page through metadata, order, exact text, dimensions and a 40-dpi PDFium RGB fingerprint. PDFium is supplied through the public pdfplumber dependency; no external rendering executable is invoked.

The guard first tries the optional external module and uses the packaged standalone fallback when unavailable. Neither runtime changes the Python module search path or climbs outside the repository. Relative artifact references resolve from the directory supplied by `--output`.

The workbook is opened read-only with openpyxl. Every detail row and typed date is compared to source-derived values. All eleven configured summary types remain visible in subset batches. COUNTIFS, SUMIFS, SUM and nested SUBSTITUTE expressions and their saved caches are reconciled independently; non-mutating overlays probe each detail row's type/status/page dependencies. Case-variant status categories are rejected. Literal wildcard status characters are interpreted using Excel's escape rules. Streamed row scans handle worksheets with absent dimension metadata and detect non-empty rows after blanks.

Run the portability acceptance once with:

```text
python tests/test_portability_delta.py
```

The suite runs the public Maker and Checker CLIs, then copies the complete repository into a fresh temporary directory with spaces and Unicode. It repeats the full source-to-output audit with Python path overrides removed, user packages disabled and an empty executable search path. It proves that the packaged guard is selected, exercises growth/loss/identity/duplicate/projection guard cases, rejects image-only and case-variant inputs, checks no-overwrite behavior and tests a four-document subset with literal wildcard statuses. Content and import scans inspect requirement declarations, repository text and workbook ZIP payloads. The runtime dependencies remain the public Python packages declared in requirements.txt; the test does not install them.

The sample's exact 15 documents, 24 pages and 11 observed types are test acceptance values. The reusable audit derives N documents and M pages from source, allowing a subset of configured types. Tables with horizontal rules and repeated text alignment are recognized without fixed coordinates or expected cell values. The milestone preserves tables visually; it does not claim arbitrary table-cell extraction accuracy.

Acceptance reports describe synthetic labeled text PDFs and a local output tree. They do not establish OCR/customer-scan accuracy, remote storage synchronization, native spreadsheet-application execution or high-resolution typography quality. Existing audit outputs are never silently overwritten.
