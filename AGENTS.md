# AGENTS.md

## Project overview

This repository converts PDF tables into a single Excel workbook using Python. The main logic lives in [pdf_to_excel.py](pdf_to_excel.py), the Flask web UI lives in [app.py](app.py), and the automated regression suite lives in [test_converter.py](test_converter.py). User-facing usage and examples are documented in [README.md](README.md).

## Repository conventions

- Keep exported data on one worksheet named `Sheet1`.
- Do not add an index column to output Excel files.
- Prefer logic that handles multi-page, continuation tables and borderless tables without breaking the single-sheet contract.
- Preserve graceful behavior for empty or scanned PDFs: skip table extraction when appropriate and log diagnostics instead of crashing.
- When adjusting extraction behavior, update or extend the relevant tests in [test_converter.py](test_converter.py) before finishing.

## Commands

Run the web app:

```bash
python app.py
```

Run the CLI converter:

```bash
python pdf_to_excel.py report.pdf
python pdf_to_excel.py report.pdf -o "C:/path/to/output.xlsx"
python pdf_to_excel.py report.pdf --include-text
python pdf_to_excel.py report.pdf --mode auto
```

Run the test suite:

```bash
python test_converter.py
```

## Key implementation notes

- `pdf_to_excel.py` contains the conversion pipeline, table-detection heuristics, and Excel export code.
- `app.py` is a Flask upload UI and should remain compatible with the converter API and file storage workflow.
- The project expects PDF table extraction to be robust across varied invoice and report layouts, including repeated headers across pages and borderless tables.
- If a change affects parsing quality, header deduplication, or output formatting, validate it against the existing regression tests.

## Working guidance for coding agents

- Prefer surgical changes in the conversion pipeline rather than rewriting the extractor.
- Reuse existing helper functions and patterns in [pdf_to_excel.py](pdf_to_excel.py) when possible.
- Preserve file naming and output-location behavior expected by the CLI and UI.
- Keep logs and exceptions informative, especially around lock conflicts, password-protected PDFs, or empty datasets.
- Before claiming a fix is complete, run the relevant test command and verify the result with fresh output.

## Relevant docs

- [README.md](README.md)
- [pdf_to_excel.py](pdf_to_excel.py)
- [app.py](app.py)
- [test_converter.py](test_converter.py)
