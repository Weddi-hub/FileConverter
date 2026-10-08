# PDF to Excel (.xlsx) Converter

A reliable, production-ready Python tool that extracts all tables and structured data from a multi-page PDF document and exports the consolidated dataset into a **single continuous sheet (`Sheet1`)** in an Excel file.

---

## Key Features

- **Strict Single-Sheet Constraint**: All extracted tables across all pages are appended vertically into **`Sheet1`**. No separate sheets are created for different pages or tables.
- **No Index Column**: Exports clean data directly to `.xlsx` with `index=False`.
- **Intelligent Header Deduplication**: Automatically detects multi-page continuation tables and drops repeated header rows so your data rows flow smoothly.
- **Dual Table Detection Strategies**:
  - Primary: Line/grid-based extraction for standard bordered tables.
  - Fallback: Text-alignment detection for borderless or whitespace-aligned tables.
- **Robust Error Handling**:
  - Gracefully skips and logs pages with no tables without failing.
  - Diagnostic warnings for scanned (image-only) PDFs where digital text is absent.
  - Catches locked-file `PermissionError` when the target `.xlsx` is open in Microsoft Excel.
  - Handles password-protected PDFs.
- **Clean Formatting**: Normalizes unwanted internal cell linebreaks/spaces and automatically adjusts column widths for immediate readability.

---

## Installation & Prerequisites

Python 3.8+ is required.

Install the necessary dependencies via `pip`:

```bash
pip install -r requirements.txt
```

Or install the packages individually:

```bash
pip install pdfplumber pandas openpyxl
```

---

## Web Interface (Interactive UI)

You can launch the browser-based UI to drag-and-drop PDFs, adjust settings, preview the continuous table, and download Excel files:

```bash
python app.py
```
Open **[http://127.0.0.1:5000](http://127.0.0.1:5000)** in your web browser.

---

## Quick Start (Command-Line Interface)

### 1. Basic Conversion
Convert a PDF to an Excel file with the same base name (e.g., `report.pdf` -> `report.xlsx`):

```bash
python pdf_to_excel.py report.pdf
```

### 2. Specify Output Path
```bash
python pdf_to_excel.py report.pdf -o "C:/path/to/output.xlsx"
```

### 3. Include Structured Text Lines
Extract non-tabular structured lines (e.g., `Key: Value` pairs) on pages without tables:

```bash
python pdf_to_excel.py report.pdf --include-text
```

### 4. Advanced Concatenation Modes
- `--mode auto` *(Default)*: Automatically detects if tables share the same columns and unifies them under a single header row while stripping duplicate continuation headers. If tables vary in shape, stacks rows vertically with column padding.
- `--mode stack`: Vertically appends all rows across all tables padded to the maximum column width.
- `--mode concat`: Converts each table to a DataFrame and merges using `pd.concat` aligning matching column names.

```bash
python pdf_to_excel.py report.pdf --mode auto
```

### 5. Password-Protected PDFs
```bash
python pdf_to_excel.py protected_document.pdf --password "secret123"
```

---

## Python API Usage

You can also import and use the converter directly in your own Python projects:

```python
from pdf_to_excel import convert_pdf_to_excel

# Basic conversion
df, output_path = convert_pdf_to_excel("input.pdf", "output.xlsx")

# Access the unified Pandas DataFrame directly
print(df.head())
print(f"Total rows extracted: {len(df)}")
```

### Advanced Options in Python:

```python
from pdf_to_excel import convert_pdf_to_excel

df, output_path = convert_pdf_to_excel(
    pdf_path="quarterly_report.pdf",
    output_excel_path="quarterly_report.xlsx",
    sheet_name="Sheet1",              # Ensures export to 'Sheet1'
    mode="auto",                      # 'auto', 'stack', or 'concat'
    drop_duplicate_headers=True,      # Omit repeated headers on continuation pages
    clean_whitespace=True,            # Strip trailing spaces & newlines inside cells
    fallback_text_strategy=True,      # Detect borderless tables if no lines exist
    include_structured_text=False,    # Extract key-value lines on pages without tables
    password=None                     # PDF password if encrypted
)
```

---

## Running the Automated Test Suite

A comprehensive test suite is included in `test_converter.py` covering:
1. Multi-page continuous tables with repeated headers.
2. Tables with continuation rows lacking repeated headers.
3. Borderless table extraction with text-strategy fallback.
4. Pages with plain text (graceful non-table handling).
5. Scanned / empty PDF diagnostics.
6. Validation of the single-sheet constraint (`Sheet1`) and omission of index column.

Run the tests with:

```bash
python test_converter.py
```
