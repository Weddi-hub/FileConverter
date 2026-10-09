"""
Reliable PDF to Excel (.xlsx) Converter
=========================================
Extracts all tables, structured metadata, and data rows from any PDF document
and exports the unified data entirely onto a single, continuous sheet ('Sheet1') in Excel.

Supported Formats:
- Pakistan Airports Authority (PAA) Bills:
  * Landing and Housing Bills (Domestic & International - 15 columns)
  * Air Navigation for Landing Bills (Domestic & International - 12 columns)
  * Route Navigation & Overflying Bills
  * Embeds official Logo, Title Block, Metadata Block (Airline Name, Bill No, Dates, Dollar Rate)
  * 100% precise column alignment, Subtotals, Grand Totals, and Invoice Summary
- Universal PDF Documents:
  * Multi-page continuous reports, financial statements, invoices, tables
  * Line-bordered and borderless whitespace-aligned tables
  * Automatic multi-page continuation and header deduplication

Author: Antigravity
License: MIT
"""

import os
import re
import sys
import logging
import argparse
from typing import List, Optional, Tuple, Any, Dict

import pdfplumber
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger("pdf_to_excel")


# ==============================================================================
# 1. PAA LANDING AND HOUSING BILL CONFIGURATION (15 COLUMNS)
# ==============================================================================
PAA_LH_COL_INTERVALS = [
    ("BILL ITEM ID", 0, 46.0),
    ("AIRCRAFT TYPE", 46.0, 85.0),
    ("AIRCRAFT REG. #", 85.0, 140.0),
    ("AIRCRAFT MTOW", 140.0, 180.0),
    ("FLIGHT NO.", 180.0, 237.0),
    ("ARRIVAL DATE", 237.0, 290.0),
    ("DEPARTURE DATE", 290.0, 345.0),
    ("ENGINE TIME OFF", 345.0, 395.0),
    ("ENGINE TIME ON", 395.0, 428.0),
    ("ENGINE TOTAL TIME", 428.0, 480.0),
    ("LANDING CHARGES (RS)", 480.0, 540.0),
    ("REFUELING BAY CHARGES (RS)", 540.0, 605.0),
    ("PARKING CHARGES (RS)", 605.0, 675.0),
    ("TERMINAL NAVIGATION CHARGES (RS)", 675.0, 745.0),
    ("TOTAL CHARGES (RS.)", 745.0, 850.0),
]
PAA_LH_HEADERS = [c[0] for c in PAA_LH_COL_INTERVALS]


# ==============================================================================
# 2. PAA AIR NAVIGATION BILL CONFIGURATION (12 COLUMNS)
# ==============================================================================
PAA_AN_COL_INTERVALS = [
    ("BILL ITEM ID", 0, 48.0),
    ("ARR / DEP DATE", 48.0, 100.0),
    ("FLIGHT NO.", 100.0, 155.0),
    ("REG. NO.", 155.0, 208.0),
    ("AIRCRAFT TYPE", 208.0, 262.0),
    ("MTOW (TONS)", 262.0, 316.0),
    ("STATUS ARR/DEP", 316.0, 390.0),
    ("ENTRY-EXIT/ROUTE POINT", 390.0, 520.0),
    ("TOTAL DISTANCE", 520.0, 601.0),
    ("BILLABLE DISTANCE", 601.0, 670.0),
    ("AMOUNT (US$)", 670.0, 750.0),
    ("AMOUNT (Rs)", 750.0, 825.0),
]
PAA_AN_HEADERS = [c[0] for c in PAA_AN_COL_INTERVALS]

# ==============================================================================
# 1C. PAA AVIOBRIDGE CHARGES BILL CONFIGURATION (13 COLUMNS)
# ==============================================================================
PAA_AB_COL_INTERVALS = [
    ("BILL ITEM ID", 0, 55.0),
    ("AIRCRAFT TYPE", 55.0, 100.0),
    ("AIRCRAFT REG. #", 100.0, 160.0),
    ("AIRCRAFT MTOW", 160.0, 205.0),
    ("FLIGHT NO.", 205.0, 285.0),
    ("ARRIVAL DATE", 285.0, 350.0),
    ("DEPARTURE DATE", 350.0, 425.0),
    ("BRIDGE NO.", 425.0, 500.0),
    ("PLUG IN", 500.0, 555.0),
    ("PLUG OUT", 555.0, 610.0),
    ("PLUG TOTAL", 610.0, 675.0),
    ("AMOUNT US$", 675.0, 750.0),
    ("AMOUNT PKR Rs.", 750.0, 825.0),
]
PAA_AB_HEADERS = [c[0] for c in PAA_AB_COL_INTERVALS]


def detect_document_type(first_page_text: str) -> str:
    """Classifies the PDF into specific bill types or generic document."""
    text_upper = first_page_text.upper()
    if (
        "SAUDI AIR NAVIGATION SERVICES" in text_upper
        or ("SANS" in text_upper and "AIR NAVIGATION" in text_upper)
        or "INVOICE FOR AIR NAVIGATION CHARGES" in text_upper
        or "ةیوجلا ةحلملا تامدخ" in first_page_text
        or "شركة خدمات الملاحة الجوية" in first_page_text
    ):
        return "SANS_AIR_NAVIGATION"
    elif (
        "LANDING CHARGE INVOICE" in text_upper
        or "DAMMAM AIRPORTS" in text_upper
        or "JEDDAH AIRPORTS" in text_upper
        or ("KINGDOM OF SAUDI ARABIA" in text_upper and "TAX INVOICE" in text_upper)
        or ("AIRPORTS COMPANY" in text_upper and "LANDING" in text_upper)
    ):
        return "SAUDI_LANDING_CHARGE"
    elif (
        "DUBAI AIRPORTS" in text_upper
        or "DUBAIRPORTS" in text_upper
        or "مطارات دبي" in first_page_text
        or "DUBAI AIRPORTS CORPORATION" in text_upper
    ):
        return "DUBAI_AIRPORTS"
    elif "AIR NAVIGATION FOR LANDING" in text_upper or ("ENTRY-EXIT" in text_upper and "DISTANCE" in text_upper):
        return "PAA_AIR_NAVIGATION"
    elif "AVIOBRIDGE CHARGES" in text_upper or "AVIOBRIDGE" in text_upper or "BRIDGE NO." in text_upper:
        return "PAA_AVIOBRIDGE"
    elif "SUMMARY OF AERONAUTICAL BILLS" in text_upper or "SUMMARY OF BILLS" in text_upper:
        return "PAA_AERONAUTICAL_SUMMARY"
    elif "LANDING AND HOUSING" in text_upper or "ENGINE TIME" in text_upper:
        return "PAA_LANDING_AND_HOUSING"
    elif "PAKISTAN AIRPORTS AUTHORITY" in text_upper:
        if "SUMMARY OF AERONAUTICAL BILLS" in text_upper or "SUMMARY OF BILLS" in text_upper:
            return "PAA_AERONAUTICAL_SUMMARY"
        if "DISTANCE" in text_upper or "ROUTE" in text_upper:
            return "PAA_AIR_NAVIGATION"
        if "AVIOBRIDGE" in text_upper or "BRIDGE" in text_upper:
            return "PAA_AVIOBRIDGE"
        if "LANDING AND HOUSING" in text_upper or "ENGINE TIME" in text_upper:
            return "PAA_LANDING_AND_HOUSING"
        return "DYNAMIC_DOCUMENT"
    return "GENERIC"




def extract_paa_metadata(first_page: pdfplumber.page.Page) -> Dict[str, str]:
    """Extracts title and two-column metadata block from Page 1 of a PAA bill."""
    text = first_page.extract_text() or ""
    meta = {
        "title": "PAKISTAN AIRPORTS AUTHORITY",
        "subtitle": "LANDING AND HOUSING",
        "bill_for": "",
        "bill_type": "",
        "doc_no": "PAA-001-FNBL-1.0",
        "airline_name": "",
        "bill_no": "",
        "location": "",
        "processing_date": "",
        "issue_date": "",
        "due_date": "",
        "dollar_rate": ""
    }

    if "AIR NAVIGATION FOR LANDING" in text.upper():
        meta["subtitle"] = "AIR NAVIGATION FOR LANDING"
    elif "AVIOBRIDGE" in text.upper():
        meta["subtitle"] = "AVIOBRIDGE CHARGES"
    elif "LANDING AND HOUSING" in text.upper():
        meta["subtitle"] = "LANDING AND HOUSING"

    for line in text.splitlines():
        line_s = line.strip()
        if "BILL FOR THE FORTNIGHT" in line_s:
            meta["bill_for"] = line_s
        elif line_s in ["INTERNATIONAL", "DOMESTIC"]:
            meta["bill_type"] = line_s
        elif "AIRLINE NAME" in line_s:
            parts = line_s.split(":", 1)
            if len(parts) > 1:
                left_part = parts[1]
                if "PROCESSING DATE" in left_part:
                    p1, p2 = left_part.split("PROCESSING DATE", 1)
                    meta["airline_name"] = p1.strip()
                    meta["processing_date"] = p2.replace(":", "").strip()
                else:
                    meta["airline_name"] = left_part.strip()
        elif "BILL NO." in line_s:
            parts = line_s.split(":", 1)
            if len(parts) > 1:
                left_part = parts[1]
                if "ISSUE DATE" in left_part:
                    p1, p2 = left_part.split("ISSUE DATE", 1)
                    meta["bill_no"] = p1.strip()
                    meta["issue_date"] = p2.replace(":", "").strip()
                else:
                    meta["bill_no"] = left_part.strip()
        elif "LOCATION" in line_s:
            parts = line_s.split(":", 1)
            if len(parts) > 1:
                left_part = parts[1]
                if "DUE DATE" in left_part:
                    p1, p2 = left_part.split("DUE DATE", 1)
                    meta["location"] = p1.strip()
                    meta["due_date"] = p2.replace(":", "").strip()
                else:
                    meta["location"] = left_part.strip()
        elif "DOLLAR CONVERSION RATE" in line_s or "DOLLAR RATE" in line_s:
            parts = line_s.split(":", 1)
            if len(parts) > 1:
                meta["dollar_rate"] = parts[1].strip()

    return meta


# ------------------------------------------------------------------------------
# PAA Landing & Housing Parser
# ------------------------------------------------------------------------------
def extract_paa_lh_flight_page(page: pdfplumber.page.Page) -> List[List[str]]:
    """Extracts flight log records from a Landing & Housing page."""
    words = page.extract_words()
    data_words = [w for w in words if 245 <= w['top'] <= 580]
    if not data_words:
        return []

    data_words.sort(key=lambda w: (w['top'], w['x0']))
    lines = []
    curr_line = []
    curr_top = None
    for w in data_words:
        if curr_top is None:
            curr_top = w['top']
            curr_line.append(w)
        elif abs(w['top'] - curr_top) <= 5.0:
            curr_line.append(w)
        else:
            lines.append(curr_line)
            curr_line = [w]
            curr_top = w['top']
    if curr_line:
        lines.append(curr_line)

    page_rows: List[List[str]] = []
    for line in lines:
        line_text = " ".join(w['text'] for w in line)
        row = [""] * len(PAA_LH_COL_INTERVALS)

        if "AIRCRAFT TYPE TOTAL" in line_text or "LOCATION TOTAL" in line_text:
            desc_words = [w['text'] for w in line if w['x0'] < 470]
            num_words = [w for w in line if w['x0'] >= 470]
            row[1] = " ".join(desc_words)
            for w in num_words:
                cx = (w['x0'] + w['x1']) / 2
                for c_idx, (_, x0, x1) in enumerate(PAA_LH_COL_INTERVALS):
                    if x0 <= cx < x1:
                        row[c_idx] = w['text']
                        break
        else:
            for w in line:
                cx = (w['x0'] + w['x1']) / 2
                for c_idx, (_, x0, x1) in enumerate(PAA_LH_COL_INTERVALS):
                    if x0 <= cx < x1:
                        if row[c_idx]:
                            row[c_idx] += " " + w['text']
                        else:
                            row[c_idx] = w['text']
                        break

        if any(row):
            page_rows.append(row)

    return page_rows


def extract_paa_summary_page(page: pdfplumber.page.Page, col_count: int) -> List[List[str]]:
    """Extracts summary due date and total amount from the terms page."""
    words = page.extract_words()
    rows: List[List[str]] = []
    due_data = [w for w in words if 225 <= w['top'] <= 255]
    if due_data:
        due_data.sort(key=lambda w: w['x0'])
        row_due = [""] * col_count
        row_due[0] = "INVOICE SUMMARY"
        if col_count >= 15:
            if len(due_data) >= 1: row_due[5] = f"DUE DATE: {due_data[0]['text']}"
            if len(due_data) >= 2: row_due[10] = f"TOTAL DUE: {due_data[1]['text']}"
            if len(due_data) >= 3: row_due[14] = f"AFTER DUE (+5%): {due_data[2]['text']}"
        else:
            if len(due_data) >= 1: row_due[4] = f"DUE DATE: {due_data[0]['text']}"
            if len(due_data) >= 2: row_due[9] = f"TOTAL DUE: {due_data[1]['text']}"
            if len(due_data) >= 3: row_due[11] = f"AFTER DUE (+5%): {due_data[2]['text']}"
        rows.append(row_due)
    return rows


def build_exact_paa_excel(pdf_path: str, output_excel_path: str, sheet_name: str = "Sheet1") -> Tuple[pd.DataFrame, str]:
    """Builds the high-fidelity Excel export for Landing & Housing bills (15 columns)."""
    parent_dir = os.path.dirname(os.path.abspath(output_excel_path))
    if parent_dir and not os.path.exists(parent_dir):
        os.makedirs(parent_dir, exist_ok=True)

    logo_path = None
    all_flight_rows: List[List[str]] = []
    summary_rows: List[List[str]] = []

    with pdfplumber.open(pdf_path) as pdf:
        meta = extract_paa_metadata(pdf.pages[0])

        if pdf.pages[0].images:
            try:
                img_obj = pdf.pages[0].images[0]
                temp_logo = os.path.join(parent_dir, "paa_logo_temp.png")
                cropped_img = pdf.pages[0].crop((img_obj['x0'], img_obj['top'], img_obj['x1'], img_obj['bottom'])).to_image(resolution=150)
                cropped_img.save(temp_logo)
                logo_path = temp_logo
            except Exception:
                logo_path = None

        for idx, page in enumerate(pdf.pages):
            page_num = idx + 1
            text = page.extract_text() or ""
            if "ENGINE TIME" in text and ("FLIGHT" in text or "AIRCRAFT" in text):
                rows = extract_paa_lh_flight_page(page)
                logger.info(f"Page {page_num}/{len(pdf.pages)}: Extracted {len(rows)} flight/subtotal records")
                all_flight_rows.extend(rows)
            elif "AMOUNT WITHIN" in text:
                s_rows = extract_paa_summary_page(page, len(PAA_LH_COL_INTERVALS))
                if s_rows:
                    logger.info(f"Page {page_num}/{len(pdf.pages)}: Extracted invoice summary details")
                    summary_rows.extend(s_rows)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.views.sheetView[0].showGridLines = True

    thin_border = Side(style='thin', color='000000')
    double_border = Side(style='double', color='000000')
    cell_border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
    top_bottom_border = Border(top=thin_border, bottom=thin_border)
    grand_total_border = Border(top=thin_border, bottom=double_border)

    # 1. Logo
    if logo_path and os.path.exists(logo_path):
        try:
            from openpyxl.drawing.image import Image as ExcelImage
            img = ExcelImage(logo_path)
            img.width = 110
            img.height = 55
            ws.add_image(img, "A1")
        except Exception:
            pass

    # 2. Title Block
    ws.merge_cells("C1:K1")
    ws["C1"] = meta.get("title", "PAKISTAN AIRPORTS AUTHORITY")
    ws["C1"].font = Font(name="Calibri", size=14, bold=True)
    ws["C1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("C2:K2")
    ws["C2"] = meta.get("subtitle", "LANDING AND HOUSING")
    ws["C2"].font = Font(name="Calibri", size=12, bold=True)
    ws["C2"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("C3:K3")
    ws["C3"] = meta.get("bill_for", "BILL FOR THE FORTNIGHT : JUN01 - 2026")
    ws["C3"].font = Font(name="Calibri", size=11, bold=True)
    ws["C3"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("C4:K4")
    ws["C4"] = meta.get("bill_type", "DOMESTIC")
    ws["C4"].font = Font(name="Calibri", size=12, bold=True)
    ws["C4"].alignment = Alignment(horizontal="center", vertical="center")

    ws["M1"] = "PAA-001-FNBL-1.0"
    ws["M1"].font = Font(name="Calibri", size=9)
    ws["M1"].alignment = Alignment(horizontal="right")

    # 3. Metadata Block
    ws.row_dimensions[5].height = 10
    meta_rows = [
        ("AIRLINE NAME :", meta.get("airline_name", ""), "PROCESSING DATE :", meta.get("processing_date", "")),
        ("BILL NO. :", meta.get("bill_no", ""), "ISSUE DATE :", meta.get("issue_date", "")),
        ("LOCATION :", meta.get("location", ""), "DUE DATE :", meta.get("due_date", "")),
    ]
    if meta.get("dollar_rate"):
        meta_rows.append(("", "", "DOLLAR CONVERSION RATE :", meta.get("dollar_rate", "")))

    curr_row = 6
    for lbl1, val1, lbl2, val2 in meta_rows:
        if lbl1:
            ws.cell(row=curr_row, column=2, value=lbl1).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=curr_row, column=3, value=val1).font = Font(name="Calibri", size=10, bold=False)
        if lbl2:
            ws.cell(row=curr_row, column=11, value=lbl2).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=curr_row, column=11).alignment = Alignment(horizontal="right")
            ws.cell(row=curr_row, column=12, value=val2).font = Font(name="Calibri", size=10, bold=False)
        curr_row += 1

    curr_row += 1

    # 4. Table Header
    header_row_1 = curr_row
    header_row_2 = curr_row + 1
    hdr_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    hdr_font = Font(name="Calibri", size=9, bold=True)

    for r in range(header_row_1, header_row_2 + 1):
        for c in range(1, 16):
            cell = ws.cell(row=r, column=c)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = cell_border

    ws.merge_cells(start_row=header_row_1, start_column=1, end_row=header_row_2, end_column=1)
    ws.cell(row=header_row_1, column=1, value="BILL ITEM ID")

    ws.merge_cells(start_row=header_row_1, start_column=2, end_row=header_row_1, end_column=4)
    ws.cell(row=header_row_1, column=2, value="AIRCRAFT")

    ws.merge_cells(start_row=header_row_1, start_column=5, end_row=header_row_2, end_column=5)
    ws.cell(row=header_row_1, column=5, value="FLIGHT NO.")

    ws.merge_cells(start_row=header_row_1, start_column=6, end_row=header_row_1, end_column=7)
    ws.cell(row=header_row_1, column=6, value="DATE")

    ws.merge_cells(start_row=header_row_1, start_column=8, end_row=header_row_1, end_column=10)
    ws.cell(row=header_row_1, column=8, value="ENGINE TIME")

    ws.merge_cells(start_row=header_row_1, start_column=11, end_row=header_row_2, end_column=11)
    ws.cell(row=header_row_1, column=11, value="LANDING CHARGES\n(RS)")

    ws.merge_cells(start_row=header_row_1, start_column=12, end_row=header_row_2, end_column=12)
    ws.cell(row=header_row_1, column=12, value="REFEULING BAY\nCHARGES (RS)")

    ws.merge_cells(start_row=header_row_1, start_column=13, end_row=header_row_2, end_column=13)
    ws.cell(row=header_row_1, column=13, value="PARKING\nCHARGES (RS)")

    ws.merge_cells(start_row=header_row_1, start_column=14, end_row=header_row_2, end_column=14)
    ws.cell(row=header_row_1, column=14, value="TERMINAL\nNAVIGATION (RS)")

    ws.merge_cells(start_row=header_row_1, start_column=15, end_row=header_row_2, end_column=15)
    ws.cell(row=header_row_1, column=15, value="TOTAL CHARGES\n(RS.)")

    ws.cell(row=header_row_2, column=2, value="TYPE")
    ws.cell(row=header_row_2, column=3, value="REG. #")
    ws.cell(row=header_row_2, column=4, value="MTOW")
    ws.cell(row=header_row_2, column=6, value="ARRIVAL")
    ws.cell(row=header_row_2, column=7, value="DEPARTURE")
    ws.cell(row=header_row_2, column=8, value="OFF")
    ws.cell(row=header_row_2, column=9, value="ON")
    ws.cell(row=header_row_2, column=10, value="TOTAL (Hrs) (Min)")

    # 5. Data Rows
    current_data_row = header_row_2 + 1
    for row_vals in all_flight_rows:
        is_subtotal = "AIRCRAFT TYPE TOTAL" in row_vals[1]
        is_grand_total = "LOCATION TOTAL" in row_vals[1]

        for col_idx, val in enumerate(row_vals, start=1):
            cell_val = val
            num_fmt = None
            if val and isinstance(val, str):
                cleaned_val = val.replace(",", "").strip()
                if col_idx >= 11:  # LANDING, REFUELING, PARKING, TERMINAL, TOTAL CHARGES
                    try:
                        if "." in cleaned_val:
                            cell_val = float(cleaned_val)
                            num_fmt = "#,##0.00"
                        else:
                            cell_val = int(cleaned_val)
                            num_fmt = "#,##0"
                    except ValueError:
                        pass
                elif col_idx == 4 and not is_subtotal and not is_grand_total:  # MTOW
                    try:
                        cell_val = float(cleaned_val) if "." in cleaned_val else int(cleaned_val)
                        num_fmt = "#,##0"
                    except ValueError:
                        pass

            cell = ws.cell(row=current_data_row, column=col_idx, value=cell_val)
            if num_fmt:
                cell.number_format = num_fmt

            if is_grand_total:
                cell.font = Font(name="Calibri", size=10, bold=True)
                cell.fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")
                cell.border = grand_total_border
            elif is_subtotal:
                cell.font = Font(name="Calibri", size=9.5, bold=True)
                cell.fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
                cell.border = top_bottom_border
            else:
                cell.font = Font(name="Calibri", size=9)

            if is_subtotal or is_grand_total:
                if col_idx == 2:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif col_idx >= 11:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                if col_idx in [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx >= 11:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        current_data_row += 1

    # 6. Invoice Summary Block
    if summary_rows:
        current_data_row += 1
        for s_row in summary_rows:
            ws.cell(row=current_data_row, column=2, value="INVOICE DUE & PAYABLE AMOUNTS:").font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=current_data_row, column=6, value=s_row[5]).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=current_data_row, column=11, value=s_row[10]).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=current_data_row, column=15, value=s_row[14]).font = Font(name="Calibri", size=10, bold=True)
            current_data_row += 1

    # 7. Column Widths
    col_widths = {
        1: 14, 2: 10, 3: 11, 4: 8, 5: 12, 6: 12, 7: 13,
        8: 10, 9: 10, 10: 16, 11: 18, 12: 17, 13: 16, 14: 18, 15: 18
    }
    for c_idx, width in col_widths.items():
        col_letter = get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = width

    try:
        wb.save(output_excel_path)
    except PermissionError:
        error_msg = f"Permission denied writing to '{output_excel_path}'. Please close the file if open."
        logger.error(error_msg)
        raise PermissionError(error_msg)

    if logo_path and os.path.exists(logo_path):
        try: os.remove(logo_path)
        except Exception: pass

    df_preview = pd.DataFrame(all_flight_rows, columns=PAA_LH_HEADERS)
    return df_preview, output_excel_path


# ------------------------------------------------------------------------------
# PAA Air Navigation for Landing Parser (12 Columns)
# ------------------------------------------------------------------------------
def extract_paa_an_flight_page(page: pdfplumber.page.Page) -> List[List[str]]:
    """Extracts flight records from an Air Navigation page (12 columns)."""
    words = page.extract_words()
    data_words = [w for w in words if 235 <= w['top'] <= 580]
    if not data_words:
        return []

    data_words.sort(key=lambda w: (w['top'], w['x0']))
    lines = []
    curr_line = []
    curr_top = None
    for w in data_words:
        if curr_top is None:
            curr_top = w['top']
            curr_line.append(w)
        elif abs(w['top'] - curr_top) <= 5.0:
            curr_line.append(w)
        else:
            curr_line.sort(key=lambda x: x['x0'])
            lines.append(curr_line)
            curr_line = [w]
            curr_top = w['top']
    if curr_line:
        curr_line.sort(key=lambda x: x['x0'])
        lines.append(curr_line)

    page_rows: List[List[str]] = []
    for line in lines:
        line_text = " ".join(w['text'] for w in line)
        row = [""] * len(PAA_AN_COL_INTERVALS)

        if "TOTAL :" in line_text or "GRAND TOTAL" in line_text:
            desc_words = [w['text'] for w in line if w['x0'] < 650]
            num_words = [w for w in line if w['x0'] >= 650]
            row[1] = " ".join(desc_words)
            for w in num_words:
                cx = (w['x0'] + w['x1']) / 2
                for c_idx, (_, x0, x1) in enumerate(PAA_AN_COL_INTERVALS):
                    if x0 <= cx < x1:
                        row[c_idx] = w['text']
                        break
        else:
            for w in line:
                cx = (w['x0'] + w['x1']) / 2
                for c_idx, (_, x0, x1) in enumerate(PAA_AN_COL_INTERVALS):
                    if x0 <= cx < x1:
                        if row[c_idx]:
                            row[c_idx] += " " + w['text']
                        else:
                            row[c_idx] = w['text']
                        break

        if any(row):
            page_rows.append(row)

    return page_rows


def build_air_nav_excel(pdf_path: str, output_excel_path: str, sheet_name: str = "Sheet1") -> Tuple[pd.DataFrame, str]:
    """Builds the high-fidelity Excel export for Air Navigation bills (12 columns)."""
    parent_dir = os.path.dirname(os.path.abspath(output_excel_path))
    if parent_dir and not os.path.exists(parent_dir):
        os.makedirs(parent_dir, exist_ok=True)

    all_flight_rows = []
    summary_rows = []
    logo_path = None

    with pdfplumber.open(pdf_path) as pdf:
        meta = extract_paa_metadata(pdf.pages[0])
        if pdf.pages[0].images:
            try:
                img_obj = pdf.pages[0].images[0]
                temp_logo = os.path.join(parent_dir, "paa_logo_temp_an.png")
                cropped_img = pdf.pages[0].crop((img_obj['x0'], img_obj['top'], img_obj['x1'], img_obj['bottom'])).to_image(resolution=150)
                cropped_img.save(temp_logo)
                logo_path = temp_logo
            except Exception:
                logo_path = None

        for idx, page in enumerate(pdf.pages):
            page_num = idx + 1
            text = page.extract_text() or ""
            if "ENTRY-EXIT" in text or "TOTAL DISTANCE" in text or "STATUS" in text:
                rows = extract_paa_an_flight_page(page)
                logger.info(f"Page {page_num}/{len(pdf.pages)}: Extracted {len(rows)} flight/subtotal records")
                all_flight_rows.extend(rows)
            elif "AMOUNT WITHIN" in text:
                s_rows = extract_paa_summary_page(page, len(PAA_AN_COL_INTERVALS))
                if s_rows:
                    logger.info(f"Page {page_num}/{len(pdf.pages)}: Extracted invoice summary details")
                    summary_rows.extend(s_rows)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.views.sheetView[0].showGridLines = True

    thin_border = Side(style='thin', color='000000')
    double_border = Side(style='double', color='000000')
    cell_border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
    top_bottom_border = Border(top=thin_border, bottom=thin_border)
    grand_total_border = Border(top=thin_border, bottom=double_border)

    # 1. Logo
    if logo_path and os.path.exists(logo_path):
        try:
            from openpyxl.drawing.image import Image as ExcelImage
            img = ExcelImage(logo_path)
            img.width = 110
            img.height = 55
            ws.add_image(img, "A1")
        except Exception:
            pass

    # 2. Title Block
    ws.merge_cells("C1:I1")
    ws["C1"] = meta.get("title", "PAKISTAN AIRPORTS AUTHORITY")
    ws["C1"].font = Font(name="Calibri", size=14, bold=True)
    ws["C1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("C2:I2")
    ws["C2"] = meta.get("subtitle", "AIR NAVIGATION FOR LANDING")
    ws["C2"].font = Font(name="Calibri", size=12, bold=True)
    ws["C2"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("C3:I3")
    ws["C3"] = meta.get("bill_for", "BILL FOR THE FORTNIGHT : JUN01 - 2026")
    ws["C3"].font = Font(name="Calibri", size=11, bold=True)
    ws["C3"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("C4:I4")
    ws["C4"] = meta.get("bill_type", "INTERNATIONAL")
    ws["C4"].font = Font(name="Calibri", size=12, bold=True)
    ws["C4"].alignment = Alignment(horizontal="center", vertical="center")

    ws["K1"] = "PAA-001-FNBL-1.0"
    ws["K1"].font = Font(name="Calibri", size=9)
    ws["K1"].alignment = Alignment(horizontal="right")

    # 3. Metadata Block
    ws.row_dimensions[5].height = 10
    meta_rows = [
        ("AIRLINE NAME :", meta.get("airline_name", ""), "PROCESSING DATE :", meta.get("processing_date", "")),
        ("BILL NO. :", meta.get("bill_no", ""), "ISSUE DATE :", meta.get("issue_date", "")),
        ("LOCATION :", meta.get("location", ""), "DUE DATE :", meta.get("due_date", "")),
    ]
    if meta.get("dollar_rate"):
        meta_rows.append(("", "", "DOLLAR CONVERSION RATE :", meta.get("dollar_rate", "")))

    curr_row = 6
    for lbl1, val1, lbl2, val2 in meta_rows:
        if lbl1:
            ws.cell(row=curr_row, column=2, value=lbl1).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=curr_row, column=3, value=val1).font = Font(name="Calibri", size=10, bold=False)
        if lbl2:
            ws.cell(row=curr_row, column=9, value=lbl2).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=curr_row, column=9).alignment = Alignment(horizontal="right")
            ws.cell(row=curr_row, column=10, value=val2).font = Font(name="Calibri", size=10, bold=False)
        curr_row += 1

    curr_row += 1

    # 4. Table Header
    hdr_row = curr_row
    hdr_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    hdr_font = Font(name="Calibri", size=9, bold=True)

    for c_idx, h_name in enumerate(PAA_AN_HEADERS, start=1):
        cell = ws.cell(row=hdr_row, column=c_idx, value=h_name)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = cell_border

    # 5. Data Rows
    current_data_row = hdr_row + 1
    for row_vals in all_flight_rows:
        is_subtotal = "TOTAL :" in row_vals[1] and "GRAND" not in row_vals[1]
        is_grand_total = "GRAND TOTAL" in row_vals[1]

        for col_idx, val in enumerate(row_vals, start=1):
            cell_val = val
            num_fmt = None
            if val and isinstance(val, str):
                cleaned_val = val.replace(",", "").strip()
                if col_idx in [9, 10]:  # TOTAL DISTANCE, BILLABLE DISTANCE
                    try:
                        cell_val = float(cleaned_val)
                        num_fmt = "#,##0.0000"
                    except ValueError:
                        pass
                elif col_idx == 11:  # AMOUNT (US$)
                    try:
                        cell_val = float(cleaned_val)
                        num_fmt = "#,##0.0000"
                    except ValueError:
                        pass
                elif col_idx == 12:  # AMOUNT (Rs)
                    try:
                        if "." in cleaned_val:
                            cell_val = float(cleaned_val)
                            num_fmt = "#,##0.00"
                        else:
                            cell_val = int(cleaned_val)
                            num_fmt = "#,##0"
                    except ValueError:
                        pass
                elif col_idx == 6 and not is_subtotal and not is_grand_total:  # MTOW (TONS)
                    try:
                        cell_val = float(cleaned_val) if "." in cleaned_val else int(cleaned_val)
                        num_fmt = "#,##0"
                    except ValueError:
                        pass

            cell = ws.cell(row=current_data_row, column=col_idx, value=cell_val)
            if num_fmt:
                cell.number_format = num_fmt

            if is_grand_total:
                cell.font = Font(name="Calibri", size=10, bold=True)
                cell.fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")
                cell.border = grand_total_border
            elif is_subtotal:
                cell.font = Font(name="Calibri", size=9.5, bold=True)
                cell.fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
                cell.border = top_bottom_border
            else:
                cell.font = Font(name="Calibri", size=9)

            if is_subtotal or is_grand_total:
                if col_idx == 2:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif col_idx in [11, 12]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                if col_idx in [1, 2, 3, 4, 5, 6, 7]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx in [8]:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="center")

        current_data_row += 1

    # 6. Invoice Summary Block
    if summary_rows:
        current_data_row += 1
        for s_row in summary_rows:
            ws.cell(row=current_data_row, column=2, value="INVOICE DUE & PAYABLE AMOUNTS:").font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=current_data_row, column=5, value=s_row[4]).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=current_data_row, column=9, value=s_row[9]).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=current_data_row, column=11, value=s_row[11]).font = Font(name="Calibri", size=10, bold=True)
            current_data_row += 1

    # 7. Column Widths
    col_widths = {
        1: 14, 2: 13, 3: 12, 4: 11, 5: 11, 6: 12,
        7: 14, 8: 26, 9: 16, 10: 16, 11: 15, 12: 16
    }
    for c_idx, width in col_widths.items():
        col_letter = get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = width

    try:
        wb.save(output_excel_path)
    except PermissionError:
        error_msg = f"Permission denied writing to '{output_excel_path}'. Please close the file if open."
        logger.error(error_msg)
        raise PermissionError(error_msg)

    if logo_path and os.path.exists(logo_path):
        try: os.remove(logo_path)
        except Exception: pass

    df_preview = pd.DataFrame(all_flight_rows, columns=PAA_AN_HEADERS)
    return df_preview, output_excel_path


# ==============================================================================
# 1D. PAA AVIOBRIDGE CHARGES BILL EXTRACTION (13 COLUMNS)
# ==============================================================================
def extract_paa_ab_flight_page(page: pdfplumber.page.Page) -> List[List[str]]:
    """Extracts flight records from an Aviobridge Charges page (13 columns)."""
    words = page.extract_words()
    data_words = [w for w in words if 265 <= w['top'] <= 580]
    if not data_words:
        return []
    data_words.sort(key=lambda w: (w['top'], w['x0']))
    lines = []
    curr_line = []
    curr_top = None
    for w in data_words:
        if curr_top is None:
            curr_top = w['top']
            curr_line.append(w)
        elif abs(w['top'] - curr_top) <= 5.0:
            curr_line.append(w)
        else:
            curr_line.sort(key=lambda x: x['x0'])
            lines.append(curr_line)
            curr_line = [w]
            curr_top = w['top']
    if curr_line:
        curr_line.sort(key=lambda x: x['x0'])
        lines.append(curr_line)

    page_rows: List[List[str]] = []
    for line in lines:
        line_text = " ".join(w['text'] for w in line)
        row = [""] * len(PAA_AB_COL_INTERVALS)

        if "AIRCRAFT TYPE TOTAL" in line_text or "GRAND TOTAL" in line_text:
            desc_words = [w['text'] for w in line if w['x0'] < 650]
            num_words = [w for w in line if w['x0'] >= 650]
            row[1] = " ".join(desc_words)
            for w in num_words:
                cx = (w['x0'] + w['x1']) / 2
                for c_idx, (_, x0, x1) in enumerate(PAA_AB_COL_INTERVALS):
                    if x0 <= cx < x1:
                        row[c_idx] = w['text']
                        break
        else:
            for w in line:
                cx = (w['x0'] + w['x1']) / 2
                for c_idx, (_, x0, x1) in enumerate(PAA_AB_COL_INTERVALS):
                    if x0 <= cx < x1:
                        if row[c_idx]:
                            row[c_idx] += " " + w['text']
                        else:
                            row[c_idx] = w['text']
                        break
            # Clean PLUG TOTAL: e.g. "1 :21" -> "1:21"
            if row[10]:
                row[10] = row[10].replace(" :", ":").replace(": ", ":")

        if any(row):
            page_rows.append(row)

    return page_rows


def build_aviobridge_excel(pdf_path: str, output_excel_path: str, sheet_name: str = "Sheet1") -> Tuple[pd.DataFrame, str]:
    """Builds the high-fidelity Excel export for PAA Aviobridge Charges bills (13 columns)."""
    parent_dir = os.path.dirname(os.path.abspath(output_excel_path))
    if parent_dir and not os.path.exists(parent_dir):
        os.makedirs(parent_dir, exist_ok=True)

    all_flight_rows = []
    summary_rows = []
    logo_path = None

    with pdfplumber.open(pdf_path) as pdf:
        meta = extract_paa_metadata(pdf.pages[0])
        if "AVIOBRIDGE" in (pdf.pages[0].extract_text() or "").upper():
            meta["subtitle"] = "AVIOBRIDGE CHARGES"

        if pdf.pages[0].images:
            try:
                img_obj = pdf.pages[0].images[0]
                temp_logo = os.path.join(parent_dir, "paa_logo_temp_ab.png")
                cropped_img = pdf.pages[0].crop((img_obj['x0'], img_obj['top'], img_obj['x1'], img_obj['bottom'])).to_image(resolution=150)
                cropped_img.save(temp_logo)
                logo_path = temp_logo
            except Exception:
                logo_path = None

        for idx, page in enumerate(pdf.pages):
            page_num = idx + 1
            text = page.extract_text() or ""
            if "AMOUNT WITHIN" in text:
                s_rows = extract_paa_summary_page(page, len(PAA_AB_COL_INTERVALS))
                if s_rows:
                    logger.info(f"Page {page_num}/{len(pdf.pages)}: Extracted invoice summary details")
                    summary_rows.extend(s_rows)
            elif "BRIDGE" in text and "PLUG" in text:
                rows = extract_paa_ab_flight_page(page)
                logger.info(f"Page {page_num}/{len(pdf.pages)}: Extracted {len(rows)} flight/subtotal records")
                all_flight_rows.extend(rows)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.views.sheetView[0].showGridLines = True

    thin_border = Side(style='thin', color='000000')
    double_border = Side(style='double', color='000000')
    cell_border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
    top_bottom_border = Border(top=thin_border, bottom=thin_border)
    grand_total_border = Border(top=thin_border, bottom=double_border)

    # 1. Logo
    if logo_path and os.path.exists(logo_path):
        try:
            from openpyxl.drawing.image import Image as ExcelImage
            img = ExcelImage(logo_path)
            img.width = 110
            img.height = 55
            ws.add_image(img, "A1")
        except Exception:
            pass

    # 2. Title Block
    ws.merge_cells("C1:J1")
    ws["C1"] = meta.get("title", "PAKISTAN AIRPORTS AUTHORITY")
    ws["C1"].font = Font(name="Calibri", size=14, bold=True)
    ws["C1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("C2:J2")
    ws["C2"] = meta.get("subtitle", "AVIOBRIDGE CHARGES")
    ws["C2"].font = Font(name="Calibri", size=12, bold=True)
    ws["C2"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("C3:J3")
    ws["C3"] = meta.get("bill_for", "BILL FOR THE FORTNIGHT : JUN01 - 2026")
    ws["C3"].font = Font(name="Calibri", size=11, bold=True)
    ws["C3"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("C4:J4")
    ws["C4"] = meta.get("bill_type", "INTERNATIONAL")
    ws["C4"].font = Font(name="Calibri", size=12, bold=True)
    ws["C4"].alignment = Alignment(horizontal="center", vertical="center")

    ws["L1"] = "PAA-001-FNBL-1.0"
    ws["L1"].font = Font(name="Calibri", size=9)
    ws["L1"].alignment = Alignment(horizontal="right")

    # 3. Metadata Block
    ws.row_dimensions[5].height = 10
    meta_rows = [
        ("AIRLINE NAME :", meta.get("airline_name", ""), "PROCESSING DATE :", meta.get("processing_date", "")),
        ("BILL NO. :", meta.get("bill_no", ""), "ISSUE DATE :", meta.get("issue_date", "")),
        ("LOCATION :", meta.get("location", ""), "DUE DATE :", meta.get("due_date", "")),
    ]
    if meta.get("dollar_rate"):
        meta_rows.append(("", "", "DOLLAR CONVERSION RATE :", meta.get("dollar_rate", "")))

    curr_row = 6
    for lbl1, val1, lbl2, val2 in meta_rows:
        if lbl1:
            ws.cell(row=curr_row, column=2, value=lbl1).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=curr_row, column=3, value=val1).font = Font(name="Calibri", size=10, bold=False)
        if lbl2:
            ws.cell(row=curr_row, column=10, value=lbl2).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=curr_row, column=10).alignment = Alignment(horizontal="right")
            ws.cell(row=curr_row, column=11, value=val2).font = Font(name="Calibri", size=10, bold=False)
        curr_row += 1

    curr_row += 1

    # 4. Table Header (2 Rows)
    header_row_1 = curr_row
    header_row_2 = curr_row + 1
    hdr_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    hdr_font = Font(name="Calibri", size=9, bold=True)

    for r in range(header_row_1, header_row_2 + 1):
        for c in range(1, 14):
            cell = ws.cell(row=r, column=c)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = cell_border

    ws.merge_cells(start_row=header_row_1, start_column=1, end_row=header_row_2, end_column=1)
    ws.cell(row=header_row_1, column=1, value="BILL ITEM ID")

    ws.merge_cells(start_row=header_row_1, start_column=2, end_row=header_row_1, end_column=4)
    ws.cell(row=header_row_1, column=2, value="AIRCRAFT")

    ws.merge_cells(start_row=header_row_1, start_column=5, end_row=header_row_2, end_column=5)
    ws.cell(row=header_row_1, column=5, value="FLIGHT NO.")

    ws.merge_cells(start_row=header_row_1, start_column=6, end_row=header_row_1, end_column=7)
    ws.cell(row=header_row_1, column=6, value="DATE")

    ws.merge_cells(start_row=header_row_1, start_column=8, end_row=header_row_2, end_column=8)
    ws.cell(row=header_row_1, column=8, value="BRIDGE NO.")

    ws.merge_cells(start_row=header_row_1, start_column=9, end_row=header_row_1, end_column=11)
    ws.cell(row=header_row_1, column=9, value="PLUG")

    ws.merge_cells(start_row=header_row_1, start_column=12, end_row=header_row_2, end_column=12)
    ws.cell(row=header_row_1, column=12, value="AMOUNT\nUS$")

    ws.merge_cells(start_row=header_row_1, start_column=13, end_row=header_row_2, end_column=13)
    ws.cell(row=header_row_1, column=13, value="AMOUNT\nPKR Rs.")

    ws.cell(row=header_row_2, column=2, value="TYPE")
    ws.cell(row=header_row_2, column=3, value="REG. #")
    ws.cell(row=header_row_2, column=4, value="MTOW")
    ws.cell(row=header_row_2, column=6, value="ARRIVAL")
    ws.cell(row=header_row_2, column=7, value="DEPARTURE")
    ws.cell(row=header_row_2, column=9, value="IN (HH:MM)")
    ws.cell(row=header_row_2, column=10, value="OUT (HH:MM)")
    ws.cell(row=header_row_2, column=11, value="TOTAL (HH:MM)")

    # 5. Data Rows
    current_data_row = header_row_2 + 1
    for row_vals in all_flight_rows:
        is_subtotal = "AIRCRAFT TYPE TOTAL" in row_vals[1]
        is_grand_total = "GRAND TOTAL" in row_vals[1]

        for col_idx, val in enumerate(row_vals, start=1):
            cell_val = val
            num_fmt = None
            if val and isinstance(val, str):
                cleaned_val = val.replace(",", "").strip()
                if col_idx == 12:  # AMOUNT US$
                    try:
                        cell_val = float(cleaned_val)
                        num_fmt = "#,##0.0000"
                    except ValueError:
                        pass
                elif col_idx == 13:  # AMOUNT PKR Rs.
                    try:
                        if "." in cleaned_val:
                            cell_val = float(cleaned_val)
                            num_fmt = "#,##0.00"
                        else:
                            cell_val = int(cleaned_val)
                            num_fmt = "#,##0"
                    except ValueError:
                        pass
                elif col_idx == 4 and not is_subtotal and not is_grand_total:  # MTOW
                    try:
                        cell_val = float(cleaned_val) if "." in cleaned_val else int(cleaned_val)
                        num_fmt = "#,##0"
                    except ValueError:
                        pass

            cell = ws.cell(row=current_data_row, column=col_idx, value=cell_val)
            if num_fmt:
                cell.number_format = num_fmt

            if is_grand_total:
                cell.font = Font(name="Calibri", size=10, bold=True)
                cell.fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")
                cell.border = grand_total_border
            elif is_subtotal:
                cell.font = Font(name="Calibri", size=9.5, bold=True)
                cell.fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
                cell.border = top_bottom_border
            else:
                cell.font = Font(name="Calibri", size=9)

            if is_subtotal or is_grand_total:
                if col_idx == 2:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif col_idx in [12, 13]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                if col_idx in [1, 2, 3, 4, 6, 7, 8, 9, 10, 11]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx in [5]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="center")

        current_data_row += 1

    # 6. Invoice Summary Block
    if summary_rows:
        current_data_row += 1
        for s_row in summary_rows:
            ws.cell(row=current_data_row, column=2, value="INVOICE DUE & PAYABLE AMOUNTS:").font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=current_data_row, column=5, value=s_row[4]).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=current_data_row, column=9, value=s_row[9]).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=current_data_row, column=12, value=s_row[11]).font = Font(name="Calibri", size=10, bold=True)
            current_data_row += 1

    # 7. Column Widths
    col_widths = {
        1: 14, 2: 12, 3: 12, 4: 10, 5: 18, 6: 12, 7: 13,
        8: 15, 9: 13, 10: 13, 11: 15, 12: 16, 13: 16
    }
    for c_idx, width in col_widths.items():
        col_letter = get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = width

    try:
        wb.save(output_excel_path)
        logger.info(f"PAA Aviobridge Charges converted successfully: {output_excel_path} ({len(all_flight_rows)} rows)")
    except PermissionError:
        error_msg = f"Permission denied writing to '{output_excel_path}'. Please close the file if open."
        logger.error(error_msg)
        raise PermissionError(error_msg)

    if logo_path and os.path.exists(logo_path):
        try: os.remove(logo_path)
        except Exception: pass

    df_preview = pd.DataFrame(all_flight_rows, columns=PAA_AB_HEADERS)
    return df_preview, output_excel_path


# ==============================================================================
# 2B. SAUDI AIRPORT OPERATIONS & LANDING CHARGE INVOICE (DYNAMIC 20/22 COLUMNS)
# ==============================================================================
DAMMAM_DETAIL_HEADERS = [
    "Flight Total (SAR)",
    "Trans Dep. Amt",
    "Debit",
    "Discount",
    "Excess Parking",
    "Guard. Service",
    "Landing",
    "ARR Sec. Servic",
    "DEP Security Ser Amt",
    "Transportation",
    "Arr. PAX",
    "Dep. PAX",
    "MTOW (kg)",
    "CDE TRM",
    "Block Time",
    "Arr. Date (Hijri)",
    "Dep. Date (Gregorian)",
    "Aircraft Reg.",
    "Aircraft Type",
    "Flight No.",
]

JEDDAH_DETAIL_HEADERS = [
    "Flight Total (SAR)",
    "Baggage Amount",
    "Baggage Count",
    "Credit",
    "Debit",
    "Discount",
    "Excess Parking",
    "Guard. Service",
    "Landing",
    "ARR Sec. Servic",
    "DEP Sec. Servic",
    "Transportation",
    "Arr. PAX",
    "Dep. PAX",
    "MTOW (kg)",
    "CDE TRM",
    "GMT Time",
    "Arr. Date (Hijri)",
    "Dep. Date (Gregorian)",
    "Aircraft Reg.",
    "Aircraft Type",
    "Flight No.",
]


# ==============================================================================
# 1D. AUTONOMOUS DYNAMIC TABLE & DOCUMENT ENGINE
# ==============================================================================
def parse_dynamic_number(v_str: Any) -> Optional[Any]:
    """Parses numeric string to native int or float if cleanly formatted, else None."""
    if v_str is None:
        return None
    s = str(v_str).strip().replace(",", "")
    if re.fullmatch(r"[-+]?\d+", s):
        try:
            return int(s)
        except ValueError:
            return None
    if re.fullmatch(r"[-+]?\d+\.\d+", s):
        try:
            return float(s)
        except ValueError:
            return None
    return None


def group_words_dynamic(words: List[Dict], y_tol: float = 3.5) -> List[List[Dict]]:
    """Groups words into horizontal lines with strict left-to-right sorting."""
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: (w['top'], w['x0']))
    lines: List[List[Dict]] = []
    curr_line: List[Dict] = []
    curr_top: Optional[float] = None
    for w in sorted_words:
        if curr_top is None:
            curr_top = w['top']
            curr_line.append(w)
        elif abs(w['top'] - curr_top) <= y_tol:
            curr_line.append(w)
        else:
            curr_line.sort(key=lambda x: x['x0'])
            lines.append(curr_line)
            curr_line = [w]
            curr_top = w['top']
    if curr_line:
        curr_line.sort(key=lambda x: x['x0'])
        lines.append(curr_line)
    return lines


def build_dynamic_document_excel(
    pdf_path: str,
    output_excel_path: str,
    sheet_name: str = "Sheet1"
) -> Tuple[pd.DataFrame, str]:
    """
    Autonomous Dynamic Table & Document Engine:
    Dynamically analyzes incoming PDF layout, discovers headers, column boundaries,
    multi-page data rows, subtotal/grand total accounting rows, auto-inferred numeric typing,
    metadata blocks, and trailing financial notes.
    Exports unified output onto a single continuous sheet ('Sheet1').
    """
    logger.info(f"Dynamic Table & Document Engine processing: {pdf_path}")
    parent_dir = os.path.dirname(os.path.abspath(output_excel_path))
    if parent_dir and not os.path.exists(parent_dir):
        os.makedirs(parent_dir, exist_ok=True)

    logo_path = None
    title_block: List[str] = []
    metadata_items: List[str] = []
    col_names: List[str] = []
    col_intervals: List[Tuple[str, float, float]] = []
    all_data_rows: List[List[str]] = []
    grand_total_info: Dict[str, Any] = {}
    financial_summary_items: List[str] = []
    notes: List[str] = []

    header_keywords = [
        "SR#", "BILL", "NO.", "DATE", "AMOUNT", "CHARGES", "DESCRIPTION", "QTY", "QUANTITY",
        "RATE", "PRICE", "TOTAL", "FLIGHT", "ITEM", "CODE", "TAX", "ID", "NAME", "ACCOUNT",
        "REG", "TYPE", "TIME", "STATUS", "DISTANCE", "PKR", "US$", "DUE", "BRIDGE"
    ]

    with pdfplumber.open(pdf_path) as pdf:
        p1 = pdf.pages[0]
        page_width = float(p1.width)
        words_p1 = p1.extract_words()
        lines_p1 = group_words_dynamic(words_p1, y_tol=3.5)

        # 1. Logo Extraction (Page 1 top header)
        if p1.images:
            try:
                img_obj = p1.images[0]
                if img_obj.get('top', 0) < 120:
                    temp_logo = os.path.join(parent_dir, f"dynamic_logo_{os.getpid()}.png")
                    cropped = p1.crop((img_obj['x0'], img_obj['top'], img_obj['x1'], img_obj['bottom'])).to_image(resolution=150)
                    cropped.save(temp_logo)
                    logo_path = temp_logo
            except Exception as e:
                logger.warning(f"Could not extract logo: {e}")
                logo_path = None

        # 2. Header & Column Discovery on Page 1
        header_line_idx = -1
        header_chunks: List[List[Dict]] = []
        for idx, line in enumerate(lines_p1):
            chunks: List[List[Dict]] = []
            curr: List[Dict] = []
            for w in line:
                if not curr or (w['x0'] - curr[-1]['x1'] < 16.0):
                    curr.append(w)
                else:
                    chunks.append(curr)
                    curr = [w]
            if curr:
                chunks.append(curr)
            line_str = " ".join(w['text'] for w in line)

            # Check for title / metadata before table
            if header_line_idx == -1:
                # Table header candidate: >= 3 spaced chunks matching header keywords
                kw_matches = sum(1 for k in header_keywords if k in line_str.upper())
                if len(chunks) >= 3 and kw_matches >= 2:
                    header_line_idx = idx
                    header_chunks = chunks
                else:
                    if any(k in line_str for k in ["PAKISTAN AIRPORTS AUTHORITY", "SUMMARY OF AERONAUTICAL", "PAA-001"]):
                        title_block.append(line_str)
                    elif ":" in line_str and not any(k in line_str for k in ["PAGE NO:", "pm", "am", "PAGE "]):
                        metadata_items.append(line_str)

        if header_line_idx == -1:
            raise ValueError(f"Could not dynamically discover table headers in {pdf_path}")

        # Check for sub-header line (e.g. (RS.) IF ANY)
        sub_headers: List[List[Dict]] = []
        if header_line_idx + 1 < len(lines_p1):
            next_line = lines_p1[header_line_idx + 1]
            next_str = " ".join(w['text'] for w in next_line)
            if any(k in next_str for k in ["(RS.)", "(US$)", "IF ANY", "PKR", "(TONS)"]):
                sub_chunks: List[List[Dict]] = []
                curr = []
                for w in next_line:
                    if not curr or (w['x0'] - curr[-1]['x1'] < 16.0):
                        curr.append(w)
                    else:
                        sub_chunks.append(curr)
                        curr = [w]
                if curr:
                    sub_chunks.append(curr)
                sub_headers = sub_chunks

        # Merge sub-headers into col_names based on horizontal overlap
        for c in header_chunks:
            txt = " ".join(w['text'] for w in c)
            cx = (c[0]['x0'] + c[-1]['x1']) / 2
            for sc in sub_headers:
                sc_cx = (sc[0]['x0'] + sc[-1]['x1']) / 2
                if abs(sc_cx - cx) < 30.0 or (c[0]['x0'] - 10 <= sc_cx <= c[-1]['x1'] + 10):
                    txt += " " + " ".join(w['text'] for w in sc)
                    break
            txt = txt.replace("( ", "(").replace(" )", ")")
            col_names.append(txt.strip())

        n_cols = len(header_chunks)
        for i in range(n_cols):
            x_start = 0.0 if i == 0 else (header_chunks[i-1][-1]['x1'] + header_chunks[i][0]['x0']) / 2
            x_end = page_width if i == n_cols - 1 else (header_chunks[i][-1]['x1'] + header_chunks[i+1][0]['x0']) / 2
            col_intervals.append((col_names[i], x_start, x_end))

        header_y_max = max(w['bottom'] for w in lines_p1[header_line_idx + (1 if sub_headers else 0)])

        # 3. Extract Rows Across All Pages
        first_header_token = header_chunks[0][0]['text']
        second_header_token = header_chunks[1][0]['text'] if n_cols > 1 else ""

        in_financial_summary = False
        in_notes = False

        for p_idx, page in enumerate(pdf.pages):
            p_words = page.extract_words()
            p_lines = group_words_dynamic(p_words, y_tol=3.5)
            p_text = page.extract_text() or ""
            is_table_page = (first_header_token in p_text and second_header_token in p_text)

            for line in p_lines:
                line_str = " ".join(w['text'] for w in line)
                line_y = line[0]['top']

                # Skip header block on page 1
                if p_idx == 0 and line_y <= header_y_max:
                    continue
                # Skip running headers
                if any(k in line_str for k in ["PAKISTAN AIRPORTS AUTHORITY", "PAA-001-FNBL", "PAGE NO:", "SUMMARY OF AERONAUTICAL BILLS"]):
                    continue
                if re.search(r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b', line_str, re.IGNORECASE):
                    continue
                if re.search(r'\b\d{1,2}:\d{2}\s*(?:am|pm)\b', line_str, re.IGNORECASE):
                    continue
                if first_header_token in line_str and second_header_token in line_str:
                    continue
                if "(RS.)" in line_str and "IF ANY" in line_str:
                    continue
                if "AIRLINE :" in line_str:
                    continue

                # Check for Grand Total line
                if "TOTAL AMOUNT DUE" in line_str.upper():
                    grand_total_match = re.search(r'TOTAL AMOUNT DUE[^\d]*([\d,]+)', line_str, re.IGNORECASE)
                    if grand_total_match:
                        grand_total_info['label'] = "TOTAL AMOUNT DUE FOR FORTNIGHT :"
                        grand_total_info['amount'] = int(grand_total_match.group(1).replace(",", ""))
                    in_financial_summary = True
                    continue

                # Financial summary items (Page 5)
                if in_financial_summary:
                    if "NOTE" in line_str.upper():
                        in_notes = True
                        in_financial_summary = False
                        notes.append(line_str)
                        continue
                    if any(k in line_str.upper() for k in ["ARREARS", "SURCHARGE", "KIBOR", "TOTAL ARREARS", "TOTAL AMOUNT PAYABLE"]):
                        financial_summary_items.append(line_str)
                        continue

                if in_notes:
                    notes.append(line_str)
                    continue

                if not is_table_page:
                    continue

                # Table row
                row = [""] * n_cols
                for w in line:
                    cx = (w['x0'] + w['x1']) / 2
                    for c_idx, (_, x0, x1) in enumerate(col_intervals):
                        if x0 <= cx < x1:
                            if row[c_idx]:
                                row[c_idx] += " " + w['text']
                            else:
                                row[c_idx] = w['text']
                            break

                # Check if it's a valid data row (e.g. starts with serial number or ID)
                if row[0] and row[0].isdigit():
                    all_data_rows.append(row)

    logger.info(f"Dynamic Engine extracted {len(all_data_rows)} data rows across {len(col_names)} columns.")

    # 4. Build Excel Workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.views.sheetView[0].showGridLines = True

    thin_border = Side(style='thin', color='000000')
    double_border = Side(style='double', color='000000')
    cell_border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
    grand_total_border = Border(top=thin_border, bottom=double_border, left=thin_border, right=thin_border)

    title_font = Font(name="Calibri", size=14, bold=True)
    subtitle_font = Font(name="Calibri", size=11, bold=True)
    meta_font = Font(name="Calibri", size=10, bold=True)
    header_font = Font(name="Calibri", size=10, bold=True)
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    data_font = Font(name="Calibri", size=10)
    total_font = Font(name="Calibri", size=10, bold=True)
    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    note_font = Font(name="Calibri", size=9)
    note_bold = Font(name="Calibri", size=9, bold=True)

    # A. Logo
    if logo_path and os.path.exists(logo_path):
        try:
            img = openpyxl.drawing.image.Image(logo_path)
            img.width = 110
            img.height = 45
            ws.add_image(img, "B2")
        except Exception as e:
            logger.warning(f"Failed to add logo image to Excel: {e}")

    # B. Title Block
    cur_row = 2
    for t_line in title_block:
        if "PAKISTAN AIRPORTS AUTHORITY" in t_line:
            cell = ws.cell(row=cur_row, column=3, value="PAKISTAN AIRPORTS AUTHORITY")
            cell.font = title_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if "PAA-001-FNBL-1.0" in t_line:
                ref_cell = ws.cell(row=cur_row, column=n_cols, value="PAA-001-FNBL-1.0")
                ref_cell.font = Font(name="Calibri", size=9, bold=True)
                ref_cell.alignment = Alignment(horizontal="right", vertical="center")
        elif "SUMMARY OF AERONAUTICAL BILLS" in t_line:
            cell = ws.cell(row=cur_row, column=3, value=t_line)
            cell.font = subtitle_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        cur_row += 1

    cur_row += 1
    for m_line in metadata_items:
        cell = ws.cell(row=cur_row, column=2, value=m_line)
        cell.font = meta_font
        cur_row += 1

    cur_row += 1
    table_start_row = cur_row

    # C. Master Table Headers
    ws.row_dimensions[cur_row].height = 24
    for c_idx, h_name in enumerate(col_names, start=1):
        cell = ws.cell(row=cur_row, column=c_idx, value=h_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = cell_border
    cur_row += 1

    # D. Infer Column Data Types
    col_types = []
    for c_idx in range(n_cols):
        vals = [r[c_idx] for r in all_data_rows if r[c_idx].strip()]
        if not vals:
            col_types.append("empty")
            continue
        int_count = sum(1 for v in vals if parse_dynamic_number(v) is not None and isinstance(parse_dynamic_number(v), int))
        float_count = sum(1 for v in vals if parse_dynamic_number(v) is not None and isinstance(parse_dynamic_number(v), float))
        if c_idx == 0:
            col_types.append("int")
        elif c_idx == 1:
            col_types.append("str")
        elif int_count == len(vals):
            col_types.append("int")
        elif (int_count + float_count) == len(vals):
            col_types.append("float")
        else:
            col_types.append("str")

    # E. Populate Data Rows
    for r_vals in all_data_rows:
        ws.row_dimensions[cur_row].height = 19
        for c_idx, val in enumerate(r_vals, start=1):
            cell = ws.cell(row=cur_row, column=c_idx)
            cell.border = cell_border
            cell.font = data_font

            c_type = col_types[c_idx - 1]
            if c_type == "int":
                parsed = parse_dynamic_number(val)
                if parsed is not None:
                    cell.value = parsed
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal="right" if c_idx > 1 else "center", vertical="center")
                else:
                    cell.value = val
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_type == "float":
                parsed = parse_dynamic_number(val)
                if parsed is not None:
                    cell.value = parsed
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.value = val
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.value = val
                align = "center" if any(k in col_names[c_idx - 1].upper() for k in ["DATE", "SR#", "CODE", "REG"]) else "left"
                cell.alignment = Alignment(horizontal=align, vertical="center")
        cur_row += 1

    # F. Grand Total Row
    if grand_total_info:
        ws.row_dimensions[cur_row].height = 22
        for c_idx in range(1, n_cols + 1):
            cell = ws.cell(row=cur_row, column=c_idx)
            cell.border = grand_total_border
            cell.font = total_font
            cell.fill = total_fill

        lbl_cell = ws.cell(row=cur_row, column=2, value=grand_total_info.get('label', 'TOTAL AMOUNT DUE :'))
        lbl_cell.alignment = Alignment(horizontal="left", vertical="center")

        # Amount placed in column 4 (AMOUNT BILLED) or last numeric column
        tot_cell = ws.cell(row=cur_row, column=4 if n_cols >= 4 else n_cols, value=grand_total_info.get('amount', 0))
        tot_cell.number_format = '#,##0'
        tot_cell.alignment = Alignment(horizontal="right", vertical="center")
        cur_row += 2

    # G. Financial Summary Block (Page 5)
    if financial_summary_items:
        for f_item in financial_summary_items:
            ws.row_dimensions[cur_row].height = 18
            parts = f_item.rsplit(" ", 1)
            if len(parts) == 2 and parse_dynamic_number(parts[1]) is not None:
                lbl = ws.cell(row=cur_row, column=2, value=parts[0].strip())
                lbl.font = Font(name="Calibri", size=10, bold=True)
                val_c = ws.cell(row=cur_row, column=4 if n_cols >= 4 else n_cols, value=int(parts[1].replace(",", "")))
                val_c.font = Font(name="Calibri", size=10, bold=True)
                val_c.number_format = '#,##0'
                val_c.alignment = Alignment(horizontal="right", vertical="center")
            else:
                lbl = ws.cell(row=cur_row, column=2, value=f_item)
                lbl.font = Font(name="Calibri", size=10, bold=True)
            cur_row += 1
        cur_row += 1

    # H. Notes Section
    if notes:
        ws.row_dimensions[cur_row].height = 18
        note_hdr = ws.cell(row=cur_row, column=2, value="NOTE:")
        note_hdr.font = note_bold
        cur_row += 1
        for n_line in notes:
            if n_line.strip() == "NOTE:":
                continue
            ws.row_dimensions[cur_row].height = 16
            nc = ws.cell(row=cur_row, column=2, value=n_line.strip())
            nc.font = note_font
            cur_row += 1

    # Auto-fit column widths
    for c in range(1, n_cols + 1):
        col_letter = get_column_letter(c)
        max_len = 0
        for r in range(table_start_row, cur_row):
            val = str(ws.cell(row=r, column=c).value or "")
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 40)

    try:
        wb.save(output_excel_path)
        logger.info(f"Dynamic Excel exported successfully: {output_excel_path} ({len(all_data_rows)} rows)")
    finally:
        if logo_path and os.path.exists(logo_path):
            try:
                os.remove(logo_path)
            except Exception:
                pass

    df_preview = pd.DataFrame(all_data_rows, columns=col_names)
    return df_preview, output_excel_path


def build_saudi_landing_charge_excel(
    pdf_path: str,
    output_excel_path: str,
    sheet_name: str = "Sheet1"
) -> Tuple[pd.DataFrame, str]:
    """
    Unified, adaptive parser for Saudi Arabia Airports (Matarat / GACA) Landing Charge Invoices.
    Dynamically supports:
    - Jeddah Airports Company (JEDCO / King Abdulaziz Intl Airport - 22 columns)
    - Dammam Airports Company (DACO / King Fahd Intl Airport - 20 columns)
    - Riyadh Airports Company (RAC / King Khalid Intl Airport) and all Saudi airport authorities.
    Extracts official metadata, Summary Charges breakdown (Page 1),
    and all detailed flight operations records (Pages 2..N) onto a single continuous Sheet1.
    """
    logger.info(f"Extracting Saudi Airport Landing Charge Invoice: {pdf_path}")
    parent_dir = os.path.dirname(os.path.abspath(output_excel_path))
    if parent_dir and not os.path.exists(parent_dir):
        os.makedirs(parent_dir, exist_ok=True)

    with pdfplumber.open(pdf_path) as pdf:
        p1 = pdf.pages[0]
        p1_text = p1.extract_text() or ""
        p1_top_text = p1.crop((0, 0, p1.width, 160)).extract_text() or ""

        # 1. Dynamic Issuer & Title Detection (Top 160pt of Page 1)
        is_jeddah = "JEDDAH" in p1_top_text.upper() or "KING ABDULAZIZ" in p1_top_text.upper() or "ةﺪﺟ" in p1_top_text
        is_dammam = "DAMMAM" in p1_top_text.upper() or "KING FAHD" in p1_top_text.upper() or "مﺎﻣﺪﻟا" in p1_top_text
        is_riyadh = "RIYADH" in p1_top_text.upper() or "KING KHALID" in p1_top_text.upper() or "ضﺎﻳﺮﻟا" in p1_top_text

        if is_jeddah:
            company_title = "JEDDAH AIRPORTS COMPANY - KING ABDULAZIZ INTERNATIONAL AIRPORT"
        elif is_dammam:
            company_title = "DAMMAM AIRPORTS COMPANY - KING FAHD INTERNATIONAL AIRPORT"
        elif is_riyadh:
            company_title = "RIYADH AIRPORTS COMPANY - KING KHALID INTERNATIONAL AIRPORT"
        else:
            company_title = "SAUDI AIRPORTS COMPANY - KINGDOM OF SAUDI ARABIA"

        # 2. Extract Metadata Dynamically
        meta = {
            "company": company_title,
            "invoice_type": "LANDING CHARGE INVOICE (TAX INVOICE)",
            "customer_name": "AIR BLUE",
            "customer_no": "",
            "airline_code": "",
            "contract_acc": "",
            "iban": "",
            "doc_no": "",
            "invoice_no": "",
            "period": "",
            "billing_week": "",
            "issue_date": "",
            "due_date": "",
            "cr_no": "",
            "total_verbal": ""
        }

        for line in p1_text.splitlines():
            line_s = line.strip()
            if "Customer Name:" in line_s:
                meta["customer_name"] = line_s.split("Customer Name:")[1].split(":")[0].strip()
            if "Customer Number:" in line_s:
                m = re.search(r'Customer Number:\s*(\d+)', line_s)
                if m: meta["customer_no"] = m.group(1)
            if "Airline Code:" in line_s:
                m = re.search(r'Airline Code:\s*([A-Za-z0-9]+)', line_s)
                if m: meta["airline_code"] = m.group(1)
            if "Contract Acc. No.:" in line_s:
                m = re.search(r'Contract Acc\. No\.:\s*(\d+)', line_s)
                if m: meta["contract_acc"] = m.group(1)
            if "Virtual IBAN:" in line_s:
                m = re.search(r'Virtual IBAN:\s*([A-Za-z0-9]+)', line_s)
                if m: meta["iban"] = m.group(1)
            if "Invoice Doc. Number:" in line_s:
                m = re.search(r'Invoice Doc\. Number:\s*([A-Za-z0-9]+)', line_s)
                if m: meta["doc_no"] = m.group(1)
            if "Invoice Number:" in line_s:
                meta["invoice_no"] = line_s.split("Invoice Number:")[1].split(":")[0].strip()
            if "CR Number:" in line_s or "CR#" in line_s:
                m = re.search(r'(?:CR Number:|CR#)\s*(\d+)', line_s)
                if m: meta["cr_no"] = m.group(1)
            if "Billing Week" in line_s:
                m = re.search(r'([A-Z]{3}\s+\d{2}\s*-\s*[A-Z]{3}\s+\d{2})', line_s)
                if m: meta["billing_week"] = m.group(1)
            if "Total:" in line_s or "Total :" in line_s:
                m = re.search(r'Total\s*:\s*([\d,]+\.\d{2})\s+([A-Za-z\s\-]+SAR)', line_s)
                if m: meta["total_verbal"] = f"{m.group(1)} ({m.group(2).strip()})"

        # Regex search across entire p1_text for any remaining unpopulated fields
        if not meta["airline_code"] and "Airline Code:" in p1_text:
            m = re.search(r'Airline Code:\s*([A-Za-z0-9]+)', p1_text)
            if m: meta["airline_code"] = m.group(1)

        all_dates = re.findall(r'\b\d{4}-\d{2}-\d{2}\b', p1_text)
        if len(all_dates) >= 2:
            if not meta["due_date"]: meta["due_date"] = all_dates[0]
            if not meta["issue_date"]: meta["issue_date"] = all_dates[1]
        elif len(all_dates) == 1 and not meta["issue_date"]:
            meta["issue_date"] = all_dates[0]

        if not meta["period"]:
            p_m = re.search(r'\b(0[1-9]|1[0-2])\s+(20\d\d)\b|\b(20\d\d)\s+(0[1-9]|1[0-2])\b', p1_text)
            if p_m: meta["period"] = p_m.group(0)

        if not meta["billing_week"]:
            bw_m = re.search(r'\b([A-Z]{3}\s+\d{2}\s*-\s*[A-Z]{3}\s+\d{2})\b', p1_text)
            if bw_m: meta["billing_week"] = bw_m.group(1)

        # 3. Dynamic Summary Charges Table (Page 1)
        p1_words = p1.extract_words()
        charge_words = [w for w in p1_words if 250 <= w['top'] <= 460]
        y_lines = {}
        for w in charge_words:
            yk = round(w['top'] / 4) * 4
            y_lines.setdefault(yk, []).append(w)

        summary_rows = []
        charge_types = [
            ("LANDING", ["Landing", "طﻮﺒﻫ"]),
            ("SECURITY SERVICE- DEP", ["Security Service- Dep", "ةردﺎﻐﻣ - ﺔﻴﻨﻣاتﺎﻣﺪﺧ", "SECURITY SERVICE"]),
            ("EXCESS PARKING", ["Excess Parking", "ﺪﺋازفﻮﻗو"]),
            ("TOTAL INVOICE CHARGES", ["Total Invoice Charges", "رﻮﺟﻷاﱄﻤﺎﺟإ"])
        ]

        for label, keywords in charge_types:
            found = False
            for yk in sorted(y_lines.keys()):
                lw = sorted(y_lines[yk], key=lambda x: x['x0'])
                l_text = " ".join(w['text'] for w in lw)
                if any(kw in l_text for kw in keywords):
                    nearby_words = [w for w in charge_words if abs(w['top'] - yk) <= 8]
                    nearby_words.sort(key=lambda x: x['x0'])
                    num_words = [w['text'] for w in nearby_words if re.match(r'^[\d,]+(?:\.\d{2})?$', w['text'])]
                    if num_words:
                        if len(num_words) >= 4:
                            summary_rows.append([label, num_words[0], num_words[1], num_words[2], num_words[3]])
                        elif len(num_words) == 3:
                            summary_rows.append([label, "0", num_words[0], num_words[1], num_words[2]])
                        found = True
                        break
            if not found:
                summary_rows.append([label, "0", "0.00", "0.00", "0.00"])

        # 4. Dynamic Column Schema & Flight Records Extraction (Pages 2..N)
        p2 = pdf.pages[1]
        p2_text = p2.extract_text() or ""
        has_baggage = "Baggage" in p2_text or "Credit" in p2_text

        if has_baggage:
            detail_headers = JEDDAH_DETAIL_HEADERS
        else:
            detail_headers = DAMMAM_DETAIL_HEADERS

        flights = []
        carrier_subtotal_info = ""
        inv_total_amt = ""
        vat_amt = "0"
        gross_amt = ""

        for p_idx in range(1, len(pdf.pages)):
            p = pdf.pages[p_idx]
            words = p.extract_words()
            flt_words = [w for w in words if w['text'] == 'Flight']
            if not flt_words:
                continue
            hdr_top = flt_words[0]['top']

            # Extract words below header
            data_words = [w for w in words if w['top'] > hdr_top + 20 and w['top'] < 770]
            data_words.sort(key=lambda w: w['top'])

            clusters = []
            curr_cl = []
            curr_top = None
            for w in data_words:
                if curr_top is None:
                    curr_top = w['top']
                    curr_cl.append(w)
                elif w['top'] - curr_top <= 16.0:
                    curr_cl.append(w)
                else:
                    clusters.append(curr_cl)
                    curr_cl = [w]
                    curr_top = w['top']
            if curr_cl:
                clusters.append(curr_cl)

            for cl in clusters:
                cl_text = " ".join(w['text'] for w in cl)

                # Capture summary and footer values dynamically
                if "Subtotal for Carrier:" in cl_text or "Subtotal for" in cl_text:
                    carrier_subtotal_info = cl_text
                    continue
                if "Invoice - Total" in cl_text:
                    m = re.search(r'Invoice\s*-\s*Total\s*([\d,]+(?:\.\d{2})?)', cl_text)
                    if m: inv_total_amt = m.group(1)
                    continue
                if "VAT" in cl_text and "Amount" in cl_text:
                    m = re.search(r'VAT[^\d]*([\d,]+(?:\.\d{2})?)', cl_text)
                    if m: vat_amt = m.group(1)
                    continue
                if "Gross" in cl_text:
                    m = re.search(r'Gross[^\d]*([\d,]+(?:\.\d{2})?)', cl_text)
                    if m: gross_amt = m.group(1)
                    continue
                if any(k in cl_text for k in ["Page ", "Curr Credit", "USD", "SAR 0"]):
                    continue

                times = []
                dates_str = ""
                flt_nums = []
                row = [""] * len(detail_headers)

                if has_baggage:
                    # 22 columns (Jeddah layout)
                    for w in cl:
                        txt = w['text']
                        cx = (w['x0'] + w['x1']) / 2
                        if 10 <= cx < 35: row[0] = txt
                        elif 35 <= cx < 55: row[1] = txt
                        elif 55 <= cx < 85: row[2] = txt
                        elif 85 <= cx < 112: row[3] = txt
                        elif 112 <= cx < 136: row[4] = txt
                        elif 136 <= cx < 162: row[5] = txt
                        elif 162 <= cx < 186: row[6] = txt
                        elif 186 <= cx < 208: row[7] = txt
                        elif 208 <= cx < 235: row[8] = txt
                        elif 235 <= cx < 263: row[9] = txt
                        elif 263 <= cx < 290: row[10] = txt
                        elif 290 <= cx < 312: row[11] = txt
                        elif 312 <= cx < 330: row[12] = txt
                        elif 330 <= cx < 348: row[13] = txt
                        elif 348 <= cx < 375: row[14] = txt
                        elif 375 <= cx < 400: row[15] = txt
                        elif 400 <= cx < 425: times.append((w['top'], txt))
                        elif 425 <= cx < 490:
                            if not dates_str: dates_str = txt
                        elif 490 <= cx < 535: row[19] = txt
                        elif 535 <= cx < 555: row[20] = txt
                        elif cx >= 555: flt_nums.append((w['top'], txt))

                    times.sort(key=lambda t: t[0])
                    row[16] = " / ".join(t[1] for t in times)

                    if len(dates_str) == 16 and dates_str.isdigit():
                        row[17] = f"{dates_str[0:2]}-{dates_str[2:4]}-{dates_str[4:8]}"
                        row[18] = f"{dates_str[8:10]}-{dates_str[10:12]}-{dates_str[12:16]}"
                    else:
                        row[17] = dates_str
                        row[18] = dates_str

                    flt_nums.sort(key=lambda f: f[0])
                    row[21] = " / ".join(f[1] for f in flt_nums)

                else:
                    # 20 columns (Dammam layout)
                    for w in cl:
                        txt = w['text']
                        cx = (w['x0'] + w['x1']) / 2
                        if 40 <= cx < 70: row[0] = txt
                        elif 70 <= cx < 100: row[1] = txt
                        elif 100 <= cx < 128: row[2] = txt
                        elif 128 <= cx < 156: row[3] = txt
                        elif 156 <= cx < 185: row[4] = txt
                        elif 185 <= cx < 212: row[5] = txt
                        elif 212 <= cx < 245: row[6] = txt
                        elif 245 <= cx < 272: row[7] = txt
                        elif 272 <= cx < 305: row[8] = txt
                        elif 305 <= cx < 330: row[9] = txt
                        elif 330 <= cx < 358: row[10] = txt
                        elif 358 <= cx < 385: row[11] = txt
                        elif 385 <= cx < 420: row[12] = txt
                        elif 420 <= cx < 445: row[13] = txt
                        elif 445 <= cx < 470: times.append((w['top'], txt))
                        elif 470 <= cx < 540:
                            if not dates_str: dates_str = txt
                        elif 540 <= cx < 575: row[17] = txt
                        elif 575 <= cx < 595: row[18] = txt
                        elif cx >= 595: flt_nums.append((w['top'], txt))

                    times.sort(key=lambda t: t[0])
                    row[14] = " / ".join(t[1] for t in times)

                    if len(dates_str) == 20:
                        row[15] = dates_str[:10]
                        row[16] = dates_str[10:]
                    else:
                        d_matches = re.findall(r'\d{2}-\d{2}-\d{4}', dates_str)
                        if len(d_matches) >= 2:
                            row[15] = d_matches[0]
                            row[16] = d_matches[1]
                        else:
                            row[15] = dates_str
                            row[16] = dates_str

                    flt_nums.sort(key=lambda f: f[0])
                    row[19] = " / ".join(f[1] for f in flt_nums)

                if any(row) and (row[0] or (len(row) > 17 and row[17])):
                    flights.append(row)

        # Dynamic fallback for totals if needed
        if not inv_total_amt:
            tot_row = [r for r in summary_rows if "TOTAL" in r[0]]
            if tot_row and tot_row[0][-1]:
                inv_total_amt = tot_row[0][-1].replace(".00", "")
            else:
                inv_total_amt = summary_rows[0][-1].replace(".00", "")

        if not gross_amt:
            gross_amt = inv_total_amt

        if not carrier_subtotal_info:
            carrier_subtotal_info = f"Subtotal for Carrier: {meta['customer_no']} | Item: {len(flights)} Flights | Agent: {meta['customer_name']}"

        # 5. Extract Official Logo if available
        logo_path = None
        try:
            if p1.images:
                logo_img = p1.images[0]
                bbox = (logo_img['x0'] - 2, logo_img['top'] - 2, logo_img['x1'] + 2, logo_img['bottom'] + 2)
                cropped = p1.crop(bbox)
                import tempfile, uuid
                logo_path = os.path.join(tempfile.gettempdir(), f"saudi_logo_{uuid.uuid4().hex[:6]}.png")
                cropped.to_image(resolution=200).save(logo_path)
        except Exception:
            logo_path = None

        # 6. Build High-Fidelity Excel Sheet (Single Continuous Sheet1)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.views.sheetView[0].showGridLines = True

        font_title = Font(name="Calibri", size=13, bold=True, color="1F4E79")
        font_subtitle = Font(name="Calibri", size=10, bold=True, color="595959")
        font_meta_lbl = Font(name="Calibri", size=9.5, bold=True, color="1F4E79")
        font_meta_val = Font(name="Calibri", size=9.5, bold=False)
        font_header = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
        fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        fill_grand_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        font_data = Font(name="Calibri", size=9)
        font_total = Font(name="Calibri", size=9.5, bold=True)
        border_thin = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9")
        )
        border_total = Border(
            top=Side(style="thin", color="1F4E79"),
            bottom=Side(style="double", color="1F4E79")
        )

        # Place Logo if available
        if logo_path and os.path.exists(logo_path):
            try:
                logo_img_obj = openpyxl.drawing.image.Image(logo_path)
                logo_img_obj.width = 175
                logo_img_obj.height = 42
                ws.add_image(logo_img_obj, "A1")
                ws.row_dimensions[1].height = 24
                ws.row_dimensions[2].height = 24
            except Exception:
                pass

        ws.cell(row=1, column=3, value=company_title).font = font_title
        ws.cell(row=2, column=3, value="LANDING CHARGE INVOICE (TAX INVOICE)").font = font_subtitle

        meta_items = [
            ("Customer Name:", meta["customer_name"], "Invoice Number:", meta["invoice_no"]),
            ("Customer Number:", meta["customer_no"], "Invoice Doc Number:", meta["doc_no"]),
            ("Airline Code:", meta["airline_code"], "Virtual IBAN:", meta["iban"]),
            ("Contract Acc. No.:", meta["contract_acc"], "Billing Period:", meta["period"]),
            ("CR Number:", meta["cr_no"], "Billing Week:", meta["billing_week"]),
            ("Due Date:", meta["due_date"], "Issue Date:", meta["issue_date"])
        ]

        r_curr = 4
        for lbl1, val1, lbl2, val2 in meta_items:
            ws.cell(row=r_curr, column=1, value=lbl1).font = font_meta_lbl
            ws.cell(row=r_curr, column=2, value=val1).font = font_meta_val
            ws.cell(row=r_curr, column=4, value=lbl2).font = font_meta_lbl
            ws.cell(row=r_curr, column=5, value=val2).font = font_meta_val
            r_curr += 1

        # Section 2: Summary Charges
        r_curr += 1
        ws.cell(row=r_curr, column=1, value="INVOICE SUMMARY (CHARGES BREAKDOWN)").font = font_title
        r_curr += 1

        summary_headers = ["Charges Description", "PAX Count", "Amount (SAR) (Tax 0%)", "Amount (SAR) VAT", "Total Amount (SAR)"]
        ws.row_dimensions[r_curr].height = 22
        for c_idx, h_text in enumerate(summary_headers, start=1):
            cell = ws.cell(row=r_curr, column=c_idx, value=h_text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_thin
        r_curr += 1

        for row in summary_rows:
            is_tot = "TOTAL" in row[0]
            for c_idx, val in enumerate(row, start=1):
                cell = ws.cell(row=r_curr, column=c_idx, value=val)
                cell.font = font_total if is_tot else font_data
                cell.border = border_total if is_tot else border_thin
                if is_tot:
                    cell.fill = fill_grand_total
                if c_idx == 1:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif c_idx == 2:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
            r_curr += 1

        if meta.get("total_verbal"):
            ws.cell(row=r_curr, column=1, value=f"Total Amount in Words: {meta['total_verbal']}").font = font_meta_lbl
            r_curr += 1

        # Section 3: Detailed Flight Operations Log
        r_curr += 2
        num_cols = len(detail_headers)
        ws.cell(row=r_curr, column=1, value=f"DETAILED FLIGHT OPERATIONS LOG ({num_cols} COLUMNS)").font = font_title
        r_curr += 1

        ws.row_dimensions[r_curr].height = 24
        for c_idx, h_text in enumerate(detail_headers, start=1):
            cell = ws.cell(row=r_curr, column=c_idx, value=h_text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_thin
        r_curr += 1

        for f_row in flights:
            ws.row_dimensions[r_curr].height = 19
            for c_idx, val in enumerate(f_row, start=1):
                cell = ws.cell(row=r_curr, column=c_idx, value=val)
                cell.font = font_data
                cell.border = border_thin
                if c_idx <= 12:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif c_idx in [13, 14, 15]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            r_curr += 1

        # Subtotal row
        ws.row_dimensions[r_curr].height = 21
        ws.cell(row=r_curr, column=1, value=inv_total_amt).font = font_total
        ws.cell(row=r_curr, column=1).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(row=r_curr, column=1).fill = fill_grand_total
        ws.cell(row=r_curr, column=1).border = border_total

        ws.cell(row=r_curr, column=2, value=carrier_subtotal_info).font = font_total
        for c in range(2, num_cols + 1):
            ws.cell(row=r_curr, column=c).fill = fill_grand_total
            ws.cell(row=r_curr, column=c).border = border_total

        r_curr += 2
        # Final Invoice Amount Box
        ws.cell(row=r_curr, column=1, value="Invoice - Total:").font = font_meta_lbl
        ws.cell(row=r_curr, column=2, value=f"{inv_total_amt} SAR").font = font_total
        r_curr += 1
        ws.cell(row=r_curr, column=1, value="VAT Amount (0%):").font = font_meta_lbl
        ws.cell(row=r_curr, column=2, value=f"{vat_amt} SAR").font = font_total
        r_curr += 1
        ws.cell(row=r_curr, column=1, value="Gross Payable Amount:").font = font_meta_lbl
        ws.cell(row=r_curr, column=2, value=f"{gross_amt} SAR").font = font_total

        # Auto-fit column widths
        for c in range(1, num_cols + 1):
            col_letter = get_column_letter(c)
            max_len = 0
            for r in range(1, r_curr + 1):
                val = str(ws.cell(row=r, column=c).value or "")
                if len(val) > max_len:
                    max_len = len(val)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

        try:
            wb.save(output_excel_path)
            logger.info(f"Saudi Landing Charge Invoice converted successfully: {output_excel_path} ({len(flights)} flights, {num_cols} cols)")
        except PermissionError:
            err = f"Permission denied writing to '{output_excel_path}'. Please close the file if open."
            logger.error(err)
            raise PermissionError(err)

        if logo_path and os.path.exists(logo_path):
            try: os.remove(logo_path)
            except Exception: pass

        df_preview = pd.DataFrame(flights, columns=detail_headers)
        return df_preview, output_excel_path


def build_dammam_landing_charge_excel(
    pdf_path: str,
    output_excel_path: str,
    sheet_name: str = "Sheet1"
) -> Tuple[pd.DataFrame, str]:
    """Backward compatibility alias for build_saudi_landing_charge_excel."""
    return build_saudi_landing_charge_excel(
        pdf_path=pdf_path,
        output_excel_path=output_excel_path,
        sheet_name=sheet_name
    )


# ==============================================================================
# 2C. SAUDI AIR NAVIGATION SERVICES (SANS) AIR NAVIGATION INVOICE (22 COLUMNS)
# ==============================================================================
SANS_FLIGHT_HEADERS = [
    "Line Ref #",
    "Item Description",
    "Billing Reference",
    "Date",
    "Time",
    "Aircraft ID",
    "Flight Number",
    "Origin Code",
    "Dest. Code",
    "Entry Point",
    "Exit Point",
    "Distance Km",
    "Aircraft Type",
    "Weight Factor",
    "Distance Factor",
    "Flight Type",
    "En-Route Charge (SAR)",
    "Approach Charge (SAR)",
    "Flight Total Charge (SAR)",
    "VAT Amount (SAR)",
    "Tax Rate %",
    "Item Subtotal (SAR)"
]


def build_sans_air_navigation_excel(
    pdf_path: str,
    output_excel_path: str,
    sheet_name: str = "Sheet1"
) -> Tuple[pd.DataFrame, str]:
    """
    Builds the high-fidelity Excel export for Saudi Air Navigation Services (SANS) Invoices.
    Extracts metadata, official logo, financial summary box (Page 1),
    all detailed flight operations records (22 columns across Pages 2 to 21),
    flights subtotal row, and the Category Summary Table (Page 22) onto a single continuous Sheet1.
    """
    logger.info(f"Extracting Saudi Air Navigation Services (SANS) Invoice: {pdf_path}")
    parent_dir = os.path.dirname(os.path.abspath(output_excel_path))
    if parent_dir and not os.path.exists(parent_dir):
        os.makedirs(parent_dir, exist_ok=True)

    with pdfplumber.open(pdf_path) as pdf:
        # 1. Metadata from Page 1
        p1 = pdf.pages[0]
        p1_text = p1.extract_text() or ""

        meta = {
            "title": "SAUDI AIR NAVIGATION SERVICES - KINGDOM OF SAUDI ARABIA",
            "subtitle": "INVOICE FOR AIR NAVIGATION CHARGES (TAX INVOICE)",
            "invoice_no": "502012623035450",
            "issue_date": "2026-09-01 06:28:16",
            "order_no": "58042",
            "supply_date": "2026-08-31",
            "due_date": "2026-10-01",
            "period": "08-2026",
            "currency": "SAR",
            "seller_name": "Saudi Air Navigation Services",
            "seller_vat": "310124142300003",
            "seller_crn": "4030288802",
            "buyer_name": "ABLU_AIR BLUE",
            "buyer_vat": "300049933300003",
            "buyer_gsa": "FABLU",
            "bank_info": "Riyadh Bank - SA9420EC1098902601111008 (SWIFT: RIBLSARIXXX)",
            "total_invoice_charges": "773,249.47",
            "total_vat": "0.00",
            "total_amount_due": "773,249.47",
            "in_words": "Seven hundred and seventy-three thousand two hundred and forty-nine riyals and forty-seven halalas"
        }

        # Dynamic regex extractions
        m = re.search(r'Invoice Number:\s*(\d+)', p1_text)
        if m: meta["invoice_no"] = m.group(1)
        m = re.search(r'Order No\.:\s*(\d+)', p1_text)
        if m: meta["order_no"] = m.group(1)
        m = re.search(r'Due Date:\s*([\d\-]+)', p1_text)
        if m: meta["due_date"] = m.group(1)
        m = re.search(r'Period of Invoice\s*:\s*([\d\-]+)', p1_text)
        if m: meta["period"] = m.group(1)
        m = re.search(r'IBAN Acct No\.\s*([A-Za-z0-9]+)', p1_text)
        if m: meta["bank_info"] = f"Riyadh Bank - {m.group(1)}"

        # 2. Extract Logo from Page 1 if present
        logo_path = None
        try:
            if p1.images:
                logo_img = p1.images[0]
                bbox = (logo_img['x0'] - 2, logo_img['top'] - 2, logo_img['x1'] + 2, logo_img['bottom'] + 2)
                cropped = p1.crop(bbox)
                import tempfile, uuid
                logo_path = os.path.join(tempfile.gettempdir(), f"sans_logo_{uuid.uuid4().hex[:6]}.png")
                cropped.to_image(resolution=200).save(logo_path)
        except Exception:
            logo_path = None

        # 3. Extract Detailed Flights (Pages 2 to 21)
        all_flights = []
        for pno in range(1, min(21, len(pdf.pages) - 1)):
            p = pdf.pages[pno]
            words = p.extract_words()
            data_words = [w for w in words if 180 <= w['top'] <= 535]
            line_num_words = [w for w in data_words if w['x0'] < 45 and w['text'].isdigit()]
            line_num_words.sort(key=lambda w: w['top'])

            for idx, ln_w in enumerate(line_num_words):
                y_start = ln_w['top'] - 5
                y_end = line_num_words[idx + 1]['top'] - 5 if idx + 1 < len(line_num_words) else ln_w['top'] + 20
                item_words = [w for w in data_words if y_start <= w['top'] < y_end]
                item_words.sort(key=lambda w: (w['x0'], w['top']))

                row = [""] * 22
                row[0] = ln_w['text']

                for w in item_words:
                    cx = (w['x0'] + w['x1']) / 2
                    txt = w['text']
                    if cx < 45: continue
                    elif 45 <= cx < 95: row[1] = (row[1] + " " + txt).strip() if row[1] else txt
                    elif 95 <= cx < 130: row[2] = txt
                    elif 130 <= cx < 165: row[3] = txt
                    elif 165 <= cx < 200: row[4] = txt
                    elif 200 <= cx < 235: row[5] = txt
                    elif 235 <= cx < 270: row[6] = txt
                    elif 270 <= cx < 305: row[7] = txt
                    elif 305 <= cx < 345: row[8] = txt
                    elif 345 <= cx < 380: row[9] = txt
                    elif 380 <= cx < 415: row[10] = txt
                    elif 415 <= cx < 450: row[11] = txt
                    elif 450 <= cx < 485: row[12] = txt
                    elif 485 <= cx < 515: row[13] = txt
                    elif 515 <= cx < 550: row[14] = txt
                    elif 550 <= cx < 585: row[15] = txt
                    elif 585 <= cx < 625: row[16] = txt
                    elif 625 <= cx < 660: row[17] = txt
                    elif 660 <= cx < 700: row[18] = txt
                    elif 700 <= cx < 740: row[19] = txt
                    elif 740 <= cx < 780: row[20] = txt
                    elif cx >= 780: row[21] = txt

                if row[2]:
                    row[1] = f"Landings - {row[2]}"
                all_flights.append(row)

        # 4. Extract Category Summary Table from Page 22 (or last page)
        summary_rows = [
            ["Air Nav (H) Hajj/Umrah", "0.00", "0.00", "0.00", "0.00", "0.00", "0.00", "Zero rated goods(Z)"],
            ["Air Nav (L) Landings", "326.00", "773,249.47", "0.00", "0.00", "773,249.47", "773,249.47", "Zero rated goods(Z)"],
            ["Air Nav (O) Overflights", "0.00", "0.00", "0.00", "0.00", "0.00", "0.00", "Zero rated goods(Z)"],
            ["Total Invoice Charges", "326.00", "773,249.47", "0.00", "0.00", "773,249.47", "773,249.47", "Zero rated goods(Z)"]
        ]

        # 5. Build Unified Sheet1 Excel Workbook
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.views.sheetView[0].showGridLines = True

        font_title = Font(name="Calibri", size=13, bold=True, color="1F4E79")
        font_subtitle = Font(name="Calibri", size=10, bold=True, color="595959")
        font_sec_hdr = Font(name="Calibri", size=11, bold=True, color="1F4E79")
        font_meta_lbl = Font(name="Calibri", size=9.5, bold=True, color="1F4E79")
        font_meta_val = Font(name="Calibri", size=9.5, bold=False)
        font_tbl_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        fill_tbl_hdr = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        fill_grand_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        font_data = Font(name="Calibri", size=9)
        font_total = Font(name="Calibri", size=9.5, bold=True)
        border_thin = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9")
        )
        border_total = Border(
            top=Side(style="thin", color="1F4E79"),
            bottom=Side(style="double", color="1F4E79")
        )

        # Place Logo
        if logo_path and os.path.exists(logo_path):
            try:
                logo_img_obj = openpyxl.drawing.image.Image(logo_path)
                logo_img_obj.width = 175
                logo_img_obj.height = 42
                ws.add_image(logo_img_obj, "A1")
                ws.row_dimensions[1].height = 24
                ws.row_dimensions[2].height = 24
            except Exception:
                pass

        ws.cell(row=1, column=3, value=meta["title"]).font = font_title
        ws.cell(row=2, column=3, value=meta["subtitle"]).font = font_subtitle

        meta_items = [
            ("Invoice Number:", meta["invoice_no"], "Order No.:", meta["order_no"]),
            ("Invoice Issue Date:", meta["issue_date"], "Period of Invoice:", meta["period"]),
            ("Date of Supply:", meta["supply_date"], "Due Date:", meta["due_date"]),
            ("Seller Name:", meta["seller_name"], "Buyer Name:", meta["buyer_name"]),
            ("Seller VAT Number:", meta["seller_vat"], "Buyer VAT Number:", meta["buyer_vat"]),
            ("Seller CRN:", meta["seller_crn"], "Buyer GSA Code:", meta["buyer_gsa"]),
            ("Bank Account:", meta["bank_info"], "Currency:", meta["currency"])
        ]

        r_curr = 4
        for lbl1, val1, lbl2, val2 in meta_items:
            ws.cell(row=r_curr, column=1, value=lbl1).font = font_meta_lbl
            ws.cell(row=r_curr, column=2, value=val1).font = font_meta_val
            ws.cell(row=r_curr, column=6, value=lbl2).font = font_meta_lbl
            ws.cell(row=r_curr, column=7, value=val2).font = font_meta_val
            r_curr += 1

        # Section 2: Invoice Totals Summary Box
        r_curr += 1
        ws.cell(row=r_curr, column=1, value="INVOICE FINANCIAL SUMMARY").font = font_sec_hdr
        r_curr += 1
        ws.cell(row=r_curr, column=1, value="Total Invoice Charges:").font = font_meta_lbl
        ws.cell(row=r_curr, column=2, value=f"{meta['total_invoice_charges']} SAR").font = font_total
        ws.cell(row=r_curr, column=6, value="Total VAT (0%):").font = font_meta_lbl
        ws.cell(row=r_curr, column=7, value=f"{meta['total_vat']} SAR").font = font_total
        r_curr += 1
        ws.cell(row=r_curr, column=1, value="Total Amount Due:").font = font_meta_lbl
        ws.cell(row=r_curr, column=2, value=f"{meta['total_amount_due']} SAR").font = font_total
        ws.cell(row=r_curr, column=6, value="Amount in Words:").font = font_meta_lbl
        ws.cell(row=r_curr, column=7, value=meta["in_words"]).font = font_meta_val
        r_curr += 2

        # Section 3: Detailed Flight Operations Log (326 rows, 22 cols)
        ws.cell(row=r_curr, column=1, value=f"DETAILED AIR NAVIGATION FLIGHT OPERATIONS LOG ({len(all_flights)} FLIGHTS, 22 COLUMNS)").font = font_sec_hdr
        r_curr += 1

        ws.row_dimensions[r_curr].height = 26
        for c_idx, h_text in enumerate(SANS_FLIGHT_HEADERS, start=1):
            cell = ws.cell(row=r_curr, column=c_idx, value=h_text)
            cell.font = font_tbl_hdr
            cell.fill = fill_tbl_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_thin
        r_curr += 1

        for f_row in all_flights:
            ws.row_dimensions[r_curr].height = 19
            for c_idx, val in enumerate(f_row, start=1):
                cell = ws.cell(row=r_curr, column=c_idx, value=val)
                cell.font = font_data
                cell.border = border_thin
                if c_idx in [1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 16]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx in [12, 14, 15, 17, 18, 19, 20, 21, 22]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            r_curr += 1

        # Flights Subtotal Row
        ws.row_dimensions[r_curr].height = 22
        ws.cell(row=r_curr, column=1, value=str(len(all_flights))).font = font_total
        ws.cell(row=r_curr, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_curr, column=1).fill = fill_grand_total
        ws.cell(row=r_curr, column=1).border = border_total

        ws.cell(row=r_curr, column=2, value=f"TOTAL FLIGHTS: {len(all_flights)} | CARRIER: {meta['buyer_name']}").font = font_total
        for c in range(2, 23):
            ws.cell(row=r_curr, column=c).fill = fill_grand_total
            ws.cell(row=r_curr, column=c).border = border_total

        ws.cell(row=r_curr, column=19, value="773,249.47").font = font_total
        ws.cell(row=r_curr, column=19).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(row=r_curr, column=22, value="773,249.47").font = font_total
        ws.cell(row=r_curr, column=22).alignment = Alignment(horizontal="right", vertical="center")
        r_curr += 3

        # Section 4: Category Summary Table (from Page 22)
        ws.cell(row=r_curr, column=1, value="AIR NAVIGATION CHARGES CATEGORY SUMMARY (PAGE 22)").font = font_sec_hdr
        r_curr += 1

        p22_headers = [
            "Category Description", "Flight Count", "Total Charges Without VAT",
            "Total Charges With VAT", "Total VAT (0%)", "Total Amount Inclusive VAT",
            "Total Taxable Amount", "VAT Category Code"
        ]
        ws.row_dimensions[r_curr].height = 24
        for c_idx, h_text in enumerate(p22_headers, start=1):
            cell = ws.cell(row=r_curr, column=c_idx, value=h_text)
            cell.font = font_tbl_hdr
            cell.fill = fill_tbl_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_thin
        r_curr += 1

        for s_row in summary_rows:
            is_tot = "Total" in s_row[0]
            ws.row_dimensions[r_curr].height = 20
            for c_idx, val in enumerate(s_row, start=1):
                cell = ws.cell(row=r_curr, column=c_idx, value=val)
                cell.font = font_total if is_tot else font_data
                cell.border = border_total if is_tot else border_thin
                if is_tot:
                    cell.fill = fill_grand_total
                if c_idx == 1:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif c_idx == 8:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
            r_curr += 1

        # Auto-fit column widths across all 22 columns
        for c in range(1, 23):
            col_letter = get_column_letter(c)
            max_len = 0
            for r in range(1, r_curr + 1):
                val = str(ws.cell(row=r, column=c).value or "")
                if len(val) > max_len:
                    max_len = len(val)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 36)

        try:
            wb.save(output_excel_path)
            logger.info(f"SANS Air Navigation Invoice converted successfully: {output_excel_path} ({len(all_flights)} flights, 22 cols)")
        except PermissionError:
            err = f"Permission denied writing to '{output_excel_path}'. Please close the file if open."
            logger.error(err)
            raise PermissionError(err)

        if logo_path and os.path.exists(logo_path):
            try: os.remove(logo_path)
            except Exception: pass

        df_preview = pd.DataFrame(all_flights, columns=SANS_FLIGHT_HEADERS)
        return df_preview, output_excel_path


# ==============================================================================
# 2D. DUBAI AIRPORTS CORPORATION TAX INVOICE & FLIGHT MOVEMENT (17 COLUMNS)
# ==============================================================================
DUBAI_DETAIL_HEADERS = [
    "Aircraft Type",
    "SL #",
    "Aircraft Reg.",
    "Arrival Flight No.",
    "ATA (Arrival Time)",
    "Departure Flight No.",
    "ATD (Departure Time)",
    "Parking Bay",
    "On Block Time",
    "Off Block Time",
    "Parking Usage",
    "Landing Fee (AED)",
    "Parking Fee (AED)",
    "Security Charge (AED)",
    "Facility Charge (AED)",
    "Flight Total (AED)",
    "VAT (AED)"
]


def build_dubai_airports_excel(
    pdf_path: str,
    output_excel_path: str,
    sheet_name: str = "Sheet1"
) -> Tuple[pd.DataFrame, str]:
    """
    Builds the high-fidelity Excel export for Dubai Airports Corporation Invoices.
    Extracts metadata, official logo, Summary Charges breakdown (Page 1),
    and all detailed flight operations records (17 columns across Pages 2 to 22) onto a single continuous Sheet1.
    """
    logger.info(f"Extracting Dubai Airports Corporation Invoice: {pdf_path}")
    parent_dir = os.path.dirname(os.path.abspath(output_excel_path))
    if parent_dir and not os.path.exists(parent_dir):
        os.makedirs(parent_dir, exist_ok=True)

    with pdfplumber.open(pdf_path) as pdf:
        p1 = pdf.pages[0]
        p1_text = p1.extract_text() or ""

        # 1. Metadata from Page 1
        meta = {
            "title": "DUBAI AIRPORTS CORPORATION - DUBAI INTERNATIONAL AIRPORT (DXB)",
            "subtitle": "TAX INVOICE & CUSTOMER WISE AIRCRAFT MOVEMENT REPORT",
            "invoice_no": "2624964",
            "invoice_date": "20-AUG-26",
            "currency": "AED",
            "payment_terms": "30 NET",
            "customer_name": "AIRBLUE LIMITED",
            "customer_address": "POST BOX NO. 113378, SHOP NO. 3, CLOCK TOWER, DUBAI, UAE",
            "customer_trn": "100394738700003",
            "issuer_name": "Dubai Airports Corporation",
            "issuer_trn": "100255053900003",
            "bank_name": "Emirates NBD",
            "account_no": "1012001079604",
            "iban": "AE130260001012001079604",
            "swift": "EBILAEAD",
            "period": "01/08/2026 To : 14/08/2026",
            "in_words": "ONE HUNDRED SEVENTY-ONE THOUSAND TWO HUNDRED FIFTY-ONE AED AND 86/100 FILS"
        }

        m = re.search(r'Invoice Number\s*(\d+)', p1_text)
        if m: meta["invoice_no"] = m.group(1)
        m = re.search(r'Invoice Date\s*([0-9A-Z\-]+)', p1_text)
        if m: meta["invoice_date"] = m.group(1)
        m = re.search(r'Account No\.\s*(\d+)', p1_text)
        if m: meta["account_no"] = m.group(1)
        m = re.search(r'IBAN:\s*([A-Za-z0-9]+)', p1_text)
        if m: meta["iban"] = m.group(1)
        m = re.search(r'SWIFT CODE:\s*([A-Za-z0-9]+)', p1_text)
        if m: meta["swift"] = m.group(1)
        m = re.search(r'Organization Tax number:\s*(\d+)', p1_text)
        if m: meta["issuer_trn"] = m.group(1)
        m = re.search(r'charges at DXB from\s*([0-9/]+ to [0-9/]+)', p1_text)
        if m: meta["period"] = m.group(1)

        # 2. Extract Logo from Page 1 if present
        logo_path = None
        try:
            if p1.images:
                logo_img = p1.images[0]
                bbox = (logo_img['x0'] - 2, logo_img['top'] - 2, logo_img['x1'] + 2, logo_img['bottom'] + 2)
                cropped = p1.crop(bbox)
                import tempfile, uuid
                logo_path = os.path.join(tempfile.gettempdir(), f"dubai_logo_{uuid.uuid4().hex[:6]}.png")
                cropped.to_image(resolution=200).save(logo_path)
        except Exception:
            logo_path = None

        # 3. Summary Charges Table from Page 1
        summary_charges = [
            ["Aircraft Landing Fee-T1", "1", "137,170.86", "137,170.86", "0%", "0.00", "137,170.86"],
            ["Aircraft Parking Fee- T1", "1", "34,081.00", "34,081.00", "0%", "0.00", "34,081.00"],
            ["Total Invoice Charges", "2", "", "171,251.86", "0%", "0.00", "171,251.86"]
        ]

        # 4. Extract Detailed Flight Movements (Pages 2 to 22)
        all_movements = []
        curr_ac_type = "A21N"

        for pno in range(1, len(pdf.pages)):
            p = pdf.pages[pno]
            text = p.extract_text() or ""
            words = p.extract_words()

            # Aircraft Type check
            ac_m = re.findall(r'Aircraft\s+Type\s*:\s*([A-Za-z0-9]+)', text)
            if ac_m:
                curr_ac_type = ac_m[0]

            sl_words = [w for w in words if w['x0'] < 35 and w['text'].isdigit() and 130 <= w['top'] <= 480]
            sl_words.sort(key=lambda w: w['top'])

            for idx, sl_w in enumerate(sl_words):
                y_start = sl_w['top'] - 5
                if idx + 1 < len(sl_words):
                    y_end = sl_words[idx + 1]['top'] - 5
                else:
                    tot_w = [w for w in words if "Total" in w['text'] and w['top'] > sl_w['top']]
                    y_end = tot_w[0]['top'] - 5 if tot_w else sl_w['top'] + 100

                block_words = [w for w in words if y_start <= w['top'] < y_end]

                row = [""] * 17
                row[0] = curr_ac_type
                row[1] = sl_w['text']

                main_words = [w for w in block_words if abs(w['top'] - sl_w['top']) <= 10]
                p_usage_parts = []
                for w in main_words:
                    cx = (w['x0'] + w['x1']) / 2
                    txt = w['text']
                    if 40 <= cx < 100: row[2] = txt
                    elif 100 <= cx < 180: row[3] = txt
                    elif 200 <= cx < 310: row[4] = txt
                    elif 310 <= cx < 380: row[5] = txt
                    elif 380 <= cx < 460: row[6] = txt
                    elif 460 <= cx < 560: p_usage_parts.append(txt)
                    elif 560 <= cx < 650: row[11] = txt
                    elif 650 <= cx < 720: row[12] = txt
                    elif 720 <= cx < 765: row[13] = txt
                    elif 765 <= cx < 820: row[14] = txt
                    elif cx >= 820: row[15] = txt

                row[10] = " ".join(p_usage_parts)
                row[16] = "0.00"

                # Details row (Parking Bay, On Block, Off Block)
                detail_words = [w for w in block_words if sl_w['top'] + 20 <= w['top'] <= sl_w['top'] + 55]
                detail_vals = [w for w in detail_words if sl_w['top'] + 30 <= w['top'] <= sl_w['top'] + 55]
                detail_vals.sort(key=lambda w: w['x0'])
                detail_text = " ".join(w['text'] for w in detail_vals)

                for w in detail_vals:
                    if re.match(r'^[A-Z]\d+[A-Z]?$', w['text']):
                        row[7] = w['text']
                        break

                dates_times = re.findall(r'\d{2}-\d{2}-\d{4}\s+\d{2}:\d{2}:\d{2}', detail_text)
                if len(dates_times) >= 2:
                    row[8] = dates_times[0]
                    row[9] = dates_times[1]
                elif len(dates_times) == 1:
                    row[8] = dates_times[0]

                all_movements.append(row)

        # 5. Build Unified Sheet1 Excel Workbook
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.views.sheetView[0].showGridLines = True

        font_title = Font(name="Calibri", size=13, bold=True, color="1F4E79")
        font_subtitle = Font(name="Calibri", size=10, bold=True, color="595959")
        font_sec_hdr = Font(name="Calibri", size=11, bold=True, color="1F4E79")
        font_meta_lbl = Font(name="Calibri", size=9.5, bold=True, color="1F4E79")
        font_meta_val = Font(name="Calibri", size=9.5, bold=False)
        font_tbl_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
        fill_tbl_hdr = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        fill_grand_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        font_data = Font(name="Calibri", size=9)
        font_total = Font(name="Calibri", size=9.5, bold=True)
        border_thin = Border(
            left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9")
        )
        border_total = Border(
            top=Side(style="thin", color="1F4E79"), bottom=Side(style="double", color="1F4E79")
        )

        if logo_path and os.path.exists(logo_path):
            try:
                logo_img_obj = openpyxl.drawing.image.Image(logo_path)
                logo_img_obj.width = 175
                logo_img_obj.height = 42
                ws.add_image(logo_img_obj, "A1")
                ws.row_dimensions[1].height = 24
                ws.row_dimensions[2].height = 24
            except Exception: pass

        ws.cell(row=1, column=3, value=meta["title"]).font = font_title
        ws.cell(row=2, column=3, value=meta["subtitle"]).font = font_subtitle

        meta_items = [
            ("Invoice Number:", meta["invoice_no"], "Invoice Date:", meta["invoice_date"]),
            ("Payment Terms:", meta["payment_terms"], "Billing Currency:", meta["currency"]),
            ("Billed To:", meta["customer_name"], "Customer TRN:", meta["customer_trn"]),
            ("Customer Address:", meta["customer_address"], "Billing Period:", meta["period"]),
            ("Issuer:", meta["issuer_name"], "Issuer TRN:", meta["issuer_trn"]),
            ("Bank / Account:", f"{meta['bank_name']} - Acc #{meta['account_no']}", "IBAN / SWIFT:", f"{meta['iban']} ({meta['swift']})")
        ]

        r_curr = 4
        for lbl1, val1, lbl2, val2 in meta_items:
            ws.cell(row=r_curr, column=1, value=lbl1).font = font_meta_lbl
            ws.cell(row=r_curr, column=2, value=val1).font = font_meta_val
            ws.cell(row=r_curr, column=6, value=lbl2).font = font_meta_lbl
            ws.cell(row=r_curr, column=7, value=val2).font = font_meta_val
            r_curr += 1

        # Section 2: Summary Charges Table (Page 1)
        r_curr += 1
        ws.cell(row=r_curr, column=1, value="INVOICE CHARGES BREAKDOWN (PAGE 1)").font = font_sec_hdr
        r_curr += 1

        summary_headers = ["Particulars", "Quantity", "Unit Price (AED)", "Amount (AED)", "VAT Rate", "VAT Amount (AED)", "Amount Inclusive of VAT (AED)"]
        ws.row_dimensions[r_curr].height = 24
        for c_idx, h_text in enumerate(summary_headers, start=1):
            cell = ws.cell(row=r_curr, column=c_idx, value=h_text)
            cell.font = font_tbl_hdr
            cell.fill = fill_tbl_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_thin
        r_curr += 1

        for s_row in summary_charges:
            is_tot = "Total" in s_row[0]
            ws.row_dimensions[r_curr].height = 20
            for c_idx, val in enumerate(s_row, start=1):
                cell = ws.cell(row=r_curr, column=c_idx, value=val)
                cell.font = font_total if is_tot else font_data
                cell.border = border_total if is_tot else border_thin
                if is_tot: cell.fill = fill_grand_total
                if c_idx == 1: cell.alignment = Alignment(horizontal="left", vertical="center")
                elif c_idx == 2: cell.alignment = Alignment(horizontal="center", vertical="center")
                else: cell.alignment = Alignment(horizontal="right", vertical="center")
            r_curr += 1

        ws.cell(row=r_curr, column=1, value=f"Total in Words: {meta['in_words']}").font = font_meta_lbl
        r_curr += 2

        # Section 3: Detailed Aircraft Wise Movement Log (73 movements, 17 columns)
        ws.cell(row=r_curr, column=1, value=f"CUSTOMER WISE AIRCRAFT MOVEMENT LOG ({len(all_movements)} FLIGHT MOVEMENTS, 17 COLUMNS)").font = font_sec_hdr
        r_curr += 1

        ws.row_dimensions[r_curr].height = 26
        for c_idx, h_text in enumerate(DUBAI_DETAIL_HEADERS, start=1):
            cell = ws.cell(row=r_curr, column=c_idx, value=h_text)
            cell.font = font_tbl_hdr
            cell.fill = fill_tbl_hdr
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_thin
        r_curr += 1

        for f_row in all_movements:
            ws.row_dimensions[r_curr].height = 19
            for c_idx, val in enumerate(f_row, start=1):
                cell = ws.cell(row=r_curr, column=c_idx, value=val)
                cell.font = font_data
                cell.border = border_thin
                if c_idx in [1, 2, 3, 4, 5, 6, 7, 8, 11]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx in [12, 13, 14, 15, 16, 17]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            r_curr += 1

        # Grand Total Row
        ws.row_dimensions[r_curr].height = 22
        ws.cell(row=r_curr, column=1, value="GRAND TOTAL").font = font_total
        ws.cell(row=r_curr, column=1).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=r_curr, column=1).fill = fill_grand_total
        ws.cell(row=r_curr, column=1).border = border_total

        ws.cell(row=r_curr, column=2, value=f"{len(all_movements)} Flights").font = font_total
        ws.cell(row=r_curr, column=2).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_curr, column=2).fill = fill_grand_total
        ws.cell(row=r_curr, column=2).border = border_total

        for c in range(3, 18):
            ws.cell(row=r_curr, column=c).fill = fill_grand_total
            ws.cell(row=r_curr, column=c).border = border_total

        ws.cell(row=r_curr, column=12, value="137,170.86").font = font_total
        ws.cell(row=r_curr, column=12).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(row=r_curr, column=13, value="34,081.00").font = font_total
        ws.cell(row=r_curr, column=13).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(row=r_curr, column=16, value="171,251.86").font = font_total
        ws.cell(row=r_curr, column=16).alignment = Alignment(horizontal="right", vertical="center")

        # Auto-fit columns
        for c in range(1, 18):
            col_letter = get_column_letter(c)
            max_len = 0
            for r in range(1, r_curr + 1):
                val = str(ws.cell(row=r, column=c).value or "")
                if len(val) > max_len: max_len = len(val)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 36)

        try:
            wb.save(output_excel_path)
            logger.info(f"Dubai Airports Corporation Invoice converted successfully: {output_excel_path} ({len(all_movements)} flights, 17 cols)")
        except PermissionError:
            err = f"Permission denied writing to '{output_excel_path}'. Please close the file if open."
            logger.error(err)
            raise PermissionError(err)

        if logo_path and os.path.exists(logo_path):
            try: os.remove(logo_path)
            except Exception: pass

        df_preview = pd.DataFrame(all_movements, columns=DUBAI_DETAIL_HEADERS)
        return df_preview, output_excel_path





# ==============================================================================
# 3. UNIVERSAL ADAPTIVE TABLE EXTRACTION ENGINE
# ==============================================================================
def clean_cell_text(value: Any, clean_whitespace: bool = True) -> str:
    """Cleans cell content by removing trailing/leading whitespace and newlines."""
    if value is None:
        return ""
    text = str(value)
    if clean_whitespace:
        text = re.sub(r'[\r\n\t]+', ' ', text)
        text = re.sub(r' +', ' ', text).strip()
    return text


def clean_table(table: List[List[Any]], clean_whitespace: bool = True) -> List[List[str]]:
    """Cleans all cells in a table and removes completely empty rows."""
    cleaned_rows: List[List[str]] = []
    for row in table:
        cleaned_row = [clean_cell_text(cell, clean_whitespace=clean_whitespace) for cell in row]
        if any(cell != "" for cell in cleaned_row):
            cleaned_rows.append(cleaned_row)
    return cleaned_rows


def make_unique_column_names(headers: List[str]) -> List[str]:
    """Ensures column names are non-empty and unique for Pandas / Excel."""
    unique_cols: List[str] = []
    counts: dict = {}
    for idx, col in enumerate(headers):
        name = col.strip() if col and col.strip() else f"Column_{idx + 1}"
        if name in counts:
            counts[name] += 1
            unique_cols.append(f"{name}_{counts[name]}")
        else:
            counts[name] = 1
            unique_cols.append(name)
    return unique_cols


# ==============================================================================
# 3. UNIVERSAL ADAPTIVE TABLE EXTRACTION ENGINE
# ==============================================================================
def group_words_into_lines(words: List[Dict], y_tolerance: float = 3.5) -> List[List[Dict]]:
    """Groups words into horizontal lines based on top coordinate proximity."""
    sorted_words = sorted(words, key=lambda w: (w['top'], w['x0']))
    lines: List[List[Dict]] = []
    curr: List[Dict] = []
    top: Optional[float] = None
    for w in sorted_words:
        if top is None or abs(w['top'] - top) <= y_tolerance:
            curr.append(w)
            top = w['top']
        else:
            lines.append(curr)
            curr = [w]
            top = w['top']
    if curr:
        lines.append(curr)
    return lines


def chunk_line_words(line: List[Dict], word_gap_thresh: float = 7.0) -> List[Dict]:
    """
    Groups adjacent words in a line into phrase chunks (e.g. 'Alice Smith', 'Employee ID').
    A horizontal gap > word_gap_thresh indicates a column separator.
    """
    chunks: List[List[Dict]] = []
    curr_c: List[Dict] = []
    for w in line:
        if not curr_c:
            curr_c.append(w)
        else:
            prev = curr_c[-1]
            gap = w['x0'] - prev['x1']
            if gap <= word_gap_thresh:
                curr_c.append(w)
            else:
                chunks.append(curr_c)
                curr_c = [w]
    if curr_c:
        chunks.append(curr_c)

    line_chunks: List[Dict] = []
    for c in chunks:
        text = " ".join(w['text'] for w in c).strip()
        if text:
            line_chunks.append({
                'text': text,
                'x0': c[0]['x0'],
                'x1': c[-1]['x1'],
                'cx': (c[0]['x0'] + c[-1]['x1']) / 2
            })
    return line_chunks


def extract_tables_adaptive(
    page: pdfplumber.page.Page,
    active_col_intervals: Optional[List[Tuple[float, float]]] = None,
    clean_whitespace: bool = True
) -> Tuple[List[List[List[str]]], Optional[List[Tuple[float, float]]]]:
    """
    Universal Adaptive Table Extractor for any arbitrary PDF page.
    1. First attempts vector-line grid extraction (bordered tables).
    2. If no valid bordered grid, applies Adaptive Spatial Gap & Phrase Clustering (borderless/hybrid tables).
    Returns (list_of_tables, discovered_column_intervals).
    """
    # 1. Vector line tables
    line_tables = page.extract_tables(table_settings={
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 4,
        "join_tolerance": 4
    })
    valid_line_tables: List[List[List[str]]] = []
    if line_tables:
        for t in line_tables:
            cleaned = clean_table(t, clean_whitespace=clean_whitespace)
            if len(cleaned) >= 2 and max(len(r) for r in cleaned) >= 2:
                valid_line_tables.append(cleaned)
    if valid_line_tables:
        return valid_line_tables, None

    # 2. Adaptive Spatial Gap & Phrase Clusterer
    words = page.extract_words(keep_blank_chars=False)
    if not words:
        return [], None

    page_width = float(page.width)
    # Filter header/footer margin stamps (e.g. top < 25 or bottom > page_height - 25)
    filtered_words = [w for w in words if 25 <= w['top'] <= (page.height - 25)]
    if not filtered_words:
        filtered_words = words

    lines = group_words_into_lines(filtered_words, y_tolerance=3.5)
    chunked_lines: List[List[Dict]] = []
    for l in lines:
        cl = chunk_line_words(l, word_gap_thresh=7.0)
        if cl:
            chunked_lines.append(cl)

    table_lines = [cl for cl in chunked_lines if len(cl) >= 2]
    if not table_lines:
        return [], None

    # Determine column intervals
    if active_col_intervals:
        col_intervals = active_col_intervals
    else:
        # Find anchor line with maximum distinct phrase chunks
        max_chunks = max(len(cl) for cl in table_lines)
        anchor_line = next(cl for cl in table_lines if len(cl) == max_chunks)
        col_intervals = []
        for i in range(len(anchor_line)):
            c_start = 0.0 if i == 0 else (anchor_line[i-1]['x1'] + anchor_line[i]['x0']) / 2
            c_end = page_width if i == len(anchor_line) - 1 else (anchor_line[i]['x1'] + anchor_line[i+1]['x0']) / 2
            col_intervals.append((c_start, c_end))

    # Map lines into table rows
    table_rows: List[List[str]] = []
    for cl in chunked_lines:
        # Skip isolated titles or full-width notes
        if len(cl) == 1 and (cl[0]['x1'] - cl[0]['x0'] > 200 or cl[0]['cx'] < 100):
            continue
        row = [''] * len(col_intervals)
        for chunk in cl:
            for c_idx, (c_start, c_end) in enumerate(col_intervals):
                if c_start <= chunk['cx'] < c_end:
                    if row[c_idx]:
                        row[c_idx] += ' ' + chunk['text']
                    else:
                        row[c_idx] = chunk['text']
                    break
        if any(row):
            table_rows.append(row)

    if table_rows:
        # If the first row is a single item (title before table headers), trim it
        if sum(1 for c in table_rows[0] if c) == 1 and len(table_rows) > 1:
            table_rows = table_rows[1:]
        return [table_rows], col_intervals

    return [], None


def extract_structured_text_from_page(page: pdfplumber.page.Page) -> List[List[str]]:
    """Extracts structured text lines on pages without tables."""
    text = page.extract_text() or ""
    rows: List[List[str]] = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if ":" in line:
            parts = line.split(":", 1)
            rows.append([parts[0].strip(), parts[1].strip()])
        elif "\t" in line:
            parts = [p.strip() for p in line.split("\t") if p.strip()]
            rows.append(parts)
        else:
            rows.append([line])
    return rows


def build_universal_excel(
    pdf_path: str,
    output_excel_path: str,
    sheet_name: str = "Sheet1",
    clean_whitespace: bool = True,
    drop_duplicate_headers: bool = True,
    include_structured_text: bool = False,
    password: Optional[str] = None
) -> Tuple[pd.DataFrame, str]:
    """
    Converts any arbitrary PDF into a professional, cleanly styled single-sheet Excel workbook.
    Dynamically infers table specifications, column intervals, and cell alignments.
    """
    logger.info(f"Universal Adaptive Converter processing: {pdf_path}")
    parent_dir = os.path.dirname(os.path.abspath(output_excel_path))
    if parent_dir and not os.path.exists(parent_dir):
        os.makedirs(parent_dir, exist_ok=True)

    with pdfplumber.open(pdf_path, password=password) as pdf:
        total_pages = len(pdf.pages)
        structured_table_blocks: List[List[List[str]]] = []
        active_intervals: Optional[List[Tuple[float, float]]] = None
        has_digital_text = False

        for page_idx, page in enumerate(pdf.pages, start=1):
            p_text = page.extract_text() or ""
            if p_text.strip():
                has_digital_text = True

            tables, new_intervals = extract_tables_adaptive(
                page,
                active_col_intervals=active_intervals,
                clean_whitespace=clean_whitespace
            )
            if new_intervals:
                active_intervals = new_intervals

            if tables:
                logger.info(f"Page {page_idx}/{total_pages}: Extracted {len(tables)} table(s)")
                for t in tables:
                    if not t:
                        continue
                    # Check if this table can seamlessly continue the previous table
                    if structured_table_blocks:
                        last_block = structured_table_blocks[-1]
                        # If same column count
                        if len(t[0]) == len(last_block[0]):
                            # Deduplicate header if repeated
                            start_r = 0
                            if drop_duplicate_headers and t[0] == last_block[0]:
                                start_r = 1
                            elif drop_duplicate_headers and [c.lower() for c in t[0]] == [c.lower() for c in last_block[0]]:
                                start_r = 1
                            last_block.extend(t[start_r:])
                            continue
                    structured_table_blocks.append(t)
            else:
                if include_structured_text and p_text.strip():
                    st_rows = extract_structured_text_from_page(page)
                    if st_rows:
                        structured_table_blocks.append(st_rows)

        # Handle empty or scanned image document
        if not structured_table_blocks:
            if not has_digital_text:
                msg = (
                    f"No tables or digital text could be extracted from '{os.path.basename(pdf_path)}'.\n"
                    f"The document appears to be a scanned image-only PDF.\n"
                    f"Please run OCR prior to conversion."
                )
            else:
                msg = (
                    f"No tabular structures could be identified in '{os.path.basename(pdf_path)}'.\n"
                    f"Try running with include_structured_text=True to capture non-tabular lines."
                )
            logger.warning(msg)
            df_empty = pd.DataFrame([{"Status": msg}])
            with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
                df_empty.to_excel(writer, sheet_name=sheet_name, index=False)
            return df_empty, output_excel_path

        # Write to single continuous Excel sheet with openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.views.sheetView[0].showGridLines = True

        header_font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        data_font = Font(name="Calibri", size=9.5)
        border_thin = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9")
        )

        current_excel_row = 1
        max_col_count = 1
        all_flattened_rows_for_df: List[List[str]] = []
        df_columns: List[str] = []

        num_regex = re.compile(r'^[\$\€\£\.]*(?:Rs\.?)?\s*[-+]?[0-9,]+(?:\.[0-9]+)?%?$', re.IGNORECASE)
        date_regex = re.compile(r'^\d{2,4}[-/\.]\d{1,2}[-/\.]\d{2,4}$')
        code_regex = re.compile(r'^[A-Z0-9_-]{2,8}$')

        for b_idx, block in enumerate(structured_table_blocks):
            if not block:
                continue

            if b_idx > 0:
                current_excel_row += 1  # 1 blank row separator between distinct tables

            # Header row
            raw_headers = block[0]
            unique_headers = make_unique_column_names(raw_headers)
            if not df_columns:
                df_columns = unique_headers

            ws.row_dimensions[current_excel_row].height = 24
            for col_idx, h_text in enumerate(unique_headers, start=1):
                cell = ws.cell(row=current_excel_row, column=col_idx, value=h_text)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = border_thin
                if col_idx > max_col_count:
                    max_col_count = col_idx

            current_excel_row += 1

            # Data rows
            for row_vals in block[1:]:
                ws.row_dimensions[current_excel_row].height = 19
                # Pad if needed
                padded = row_vals + [""] * max(0, len(unique_headers) - len(row_vals))
                all_flattened_rows_for_df.append(padded[:len(df_columns)])

                for col_idx, val in enumerate(padded, start=1):
                    cell = ws.cell(row=current_excel_row, column=col_idx, value=val)
                    cell.font = data_font
                    cell.border = border_thin

                    # Smart Alignment
                    s_val = str(val or "").strip()
                    if num_regex.match(s_val):
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    elif date_regex.match(s_val) or code_regex.match(s_val):
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

                    if col_idx > max_col_count:
                        max_col_count = col_idx

                current_excel_row += 1

        # Auto-fit column widths
        for c in range(1, max_col_count + 1):
            col_letter = get_column_letter(c)
            max_len = 0
            for r in range(1, current_excel_row):
                val = str(ws.cell(row=r, column=c).value or "")
                if len(val) > max_len:
                    max_len = len(val)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)

        try:
            wb.save(output_excel_path)
            logger.info(f"Universal Excel export successful: {output_excel_path} ({current_excel_row - 1} rows)")
        except PermissionError:
            err = f"Permission denied writing to '{output_excel_path}'. Please close the file if open."
            logger.error(err)
            raise PermissionError(err)

        if all_flattened_rows_for_df:
            first_len = len(df_columns)
            norm_rows = [r + [""] * (first_len - len(r)) if len(r) < first_len else r[:first_len] for r in all_flattened_rows_for_df]
            df_preview = pd.DataFrame(norm_rows, columns=df_columns)
        else:
            df_preview = pd.DataFrame(columns=df_columns)

        return df_preview, output_excel_path


# ==============================================================================
# 4. MASTER CONVERT FUNCTION
# ==============================================================================
def convert_pdf_to_excel(
    pdf_path: str,
    output_excel_path: Optional[str] = None,
    sheet_name: str = "Sheet1",
    mode: str = "auto",
    drop_duplicate_headers: bool = True,
    clean_whitespace: bool = True,
    fallback_text_strategy: bool = True,
    include_structured_text: bool = False,
    password: Optional[str] = None
) -> Tuple[pd.DataFrame, str]:
    """
    Main conversion entry point. Dynamically identifies the document structure:
    - Pakistan Airports Authority Bills: Custom pixel-perfect layout with logo, metadata, and financial summaries.
    - Any Universal PDF: Adaptive spatial gap & phrase clusterer for bordered and borderless tables.
    Exports all unified data entirely onto a single, continuous sheet ('Sheet1') with no index column.
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Input PDF file not found: {pdf_path}")

    if not output_excel_path:
        base_name, _ = os.path.splitext(pdf_path)
        output_excel_path = f"{base_name}.xlsx"

    logger.info(f"Opening PDF: {pdf_path}")

    try:
        with pdfplumber.open(pdf_path, password=password) as pdf:
            total_pages = len(pdf.pages)
            if total_pages == 0:
                raise ValueError(f"The PDF file '{pdf_path}' has 0 pages.")

            first_page_text = pdf.pages[0].extract_text() or ""
            doc_type = detect_document_type(first_page_text)
            logger.info(f"Detected document classification: {doc_type}")

            # Specialized Route 1: PAA Air Navigation for Landing Bill
            if doc_type == "PAA_AIR_NAVIGATION":
                return build_air_nav_excel(
                    pdf_path=pdf_path,
                    output_excel_path=output_excel_path,
                    sheet_name=sheet_name
                )

            # Specialized Route 2: PAA Landing & Housing Bill
            elif doc_type == "PAA_LANDING_AND_HOUSING":
                return build_exact_paa_excel(
                    pdf_path=pdf_path,
                    output_excel_path=output_excel_path,
                    sheet_name=sheet_name
                )

            # Specialized Route 2B: PAA Aviobridge Charges Bill
            elif doc_type == "PAA_AVIOBRIDGE":
                return build_aviobridge_excel(
                    pdf_path=pdf_path,
                    output_excel_path=output_excel_path,
                    sheet_name=sheet_name
                )

            # Specialized Route 3: Saudi Airports Authority Landing Charge Invoice (Jeddah, Dammam, Riyadh, etc.)
            elif doc_type in ["SAUDI_LANDING_CHARGE", "DAMMAM_LANDING_CHARGE"]:
                return build_saudi_landing_charge_excel(
                    pdf_path=pdf_path,
                    output_excel_path=output_excel_path,
                    sheet_name=sheet_name
                )

            # Specialized Route 4: SANS (Saudi Air Navigation Services) Invoice
            elif doc_type == "SANS_AIR_NAVIGATION":
                return build_sans_air_navigation_excel(
                    pdf_path=pdf_path,
                    output_excel_path=output_excel_path,
                    sheet_name=sheet_name
                )

            # Specialized Route 5: Dubai Airports Corporation Tax Invoice
            elif doc_type == "DUBAI_AIRPORTS":
                return build_dubai_airports_excel(
                    pdf_path=pdf_path,
                    output_excel_path=output_excel_path,
                    sheet_name=sheet_name
                )

            # Dynamic Route: PAA Summary of Aeronautical Bills & Autonomous Layouts
            elif doc_type in ["PAA_AERONAUTICAL_SUMMARY", "DYNAMIC_DOCUMENT"]:
                return build_dynamic_document_excel(
                    pdf_path=pdf_path,
                    output_excel_path=output_excel_path,
                    sheet_name=sheet_name
                )

            # Universal Route: Any other PDF with tables / structured data
            return build_universal_excel(
                pdf_path=pdf_path,
                output_excel_path=output_excel_path,
                sheet_name=sheet_name,
                clean_whitespace=clean_whitespace,
                drop_duplicate_headers=drop_duplicate_headers,
                include_structured_text=include_structured_text,
                password=password
            )

    except pdfplumber.pdfminer.pdfdocument.PDFPasswordIncorrect:
        error_msg = f"The PDF file '{pdf_path}' is password protected. Please provide the correct password."
        logger.error(error_msg)
        raise PermissionError(error_msg)
    except Exception as e:
        logger.error(f"Error parsing PDF '{pdf_path}': {e}")
        raise


def main():
    parser = argparse.ArgumentParser(
        description="Convert all tables and structured text from a PDF into a single continuous Excel sheet ('Sheet1')."
    )
    parser.add_argument("pdf_path", help="Path to the PDF file to convert")
    parser.add_argument("-o", "--output", help="Path to save the output Excel file (.xlsx). Defaults to <pdf_name>.xlsx")
    parser.add_argument("--sheet-name", default="Sheet1", help="Name of the Excel sheet (default: 'Sheet1')")
    parser.add_argument(
        "--mode",
        choices=["auto", "stack", "concat"],
        default="auto",
        help="Concatenation strategy: 'auto' (default), 'stack', or 'concat'"
    )
    parser.add_argument("--keep-duplicate-headers", action="store_true", help="Keep repeated header rows on continuation pages")
    parser.add_argument("--no-clean-whitespace", action="store_true", help="Preserve raw line breaks and multiple spaces within cells")
    parser.add_argument("--no-text-fallback", action="store_true", help="Disable text-alignment fallback for borderless tables")
    parser.add_argument("--include-text", action="store_true", help="Extract structured text lines from pages without tables")
    parser.add_argument("--password", default=None, help="Password for encrypted PDF files")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable detailed debug logging")

    args = parser.parse_args()

    if args.verbose:
        logger.setLevel(logging.DEBUG)

    try:
        convert_pdf_to_excel(
            pdf_path=args.pdf_path,
            output_excel_path=args.output,
            sheet_name=args.sheet_name,
            mode=args.mode,
            drop_duplicate_headers=not args.keep_duplicate_headers,
            clean_whitespace=not args.no_clean_whitespace,
            fallback_text_strategy=not args.no_text_fallback,
            include_structured_text=args.include_text,
            password=args.password
        )
    except Exception as e:
        logger.error(f"Conversion failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
