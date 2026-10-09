import os
import re
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

def parse_number(val_str):
    """Parses numeric string to int or float if applicable, else None."""
    s = val_str.strip().replace(",", "")
    # Check if purely an integer
    if re.fullmatch(r"[-+]?\d+", s):
        try:
            return int(s)
        except ValueError:
            return None
    # Check if a float
    if re.fullmatch(r"[-+]?\d+\.\d+", s):
        try:
            return float(s)
        except ValueError:
            return None
    return None

def group_words(words, y_tol=4.0):
    if not words:
        return []
    sorted_words = sorted(words, key=lambda w: (w['top'], w['x0']))
    lines = []
    curr_line = []
    curr_top = None
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

def analyze_and_extract_dynamic_table(pdf_path):
    """
    Dynamically analyzes PDF pages, discovers headers, column intervals,
    data rows, total rows, and summary blocks without hardcoded column lists.
    """
    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        p1 = pdf.pages[0]
        words_p1 = p1.extract_words()
        
        # 1. Identify Document Header & Metadata
        # Group words on page 1 by y-coordinate (tolerance 4.0)
        lines_p1 = group_words(words_p1, y_tol=4.0)
        
        # Find potential table header:
        # A line with multiple words spaced across width, containing typical header keywords or high token spread
        header_line_idx = -1
        header_lines = []
        for idx, line in enumerate(lines_p1):
            line_text = " ".join(w['text'] for w in line)
            # Table headers often have SR#, NO, DATE, AMOUNT, BILLED, CHARGES, FLIGHT, ITEM, CODE, DESCRIPTION, etc.
            # or contain >= 3 distinct spaced tokens spanning > 50% of page width
            x_min = min(w['x0'] for w in line)
            x_max = max(w['x1'] for w in line)
            span = x_max - x_min
            if len(line) >= 3 and span > p1.width * 0.5:
                # Check if it has header-like words or is followed by another header line
                if any(k in line_text.upper() for k in ["SR#", "BILL", "DATE", "AMOUNT", "ITEM", "FLIGHT", "NO.", "CHARGES", "DESCRIPTION", "TOTAL"]):
                    header_line_idx = idx
                    header_lines.append(line)
                    # Check if next line is a subheader (e.g. (RS.) IF ANY)
                    if idx + 1 < len(lines_p1):
                        next_line = lines_p1[idx + 1]
                        next_text = " ".join(w['text'] for w in next_line)
                        if any(k in next_text for k in ["(RS.)", "(US$)", "IF ANY", "PKR", "(TONS)", "TIME"]):
                            header_lines.append(next_line)
                    break
        
        print(f"Discovered header line idx: {header_line_idx}")
        for hl in header_lines:
            print("  Header line:", " ".join(w['text'] for w in hl))

if __name__ == "__main__":
    analyze_and_extract_dynamic_table('tests/samples/PAA_Aeronautical_Bills_Summary.pdf')
