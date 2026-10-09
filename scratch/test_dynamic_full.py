import os
import re
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

def clean_val(v):
    return re.sub(r'\s+', ' ', str(v or "")).strip()

def parse_num(v_str):
    s = str(v_str or "").strip().replace(",", "")
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

def group_words_by_y(words, y_tol=4.0):
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

def test_full_dynamic(pdf_path, output_excel_path):
    print(f"Analyzing {pdf_path} dynamically...")
    with pdfplumber.open(pdf_path) as pdf:
        p1 = pdf.pages[0]
        words_p1 = p1.extract_words()
        lines_p1 = group_words_by_y(words_p1, y_tol=3.5)
        
        # 1. Discover header line
        header_line_idx = -1
        header_chunks = []
        for idx, line in enumerate(lines_p1):
            chunks = []
            curr = []
            for w in line:
                if not curr or (w['x0'] - curr[-1]['x1'] < 16.0):
                    curr.append(w)
                else:
                    chunks.append(curr)
                    curr = [w]
            if curr:
                chunks.append(curr)
            line_str = " ".join(w['text'] for w in line)
            if len(chunks) >= 4 and any(k in line_str.upper() for k in ["SR#", "BILL", "NO.", "DATE", "AMOUNT"]):
                header_line_idx = idx
                header_chunks = chunks
                break
        
        print(f"Header discovered on line {header_line_idx}: {len(header_chunks)} columns")
        
        # Check if next line contains sub-headers like (RS.) IF ANY
        sub_headers = []
        if header_line_idx + 1 < len(lines_p1):
            next_line = lines_p1[header_line_idx + 1]
            next_str = " ".join(w['text'] for w in next_line)
            if any(k in next_str for k in ["(RS.)", "(US$)", "IF ANY", "PKR"]):
                sub_chunks = []
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
        
        # Build column definitions: (Header Name, x0, x1)
        # Combine main header with sub-header
        col_names = []
        for i, c in enumerate(header_chunks):
            txt = " ".join(w['text'] for w in c)
            cx = (c[0]['x0'] + c[-1]['x1']) / 2
            # Check if any sub_header falls near cx
            for sc in sub_headers:
                sc_cx = (sc[0]['x0'] + sc[-1]['x1']) / 2
                if abs(sc_cx - cx) < 30.0 or (c[0]['x0'] - 10 <= sc_cx <= c[-1]['x1'] + 10):
                    txt += " " + " ".join(w['text'] for w in sc)
                    break
            col_names.append(txt)
        
        print("Column Names:")
        for i, h in enumerate(col_names):
            print(f"  Col {i}: {h}")
            
        # Determine column intervals across page width
        col_intervals = []
        n_cols = len(header_chunks)
        page_width = float(p1.width)
        for i in range(n_cols):
            x_start = 0.0 if i == 0 else (header_chunks[i-1][-1]['x1'] + header_chunks[i][0]['x0']) / 2
            x_end = page_width if i == n_cols - 1 else (header_chunks[i][-1]['x1'] + header_chunks[i+1][0]['x0']) / 2
            col_intervals.append((col_names[i], x_start, x_end))
            
        # Extract rows across all pages
        all_data_rows = []
        summary_rows = []
        header_y_max = max(w['bottom'] for w in lines_p1[header_line_idx + (1 if sub_headers else 0)])
        
        for p_idx, page in enumerate(pdf.pages):
            p_words = page.extract_words()
            p_lines = group_words_by_y(p_words, y_tol=3.5)
            
            # Check if this page has the table header or table columns
            p_text = page.extract_text() or ""
            is_table_page = "SR#" in p_text and "BILL NO." in p_text
            
            for line in p_lines:
                line_str = " ".join(w['text'] for w in line)
                line_y = line[0]['top']
                
                # Skip header block on page 1, or running headers
                if p_idx == 0 and line_y <= header_y_max:
                    continue
                if any(k in line_str for k in ["PAKISTAN AIRPORTS AUTHORITY", "PAA-001-FNBL", "PAGE NO:", "SUMMARY OF AERONAUTICAL BILLS", "June 17, 2026", "5:06 pm", "AIRLINE :"]):
                    continue
                if "SR#" in line_str and "BILL NO." in line_str:
                    continue
                if "(RS.)" in line_str and "IF ANY" in line_str:
                    continue
                    
                # Check if this line is in the financial summary block
                if any(k in line_str.upper() for k in ["TOTAL AMOUNT DUE", "ARREARS AS ON", "SURCHARGE AS ON", "2% + KIBOR", "TOTAL ARREARS", "TOTAL AMOUNT PAYABLE", "NOTE :"]):
                    summary_rows.append(line_str)
                    continue
                
                if not is_table_page:
                    summary_rows.append(line_str)
                    continue
                    
                # Check for table row
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
                            
                # Validate if it's a data row
                if row[0] and row[0].isdigit():
                    all_data_rows.append(row)
                elif any(row):
                    summary_rows.append(line_str)
                    
        print(f"Total data rows extracted: {len(all_data_rows)}")
        if all_data_rows:
            print("First row:", all_data_rows[0])
            print("Last row:", all_data_rows[-1])
            # Check sum of column 3 (AMOUNT BILLED)
            col3_sum = sum(int(r[3].replace(",", "")) for r in all_data_rows if r[3])
            print(f"Sum of Column 3 (AMOUNT BILLED): {col3_sum:,} Rs")
            
        print("Summary rows captured:")
        for sr in summary_rows:
            print(" ", sr)

if __name__ == "__main__":
    test_full_dynamic(
        'tests/samples/PAA_Aeronautical_Bills_Summary.pdf',
        'tests/samples/test_dynamic_summary.xlsx'
    )
