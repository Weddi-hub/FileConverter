import os
import re
import pdfplumber

def group_words_by_y(words, y_tol=3.5):
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

header_keywords = [
    "SR#", "BILL", "NO.", "DATE", "AMOUNT", "CHARGES", "DESCRIPTION", "QTY", "QUANTITY",
    "RATE", "PRICE", "TOTAL", "FLIGHT", "ITEM", "CODE", "TAX", "ID", "NAME", "ACCOUNT",
    "REG", "TYPE", "TIME", "STATUS", "DISTANCE", "PKR", "US$", "DUE", "BRIDGE"
]

def discover_headers(pdf_path):
    print("=" * 60)
    print(f"Discovering headers in: {os.path.basename(pdf_path)}")
    with pdfplumber.open(pdf_path) as pdf:
        p1 = pdf.pages[0]
        words = p1.extract_words()
        lines = group_words_by_y(words, y_tol=3.5)
        
        found = False
        for idx, line in enumerate(lines):
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
            
            # Check if this line qualifies as a table header
            matches = sum(1 for k in header_keywords if k in line_str.upper())
            if len(chunks) >= 3 and matches >= 2:
                print(f"Discovered Table Header at line {idx}: {len(chunks)} columns")
                for c_idx, c in enumerate(chunks):
                    txt = " ".join(w['text'] for w in c)
                    print(f"   Col {c_idx}: '{txt}' ({c[0]['x0']:.1f}-{c[-1]['x1']:.1f})")
                found = True
                break
        if not found:
            print("No table header discovered with criteria.")

for sample in [
    'tests/samples/PAA_Aeronautical_Bills_Summary.pdf',
    'tests/samples/PAA_Aviobridge_Bill.pdf',
    'tests/samples/8pages_air_navigation.pdf',
    'tests/samples/PAA_Domestic_Bill.pdf'
]:
    discover_headers(sample)
