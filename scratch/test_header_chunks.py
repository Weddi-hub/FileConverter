import os
import re
import pdfplumber

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

def discover_table_columns(header_lines, page_width):
    # Combine words across header lines and group into column clusters
    all_header_words = []
    for line in header_lines:
        all_header_words.extend(line)
    
    # Also inspect data rows immediately below the header to see data column centers!
    # That gives us even stronger confirmation of column locations.
    pass

with pdfplumber.open('tests/samples/PAA_Aeronautical_Bills_Summary.pdf') as pdf:
    p1 = pdf.pages[0]
    words_p1 = p1.extract_words()
    lines_p1 = group_words(words_p1, y_tol=4.0)
    
    # Header row is index 6
    hl = lines_p1[6]
    # Look at word gaps in the header line:
    chunks = []
    curr_chunk = []
    for w in hl:
        if not curr_chunk:
            curr_chunk.append(w)
        else:
            gap = w['x0'] - curr_chunk[-1]['x1']
            if gap < 20.0:  # words in same column title
                curr_chunk.append(w)
            else:
                chunks.append(curr_chunk)
                curr_chunk = [w]
    if curr_chunk:
        chunks.append(curr_chunk)
        
    print(f"Header chunks found: {len(chunks)}")
    for i, c in enumerate(chunks):
        txt = " ".join(w['text'] for w in c)
        print(f"  Col {i}: '{txt}' (x0={c[0]['x0']:.1f}, x1={c[-1]['x1']:.1f})")
