import os
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

def test_file(path):
    print("=" * 60)
    print(f"Testing: {os.path.basename(path)}")
    with pdfplumber.open(path) as pdf:
        p1 = pdf.pages[0]
        words = p1.extract_words()
        lines = group_words(words, y_tol=3.5)
        # Look for table header candidates
        candidates = []
        for idx, line in enumerate(lines):
            line_str = " ".join(w['text'] for w in line)
            # Find lines with >= 3 distinct spaced chunks
            chunks = []
            curr = []
            for w in line:
                if not curr or (w['x0'] - curr[-1]['x1'] < 15.0):
                    curr.append(w)
                else:
                    chunks.append(curr)
                    curr = [w]
            if curr:
                chunks.append(curr)
            if len(chunks) >= 4:
                candidates.append((idx, len(chunks), line_str, chunks))
        
        for idx, n_chunks, line_str, chunks in candidates[:3]:
            print(f"Line {idx} ({n_chunks} chunks): {line_str[:90]}...")
            for i, c in enumerate(chunks):
                txt = " ".join(w['text'] for w in c)
                print(f"   Col {i}: '{txt}' [{c[0]['x0']:.1f}-{c[-1]['x1']:.1f}]")

for sample in [
    'tests/samples/PAA_Aeronautical_Bills_Summary.pdf',
    'tests/samples/PAA_Aviobridge_Bill.pdf',
    'tests/samples/8pages_air_navigation.pdf'
]:
    test_file(sample)
