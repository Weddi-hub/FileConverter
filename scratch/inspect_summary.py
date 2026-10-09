import pdfplumber

with pdfplumber.open('tests/samples/PAA_Aeronautical_Bills_Summary.pdf') as pdf:
    p = pdf.pages[0]
    words = p.extract_words()
    top_words = [w for w in words if w['top'] <= 210]
    top_words.sort(key=lambda w: (round(w['top'], 1), w['x0']))
    curr_top = None
    for w in top_words:
        tr = round(w['top'], 1)
        if curr_top != tr:
            curr_top = tr
            print(f"\n--- top = {tr} ---")
        print(f"  {w['text']:<25} x0={w['x0']:<6.1f} x1={w['x1']:<6.1f}")
