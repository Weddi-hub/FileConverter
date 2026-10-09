import pdfplumber

with pdfplumber.open('tests/samples/PAA_Aeronautical_Bills_Summary.pdf') as pdf:
    p5 = pdf.pages[4]
    words = [w for w in p5.extract_words() if 100 <= w['top'] <= 350]
    words.sort(key=lambda w: (round(w['top'], 1), w['x0']))
    curr_top = None
    for w in words:
        tr = round(w['top'], 1)
        if curr_top != tr:
            curr_top = tr
            print(f"\n--- top = {tr} ---")
        print(f"  {w['text']:<35} x0={w['x0']:<6.1f} x1={w['x1']:<6.1f}")
