import pdfplumber

with pdfplumber.open('tests/samples/PAA_Aeronautical_Bills_Summary.pdf') as pdf:
    p1 = pdf.pages[0]
    header_words = [w for w in p1.extract_words() if 130 <= w['top'] <= 160]
    header_words.sort(key=lambda w: w['x0'])
    for w in header_words:
        print(f"{w['text']:<25} x0={w['x0']:<6.1f} x1={w['x1']:<6.1f} top={w['top']:<6.1f}")
