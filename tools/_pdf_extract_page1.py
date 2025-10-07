import sys
import os
from PyPDF2 import PdfReader

p = r'C:\Users\alex\Documents\Python\VSCode\CSInfo\test_report.pdf'
if not os.path.exists(p):
    print('PDF not found:', p)
    sys.exit(2)
reader = PdfReader(p)
text = ''
if len(reader.pages) > 0:
    try:
        text = reader.pages[0].extract_text() or ''
    except Exception:
        text = ''
print('PAGE 1 CHARS:', len(text))
preview = text.replace('\n', ' ')[:800]
print('PREVIEW:', preview)
norm = preview.lower()
print("contains 'índice'?:", 'índice' in norm or 'indice' in norm)
