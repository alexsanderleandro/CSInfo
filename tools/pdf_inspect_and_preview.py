from pathlib import Path
import fitz
from PyPDF2 import PdfReader

p = Path(r'C:\Users\alex\Documents\Python\VSCode\CSInfo\Info_maquina_CEOSOFT-059.pdf')
if not p.exists():
    print('PDF missing', p)
    raise SystemExit(1)

print('PDF size', p.stat().st_size)

# extrair texto por página com PyPDF2
reader = PdfReader(str(p))
for i,page in enumerate(reader.pages[:5]):
    t = page.extract_text() or ''
    print('--- PAGE', i+1, 'chars', len(t))
    print(t[:400].replace('\n',' '))
    print()

# gerar imagens das primeiras 2 páginas com PyMuPDF
doc = fitz.open(str(p))
for i in range(min(2, doc.page_count)):
    page = doc.load_page(i)
    mat = fitz.Matrix(2,2)  # 2x zoom
    pix = page.get_pixmap(matrix=mat, alpha=False)
    out = Path(f'.\\pdf_preview_page_{i+1}.png')
    pix.save(str(out))
    print('Wrote', out, 'size', out.stat().st_size)
