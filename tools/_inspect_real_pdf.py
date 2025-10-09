import sys, os, pathlib
from PyPDF2 import PdfReader

pdf = r'C:\Users\alex\Documents\Python\VSCode\CSInfo\Info_maquina_CEOSOFT-059.pdf'
if not os.path.exists(pdf):
    print('PDF não encontrado:', pdf)
    sys.exit(2)
reader = PdfReader(pdf)
print('page_count:', len(reader.pages))
try:
    outlines = reader.outline
except Exception:
    try:
        outlines = reader.getOutlines()
    except Exception:
        outlines = None
print('outlines:', outlines)
if len(reader.pages) > 0:
    txt = reader.pages[0].extract_text() or ''
    print('PAGE0 preview:', txt.replace('\n',' ')[:800])
