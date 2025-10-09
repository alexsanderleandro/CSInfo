from PyPDF2 import PdfReader
import os, sys, pathlib
project_root = str(pathlib.Path(__file__).resolve().parents[1])
pdf = os.path.join(project_root, 'test_report.pdf')
if not os.path.exists(pdf):
    print('no pdf')
    sys.exit(1)
reader = PdfReader(pdf)
try:
    outlines = reader.outline
except Exception:
    try:
        outlines = reader.getOutlines()
    except Exception:
        outlines = None
print('outlines:', outlines)
