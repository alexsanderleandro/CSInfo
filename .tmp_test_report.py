import os, sys
sys.path.insert(0, os.getcwd())
import csinfo
from datetime import datetime
p_txt = os.path.join(os.getcwd(), 'Relatorio', 'TXT', 'test_header.txt')
pdf_dir = os.path.join(os.getcwd(), 'Relatorio', 'PDF')
os.makedirs(os.path.dirname(p_txt), exist_ok=True)
os.makedirs(pdf_dir, exist_ok=True)
lines = ['Linha1', 'Linha2']
try:
    csinfo.__version__ = getattr(csinfo, '__version__', '0.0.0')
    csinfo.__logo_path__ = os.path.join(os.getcwd(), 'assets', 'ico.png')
    csinfo.__app_name__ = 'CSInfo'
    # write txt
    try:
        csinfo.write_report(p_txt, lines)
        print('TXT written:', p_txt)
    except Exception as e:
        print('TXT failed:', e)
    # attempt pdf write if available
    try:
        p_pdf = os.path.join(pdf_dir, 'test_header.pdf')
        if hasattr(csinfo, 'write_pdf_report'):
            csinfo.write_pdf_report(p_pdf, lines, 'MACHINE')
            print('PDF written:', p_pdf)
        else:
            print('write_pdf_report not available')
    except Exception as e:
        print('PDF failed:', e)
except Exception as e:
    print('Overall fail', e)
