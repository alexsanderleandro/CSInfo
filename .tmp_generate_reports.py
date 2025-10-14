# Test script to generate TXT and PDF using csinfo backend
import os, sys
sys.path.insert(0, os.getcwd())
import csinfo

# ensure front-end metadata is propagated
try:
    import version
    csinfo.__version__ = getattr(version, '__version__', None)
except Exception:
    pass
try:
    csinfo.__logo_path__ = os.path.join(os.getcwd(), 'assets', 'ico.png')
except Exception:
    pass
try:
    csinfo.__app_name__ = 'CSInfo'
except Exception:
    pass

base_cwd = os.getcwd()
pdf_folder = os.path.join(base_cwd, 'Relatorio', 'PDF')
txt_folder = os.path.join(base_cwd, 'Relatorio', 'TXT')
os.makedirs(pdf_folder, exist_ok=True)
os.makedirs(txt_folder, exist_ok=True)

sample_lines = [
    'Relatório de teste',
    'Nome do computador: TEST-MACHINE',
    'Versão do sistema: TestOS 1.0',
    'Hardware: exemplo',
]

base = 'test_header'
ptxt = os.path.join(txt_folder, base + '.txt')
ppdf = os.path.join(pdf_folder, base + '.pdf')

print('Writing TXT ->', ptxt)
try:
    csinfo.write_report(ptxt, sample_lines)
    print('TXT written')
except Exception as e:
    print('TXT error', e)

print('Writing PDF ->', ppdf)
try:
    # some backends accept (path, lines, computer_name)
    try:
        csinfo.write_pdf_report(ppdf, sample_lines, 'TEST-MACHINE')
    except TypeError:
        csinfo.write_pdf_report(ppdf, sample_lines)
    print('PDF written')
except Exception as e:
    print('PDF error', e)

print('done')
