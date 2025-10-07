import sys
import os
import pathlib

# ensure project root is on sys.path so `import csinfo` works when running from tools/
project_root = str(pathlib.Path(__file__).resolve().parents[1])
if project_root not in sys.path:
    sys.path.insert(0, project_root)

import csinfo._impl as impl

lines = [
    'IDENTIFICAÇÃO',
    'Nome do computador: TEST-PC',
    '',
    'INFORMAÇÕES DO SISTEMA',
    'Versão do sistema operacional: Windows 10',
    '',
    'INFORMAÇÕES DE HARDWARE',
    'Processador: Intel',
    '',
    'INFORMAÇÕES DE REDE',
    'IP: 192.168.1.100',
]

out_path = r"C:\Users\alex\Documents\Python\VSCode\CSInfo\test_report.pdf"
ok = impl.write_pdf_report(out_path, lines, 'TEST-PC')
print('write_pdf_report returned', ok)
print('generated file exists?', os.path.exists(out_path))
