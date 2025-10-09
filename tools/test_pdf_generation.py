import os
import tempfile
from csinfo._impl import write_pdf_report

# construir conteúdo mínimo com as seções que queremos testar
lines = [
    "INFORMAÇÕES DO SISTEMA",
    "Versão do sistema operacional: TestOS 1.0",
    "INFORMAÇÕES DE HARDWARE",
    "Processador: TestCPU",
    "ADMINISTRADORES",
    "Usuario: admin",
]

out_dir = tempfile.gettempdir()
output_pdf = os.path.join(out_dir, 'csinfo_test_output.pdf')

print('Gerando PDF de teste em:', output_pdf)
res = write_pdf_report(output_pdf, lines, computer_name='TEST-PC')
print('write_pdf_report returned:', res)

# inspecionar anotações usando pypdf
try:
    from pypdf import PdfReader
    rdr = PdfReader(output_pdf)
    print('Número de páginas:', len(rdr.pages))
    nd = getattr(rdr, 'named_destinations', None) or {}
    print('Named destinations found:', list(nd.keys())[:50])
    # listar anotações por página
    for i, p in enumerate(rdr.pages):
        ann = p.get('/Annots') or []
        print(f'Page {i+1} annots:', len(ann))
        for a in ann:
            try:
                obj = a.get_object()
                print(' - subtype:', obj.get('/Subtype'), 'rect:', obj.get('/Rect'))
            except Exception:
                pass
except Exception as e:
    print('pypdf not available or error inspecting:', e)
