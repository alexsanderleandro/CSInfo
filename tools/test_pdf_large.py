import os
import tempfile
from csinfo._impl import write_pdf_report

# construir conteúdo mais longo com muitas seções para forçar múltiplas páginas
sections = [
    'INFORMAÇÕES DO SISTEMA',
    'INFORMAÇÕES DE HARDWARE',
    'INFORMAÇÕES DE REDE',
    'SEGURANÇA DO SISTEMA',
    'ADMINISTRADORES',
    'SOFTWARES INSTALADOS',
    'IDENTIFICAÇÃO'
]

lines = []
for i in range(12):
    for sec in sections:
        lines.append(sec)
        # adicionar algumas linhas de conteúdo para cada seção
        for j in range(6):
            lines.append(f"{sec} linha exemplo {i}-{j}")

out_dir = tempfile.gettempdir()
output_pdf = os.path.join(out_dir, 'csinfo_test_large_output.pdf')

print('Gerando PDF grande de teste em:', output_pdf)
res = write_pdf_report(output_pdf, lines, computer_name='TEST-PC-LARGE')
print('write_pdf_report returned:', res)

# inspecionar anotações usando pypdf
try:
    from pypdf import PdfReader
    rdr = PdfReader(output_pdf)
    print('Número de páginas:', len(rdr.pages))
    nd = getattr(rdr, 'named_destinations', None) or {}
    print('Named destinations count:', len(nd))
    # contar anotações totais
    total_ann = 0
    for i, p in enumerate(rdr.pages):
        ann = p.get('/Annots') or []
        total_ann += len(ann)
    print('Total annots in document:', total_ann)
except Exception as e:
    print('pypdf not available or error inspecting:', e)
