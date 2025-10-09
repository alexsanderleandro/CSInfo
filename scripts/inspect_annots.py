"""
Inspeciona anotações de link em test_report.pdf e imprime um resumo
Tenta usar pypdf/PyPDF2 se disponível; se não, faz uma inspeção binária básica.
"""
import sys
from pathlib import Path
p = Path('test_report.pdf')
if not p.exists():
    print('no file test_report.pdf')
    sys.exit(2)

# tentar pypdf / PyPDF2
reader = None
try:
    from pypdf import PdfReader
    reader = PdfReader(str(p))
    backend = 'pypdf'
except Exception:
    try:
        import PyPDF2
        reader = PyPDF2.PdfReader(str(p))
        backend = 'PyPDF2'
    except Exception:
        reader = None
        backend = None

print('using backend:', backend)

if reader is not None:
    annots_found = 0
    for i, page in enumerate(reader.pages, start=1):
        ann = page.get('/Annots') or page.get('/ANNOTS') or page.get('/annots')
        if not ann:
            continue
        if isinstance(ann, list):
            ann_list = ann
        else:
            try:
                ann_list = list(ann)
            except Exception:
                ann_list = [ann]
        for j, a in enumerate(ann_list, start=1):
            annots_found += 1
            try:
                # resolve indirect objects if backend supports
                if hasattr(a, 'get_object'):
                    aobj = a.get_object()
                else:
                    aobj = a
                keys = list(aobj.keys())
                print(f'Page {i} ann[{j}]: keys={keys}')
                # Dest
                if '/Dest' in aobj:
                    dest = aobj['/Dest']
                    print('  /Dest ->', repr(dest))
                # Action
                if '/A' in aobj:
                    action = aobj['/A']
                    print('  /A ->', action)
                    if isinstance(action, dict) and '/D' in action:
                        print('   /A /D ->', action.get('/D'))
                # Subtype
                st = aobj.get('/Subtype')
                if st:
                    print('  /Subtype =', st)
            except Exception as e:
                print('  error reading annot', e)
    if annots_found == 0:
        print('No annotations found via PDF library')
    else:
        print('Total annotations inspected:', annots_found)
else:
    # fallback: inspeção binária simples
    b = p.read_bytes()
    for token in [b'/Subtype /Link', b'/Dest', b'/A', b'/Annots']:
        print(token.decode(), 'count', b.count(token))
    print('size', len(b))

