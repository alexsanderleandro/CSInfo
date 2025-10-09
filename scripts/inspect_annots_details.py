"""Inspeciona anotações no PDF e resolve /Dest para o índice de página alvo.
Uso: python scripts/inspect_annots_details.py test_report.pdf
"""
import sys
p = sys.argv[1] if len(sys.argv) > 1 else 'test_report.pdf'
try:
    from pypdf import PdfReader
    rdr = PdfReader(p)
    backend = 'pypdf'
except Exception:
    try:
        import PyPDF2
        rdr = PyPDF2.PdfReader(p)
        backend = 'PyPDF2'
    except Exception as e:
        print('Nenhuma biblioteca PDF disponível:', e)
        sys.exit(2)

print('using backend:', backend)
# build a map of page object idnum -> page number
page_map = {}
for i, pg in enumerate(rdr.pages, start=1):
    ir = getattr(pg, 'indirect_ref', None) or getattr(pg, 'indirectReference', None) or getattr(pg, 'indirect', None)
    idnum = None
    try:
        idnum = getattr(ir, 'idnum', None) or getattr(ir, 'id', None)
    except Exception:
        idnum = None
    page_map[idnum] = i

for pnum, pg in enumerate(rdr.pages, start=1):
    annots = None
    try:
        annots = pg.get('/Annots') if backend == 'PyPDF2' else pg.get('/Annots')
    except Exception:
        try:
            annots = pg.get('/Annots')
        except Exception:
            annots = None
    if not annots:
        continue
    for i, a in enumerate(annots, start=1):
        try:
            obj = a.get_object()
        except Exception:
            try:
                obj = a
            except Exception:
                obj = None
        keys = list(obj.keys()) if obj else []
        print(f'Page {pnum} ann[{i}]: keys={keys}')
        dest = obj.get('/Dest') if obj else None
        if dest:
            first = dest[0]
            idnum = getattr(first, 'idnum', None) or getattr(first, 'id', None)
            print('  /Dest first idnum:', idnum)
            target_page = page_map.get(idnum)
            print('  -> target page index:', target_page)
        else:
            print('  no /Dest')

print('page_map snapshot:', page_map)
