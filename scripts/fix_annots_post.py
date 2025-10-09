"""Tenta corrigir anotações /Dest no PDF substituindo referências PageN por referências indiretas reais.
Uso: python scripts/fix_annots_post.py test_report.pdf
"""
import sys, os
p = sys.argv[1] if len(sys.argv) > 1 else 'test_report.pdf'
print('processing', p)
# try pypdf first
try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import NameObject, ArrayObject
    rdr = PdfReader(p)
    writer = PdfWriter()
    # build map of page indirect ref -> page index
    page_refs = [getattr(pg, 'indirect_ref', None) for pg in rdr.pages]
    print('page refs ids:', [getattr(r, 'idnum', None) for r in page_refs])
    changed = 0
    for i, pg in enumerate(rdr.pages):
        annots = pg.get('/Annots')
        if not annots:
            writer.add_page(pg)
            continue
        # annots is a list of indirect references
        for a in annots:
            try:
                obj = a.get_object()
            except Exception:
                obj = a
            dest = obj.get('/Dest')
            cont = obj.get('/Contents')
            rect = obj.get('/Rect')
            if dest:
                print('page', i+1, 'annot contents:', cont, 'rect:', rect, 'dest:', dest)
                try:
                    first = dest[0]
                    # if first is name like 'Page1' (string), try to map by stripping 'Page' prefix
                    if hasattr(first, 'rstrip') and isinstance(first, str) and first.startswith('Page'):
                        n = int(first.replace('Page',''))
                        target_ref = page_refs[n-1]
                        if target_ref:
                            newdest = ArrayObject([target_ref, NameObject('/Fit')])
                            obj[NameObject('/Dest')] = newdest
                            changed += 1
                except Exception as e:
                    pass
        writer.add_page(pg)
    out = p + '.fixed'
    with open(out, 'wb') as f:
        writer.write(f)
    print('changed', changed, 'annotations')
    os.replace(out, p)
    print('replaced')
except Exception as e:
    print('pypdf failed:', e)
    try:
        import PyPDF2
        from PyPDF2 import PdfReader as P2Reader, PdfWriter as P2Writer
        from PyPDF2.generic import NameObject as P2NameObject, ArrayObject as P2ArrayObject
        rdr = P2Reader(p)
        writer = P2Writer()
        page_refs = [getattr(pg, 'indirect_ref', None) for pg in rdr.pages]
        changed = 0
        for i, pg in enumerate(rdr.pages):
            annots = pg.get('/Annots')
            if not annots:
                writer.add_page(pg)
                continue
            for a in annots:
                try:
                    obj = a.get_object()
                except Exception:
                    obj = a
                dest = obj.get('/Dest')
                cont = obj.get('/Contents')
                rect = obj.get('/Rect')
                if dest:
                    print('page', i+1, 'annot contents:', cont, 'rect:', rect, 'dest:', dest)
                    try:
                        first = dest[0]
                        if hasattr(first, 'rstrip') and isinstance(first, str) and first.startswith('Page'):
                            n = int(first.replace('Page',''))
                            target_ref = page_refs[n-1]
                            if target_ref:
                                newdest = P2ArrayObject([target_ref, P2NameObject('/Fit')])
                                obj[P2NameObject('/Dest')] = newdest
                                changed += 1
                    except Exception:
                        pass
        out = p + '.fixed'
        with open(out, 'wb') as f:
            writer.write(f)
        os.replace(out, p)
        print('changed', changed, 'annotations, replaced')
    except Exception as e2:
        print('PyPDF2 failed too:', e2)
        sys.exit(2)
