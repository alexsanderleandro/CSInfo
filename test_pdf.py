import traceback
import csinfo._impl as impl
import os

print('START')
try:
    # include_debug_on_export=True ajuda a anexar logs internos caso algo falhe
    resultado = impl.main(export_type='pdf', include_debug_on_export=True)
    print('DONE', resultado)
    # listar arquivo PDF gerado mais recente
    cwd = os.getcwd()
    pdfs = [f for f in os.listdir(cwd) if f.lower().startswith('info_maquina_') and f.lower().endswith('.pdf')]
    if pdfs:
        pdfs_sorted = sorted(pdfs, key=lambda n: os.path.getmtime(os.path.join(cwd,n)), reverse=True)
        recent = pdfs_sorted[0]
        path = os.path.join(cwd, recent)
        size = os.path.getsize(path)
        print('PDF:', path, 'SIZE:', size)
    else:
        print('PDF: none')
except Exception:
    traceback.print_exc()
