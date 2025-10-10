import os
import sys
import tempfile

# Ajuste este caminho se você tiver um PDF em outro local
PDF_PATH = r"C:\Users\alex\AppData\Local\Temp\csinfo_test_large_output.pdf"
OUT_DIR = tempfile.gettempdir()
MAX_PAGES = 3

def try_fitz(pdf_path, out_dir):
    try:
        import fitz
    except Exception as e:
        print('PyMuPDF (fitz) não disponível:', e)
        return False, []
    created = []
    try:
        doc = fitz.open(pdf_path)
        n = min(len(doc), MAX_PAGES)
        for i in range(n):
            page = doc.load_page(i)
            mat = fitz.Matrix(2, 2)  # scale for better resolution
            pix = page.get_pixmap(matrix=mat, alpha=False)
            out_path = os.path.join(out_dir, f"csinfo_sample_page_{i+1}.png")
            pix.save(out_path)
            created.append(out_path)
        doc.close()
        return True, created
    except Exception as e:
        print('Erro ao renderizar com PyMuPDF:', e)
        return False, []


def try_pdf2image(pdf_path, out_dir):
    try:
        from pdf2image import convert_from_path
    except Exception as e:
        print('pdf2image não disponível:', e)
        return False, []
    created = []
    try:
        # convert_from_path requer poppler; assume disponível no PATH
        images = convert_from_path(pdf_path, dpi=150, first_page=1, last_page=MAX_PAGES)
        for i, img in enumerate(images):
            out_path = os.path.join(out_dir, f"csinfo_sample_page_{i+1}.png")
            img.save(out_path, 'PNG')
            created.append(out_path)
        return True, created
    except Exception as e:
        print('Erro ao renderizar com pdf2image:', e)
        return False, []


def main():
    pdf_path = PDF_PATH
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    if not os.path.exists(pdf_path):
        print('PDF não encontrado em:', pdf_path)
        sys.exit(2)
    print('Usando PDF:', pdf_path)
    ok, created = try_fitz(pdf_path, OUT_DIR)
    if not ok:
        ok, created = try_pdf2image(pdf_path, OUT_DIR)
    if not ok:
        print('\nNenhuma biblioteca de renderização disponível. Opções:')
        print('- Instalar PyMuPDF: pip install pymupdf')
        print('- Ou instalar pdf2image + poppler (e adicionar poppler ao PATH) e pillow: pip install pdf2image pillow')
        sys.exit(3)
    print('\nImagens criadas:')
    for p in created:
        print(p)
    print('\nDiretório de saída:', OUT_DIR)

if __name__ == '__main__':
    main()
