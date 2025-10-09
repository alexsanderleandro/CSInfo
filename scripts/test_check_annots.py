"""Teste automatizado mínimo para validar as anotações produzidas pelo gerador.

- Gera o PDF chamando run_test_pdf.py
- Abre test_report.pdf com pypdf
- Verifica que existem anotações de /Subtype /Link com /A {'/S':'/GoTo','/D': <nome>} e que cada nome aparece em reader.named_destinations

Retorna código 0 em sucesso, 2 em falha de verificação, 1 em erro inesperado.
"""
import subprocess
import sys
import os
import json

PDF_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'test_report.pdf'))
RUNNER = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'run_test_pdf.py'))

print('Gerando PDF via runner...')
res = subprocess.run([sys.executable, RUNNER], capture_output=True, text=True)
print(res.stdout)
if res.returncode != 0:
    print('Runner falhou:', res.returncode)
    print(res.stderr)
    sys.exit(1)

if not os.path.exists(PDF_PATH):
    print('PDF não foi gerado:', PDF_PATH)
    sys.exit(1)

try:
    from pypdf import PdfReader
    from pypdf.generic import NameObject, TextStringObject
except Exception as exc:
    print('pypdf ausente ou erro ao importar:', exc)
    sys.exit(1)

rdr = PdfReader(PDF_PATH)
print('Páginas: ', len(rdr.pages))

annots_found = 0
missing_named = []
used_names = set()

for pi, pg in enumerate(rdr.pages, start=1):
    annots = pg.get('/Annots') or []
    if not annots:
        continue
    for a in annots:
        try:
            obj = a.get_object()
        except Exception:
            obj = a
        st = obj.get('/Subtype')
        if st != '/Link':
            continue
        annots_found += 1
        A = obj.get('/A')
        if not A:
            print(f'Page {pi} ann has no /A action')
            missing_named.append((pi, None))
            continue
        S = A.get('/S')
        D = A.get('/D')
        if S != '/GoTo':
            print(f'Page {pi} ann has action not GoTo: {S}')
            missing_named.append((pi, D))
            continue
        # D may be NameObject('/name') or TextStringObject('name')
        if isinstance(D, NameObject):
            name = str(D)
        else:
            name = str(D)
        # normalize leading slash if present
        if name.startswith('/'):
            name = name[1:]
        used_names.add(name)

# confirm named destinations exist
named = getattr(rdr, 'named_destinations', {}) or {}

for n in used_names:
    if n not in named:
        missing_named.append(('missing_named', n))

print('Anotações link encontradas:', annots_found)
print('Destinos nomeados usados nas anotações:', used_names)
print('Destinos nomeados no documento:', set(named.keys()))

if not annots_found:
    print('ERRO: nenhuma anotação de link encontrada')
    sys.exit(2)

if missing_named:
    print('ERRO: destinos nomeados ausentes ou problemas detectados:')
    print(missing_named)
    sys.exit(2)

print('OK: anotações com /A GoTo e destinos nomeados detectados')
sys.exit(0)
