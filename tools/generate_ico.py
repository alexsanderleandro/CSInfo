"""Pequeno utilitário para gerar assets/app.ico a partir de assets/ico.png usando Pillow.
Uso: python tools/generate_ico.py
"""
from PIL import Image
import os

ROOT = os.path.dirname(os.path.dirname(__file__))
PNG = os.path.join(ROOT, 'assets', 'ico.png')
ICO = os.path.join(ROOT, 'assets', 'app.ico')

if not os.path.exists(PNG):
    print('Arquivo PNG não encontrado em', PNG)
    raise SystemExit(1)

os.makedirs(os.path.dirname(ICO), exist_ok=True)
img = Image.open(PNG)
if img.mode not in ('RGBA', 'RGB'):
    img = img.convert('RGBA')
sizes = [(16,16),(32,32),(48,48),(256,256)]
img.save(ICO, format='ICO', sizes=sizes)
print('Gerado', ICO)
