from PIL import Image
import os

SRC = os.path.join('assets', 'ico.png')
DST = os.path.join('assets', 'app.ico')

try:
    img = Image.open(SRC)
    sizes = [(256,256),(128,128),(64,64),(48,48),(32,32),(16,16)]
    img.save(DST, sizes=sizes)
    print(f'Wrote {DST} with sizes {sizes}')
except Exception as e:
    print('ERROR:', e)
    raise
