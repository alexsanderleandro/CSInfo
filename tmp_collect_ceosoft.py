import os, json, time
from datetime import datetime
# Forcamos o debug
os.environ['CSINFO_DEBUG'] = '1'

import csinfo
# Preferir o módulo _impl caso csinfo reexporte diferente
try:
    from csinfo import _impl as impl
except Exception:
    import csinfo._impl as impl

TARGET = 'ceosoft-031'
requested = ['sql_server','antivirus','memory_modules','processor','disks_short','monitors','kbd_mouse','nics','video','installed','winupdate']

def find_collector(mod, base):
    # tenta vários sufixos/padrões
    candidates = [
        f'get_{base}',
        f'get_{base}_info',
        f'get_{base}s',
        f'get_{base}es',
        f'get_{base.replace("_","")}',
    ]
    for c in candidates:
        if hasattr(mod, c):
            return c
    return None

collectors = []
for name in requested:
    fn = find_collector(impl, name)
    collectors.append((name, fn))

results = {}
for name, fn in collectors:
    if not fn:
        results[name] = {'error': 'not found'}
        continue
    f = getattr(impl, fn, None)
    if not f:
        results[name] = {'error': 'not found'}
        continue
    try:
        t0 = time.time()
        v = f(TARGET)
        dur = time.time() - t0
        results[name] = {'duration': dur, 'result': v}
    except Exception as e:
        results[name] = {'error': repr(e)}

print('=== COLLECTION RESULTS', datetime.now().isoformat(), '===')
print(json.dumps(results, indent=2, ensure_ascii=False))

# listar logs de debug gerados no temp
import glob, tempfile
temp = tempfile.gettempdir()
pattern = os.path.join(temp, 'csinfo_debug_*')
logs = sorted(glob.glob(pattern), key=os.path.getmtime, reverse=True)[:10]
print('\n=== FOUND DEBUG LOG FILES ===')
for p in logs:
    print(p)

# Print last 5 logs content (trimmed to 2000 chars each)
print('\n=== LAST LOGS CONTENT (trimmed) ===')
for p in logs[:5]:
    try:
        with open(p, 'r', encoding='utf-8', errors='replace') as fh:
            data = fh.read()
        print('\n---', p, '---')
        print(data[:2000])
    except Exception as e:
        print('Could not read', p, e)

print('\n=== DONE ===')
