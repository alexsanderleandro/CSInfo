from teste_sistema import get_memory_modules
import json
res = get_memory_modules()
print('TYPE:', type(res), 'LEN:', len(res))
for i, r in enumerate(res[:10],1):
    print(i, type(r), repr(r)[:200])
    if isinstance(r, dict):
        print('keys:', list(r.keys()))
