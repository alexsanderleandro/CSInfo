from csinfo._impl import get_memory_modules_info
import json
items = get_memory_modules_info()
print('LEN:', len(items))
for i, it in enumerate(items[:3],1):
    print(i, type(it), repr(it)[:200])
    try:
        print('keys:', list(it.keys()))
    except Exception as e:
        print('not dict:', e)
