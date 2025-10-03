from csinfo import _impl
items = _impl.get_memory_modules()
print('LEN:', len(items))
for i, it in enumerate(items[:10],1):
    print(i, type(it), repr(it)[:400])
    if isinstance(it, dict):
        print('keys:', list(it.keys()))
