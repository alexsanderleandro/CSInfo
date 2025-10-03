from teste_sistema import get_memory_modules
mem = get_memory_modules()
print('type(mem)=', type(mem))
try:
    for i, m in enumerate(mem,1):
        print(i, type(m), repr(m)[:200])
except Exception as e:
    print('Error iterating mem:', e)
