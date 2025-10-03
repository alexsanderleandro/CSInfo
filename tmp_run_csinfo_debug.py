import os
import tempfile
import glob
import json
import time

# Habilitar debug caso não esteja
os.environ.setdefault('CSINFO_DEBUG', '1')

print('CSINFO_DEBUG=', os.environ.get('CSINFO_DEBUG'))

try:
    import csinfo._impl as impl
except Exception as e:
    print('Erro ao importar csinfo._impl:', e)
    raise


def run_and_print(name, func, *args, **kwargs):
    print('\n' + '='*40)
    print('CALL:', name)
    start = time.time()
    try:
        res = func(*args, **kwargs)
        dur = time.time() - start
        print('DURATION:', f'{dur:.3f}s')
        print('TYPE:', type(res))
        try:
            # imprimir JSON formatado, proteger-se contra non-serializable
            print(json.dumps(res, default=str, ensure_ascii=False, indent=2))
        except Exception:
            print(repr(res))
    except Exception as e:
        import traceback
        print('EXCEPTION calling', name, e)
        traceback.print_exc()


# Chamadas de teste rápidas
run_and_print('run_powershell_echo', impl.run_powershell, 'Write-Output "CSINFO_TEST"', timeout=10)
run_and_print('memory_modules', impl.get_memory_modules_info)
run_and_print('video_cards', impl.get_video_cards_info)
run_and_print('physical_disks_short', impl.get_physical_disks_short)
run_and_print('disk_info', impl.get_disk_info)

# Listar arquivos de debug no temp
print('\n' + '='*40)
print('Procurando logs em temp...')
log_dir = tempfile.gettempdir()
patterns = ['csinfo_debug_*.log', 'csinfo_debug_fallback_*.log']
found = []
for p in patterns:
    found.extend(glob.glob(os.path.join(log_dir, p)))

if not found:
    print('Nenhum log encontrado em', log_dir)
else:
    for f in sorted(found):
        print('\n' + '#' * 60)
        print('LOG FILE:', f)
        print('#' * 60)
        try:
            with open(f, 'r', encoding='utf-8', errors='replace') as fh:
                content = fh.read()
                # Limitar tamanho para evitar floods, mas mostrar bastante
                maxlen = 200000
                if len(content) > maxlen:
                    print(content[:maxlen])
                    print('\n... (log truncated, length=', len(content), ')')
                else:
                    print(content)
        except Exception as e:
            print('Erro ao abrir log', f, e)

print('\nFim do run.')