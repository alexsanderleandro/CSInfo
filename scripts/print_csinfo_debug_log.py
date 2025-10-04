import os, sys
# garantir que o diretório do projeto (pai de scripts/) esteja no sys.path
proj_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if proj_root not in sys.path:
    sys.path.insert(0, proj_root)
import csinfo
p = None
try:
    p = csinfo.get_debug_session_log()
except Exception as e:
    print('get_debug_session_log() error:', e)
if not p:
    try:
        p = getattr(csinfo.run_powershell, '_csinfo_session_log', None)
    except Exception:
        p = None
print('DEBUG LOG PATH:', p)
if p and os.path.exists(p):
    try:
        with open(p, 'r', encoding='utf-8', errors='replace') as f:
            data = f.read()
            print('\n---- LAST 8000 CHARS OF DEBUG LOG ----\n')
            print(data[-8000:])
    except Exception as e:
        print('error reading log:', e)
else:
    print('No debug session log found or path invalid')
