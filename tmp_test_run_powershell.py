import os
from csinfo._impl import run_powershell

# ativar debug
os.environ['CSINFO_DEBUG'] = '1'

out = run_powershell('Write-Output "hello from ps"', timeout=5)
print('OUT:', repr(out))

# listar logs no temp
import tempfile, glob
pattern = os.path.join(tempfile.gettempdir(), 'csinfo_debug_*.log')
logs = glob.glob(pattern)
print('LOGS:', logs)
if logs:
    with open(logs[-1], 'r', encoding='utf-8', errors='replace') as f:
        print('--- LOG CONTENT ---')
        print(f.read())
