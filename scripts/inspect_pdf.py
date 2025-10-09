import sys
from pathlib import Path
p = Path('test_report.pdf')
if not p.exists():
    print('no file')
    sys.exit(2)
b = p.read_bytes()
for key in [b'sec_', b'detail_sec', b'sec_IDENTIFICACAO']:
    idx = b.find(key)
    print(key.decode(), 'found at', idx)
# try to print nearby context for first occurrence
for key in [b'detail_sec', b'sec_']:
    i = b.find(key)
    if i!=-1:
        start = max(0, i-40)
        print('context for', key.decode(), b'...', b[start:start+120])

# print names in /Names or /Dests if present
s = b
for token in [b'/Dests', b'/Names', b'/Outline', b'/Dest']:
    i = s.find(token)
    print(token.decode(), 'at', i)

print('size', len(b))
