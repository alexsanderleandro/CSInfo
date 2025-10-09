import ast, sys
p = r'c:\Users\alex\Documents\Python\VSCode\CSInfo\csinfo\_impl.py'
try:
    s = open(p, 'r', encoding='utf-8').read()
    ast.parse(s)
    print('OK')
except SyntaxError as e:
    print('SyntaxError:', e.msg)
    print('File:', e.filename)
    print('Line:', e.lineno)
    print('Offset:', e.offset)
    # print context
    with open(p, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    start = max(0, e.lineno - 6)
    end = min(len(lines), e.lineno + 4)
    for i in range(start, end):
        prefix = '->' if i+1 == e.lineno else '  '
        print(f"{prefix} {i+1:4}: {lines[i].rstrip()}")
except Exception as e:
    print('Other error:', e)
