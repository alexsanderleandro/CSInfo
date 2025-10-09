import os
p = r'C:\Users\alex\Documents\Python\VSCode\CSInfo\\.tmp_inspect_pypdf.py'
print('checking', p)
print('exists?', os.path.exists(p))
if os.path.exists(p):
    try:
        os.remove(p)
        print('removed')
    except Exception as e:
        print('remove error', e)
else:
    print('no file to remove')
