import sys, os, pathlib
project_root = str(pathlib.Path(__file__).resolve().parents[1])
if project_root not in sys.path:
    sys.path.insert(0, project_root)
import csinfo._impl as impl

print('Running real report for CEOSOFT-059...')
res = impl.main(export_type='pdf', computer_name='CEOSOFT-059')
print('Result keys:', res.keys())
print('pdf path:', res.get('pdf'))
