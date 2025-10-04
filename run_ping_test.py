from csinfo_gui import CSInfoGUI
import time

app = CSInfoGUI()
name = 'CEOSOFT-012'
print('ping_host result:', app._ping_host(name))
# run ping worker (updates machine_list and saves file)
app._ping_worker()
print('machine_list after ping:')
for m in app.machine_list:
    print(m)
print('\nmachine_json_path:', app.machine_json_path)
# read file
try:
    with open(app.machine_json_path, 'r', encoding='utf-8') as fh:
        print('\nfile content:')
        print(fh.read())
except Exception as e:
    print('failed to read file:', e)
# close the app window properly
try:
    app.destroy()
except Exception:
    pass
