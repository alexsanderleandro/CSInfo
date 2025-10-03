from teste_sistema import get_memory_modules, get_disk_info, get_network_info, get_gpu_info, get_usb_devices, get_installed_programs

print('Inspecting get_memory_modules()')
mem = get_memory_modules()
print('Type:', type(mem))
try:
    print('Len:', len(mem))
    for i, e in enumerate(mem[:10], 1):
        print(i, type(e), repr(e)[:200])
except Exception as ex:
    print('Error iterating memory result:', ex)

print('\nInspecting get_disk_info()')
discs = get_disk_info()
print(type(discs), len(discs) if hasattr(discs, '__len__') else 'no len')
for i, d in enumerate(discs[:5],1):
    print(i, type(d), repr(d)[:200])

print('\nInspecting get_network_info()')
net = get_network_info()
print(type(net))
if isinstance(net, list):
    for i, n in enumerate(net[:5],1):
        print(i, type(n), repr(n)[:200])
else:
    print('Network result repr:', repr(net)[:300])

print('\nInspecting get_gpu_info()')
print(type(get_gpu_info()))
print('\nInspecting usb')
print(type(get_usb_devices()))
print('\nInstalled programs sample')
print(type(get_installed_programs()))
