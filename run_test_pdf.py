import csinfo._impl as impl
lines = [
    'IDENTIFICAÇÃO',
    'Nome do computador: TEST-PC',
    '',
    'INFORMAÇÕES DO SISTEMA',
    'Versão do sistema operacional: Windows 10',
    '',
    'INFORMAÇÕES DE HARDWARE',
    'Processador: Intel',
    '',
    'INFORMAÇÕES DE REDE',
    'IP: 192.168.1.100',
]
ok = impl.write_pdf_report('C:\\Users\\alex\\Documents\\Python\\VSCode\\CSInfo\\test_report.pdf', lines, 'TEST-PC')
print('write_pdf_report returned', ok)
