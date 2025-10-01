import csinfo
lines = [
    'CSInfo - Resumo Técnico do Dispositivo',
    'Nome do computador: AUTO_TEST',
    'Tipo: Desktop',
    '',
    'INFORMAÇÕES DO SISTEMA',
    'Versão do sistema operacional: TestOS 1.0',
    '',
    'INFORMAÇÕES DE HARDWARE',
    'Memória RAM total: 8 GB',
    '',
    'CSInfo by CEOsoftware'
]

txt = 'auto_test_report.txt'
pdf = 'auto_test_report.pdf'
csinfo.write_report(txt, lines)
ok = csinfo.write_pdf_report(pdf, lines, 'AUTO_TEST')
print('WRITTEN', txt, pdf, ok)
