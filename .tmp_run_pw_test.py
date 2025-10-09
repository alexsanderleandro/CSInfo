from csinfo._impl import run_powershell

# comando que imprime texto UTF-8 e depois um byte 0xA2 para simular saída com encoding diferente
cmd = 'Write-Output "á é ç"; [Console]::Out.Write([char]0xA2)'
print('running...')
try:
    out = run_powershell(cmd)
    print('---OUTPUT---')
    print(out)
    print('---END---')
except Exception as e:
    print('ERROR', e)
