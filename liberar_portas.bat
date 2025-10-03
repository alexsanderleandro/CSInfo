@echo off
REM libera_portas_firewall.bat
REM Executar como Administrador

REM Verifica se está executando como administrador
NET SESSION >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
  echo Este script precisa ser executado como Administrador.
  echo Clique com o botao direito e escolha "Executar como administrador".
  pause
  exit /B 1
)

echo Removendo regras antigas (se existirem)...
netsh advfirewall firewall delete rule name="Allow WinRM HTTP 5985" protocol=TCP localport=5985 >nul 2>&1
netsh advfirewall firewall delete rule name="Allow WinRM HTTPS 5986" protocol=TCP localport=5986 >nul 2>&1
netsh advfirewall firewall delete rule name="Allow SMB 445" protocol=TCP localport=445 >nul 2>&1
netsh advfirewall firewall delete rule name="Allow RPC Endpoint Mapper 135" protocol=TCP localport=135 >nul 2>&1

echo Criando novas regras de firewall...
netsh advfirewall firewall add rule name="Allow WinRM HTTP 5985" dir=in action=allow protocol=TCP localport=5985 profile=domain,private description="Permite WinRM HTTP (porta 5985)"
netsh advfirewall firewall add rule name="Allow WinRM HTTPS 5986" dir=in action=allow protocol=TCP localport=5986 profile=domain,private description="Permite WinRM HTTPS (porta 5986)"
netsh advfirewall firewall add rule name="Allow SMB 445" dir=in action=allow protocol=TCP localport=445 profile=domain,private description="Permite SMB (porta 445)"
netsh advfirewall firewall add rule name="Allow RPC Endpoint Mapper 135" dir=in action=allow protocol=TCP localport=135 profile=domain,private description="Permite WMI/RPC (porta 135)"

echo Regras adicionadas com sucesso.
echo.

REM Habilitar e iniciar WinRM (se desejar)
echo Configurando WinRM (servico)...
sc config WinRM start= auto >nul 2>&1
sc start WinRM >nul 2>&1
REM winrm quickconfig com -q para rodar sem prompts
winrm quickconfig -q >nul 2>&1

echo WinRM configurado (se disponivel no sistema).
echo.

echo Operacao concluida.
pause
exit /B 0
