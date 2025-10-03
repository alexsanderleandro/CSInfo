"""
Diagnóstico remoto leve para investigar falhas de coleta de hardware.

Usage:
    python tmp_remote_diag.py <target>

Exemplo:
    python tmp_remote_diag.py ceosoft-031

O script executa várias chamadas PowerShell via `run_powershell` (do pacote `csinfo`)
para reproduzir os tipos de comandos que o coletor usa, e também testa portas de
rede relevantes (WinRM/WSMan e WMI/SMB). Gera um arquivo de log único em %TEMP%.

Quando executar na sua estação contra a máquina remota (ceosoft-031), envie o
arquivo de log gerado para que eu analise e proponha a causa/extrato de correções.
"""
import sys
import os
import socket
import tempfile
import time
from datetime import datetime
from csinfo._impl import run_powershell


def tcp_check(host, port, timeout=3):
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True, None
    except Exception as e:
        return False, str(e)


def add_section(f, title, content):
    f.write('\n' + ('=' * 80) + '\n')
    f.write(title + '\n')
    f.write(('=' * 80) + '\n')
    f.write(content + '\n')


def main():
    if len(sys.argv) < 2:
        print('Usage: python tmp_remote_diag.py <target>')
        sys.exit(1)
    target = sys.argv[1]
    debug_env = os.getenv('CSINFO_DEBUG', '')

    ts = datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')
    fn = f'csinfo_remote_diag_{target}_{ts}.log'
    path = os.path.join(tempfile.gettempdir(), fn)

    commands = [
        ("Check-WSMan (Test-WSMan)", f"Test-WSMan -ComputerName {target} -ErrorAction SilentlyContinue"),
        ("Invoke-Command (echo)", f"Invoke-Command -ComputerName {target} -ScriptBlock {{ Write-Output \"CSINFO_INVOKE_OK\" }} -ErrorAction SilentlyContinue"),
        ("Get-CimInstance ComputerSystem", f"Get-CimInstance Win32_ComputerSystem -ComputerName {target} | Select-Object -Property Manufacturer,Model,TotalPhysicalMemory | ConvertTo-Json -Compress"),
        ("Get-CimInstance Processor", f"Get-CimInstance Win32_Processor -ComputerName {target} | Select-Object -Property Name,NumberOfCores,NumberOfLogicalProcessors | ConvertTo-Json -Compress"),
        ("Get-CimInstance BIOS", f"Get-CimInstance Win32_BIOS -ComputerName {target} | Select-Object -Property Manufacturer,SMBIOSBIOSVersion,ReleaseDate | ConvertTo-Json -Compress"),
        ("Get-WmiObject ComputerSystem (legacy)", f"Get-WmiObject Win32_ComputerSystem -ComputerName {target} | Select-Object Manufacturer,Model,TotalPhysicalMemory | ConvertTo-Json -Compress"),
        ("Disk drives (Get-CimInstance Win32_DiskDrive)", f"Get-CimInstance Win32_DiskDrive -ComputerName {target} | Select-Object Model,Size,InterfaceType | ConvertTo-Json -Compress"),
        ("Get-BitLockerVolume", f"Try {{ Get-BitLockerVolume -ComputerName {target} | ConvertTo-Json -Compress }} Catch {{ 'EXC' }}"),
    ]

    ports = [5985, 5986, 135, 445]

    with open(path, 'w', encoding='utf-8', errors='replace') as f:
        f.write('CSInfo remote diagnostic log\n')
        f.write(f'TARGET: {target}\n')
        f.write(f'TIMESTAMP_UTC: {datetime.utcnow().isoformat()}Z\n')
        f.write(f'CSINFO_DEBUG env: {debug_env}\n')
        f.write('\n')

        # TCP checks
        add_section(f, 'TCP PORT CHECKS', '')
        for p in ports:
            ok, err = tcp_check(target, p, timeout=4)
            f.write(f'Port {p}: {"OPEN" if ok else "CLOSED"} - {err or ""}\n')

        # Short pause
        f.write('\n')
        add_section(f, 'POWERSHELL COMMANDS (via run_powershell)', '')
        for (label, cmd) in commands:
            f.write('\n')
            f.write('--- ' + label + ' ---\n')
            f.write('COMMAND: ' + cmd + '\n')
            f.write('START: ' + datetime.utcnow().isoformat() + 'Z\n')
            try:
                out = run_powershell(cmd, timeout=15, computer_name=None)
                # Note: we pass the computer name inside the command because run_powershell wraps
                # with Invoke-Command if computer_name param is used; here the cmd already uses -ComputerName
                if not out:
                    f.write('OUTPUT: (empty)\n')
                else:
                    # truncate long outputs but keep useful portion
                    out_trim = out.strip()
                    if len(out_trim) > 20000:
                        f.write('OUTPUT (truncated 20k):\n')
                        f.write(out_trim[:20000])
                        f.write('\n...TRUNCATED...\n')
                    else:
                        f.write('OUTPUT:\n')
                        f.write(out_trim + '\n')
            except Exception as e:
                f.write('EXCEPTION: ' + repr(e) + '\n')
            f.write('END: ' + datetime.utcnow().isoformat() + 'Z\n')

        # Extra: try calling run_powershell using the computer_name parameter (Invoke-Command wrapper)
        add_section(f, 'POWERSHELL (Invoke-Command wrapper via computer_name param)', '')
        f.write('Trying simple echo via run_powershell(computer_name=target)\n')
        f.write('START: ' + datetime.utcnow().isoformat() + 'Z\n')
        try:
            out2 = run_powershell('Write-Output "CSINFO_WRAPPER_OK"', timeout=15, computer_name=target)
            f.write('OUTPUT: ' + (out2 or '(empty)') + '\n')
        except Exception as e:
            f.write('EXCEPTION: ' + repr(e) + '\n')
        f.write('END: ' + datetime.utcnow().isoformat() + 'Z\n')

    print('Diagnostic log written to:', path)
    print('\nSummary (head):\n')
    with open(path, 'r', encoding='utf-8', errors='replace') as f:
        lines = f.readlines()
        print(''.join(lines[:200]))


if __name__ == '__main__':
    main()
