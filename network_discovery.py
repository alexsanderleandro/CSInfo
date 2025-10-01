import socket
import os
import subprocess
import re
import getpass

def get_logged_user(machine):
    # Usa PowerShell ao invés de wmic
    try:
        ps = f"Get-WmiObject -Class Win32_ComputerSystem -ComputerName '{machine}' | Select-Object -ExpandProperty UserName"
        result = subprocess.check_output(["powershell", "-Command", ps], encoding='utf-8', errors='ignore')
        user = result.strip()
        return user if user else 'Desconhecido'
    except Exception:
        return 'Desconhecido'

def get_local_ip():
    try:
        # Obtém o nome da máquina
        hostname = socket.gethostname()
        # Obtém o endereço IP associado ao nome da máquina
        local_ip = socket.gethostbyname(hostname)
        return local_ip
    except Exception:
        return 'Desconhecido'

def get_machine_info(machine):
    try:
        local_ip = get_local_ip()
        local_name = socket.gethostname()
        local_user = getpass.getuser()
        if machine == local_name or machine == local_ip:
            user = local_user
            print(f"Usuário logado localmente: {user}")
            return (machine, user)
        user = get_logged_user(machine)
        if user == 'Desconhecido':
            print(f"Não foi possível resolver ou acessar a máquina '{machine}'. Verifique o nome informado e as permissões de rede.")
        return (machine, user)
    except Exception:
        print(f"Não foi possível resolver ou acessar a máquina '{machine}'. Verifique o nome informado e as permissões de rede.")
        return (machine, 'Desconhecido')

if __name__ == "__main__":
    machine = input("Informe o nome da máquina para buscar: ").strip()
    info = get_machine_info(machine)
    print(f"{info[0]} - {info[1]}")
