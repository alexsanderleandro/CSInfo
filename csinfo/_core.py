# --- NOVAS FUNÇÕES DE ANÁLISE AVANÇADA ---
import socket
import json

def get_network_details(computer_name=None):
    ps = r'''
    $adapters = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }
    $out = @()
    foreach ($a in $adapters) {
        $ip = $a.IPAddress[0]
        $gw = $a.DefaultIPGateway[0]
        $dns = $a.DNSServerSearchOrder -join ", "
        $mac = $a.MACAddress
        $desc = $a.Description
        $out += [PSCustomObject]@{
            IP = $ip; Gateway = $gw; DNS = $dns; MAC = $mac; Descricao = $desc
        }
    }
    $out | ConvertTo-Json -Compress
    '''
    out = run_powershell(ps, computer_name=computer_name)
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, dict):
            parsed = [parsed]
        return parsed if parsed else []
    except Exception:
        return []

def get_firewall_status(computer_name=None):
    ps = r'''
    try {
        $profiles = Get-NetFirewallProfile
        $out = @()
        foreach ($p in $profiles) {
            $out += [PSCustomObject]@{
                Perfil = $p.Name; Ativado = $p.Enabled
            }
        }
        $out | ConvertTo-Json -Compress
    } catch { "NÃO OBTIDO" }
    '''
    out = run_powershell(ps, computer_name=computer_name)
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, dict):
            parsed = [parsed]
        return parsed if parsed else []
    except Exception:
        return []

# ... (o resto do conteúdo de csinfo.py segue inalterado) ...

from ._core_impl import *
