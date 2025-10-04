from csinfo._impl import get_processor_info, run_powershell
import json
import platform
from datetime import datetime

class Cores:
    HEADER = '\033[95m'
    AZUL = '\033[94m'
    CIANO = '\033[96m'
    VERDE = '\033[92m'
    AMARELO = '\033[93m'
    VERMELHO = '\033[91m'
    NEGRITO = '\033[1m'
    SUBLINHADO = '\033[4m'
    FIM = '\033[0m'

def formatar_tamanho(tamanho_bytes):
    """Formata um número de bytes para uma string legível (KB, MB, GB, ...)."""
    try:
        n = float(tamanho_bytes)
    except Exception:
        return str(tamanho_bytes)
    for unit in ['bytes', 'KB', 'MB', 'GB', 'TB', 'PB']:
        if n < 1024.0 or unit == 'PB':
            if unit == 'bytes':
                return f"{int(n)} {unit}"
            return f"{n:.2f} {unit}"
        n /= 1024.0

def get_antivirus_info():
    """Obtém informações sobre o antivírus instalado"""
    ps_script = """
    $result = @{
        "Antivírus" = @()
    }
    
    try {
        $antivirus = Get-CimInstance -Namespace "root\\SecurityCenter2" -ClassName "AntivirusProduct" -ErrorAction SilentlyContinue
        
        if ($antivirus) {
            $antivirusList = @()
            $antivirus | ForEach-Object {
                $status = switch ($_.productState) {
                    { $_ -ge 0x1000 } { "Atualizado"; break }
                    default { "Desatualizado" }
                }
                
                $antivirusList += @{
                    Nome = $_.displayName
                    Status = $status
                }
            }
            $result["Antivírus"] = $antivirusList
        }
    }
    catch {
        Write-Debug "Erro ao obter informações do antivírus: $_"
    }
    
    return $result | ConvertTo-Json -Depth 5 -Compress
    """
    
    try:
        result = run_powershell(ps_script)
        if result and result.strip() and result.strip() != '{}':
            try:
                return json.loads(result)
            except UnicodeDecodeError:
                return json.loads(result.encode('latin1').decode('utf-8', errors='ignore'))
            except json.JSONDecodeError as e:
                print(f"{Cores.VERMELHO}Erro ao decodificar informações do antivírus: {e}{Cores.FIM}")
        return {}
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao obter informações do antivírus: {e}{Cores.FIM}")
        return {}

def get_monitor_info():
    """Obtém informações sobre os monitores conectados"""
    ps_script = """
    $ErrorActionPreference = 'Stop'
    $monitors = @()
    
    try {
        # Obtém informações básicas dos monitores
        $displayDevices = Get-CimInstance -Namespace "root/wmi" -ClassName WmiMonitorBasicDisplayParams -ErrorAction SilentlyContinue
        $displayIDs = Get-CimInstance -Namespace "root/wmi" -ClassName WmiMonitorID -ErrorAction SilentlyContinue
        $displayDescriptions = Get-CimInstance -Namespace "root/wmi" -ClassName WmiMonitorDescriptorMethods -ErrorAction SilentlyContinue
        $displaySettings = Get-CimInstance -Namespace "root/wmi" -ClassName WmiMonitorListedSupportedSourceModes -ErrorAction SilentlyContinue
        
        # Obtém a resolução atual de cada monitor
        Add-Type -AssemblyName System.Windows.Forms
        $screens = [System.Windows.Forms.Screen]::AllScreens
        
        # ... (rest of the monitor info function)
        
        return @{ Monitores = $monitors } | ConvertTo-Json -Depth 5 -Compress
    }
    catch {
        Write-Debug "Erro ao obter informações dos monitores: $_"
        return @{ Monitores = @() } | ConvertTo-Json -Compress
    }
    """
    
    try:
        result = run_powershell(ps_script)
        if result and result.strip() and result.strip() != '{}':
            try:
                return json.loads(result)
            except UnicodeDecodeError:
                return json.loads(result.encode('latin1').decode('utf-8', errors='ignore'))
            except json.JSONDecodeError as e:
                print(f"{Cores.VERMELHO}Erro ao decodificar informações dos monitores: {e}{Cores.FIM}")
        return {}
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao obter informações dos monitores: {e}{Cores.FIM}")
        return {}

def get_sqlserver_info():
    """Obtém informações sobre as instâncias do SQL Server instaladas"""
    ps_script = """
    $result = @{
        "Versoes" = @()
        "Instancias" = @()
        "Servicos" = @()
    }
    
    try {
        # Mapeamento de versões do SQL Server
        $sqlVersionMap = @{
            # ... (existing version mapping)
        }
        
        # Mapeamento de builds para versões e atualizações
        $sqlBuildMap = @{
            # ... (existing build mapping)
        }
        
        # Obtém versões instaladas
        $uninstallKeys = @(
            "HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*",
            "HKLM:\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*"
        )
        
        $items = Get-ItemProperty -Path $uninstallKeys -ErrorAction SilentlyContinue | 
                 Where-Object { $_.DisplayName -match "SQL Server" -or $_.DisplayName -match "Microsoft SQL" }
        
        foreach ($item in $items) {
            $version = (Get-ItemProperty -Path $item.PSPath -Name "DisplayVersion" -ErrorAction SilentlyContinue).DisplayVersion
            $build = (Get-ItemProperty -Path $item.PSPath -Name "Version" -ErrorAction SilentlyContinue).Version
            
            # Determina a versão do SQL Server
            $sqlVersion = "Desconhecida"
            $versionParts = $version -split '\\.'
            
            # ... (rest of the SQL Server function)
            
        }
        
        # ... (rest of the SQL Server function)
        
        return $result | ConvertTo-Json -Depth 10 -Compress
    }
    catch {
        Write-Debug "Erro ao obter informações do SQL Server: $_"
        return $result | ConvertTo-Json -Compress
    }
    """
    
    try:
        result = run_powershell(ps_script)
        if result and result.strip() and result.strip() != '{}':
            try:
                return json.loads(result)
            except UnicodeDecodeError:
                return json.loads(result.encode('latin1').decode('utf-8', errors='ignore'))
            except json.JSONDecodeError as e:
                print(f"{Cores.VERMELHO}Erro ao decodificar informações do SQL Server: {e}{Cores.FIM}")
        return {}
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao obter informações do SQL Server: {e}{Cores.FIM}")
        return {}
