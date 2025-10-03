# Importações
from csinfo._impl import get_processor_info, run_powershell
import json
import platform
from datetime import datetime

# Cores para o terminal
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

# Funções auxiliares
def formatar_tamanho(tamanho_bytes):
    """Formata o tamanho em bytes para uma representação legível."""
    if tamanho_bytes is None or tamanho_bytes == 0:
        return "N/A"
    try:
        tamanho_bytes = float(tamanho_bytes)
        for unidade in ['B', 'KB', 'MB', 'GB', 'TB']:
            if tamanho_bytes < 1024.0 or unidade == 'TB':
                return f"{tamanho_bytes:.2f} {unidade}"
            tamanho_bytes /= 1024.0
    except (ValueError, TypeError):
        pass
    return "N/A"

def get_system_info():
    """Obtém informações básicas do sistema"""
    return {
        "Sistema Operacional": f"{platform.system()} {platform.release()}",
        "Versão": platform.version(),
        "Arquitetura": platform.architecture()[0],
        "Processador": platform.processor(),
        "Máquina": platform.machine(),
        "Node": platform.node(),
        "Data/Hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

# Funções de exibição
def exibir_cabecalho(titulo):
    """Exibe um cabeçalho formatado"""
    print(f"\n{Cores.HEADER}{Cores.NEGRITO}=== {titulo.upper()} ==={Cores.FIM}")

def exibir_chave_valor(chave, valor, nivel=0):
    """Exibe um par chave-valor formatado"""
    indent = "  " * nivel
    if isinstance(valor, dict):
        print(f"{indent}{Cores.AZUL}{chave}:{Cores.FIM}")
        for k, v in valor.items():
            exibir_chave_valor(k, v, nivel + 1)
    elif isinstance(valor, list):
        print(f"{indent}{Cores.AZUL}{chave}:{Cores.FIM}")
        for i, item in enumerate(valor, 1):
            if isinstance(item, dict):
                print(f"{indent}  {Cores.VERDE}Item {i}:{Cores.FIM}")
                for k, v in item.items():
                    exibir_chave_valor(k, v, nivel + 2)
            else:
                print(f"{indent}  - {item}")
    else:
        if isinstance(valor, (int, float)):
            if any(x in chave.lower() for x in ['size', 'tamanho', 'bytes', 'gb', 'mb', 'kb']):
                valor = formatar_tamanho(valor)
        print(f"{indent}{Cores.AZUL}{chave}:{Cores.FIM} {Cores.CIANO}{valor}{Cores.FIM}")

# Funções principais
def exibir_info_sistema():
    exibir_cabecalho("Informações do Sistema")
    for chave, valor in get_system_info().items():
        exibir_chave_valor(chave, valor)

def exibir_info_processador():
    exibir_cabecalho("Informações do Processador")
    try:
        processadores = get_processor_info()
        if not processadores:
            print(f"{Cores.AMARELO}Nenhum processador encontrado.{Cores.FIM}")
            return
            
        print(f"Total de processadores: {Cores.VERDE}{len(processadores)}{Cores.FIM}")
        for i, proc in enumerate(processadores, 1):
            print(f"\n{Cores.VERDE}Processador {i}:{Cores.FIM}")
            for chave, valor in proc.items():
                exibir_chave_valor(chave, valor, 1)
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao obter informações do processador: {e}{Cores.FIM}")

def get_memory_modules(computer_name=None):
    """Obtém informações detalhadas sobre os módulos de memória"""
    ps_script = """
    $memoryModules = @()
    try {
        $memoryModules = Get-CimInstance -ClassName Win32_PhysicalMemory | ForEach-Object {
            [PSCustomObject]@{
                Tamanho = $_.Capacity
                Bancada = $_.BankLabel
                Tipo = switch ($_.MemoryType) {
                    20 { "DDR" }
                    21 { "DDR2" }
                    24 { "DDR3" }
                    26 { "DDR4" }
                    default { "Desconhecido" }
                }
                Velocidade = if ($_.Speed) { "$($_.Speed) MHz" } else { "N/A" }
                Fabricante = $_.Manufacturer
                NumeroSerie = $_.SerialNumber
                PartNumber = $_.PartNumber
                Posicao = $_.DeviceLocator
            }
        } | ConvertTo-Json -Depth 5 -Compress
        $memoryModules
    } catch {
        Write-Error $_.Exception.Message
        "[]"
    }
    """
    
    try:
        result = run_powershell(ps_script, computer_name=computer_name)
        if result and result.strip():
            parsed = json.loads(result)
            if isinstance(parsed, dict):
                parsed = [parsed]
            return parsed
        return []
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao processar informações dos módulos de memória: {e}{Cores.FIM}")
        return []

def exibir_modulos_memoria():
    exibir_cabecalho("Módulos de Memória")
    try:
        modulos = get_memory_modules()
        if not modulos:
            print(f"{Cores.AMARELO}Nenhum módulo de memória encontrado.{Cores.FIM}")
            return
            
        print(f"Total de módulos: {Cores.VERDE}{len(modulos)}{Cores.FIM}")
        for i, modulo in enumerate(modulos, 1):
            print(f"\n{Cores.VERDE}Módulo {i}:{Cores.FIM}")
            for chave, valor in modulo.items():
                if chave == "Tamanho" and valor:
                    print(f"  {chave}: {formatar_tamanho(valor)}")
                else:
                    print(f"  {chave}: {valor or 'N/A'}")
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao exibir informações de memória: {e}{Cores.FIM}")

def get_disk_info(computer_name=None):
    """Obtém informações dos discos"""
    ps_script = """
    $disks = @()
    try {
        $disks = Get-PhysicalDisk | ForEach-Object {
            $sizeGB = [math]::Round($_.Size/1GB, 2)
            [PSCustomObject]@{
                Modelo = $_.FriendlyName
                TipoMidia = $_.MediaType
                TamanhoGB = $sizeGB
                TamanhoBytes = $_.Size
                Saude = $_.HealthStatus
                Status = $_.OperationalStatus
                NumeroSerie = $_.SerialNumber
                Interface = $_.BusType
                Firmware = $_.FirmwareVersion
            }
        } | ConvertTo-Json -Depth 5 -Compress
        $disks
    } catch {
        Write-Error $_.Exception.Message
        "[]"
    }
    """
    
    try:
        result = run_powershell(ps_script, computer_name=computer_name)
        if result and result.strip():
            parsed = json.loads(result)
            if isinstance(parsed, dict):
                parsed = [parsed]
            return parsed
        return []
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao processar informações dos discos: {e}{Cores.FIM}")
        return []

def exibir_info_discos():
    exibir_cabecalho("Informações dos Discos")
    try:
        discos = get_disk_info()
        if not discos:
            print(f"{Cores.AMARELO}Nenhum disco encontrado.{Cores.FIM}")
            return
            
        print(f"Total de discos: {Cores.VERDE}{len(discos)}{Cores.FIM}")
        for i, disco in enumerate(discos, 1):
            print(f"\n{Cores.VERDE}Disco {i}:{Cores.FIM}")
            for chave, valor in disco.items():
                if chave in ['TamanhoBytes', 'TamanhoGB']:
                    print(f"  {chave}: {formatar_tamanho(valor)}")
                else:
                    print(f"  {chave}: {valor or 'N/A'}")
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao exibir informações dos discos: {e}{Cores.FIM}")

def get_network_info(computer_name=None):
    """Obtém informações de rede"""
    ps_script = """
    $adapters = @()
    try {
        $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | ForEach-Object {
            $ipConfig = Get-NetIPConfiguration -InterfaceIndex $_.ifIndex -ErrorAction SilentlyContinue
            [PSCustomObject]@{
                Nome = $_.Name
                Descricao = $_.InterfaceDescription
                Status = $_.Status
                Velocidade = if ($_.LinkSpeed) { $_.LinkSpeed } else { "Desconhecido" }
                MacAddress = $_.MacAddress
                IPs = @(if ($ipConfig.IPv4Address.IPAddress) { $ipConfig.IPv4Address.IPAddress })
                Gateway = if ($ipConfig.IPv4DefaultGateway) { $ipConfig.IPv4DefaultGateway.NextHop } else { "Nenhum" }
                DNS = @(if ($ipConfig.DNSServer.ServerAddresses) { $ipConfig.DNSServer.ServerAddresses })
                DHCP = if ($ipConfig.NetIPv4Interface.Dhcp -eq 'Enabled') { "Ativado" } else { "Desativado" }
            }
        } | ConvertTo-Json -Depth 5 -Compress
        $adapters
    } catch {
        Write-Error $_.Exception.Message
        "[]"
    }
    """
    
    try:
        result = run_powershell(ps_script, computer_name=computer_name)
        if result and result.strip():
            parsed = json.loads(result)
            if isinstance(parsed, dict):
                parsed = [parsed]
            return parsed
        return []
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao processar informações de rede: {e}{Cores.FIM}")
        return []

def exibir_info_rede():
    exibir_cabecalho("Informações de Rede")
    try:
        adapters = get_network_info()
        if not adapters:
            print(f"{Cores.AMARELO}Nenhum adaptador de rede ativo encontrado.{Cores.FIM}")
            return
            
        print(f"Total de adaptadores ativos: {Cores.VERDE}{len(adapters)}{Cores.FIM}")
        for i, adapter in enumerate(adapters, 1):
            print(f"\n{Cores.VERDE}Adaptador {i}: {adapter.get('Nome', 'Desconhecido')}{Cores.FIM}")
            for chave, valor in adapter.items():
                if chave != 'Nome':
                    if isinstance(valor, list):
                        print(f"  {chave}:")
                        for item in valor:
                            print(f"    - {item}")
                    else:
                        print(f"  {chave}: {valor or 'N/A'}")
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao exibir informações de rede: {e}{Cores.FIM}")

def get_gpu_info(computer_name=None):
    """Obtém informações sobre placas de vídeo"""
    ps_script = """
    $gpus = @()
    try {
        $gpus = Get-CimInstance -ClassName Win32_VideoController | ForEach-Object {
            [PSCustomObject]@{
                Nome = $_.Name
                Descricao = $_.VideoProcessor
                RAM = $_.AdapterRAM
                Resolucao = "$($_.CurrentHorizontalResolution)x$($_.CurrentVerticalResolution)"
                DriverVersao = $_.DriverVersion
                DriverData = $_.DriverDate
                Status = $_.Status
            }
        } | ConvertTo-Json -Depth 5 -Compress
        $gpus
    } catch {
        Write-Error $_.Exception.Message
        "[]"
    }
    """
    
    try:
        result = run_powershell(ps_script, computer_name=computer_name)
        if result and result.strip():
            return json.loads(result)
        return []
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao processar informações das placas de vídeo: {e}{Cores.FIM}")
        return []

def exibir_placas_video():
    exibir_cabecalho("Placas de Vídeo")
    try:
        placas = get_gpu_info()
        if not placas:
            print(f"{Cores.AMARELO}Nenhuma placa de vídeo encontrada.{Cores.FIM}")
            return
            
        print(f"Total de placas: {Cores.VERDE}{len(placas)}{Cores.FIM}")
        for i, placa in enumerate(placas, 1):
            print(f"\n{Cores.VERDE}Placa {i}:{Cores.FIM}")
            for chave, valor in placa.items():
                if chave == "RAM" and valor:
                    print(f"  {chave}: {formatar_tamanho(valor)}")
                else:
                    print(f"  {chave}: {valor or 'N/A'}")
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao exibir informações das placas de vídeo: {e}{Cores.FIM}")

def get_usb_devices(computer_name=None):
    """Obtém informações sobre dispositivos USB"""
    ps_script = """
    $devices = @()
    try {
        $devices = Get-PnpDevice -Class USB | Where-Object { $_.Status -eq 'OK' } | ForEach-Object {
            [PSCustomObject]@{
                Nome = $_.FriendlyName
                Descricao = $_.DeviceDesc
                Status = $_.Status
                Classe = $_.Class
                Fabricante = $_.Manufacturer
                ID = $_.DeviceID
            }
        } | ConvertTo-Json -Depth 5 -Compress
        $devices
    } catch {
        Write-Error $_.Exception.Message
        "[]"
    }
    """
    
    try:
        result = run_powershell(ps_script, computer_name=computer_name)
        if result and result.strip():
            return json.loads(result)
        return []
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao processar dispositivos USB: {e}{Cores.FIM}")
        return []

def exibir_dispositivos_usb():
    exibir_cabecalho("Dispositivos USB")
    try:
        dispositivos = get_usb_devices()
        if not dispositivos:
            print(f"{Cores.AMARELO}Nenhum dispositivo USB encontrado.{Cores.FIM}")
            return
            
        print(f"Total de dispositivos: {Cores.VERDE}{len(dispositivos)}{Cores.FIM}")
        for i, dispositivo in enumerate(dispositivos, 1):
            print(f"\n{Cores.VERDE}Dispositivo {i}:{Cores.FIM}")
            for chave, valor in dispositivo.items():
                print(f"  {chave}: {valor or 'N/A'}")
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao exibir dispositivos USB: {e}{Cores.FIM}")

def get_installed_programs(computer_name=None, limite=50):
    """Obtém a lista de programas instalados"""
    ps_script = f"""
    $programs = @()
    try {{
        # Programas de 64 bits
        $programs += Get-ItemProperty "HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*" | 
            Where-Object {{ $_.DisplayName -and -not $_.SystemComponent }} | 
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, EstimatedSize,
                        UninstallString, InstallLocation, URLInfoAbout
        
        # Programas de 32 bits (se for sistema 64 bits)
        if ([System.Environment]::Is64BitOperatingSystem) {{
            $programs += Get-ItemProperty "HKLM:\\Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*" | 
                Where-Object {{ $_.DisplayName -and -not $_.SystemComponent }} | 
                Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, EstimatedSize,
                            UninstallString, InstallLocation, URLInfoAbout
        }}
        
        # Formata a saída
        $programs | Sort-Object DisplayName | Select-Object -First {limite} | ForEach-Object {{
            [PSCustomObject]@{{
                Nome = $_.DisplayName
                Versao = if ($_.DisplayVersion) {{ $_.DisplayVersion }} else {{ "N/A" }}
                Publicador = if ($_.Publisher) {{ $_.Publisher }} else {{ "N/A" }}
                DataInstalacao = if ($_.InstallDate) {{ $_.InstallDate }} else {{ "N/A" }}
                TamanhoMB = if ($_.EstimatedSize) {{ [math]::Round($_.EstimatedSize, 2) }} else {{ "N/A" }}
                Local = if ($_.InstallLocation) {{ $_.InstallLocation }} else {{ "N/A" }}
            }}
        }} | ConvertTo-Json -Depth 5 -Compress
    }} catch {{
        Write-Error $_.Exception.Message
        "[]"
    }}
    """
    
    try:
        result = run_powershell(ps_script, computer_name=computer_name)
        if result and result.strip():
            return json.loads(result)
        return []
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao obter programas instalados: {e}{Cores.FIM}")
        return []

def exibir_programas_instalados(limite=50):
    exibir_cabecalho("Programas Instalados")
    try:
        programas = get_installed_programs(limite=limite)
        if not programas:
            print(f"{Cores.AMARELO}Não foi possível obter a lista de programas.{Cores.FIM}")
            return
            
        print(f"Total de programas: {Cores.VERDE}{len(programas)}{Cores.FIM} (limitado a {limite})")
        for i, programa in enumerate(programas, 1):
            print(f"\n{Cores.VERDE}{i}. {programa.get('Nome', 'Desconhecido')}{Cores.FIM}")
            for chave, valor in programa.items():
                if chave != 'Nome' and valor and valor != 'N/A':
                    print(f"   {chave}: {valor}")
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao exibir programas instalados: {e}{Cores.FIM}")

def get_windows_update_status(computer_name=None):
    """Obtém o status das atualizações do Windows"""
    ps_script = """
    $updateStatus = @{}
    try {
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $updates = $searcher.GetTotalHistoryCount()
        $lastUpdate = $searcher.GetTotalHistory(1) | Select-Object -First 1
        
        $updateStatus = @{
            "Atualizações Instaladas" = $updates
            "Última Atualização" = if ($lastUpdate) { $lastUpdate.Date } else { "Nunca" }
            "Status" = if ($lastUpdate) { $lastUpdate.Operation } else { "Nunca atualizado" }
        }
        $updateStatus | ConvertTo-Json -Depth 5 -Compress
    } catch {
        Write-Error $_.Exception.Message
        "{}"
    }
    """
    
    try:
        result = run_powershell(ps_script, computer_name=computer_name)
        if result and result.strip():
            return json.loads(result)
        return {}
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao verificar atualizações do Windows: {e}{Cores.FIM}")
        return {}

def exibir_info_atualizacoes():
    exibir_cabecalho("Atualizações do Windows")
    try:
        atualizacoes = get_windows_update_status()
        if not atualizacoes:
            print(f"{Cores.AMARELO}Não foi possível verificar as atualizações do Windows.{Cores.FIM}")
            return
            
        for chave, valor in atualizacoes.items():
            print(f"{chave}: {valor}")
    except Exception as e:
        print(f"{Cores.VERMELHO}Erro ao exibir informações de atualizações: {e}{Cores.FIM}")

# Ponto de entrada principal
if __name__ == "__main__":
    try:
        exibir_info_sistema()
        exibir_info_processador()
        exibir_modulos_memoria()
        exibir_placas_video()
        exibir_info_discos()
        exibir_info_rede()
        exibir_dispositivos_usb()
        exibir_info_atualizacoes()
        exibir_programas_instalados(50)
    except Exception as e:
        print(f"\n{Cores.VERMELHO}Erro durante a execução: {e}{Cores.FIM}")
    finally:
        input("\nPressione Enter para sair...")