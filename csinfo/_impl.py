"""Implementação interna do pacote csinfo.

Este arquivo contém a implementação completa que antes vivia em `csinfo.py` na raiz
do projeto. Foi movida para dentro do pacote para permitir importação limpa
(`import csinfo`) e facilitar o empacotamento com PyInstaller.
"""

# Conteúdo migrado de csinfo.py
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

def get_bitlocker_status(computer_name=None):
    ps = r'''
    try {
        $vols = Get-BitLockerVolume -ErrorAction SilentlyContinue
        $out = @()
        foreach ($v in $vols) {
            $out += [PSCustomObject]@{
                Unidade = $v.VolumeLetter; Protegido = $v.ProtectionStatus; Status = $v.LockStatus
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

def get_windows_update_status(computer_name=None):
    ps = r'''
    try {
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $result = $searcher.Search("IsInstalled=0")
        $count = $result.Updates.Count
        if ($count -eq 0) { "Nenhuma atualização pendente" }
        else { "$count atualizações pendentes" }
    } catch { "NÃO OBTIDO" }
    '''
    out = run_powershell(ps, computer_name=computer_name)
    return out.strip() if out else "NÃO OBTIDO"

def get_running_processes(computer_name=None):
    ps = r'''
    Get-Process | Select-Object -First 10 Name,Id,CPU | ConvertTo-Json -Compress
    '''
    out = run_powershell(ps, computer_name=computer_name)
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, dict):
            parsed = [parsed]
        return parsed if parsed else []
    except Exception:
        return []

def get_critical_services(computer_name=None):
    ps = r'''
    $names = @("wuauserv", "WinDefend", "BITS", "Spooler", "LanmanServer", "LanmanWorkstation")
    $out = @()
    foreach ($n in $names) {
        $s = Get-Service -Name $n -ErrorAction SilentlyContinue
        if ($s) {
            $out += [PSCustomObject]@{Nome=$s.Name; Status=$s.Status}
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

def get_firewall_controller(computer_name=None):
    ps = r'''
    try {
        $namespace = "root\\SecurityCenter2"
        $firewallProducts = Get-WmiObject -Namespace $namespace -Class FirewallProduct -ErrorAction SilentlyContinue
        $out = @()
        foreach ($fw in $firewallProducts) {
            $out += $fw.displayName
        }
        if ($out.Count -gt 0) {
            $out -join ", "
        } else {
            "Windows Firewall (padrão)"
        }
    } catch {
        "NÃO OBTIDO"
    }
    '''
    out = run_powershell(ps, computer_name=computer_name)
    return out.strip() if out else "NÃO OBTIDO"

# coletar_info_maquina.py (restante do conteúdo migrado)
import subprocess
import platform
import os
import re
import sys
import winreg
from datetime import datetime
import time
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.pdfbase import pdfutils
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

def run_powershell(cmd, timeout=20, computer_name=None):
    if computer_name:
        # Adicionar -ComputerName para execução remota
        cmd = f"Invoke-Command -ComputerName {computer_name} -ScriptBlock {{ {cmd} }} -ErrorAction SilentlyContinue"
    # Forçar codificação UTF-8 no PowerShell
    cmd_with_encoding = f"[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; {cmd}"
    full = ['powershell', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', cmd_with_encoding]
    try:
        if os.name == 'nt':
            creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000)
            out = subprocess.check_output(full, stderr=subprocess.STDOUT, text=True, encoding='utf-8', timeout=timeout, creationflags=creationflags)
        else:
            out = subprocess.check_output(full, stderr=subprocess.STDOUT, text=True, encoding='utf-8', timeout=timeout)
        return out.strip()
    except subprocess.CalledProcessError:
        return ""
    except Exception:
        return ""

def safe_filename(s):
    # remove caracteres inválidos para nome de arquivo
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', s)

def get_machine_name(computer_name=None):
    if computer_name:
        return computer_name
    return platform.node() or "Desconhecido"

def get_os_version(computer_name=None):
    cmd = "(Get-CimInstance Win32_OperatingSystem | Select-Object -Property Caption,Version,OSArchitecture | ConvertTo-Json -Compress)"
    out = run_powershell(cmd, computer_name=computer_name)
    m_caption = re.search(r'"Caption"\s*:\s*"([^"]+)"', out)
    m_version = re.search(r'"Version"\s*:\s*"([^"]+)"', out)
    m_arch = re.search(r'"OSArchitecture"\s*:\s*"([^"]+)"', out)
    caption = m_caption.group(1) if m_caption else ""
    version = m_version.group(1) if m_version else ""
    arch = m_arch.group(1) if m_arch else ""
    if caption:
        result = f"{caption} (Version {version})" if version else caption
        if arch:
            result += f" - {arch}"
        return result
    return "NÃO OBTIDO"

def get_memory_info(computer_name=None):
    cmd = "(Get-CimInstance Win32_ComputerSystem | Select-Object -Property TotalPhysicalMemory | ConvertTo-Json -Compress)"
    out = run_powershell(cmd, computer_name=computer_name)
    try:
        m_memory = re.search(r'"TotalPhysicalMemory"\s*:\s*(\d+)', out)
        if m_memory:
            total_bytes = int(m_memory.group(1))
            total_gb = round(total_bytes / (1024**3), 2)
            return f"{total_gb} GB"
    except Exception:
        pass
    return "NÃO OBTIDO"

def get_disk_info(computer_name=None):
    ps = r"""
    $disks = @()
    try {
        $physicalDisks = Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue
        foreach ($disk in $physicalDisks) {
            $model = $disk.Model
            $size = [math]::Round($disk.Size / 1GB, 2)
            $mediaType = $disk.MediaType
            $interface = $disk.InterfaceType
            
            # Tentar determinar se é SSD ou HDD
            $diskType = "HDD"
            if ($mediaType -like "*SSD*" -or $model -like "*SSD*" -or $model -like "*Solid State*") {
                $diskType = "SSD"
            } elseif ($mediaType -like "*Fixed*" -or $interface -eq "SCSI") {
                # Usar SMART para detectar SSD (método alternativo)
                try {
                    $smartData = Get-WmiObject -Namespace root\wmi -Class MSStorageDriver_FailurePredictData -ErrorAction SilentlyContinue | Where-Object {$_.InstanceName -like "*$($disk.PNPDeviceID)*"}
                    if ($smartData) {
                        $diskType = "SSD"
                    }
                } catch {}
            }
            
            # Obter informações de espaço das partições associadas ao disco físico
            $totalUsed = 0
            $totalFree = 0
            $partitions = ""
            
            try {
                # Método mais direto: obter partições pelo índice do disco
                $associatedPartitions = Get-CimInstance -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='$($disk.DeviceID)'} WHERE AssocClass=Win32_DiskDriveToDiskPartition" -ErrorAction SilentlyContinue
                
                foreach ($partition in $associatedPartitions) {
                    $logicalDisks = Get-CimInstance -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($partition.DeviceID)'} WHERE AssocClass=Win32_LogicalDiskToDiskPartition" -ErrorAction SilentlyContinue
                    
                    foreach ($logicalDisk in $logicalDisks) {
                        if ($logicalDisk.Size) {
                            $driveSize = [math]::Round($logicalDisk.Size / 1GB, 2)
                            $driveFree = [math]::Round($logicalDisk.FreeSpace / 1GB, 2)
                            $driveUsed = $driveSize - $driveFree
                            $totalUsed += $driveUsed
                            $totalFree += $driveFree
                            if ($partitions) { $partitions += ", " }
                            $partitions += "$($logicalDisk.DeviceID) ($driveFree GB livre de $driveSize GB)"
                        }
                    }
                }
                
                # Se o método acima não funcionou, tentar método alternativo
                if ($totalFree -eq 0 -and $totalUsed -eq 0) {
                    $diskPartitions = Get-CimInstance Win32_DiskPartition -ErrorAction SilentlyContinue | Where-Object { $_.DiskIndex -eq $disk.Index }
                    foreach ($partition in $diskPartitions) {
                        $logicalDisks = Get-CimInstance Win32_LogicalDiskToDiskPartition -ErrorAction SilentlyContinue | Where-Object { $_.Antecedent -like "*$($partition.DeviceID)*" }
                        foreach ($logicalDiskRel in $logicalDisks) {
                            $deviceId = ($logicalDiskRel.Dependent -split '"')[1]
                            $drive = Get-CimInstance Win32_LogicalDisk -ErrorAction SilentlyContinue | Where-Object { $_.DeviceID -eq $deviceId }
                            if ($drive -and $drive.Size) {
                                $driveSize = [math]::Round($drive.Size / 1GB, 2)
                                $driveFree = [math]::Round($drive.FreeSpace / 1GB, 2)
                                $driveUsed = $driveSize - $driveFree
                                $totalUsed += $driveUsed
                                $totalFree += $driveFree
                                if ($partitions) { $partitions += ", " }
                                $partitions += "$($drive.DeviceID) ($driveFree GB livre de $driveSize GB)"
                            }
                        }
                    }
                }
            } catch {}
            
            $espacoLivre = if ($totalFree -gt 0) { "$totalFree GB" } else { "N/A" }
            $espacoUsado = if ($totalUsed -gt 0) { "$totalUsed GB" } else { "N/A" }
            $particoesInfo = if ($partitions) { $partitions } else { "N/A" }
            
            $disks += [PSCustomObject]@{
                Modelo = $model
                Tamanho = "$size GB"
                EspacoUsado = $espacoUsado
                EspacoLivre = $espacoLivre
                Particoes = $particoesInfo
                Tipo = $diskType
                Interface = $interface
            }
        }
    } catch {}
    $disks | ConvertTo-Json -Compress
    """
    out = run_powershell(ps, computer_name)
    items = []
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, dict):
            parsed = [parsed]
        for disk in parsed:
            modelo = disk.get('Modelo', 'NÃO OBTIDO')
            tamanho = disk.get('Tamanho', 'NÃO OBTIDO')
            espaco_usado = disk.get('EspacoUsado', 'NÃO OBTIDO')
            espaco_livre = disk.get('EspacoLivre', 'NÃO OBTIDO')
            particoes = disk.get('Particoes', 'NÃO OBTIDO')
            tipo = disk.get('Tipo', 'NÃO OBTIDO')
            interface = disk.get('Interface', 'NÃO OBTIDO')
            items.append((modelo, tamanho, espaco_usado, espaco_livre, particoes, tipo, interface))
    except Exception:
        pass
    return items or [("NÃO OBTIDO", "NÃO OBTIDO", "NÃO OBTIDO", "NÃO OBTIDO", "NÃO OBTIDO", "NÃO OBTIDO", "NÃO OBTIDO")]

def get_windows_activation_status(computer_name=None):
    try:
        ps_command = r'''
        try {
            # Usar slmgr diretamente
            $slmgrResult = & cscript //nologo C:\Windows\System32\slmgr.vbs /xpr
            $output = $slmgrResult -join "`n"
            
            # Verificar diferentes padrões de texto para ativação
            if ($output -match "ativada permanentemente|permanently activated|permanently|permanente") {
                "Ativado"
            } elseif ($output -match "grace period|período de carência|carência") {
                "Período de carência"
            } elseif ($output -match "notification|notificação") {
                "Período de notificação"
            } elseif ($output -match "not activated|não ativado|não está ativado") {
                "Não ativado"
            } else {
                # Se não conseguir interpretar, retornar a saída original limpa
                ($output -replace "`r", "" -replace "`n", " ").Trim()
            }
        } catch {
            try {
                # Método alternativo usando Get-WmiObject
                $licenses = Get-WmiObject -Class SoftwareLicensingProduct | Where-Object {
                    $_.Name -like "*Windows*" -and $_.PartialProductKey -ne $null
                } | Select-Object -First 1
                
                if ($licenses) {
                    switch ($licenses.LicenseStatus) {
                        1 { "Ativado" }
                        0 { "Não licenciado" }
                        2 { "Período de carência" }
                        3 { "OOT (Out of Tolerance)" }
                        4 { "OOB (Out of Box)" }
                        5 { "Notificação" }
                        6 { "Carência estendida" }
                        default { "Status desconhecido: $($licenses.LicenseStatus)" }
                    }
                } else {
                    "Nenhuma licença encontrada"
                }
            } catch {
                "Erro na verificação"
            }
        }
        '''
        
        result = run_powershell(ps_command, computer_name=computer_name)
        return result.strip() if result and result.strip() else "NÃO OBTIDO"
    except Exception as e:
        return f"ERRO: {str(e)}"

def get_office_activation_status(computer_name=None):
    try:
        ps_command = '''
        try {
            # Verificar Office através do registro
            $officeApps = @("Word", "Excel", "PowerPoint", "Outlook")
            $activated = $false
            
            foreach ($app in $officeApps) {
                try {
                    $comObject = New-Object -ComObject "$app.Application"
                    if ($comObject) {
                        $activated = $true
                        $comObject.Quit()
                        break
                    }
                } catch {}
            }
            
            if ($activated) {
                "Ativado"
            } else {
                # Método alternativo: verificar chaves de ativação no registro
                $regPaths = @(
                    "HKLM:\\SOFTWARE\\Microsoft\\Office\\*\\Registration",
                    "HKLM:\\SOFTWARE\\WOW6432Node\\Microsoft\\Office\\*\\Registration"
                )
                
                $foundActivation = $false
                foreach ($path in $regPaths) {
                    try {
                        $items = Get-ChildItem $path -ErrorAction SilentlyContinue
                        if ($items) {
                            $foundActivation = $true
                            break
                        }
                    } catch {}
                }
                
                if ($foundActivation) {
                    "Possivelmente ativado"
                } else {
                    "Não detectado"
                }
            }
        } catch {
            "NÃO OBTIDO"
        }
        '''
        
        result = run_powershell(ps_command, computer_name=computer_name)
        return result.strip() if result and result.strip() else "NÃO OBTIDO"
    except:
        return "NÃO OBTIDO"

def is_domain_computer(computer_name=None):
    try:
        ps_command = '''
        $computer = Get-CimInstance -ClassName Win32_ComputerSystem
        if ($computer.PartOfDomain) {
            "Domínio: $($computer.Domain)"
        } else {
            "Workgroup: $($computer.Workgroup)"
        }
        '''
        
        result = run_powershell(ps_command, computer_name=computer_name)
        return result.strip() if result and result.strip() else "NÃO OBTIDO"
    except:
        return "NÃO OBTIDO"

def get_office_version(computer_name=None):
    if computer_name:
        # Para máquinas remotas, usar PowerShell para verificar registry
        cmd = """
        $apps = Get-ItemProperty HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* -ErrorAction SilentlyContinue | Where-Object {$_.DisplayName -like "*Office*" -or $_.DisplayName -like "*Microsoft 365*"}
        if ($apps) { ($apps | Select-Object -First 1).DisplayName + " " + ($apps | Select-Object -First 1).DisplayVersion } else { "Não encontrado" }
        """
        return run_powershell(cmd, computer_name=computer_name) or "NÃO OBTIDO"
    
    # Código original para máquina local
    keys = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]
    results = []
    for base in (winreg.HKEY_LOCAL_MACHINE,):
        for key in keys:
            try:
                reg = winreg.OpenKey(base, key)
            except FileNotFoundError:
                continue
            i = 0
            while True:
                try:
                    sub = winreg.EnumKey(reg, i)
                except OSError:
                    break
                i += 1
                try:
                    sk = winreg.OpenKey(reg, sub)
                    try:
                        name = winreg.QueryValueEx(sk, "DisplayName")[0]
                    except OSError:
                        name = ""
                    try:
                        ver = winreg.QueryValueEx(sk, "DisplayVersion")[0]
                    except OSError:
                        ver = ""
                    if any(k in name for k in ("Office", "Microsoft 365", "Microsoft Office", "Word", "Excel")):
                        results.append(f"{name} {ver}".strip())
                except OSError:
                    pass
    if not results:
        cmd = "try { $w = New-Object -ComObject Word.Application; $v = $w.Version; $w.Quit(); $v } catch { '' }"
        out = run_powershell(cmd)
        if out:
            results.append("Word.Application version " + out.strip())
    return results[0] if results else "NÃO OBTIDO"

def get_motherboard_info(computer_name=None):
    cmd = "(Get-CimInstance Win32_BaseBoard | Select-Object Manufacturer,Product,SerialNumber | ConvertTo-Json -Compress)"
    out = run_powershell(cmd, computer_name=computer_name)
    try:
        info = json.loads(out) if out else {}
        if isinstance(info, list):
            info = info[0] if info else {}
        fabricante = info.get('Manufacturer', '')
        modelo = info.get('Product', '')
        serial = info.get('SerialNumber', '')
        return fabricante, modelo, serial
    except Exception:
        return "", "", "NÃO OBTIDO"

def get_monitor_infos(computer_name=None):
        ps = r"""
        $s = @()
        try {
            $mons = Get-WmiObject -Namespace root\wmi -Class WmiMonitorID -ErrorAction SilentlyContinue
            foreach ($m in $mons) {
                $arr = $m.SerialNumberID
                $manuf = ($m.ManufacturerName | ForEach-Object {[char]$_}) -join ''
                $model = ($m.UserFriendlyName | ForEach-Object {[char]$_}) -join ''
                $serial = ($arr | ForEach-Object {[char]$_}) -join ''
                $s += [PSCustomObject]@{Fabricante=$manuf; Modelo=$model; Serial=$serial}
            }
        } catch {}
        $s | ConvertTo-Json -Compress
        """
        out = run_powershell(ps, computer_name)
        try:
                parsed = json.loads(out) if out else []
                if isinstance(parsed, dict):
                        parsed = [parsed]
                return parsed if parsed else [{"Fabricante":"","Modelo":"","Serial":"NÃO OBTIDO"}]
        except Exception:
                return [{"Fabricante":"","Modelo":"","Serial":"NÃO OBTIDO"}]

def get_devices_by_class(devclass, computer_name=None):
    ps = r"""
    $out = @()
    try {{
        $devs = Get-PnpDevice -Class {cls} -ErrorAction SilentlyContinue
        foreach ($d in $devs) {{
            $id = $d.InstanceId
            $name = $d.FriendlyName
            if (-not $name) {{ $name = $d.Name }}
            $serial = ""
            $manuf = ""
            $model = ""
            try {{
                $prop = Get-PnpDeviceProperty -InstanceId $id -KeyName 'DEVPKEY_Device_SerialNumber' -ErrorAction SilentlyContinue
                if ($prop) {{ $serial = $prop.Data }}
                $manufprop = Get-PnpDeviceProperty -InstanceId $id -KeyName 'DEVPKEY_Device_Manufacturer' -ErrorAction SilentlyContinue
                if ($manufprop) {{ $manuf = $manufprop.Data }}
                $modelprop = Get-PnpDeviceProperty -InstanceId $id -KeyName 'DEVPKEY_Device_Model' -ErrorAction SilentlyContinue
                if ($modelprop) {{ $model = $modelprop.Data }}
            }} catch {{}}
            if (-not $serial) {{ $serial = $id }}
            $out += [PSCustomObject]@{{ Name = $name; Serial = $serial; Fabricante = $manuf; Modelo = $model }}
        }}
    }} catch {{}}
    $out | ConvertTo-Json -Compress
    """
    out = run_powershell(ps.format(cls=devclass), computer_name=computer_name)
    items = []
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, dict):
            parsed = [parsed]
        for p in parsed:
            name = p.get("Name") or ""
            serial = p.get("Serial") or ""
            fabricante = p.get("Fabricante") or ""
            modelo = p.get("Modelo") or ""
            items.append((name, serial, fabricante, modelo))
    except Exception:
        for l in out.splitlines():
            if l.strip():
                items.append((l, "", "", ""))
    return items or [("NÃO OBTIDO", "", "", "")]

def get_sql_server_info(computer_name=None):
    ps = r"""
    $sqlInstances = @()
    
    # Buscar serviços do SQL Server
    try {
        $services = Get-WmiObject -Class Win32_Service -Filter "Name LIKE '%SQL%'" -ErrorAction SilentlyContinue
        $sqlServices = $services | Where-Object { $_.Name -match "MSSQL\$" -or $_.Name -eq "MSSQLSERVER" }
        
        foreach ($service in $sqlServices) {
            $instanceName = if ($service.Name -eq "MSSQLSERVER") { "Default" } else { $service.Name -replace "MSSQL\$", "" }
            $status = $service.State
            
            # Tentar obter versão do registro
            $version = ""
            try {
                if ($instanceName -eq "Default") {
                    $regPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL*\MSSQLServer\CurrentVersion"
                } else {
                    $regPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL*\MSSQLServer\CurrentVersion"
                }
                
                $versionKeys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\" -ErrorAction SilentlyContinue | 
                               Where-Object { $_.Name -match "MSSQL\d+\." }
                
                foreach ($key in $versionKeys) {
                    $currentVersionPath = Join-Path $key.PSPath "MSSQLServer\CurrentVersion"
                    if (Test-Path $currentVersionPath) {
                        $versionReg = Get-ItemProperty $currentVersionPath -ErrorAction SilentlyContinue
                        if ($versionReg.CurrentVersion) {
                            $version = $versionReg.CurrentVersion
                            break
                        }
                    }
                }
            } catch {}
            
            $sqlInstances += [PSCustomObject]@{
                Instance = $instanceName
                Status = $status
                Version = $version
            }
        }
    } catch {}
    
    # Se não encontrou serviços, tentar pelo registro
    if ($sqlInstances.Count -eq 0) {
        try {
            $regKeys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\" -ErrorAction SilentlyContinue | 
                       Where-Object { $_.Name -match "MSSQL\d+\." }
            
            foreach ($key in $regKeys) {
                $setupPath = Join-Path $key.PSPath "Setup"
                if (Test-Path $setupPath) {
                    $setup = Get-ItemProperty $setupPath -ErrorAction SilentlyContinue
                    if ($setup.SqlProgramDir) {
                        $sqlInstances += [PSCustomObject]@{
                            Instance = $setup.Edition
                            Status = "Unknown"
                            Version = $setup.Version
                        }
                    }
                }
            }
        } catch {}
    }
    
    $sqlInstances | ConvertTo-Json -Compress
    """
    out = run_powershell(ps, computer_name)
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, dict):
            parsed = [parsed]
        return parsed if parsed else []
    except Exception:
        return []

def get_network_adapters_info(computer_name=None):
    ps = r"""
    $networkAdapters = @()
    
    try {
        # Obter adaptadores de rede ativos
        $adapters = Get-CimInstance Win32_NetworkAdapter -Filter "NetConnectionStatus=2" -ErrorAction SilentlyContinue
        
        foreach ($adapter in $adapters) {
            $name = $adapter.Name
            $manufacturer = $adapter.Manufacturer
            $speed = "N/A"
            $macAddress = $adapter.MACAddress
            
            # Tentar obter velocidade
            if ($adapter.Speed) {
                $speedMbps = [math]::Round($adapter.Speed / 1000000, 0)
                $speed = "$speedMbps Mbps"
            }
            
            # Verificar se é adaptador físico (não virtual)
            if ($adapter.PhysicalAdapter -eq $true -or $name -notlike "*Virtual*" -and $name -notlike "*Loopback*" -and $macAddress) {
                $networkAdapters += [PSCustomObject]@{
                    Name = $name
                    Manufacturer = $manufacturer
                    Speed = $speed
                    MACAddress = $macAddress
                }
            }
        }
        
        # Se não encontrou pelo método acima, tentar método alternativo
        if ($networkAdapters.Count -eq 0) {
            $netAdapters = Get-NetAdapter -Physical -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq "Up" }
            foreach ($net in $netAdapters) {
                $speed = if ($net.LinkSpeed) { 
                    $speedValue = $net.LinkSpeed
                    if ($speedValue -ge 1000000000) {
                        [math]::Round($speedValue / 1000000000, 1).ToString() + " Gbps"
                    } else {
                        [math]::Round($speedValue / 1000000, 0).ToString() + " Mbps"
                    }
                } else { "N/A" }
                
                $networkAdapters += [PSCustomObject]@{
                    Name = $net.Name
                    Manufacturer = $net.DriverProvider
                    Speed = $speed
                    MACAddress = $net.MacAddress
                }
            }
        }
    } catch {}
    
    $networkAdapters | ConvertTo-Json -Compress
    """
    out = run_powershell(ps, computer_name)
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, dict):
            parsed = [parsed]
        return parsed if parsed else []
    except Exception:
        return []

def get_processor_info(computer_name=None):
    ps = r"""
    $processorInfo = @()
    
    try {
        $processors = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue
        
        foreach ($cpu in $processors) {
            $name = $cpu.Name
            $manufacturer = $cpu.Manufacturer
            $architecture = $cpu.Architecture
            $cores = $cpu.NumberOfCores
            $logicalProcessors = $cpu.NumberOfLogicalProcessors
            $maxClockSpeed = $cpu.MaxClockSpeed
            $currentClockSpeed = $cpu.CurrentClockSpeed
            $l2CacheSize = $cpu.L2CacheSize
            $l3CacheSize = $cpu.L3CacheSize
            $socket = $cpu.SocketDesignation
            
            # Converter arquitetura para texto
            $archText = switch ($architecture) {
                0 { "x86" }
                1 { "MIPS" }
                2 { "Alpha" }
                3 { "PowerPC" }
                6 { "Intel Itanium" }
                9 { "x64" }
                default { "Desconhecida ($architecture)" }
            }
            
            # Converter velocidades de MHz para GHz
            $maxSpeed = if ($maxClockSpeed) { [math]::Round($maxClockSpeed / 1000, 2).ToString() + " GHz" } else { "N/A" }
            $currentSpeed = if ($currentClockSpeed) { [math]::Round($currentClockSpeed / 1000, 2).ToString() + " GHz" } else { "N/A" }
            
            # Converter cache para KB/MB
            $l2Cache = if ($l2CacheSize) { 
                if ($l2CacheSize -ge 1024) { [math]::Round($l2CacheSize / 1024, 2).ToString() + " MB" }
                else { $l2CacheSize.ToString() + " KB" }
            } else { "N/A" }
            
            $l3Cache = if ($l3CacheSize) { 
                if ($l3CacheSize -ge 1024) { [math]::Round($l3CacheSize / 1024, 2).ToString() + " MB" }
                else { $l3CacheSize.ToString() + " KB" }
            } else { "N/A" }
            
            $processorInfo += [PSCustomObject]@{
                Name = $name
                Manufacturer = $manufacturer
                Architecture = $archText
                Cores = $cores
                LogicalProcessors = $logicalProcessors
                MaxSpeed = $maxSpeed
                CurrentSpeed = $currentSpeed
                L2Cache = $l2Cache
                L3Cache = $l3Cache
                Socket = $socket
            }
        }
    } catch {}
    
    $processorInfo | ConvertTo-Json -Compress
    """
    out = run_powershell(ps, computer_name)
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, dict):
            parsed = [parsed]
        return parsed if parsed else []
    except Exception:
        return []

def get_memory_modules_info(computer_name=None):
    ps = r"""
    $memoryModules = @()
    
    try {
        $modules = Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue
        
        foreach ($module in $modules) {
            $manufacturer = $module.Manufacturer
            $partNumber = $module.PartNumber
            $serialNumber = $module.SerialNumber
            $capacity = $module.Capacity
            $speed = $module.Speed
            $memoryType = $module.MemoryType
            $formFactor = $module.FormFactor
            $deviceLocator = $module.DeviceLocator
            $bankLabel = $module.BankLabel
            
            # Converter capacidade para GB
            $capacityGB = if ($capacity) { [math]::Round($capacity / 1GB, 2).ToString() + " GB" } else { "N/A" }
            
            # Converter velocidade
            $speedMHz = if ($speed) { $speed.ToString() + " MHz" } else { "N/A" }
            
            # Converter tipo de memória
            $memTypeText = switch ($memoryType) {
                20 { "DDR" }
                21 { "DDR2" }
                22 { "DDR2 FB-DIMM" }
                24 { "DDR3" }
                26 { "DDR4" }
                34 { "DDR5" }
                default { 
                    # Tentar deduzir pelo speed (heurística) se $memoryType for 0, null ou desconhecido
                    if ($speed -and $speed -ge 2133) { "DDR4" }
                    elseif ($speed -and $speed -ge 1066) { "DDR3" } 
                    elseif ($speed -and $speed -ge 533) { "DDR2" }
                    elseif ($speed -and $speed -ge 200) { "DDR" }
                    else { "Desconhecido" }
                }
            }
            
            # Converter fator de forma
            $formFactorText = switch ($formFactor) {
                8 { "DIMM" }
                12 { "SO-DIMM" }
                13 { "Micro-DIMM" }
                default { "Form $formFactor" }
            }
            
            $memoryModules += [PSCustomObject]@{
                Manufacturer = if ($manufacturer) { $manufacturer.Trim() } else { "N/A" }
                PartNumber = if ($partNumber) { $partNumber.Trim() } else { "N/A" }
                SerialNumber = if ($serialNumber) { $serialNumber.Trim() } else { "N/A" }
                Capacity = $capacityGB
                Speed = $speedMHz
                MemoryType = $memTypeText
                FormFactor = $formFactorText
                Location = if ($deviceLocator) { $deviceLocator } else { $bankLabel }
            }
        }
    } catch {}
    
    $memoryModules | ConvertTo-Json -Compress
    """
    out = run_powershell(ps, computer_name=computer_name)
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, dict):
            parsed = [parsed]
        return parsed if parsed else []
    except Exception:
        return []

def get_video_cards_info(computer_name=None):
    ps = r"""
    $videoCards = @()
    
    try {
        $cards = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue
        
        foreach ($card in $cards) {
            $name = $card.Name
            $manufacturer = $card.AdapterCompatibility
            $memory = "N/A"
            $driver = $card.DriverVersion
            $type = "Desconhecido"
            
            # Determinar memória de vídeo
            if ($card.AdapterRAM -and $card.AdapterRAM -gt 0) {
                $memoryGB = [math]::Round($card.AdapterRAM / 1GB, 2)
                $memory = "$memoryGB GB"
            }
            
            # Tentar determinar se é onboard ou offboard
            if ($name -like "*Intel*" -and ($name -like "*HD Graphics*" -or $name -like "*UHD Graphics*" -or $name -like "*Iris*")) {
                $type = "Onboard (Integrada)"
            } elseif ($name -like "*AMD*" -and $name -like "*Radeon*" -and ($name -like "*Vega*" -or $name -like "*APU*")) {
                $type = "Onboard (Integrada)"
            } elseif ($name -like "*NVIDIA*" -or $name -like "*AMD Radeon RX*" -or $name -like "*GeForce*" -or $name -like "*Quadro*") {
                $type = "Offboard (Dedicada)"
            } elseif ($card.PNPDeviceID -like "*PCI\VEN_*") {
                $type = "Offboard (Dedicada)"
            } else {
                $type = "Onboard (Integrada)"
            }
            
            $videoCards += [PSCustomObject]@{
                Name = $name
                Manufacturer = $manufacturer
                Memory = $memory
                Driver = $driver
                Type = $type
            }
        }
    } catch {}
    
    $videoCards | ConvertTo-Json -Compress
    """
    out = run_powershell(ps, computer_name)
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, dict):
            parsed = [parsed]
        return parsed if parsed else []
    except Exception:
        return []

def get_logical_drives_info(computer_name=None):
    ps = r"""
    $drives = @()
    
    try {
        $logicalDrives = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue
        foreach ($drive in $logicalDrives) {
            $deviceId = $drive.DeviceID
            $size = if ($drive.Size) { [math]::Round($drive.Size / 1GB, 2) } else { 0 }
            $freeSpace = if ($drive.FreeSpace) { [math]::Round($drive.FreeSpace / 1GB, 2) } else { 0 }
            $usedSpace = $size - $freeSpace
            $fileSystem = $drive.FileSystem
            $volumeName = $drive.VolumeName
            
            $drives += [PSCustomObject]@{
                Drive = $deviceId
                Size = $size
                Used = $usedSpace
                Free = $freeSpace
                FileSystem = $fileSystem
                Label = if ($volumeName) { $volumeName } else { "Sem rótulo" }
            }
        }
        
        # Ordenar por letra da unidade
        $sortedDrives = $drives | Sort-Object Drive
        
    } catch {
        $sortedDrives = @()
    }
    
    $sortedDrives | ConvertTo-Json -Compress
    """
    out = run_powershell(ps, computer_name)
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, dict):
            parsed = [parsed]
        return parsed if parsed else []
    except Exception:
        return []

def get_admin_users(computer_name=None):
    ps = r"""
    $adminUsers = @()
    
    try {
        # Obter membros do grupo de administradores
        $adminGroup = Get-LocalGroup -Name "Administradores" -ErrorAction SilentlyContinue
        if (-not $adminGroup) {
            $adminGroup = Get-LocalGroup -Name "Administrators" -ErrorAction SilentlyContinue
        }
        
        if ($adminGroup) {
            $members = Get-LocalGroupMember -Group $adminGroup -ErrorAction SilentlyContinue
            foreach ($member in $members) {
                $name = $member.Name
                $objectClass = $member.ObjectClass
                $principalSource = $member.PrincipalSource
                
                # Remover o nome do computador/domínio do nome do usuário
                if ($name -like "*\\*") {
                    $name = ($name -split "\\")[-1]
                }
                
                $adminUsers += $name
            }
        }
        
        # Ordenar por nome em ordem crescente e remover duplicatas
        $sortedUsers = $adminUsers | Sort-Object | Get-Unique
        
    } catch {
        # Método alternativo usando WMI se o método acima falhar
        try {
            $group = Get-WmiObject -Class Win32_Group -Filter "Name='Administrators' OR Name='Administradores'" -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($group) {
                $members = Get-WmiObject -Class Win32_GroupUser -ErrorAction SilentlyContinue | Where-Object { $_.GroupComponent -like "*$($group.Name)*" }
                foreach ($member in $members) {
                    $userPath = $member.PartComponent
                    if ($userPath -match 'Name="([^"]+)"') {
                        $userName = $matches[1]
                        $sortedUsers += $userName
                    }
                }
                $sortedUsers = $sortedUsers | Sort-Object | Get-Unique
            }
        } catch {
            $sortedUsers = @()
        }
    }
    
    $sortedUsers | ConvertTo-Json -Compress
    """
    out = run_powershell(ps, computer_name=computer_name)
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, str):
            return [parsed]  # Se retornou apenas um usuário como string
        return parsed if parsed else []
    except Exception:
        return []

def get_installed_software(computer_name=None):
    ps = r"""
    $softwareList = @()
    
    try {
        $uninstallKeys = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        foreach ($key in $uninstallKeys) {
            $programs = Get-ItemProperty $key -ErrorAction SilentlyContinue
            foreach ($program in $programs) {
                $displayName = $program.DisplayName
                $version = $program.DisplayVersion
                $publisher = $program.Publisher
                
                if ($displayName -and $displayName.Trim() -ne "") {
                    # Filtrar alguns itens desnecessários
                    if ($displayName -notlike "*Update*" -and 
                        $displayName -notlike "*Hotfix*" -and
                        $displayName -notlike "*Security Update*" -and
                        $displayName -notlike "KB*" -and
                        $displayName -ne "Microsoft Visual C++ 2019 X64 Additional Runtime" -and
                        $displayName -notlike "*Redistributable*") {
                        
                        $softwareList += [PSCustomObject]@{
                            Name = $displayName
                            Version = if ($version) { $version } else { "N/A" }
                            Publisher = if ($publisher) { $publisher } else { "N/A" }
                        }
                    }
                }
            }
        }
        
        # Remover duplicatas e ordenar por nome
        $uniqueSoftware = $softwareList | Group-Object Name | ForEach-Object { $_.Group | Select-Object -First 1 }
        $sortedSoftware = $uniqueSoftware | Sort-Object Name
        
    } catch {
        $sortedSoftware = @()
    }
    
    $sortedSoftware | ConvertTo-Json -Compress
    """
    out = run_powershell(ps, computer_name)
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, dict):
            parsed = [parsed]
        return parsed if parsed else []
    except Exception:
        return []

def get_antivirus_info(computer_name=None):
    ps = r"""
    $antivirusList = @()
    
    # Método 1: Windows Security Center (funciona no Windows 10/11)
    try {
        $namespace = "root\SecurityCenter2"
        $antivirusProducts = Get-WmiObject -Namespace $namespace -Class AntiVirusProduct -ErrorAction SilentlyContinue
        
        foreach ($av in $antivirusProducts) {
            $name = $av.displayName
            $state = $av.productState
            
            # Decodificar o estado do produto
            $hex = [Convert]::ToString($state, 16).PadLeft(6, "0")
            $s2 = $hex.Substring(2,2) 
            
            $enabled = "Desconhecido"
            
            # Verificar se está ativado (bit 4 e 5)
            if ([Convert]::ToInt32($s2, 16) -band 0x10) {
                $enabled = "Ativado"
            } elseif ([Convert]::ToInt32($s2, 16) -band 0x00) {
                $enabled = "Desativado"
            }
            
            $antivirusList += [PSCustomObject]@{
                Name = $name
                Enabled = $enabled
            }
        }
    } catch {}
    
    # Método 2: Verificar pelo registro (programas instalados)
    if ($antivirusList.Count -eq 0) {
        try {
            $uninstallKeys = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
            )
            
            $antivirusNames = @(
                "Avast", "AVG", "Kaspersky", "Norton", "McAfee", "Bitdefender", 
                "ESET", "Trend Micro", "F-Secure", "Panda", "Sophos", "Malwarebytes",
                "Windows Defender", "Microsoft Defender", "Symantec", "Avira",
                "Quick Heal", "Comodo", "360 Total Security", "Baidu Antivirus"
            )
            
            foreach ($key in $uninstallKeys) {
                $programs = Get-ItemProperty $key -ErrorAction SilentlyContinue
                foreach ($program in $programs) {
                    $displayName = $program.DisplayName
                    if ($displayName) {
                        foreach ($avName in $antivirusNames) {
                            if ($displayName -like "*$avName*") {
                                $antivirusList += [PSCustomObject]@{
                                    Name = $displayName
                                    Enabled = "Instalado"
                                }
                                break
                            }
                        }
                    }
                }
            }
        } catch {}
    }
    
    # Método 3: Verificar Windows Defender especificamente
    if ($antivirusList.Count -eq 0) {
        try {
            $defenderStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
            if ($defenderStatus) {
                $enabled = if ($defenderStatus.RealTimeProtectionEnabled) { "Ativado" } else { "Desativado" }
                
                $antivirusList += [PSCustomObject]@{
                    Name = "Windows Defender"
                    Enabled = $enabled
                }
            }
        } catch {}
    }
    
    # Se nada foi encontrado, verificar se há pelo menos o Windows Defender básico
    if ($antivirusList.Count -eq 0) {
        try {
            $defenderService = Get-Service -Name "WinDefend" -ErrorAction SilentlyContinue
            if ($defenderService) {
                $status = if ($defenderService.Status -eq "Running") { "Em execução" } else { $defenderService.Status }
                $antivirusList += [PSCustomObject]@{
                    Name = "Windows Defender (Serviço)"
                    Enabled = $status
                }
            }
        } catch {}
    }
    
    # Remover duplicatas baseado no nome
    $uniqueList = @()
    $seenNames = @()
    foreach ($av in $antivirusList) {
        $cleanName = $av.Name -replace '\s*\d+.*$', '' -replace '\s*\(.*\)', '' -replace '\s+', ' '
        $cleanName = $cleanName.Trim()
        
        if ($seenNames -notcontains $cleanName) {
            $seenNames += $cleanName
            $uniqueList += [PSCustomObject]@{
                Name = $cleanName
                Enabled = $av.Enabled
            }
        }
    }
    
    $uniqueList | ConvertTo-Json -Compress
    """
    out = run_powershell(ps, computer_name)
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, dict):
            parsed = [parsed]
        return parsed if parsed else []
    except Exception:
        return []

def get_keyboard_mouse_status(computer_name=None):
    ps = r"""
    $hasKeyboard = $false
    $hasMouse = $false
    
    try {
        $keyboards = Get-WmiObject -Class Win32_Keyboard -ErrorAction SilentlyContinue
        if ($keyboards) { $hasKeyboard = $true }
        
        $mice = Get-WmiObject -Class Win32_PointingDevice -ErrorAction SilentlyContinue
        if ($mice) { $hasMouse = $true }
    } catch {}
    
    [PSCustomObject]@{
        HasKeyboard = $hasKeyboard
        HasMouse = $hasMouse
    } | ConvertTo-Json -Compress
    """
    out = run_powershell(ps, computer_name)
    try:
        parsed = json.loads(out) if out else {}
        return parsed.get('HasKeyboard', False), parsed.get('HasMouse', False)
    except Exception:
        return False, False

def get_printers(computer_name=None):
        ps = r"""
        $o = @()
        try {
            $pr = Get-CimInstance Win32_Printer -ErrorAction SilentlyContinue
            foreach ($p in $pr) {
                $name = $p.Name
                $pnp = $p.PNPDeviceID
                $serial = ""
                $manuf = $p.Manufacturer
                $model = $p.DriverName
                if ($pnp) {
                    try {
                        $prop = Get-PnpDeviceProperty -InstanceId $pnp -KeyName 'DEVPKEY_Device_SerialNumber' -ErrorAction SilentlyContinue
                        if ($prop) { $serial = $prop.Data }
                    } catch {}
                }
                if (-not $serial) { $serial = $pnp }
                $o += [PSCustomObject]@{Name=$name; Serial=$serial; Fabricante=$manuf; Modelo=$model}
            }
        } catch {}
        $o | ConvertTo-Json -Compress
        """
        out = run_powershell(ps, computer_name=computer_name)
        items = []
        try:
                parsed = json.loads(out) if out else []
                if isinstance(parsed, dict):
                        parsed = [parsed]
                for p in parsed:
                        items.append((p.get('Name'), p.get('Serial'), p.get('Fabricante'), p.get('Modelo')))
        except Exception:
                for l in out.splitlines():
                        if l.strip():
                                items.append((l, "", "", ""))
        return items or [("NÃO OBTIDO", "", "", "")]

def is_laptop(computer_name=None):
    # 1) verificar Win32_Battery (se existir, provavelmente notebook)
    out = run_powershell("Get-CimInstance Win32_Battery | ConvertTo-Json -Compress", computer_name=computer_name)
    if out and out.strip() != "null":
        return True
    # 2) verificar ChassisTypes em Win32_SystemEnclosure (valores que indicam portátil: 8,9,10,14)
    cmd = "(Get-CimInstance Win32_SystemEnclosure | Select-Object -ExpandProperty ChassisTypes | ConvertTo-Json -Compress)"
    out2 = run_powershell(cmd, computer_name=computer_name)
    try:
        arr = json.loads(out2) if out2 else []
        if isinstance(arr, int):
            arr = [arr]
        for v in arr:
            if int(v) in (8,9,10,14):  # Portable/Laptop/Notebook/SubNotebook
                return True
    except Exception:
        pass
    return False

def remove_duplicate_lines(lines):
    """Remove linhas duplicadas consecutivas, mantendo apenas a primeira ocorrência"""
    if not lines:
        return lines
    
    filtered_lines = [lines[0]]  # Sempre manter a primeira linha
    
    for i in range(1, len(lines)):
        current_line = lines[i].strip()
        previous_line = lines[i-1].strip()
        
        # Não remover linhas em branco ou linhas com "==="
        if not current_line or "===" in current_line:
            filtered_lines.append(lines[i])
            continue
        
        # Remover linha se for idêntica à anterior
        if current_line != previous_line:
            filtered_lines.append(lines[i])
    
    return filtered_lines

def write_report(path, lines):
    # Remove duplicidades antes de escrever
    filtered_lines = remove_duplicate_lines(lines)
    # Cabeçalho com versão e data
    try:
        import csinfo as _cs
        ver = getattr(_cs, '__version__', None)
    except Exception:
        ver = None
    # Novo cabeçalho conforme solicitado:
    # CEOsoftware Sistemas
    # CSInfo - Inventário de hardware e software - v{versão}
    header = [
        "CEOsoftware Sistemas",
        f"CSInfo - Inventário de hardware e software - v{ver if ver else 'desconhecida'}",
        "",
    ]
    with open(path, 'w', encoding='utf-8-sig') as f:
        f.write("\n".join(header + filtered_lines))

def organize_pdf_data(lines, computer_name):
    """Organiza os dados para o PDF conforme a nova estrutura solicitada"""
    # Remover duplicatas
    filtered_lines = remove_duplicate_lines(lines)
    
    # Dicionários para organizar as informações
    sistema_info = {}
    hardware_info = []
    admin_info = []
    software_info = []
    
    # Variáveis de controle
    current_section = ""
    report_date = ""
    machine_type = ""
    
    for line in filtered_lines:
        line = line.strip()
        if not line:
            continue
            
        # Cabeçalho e informações básicas
        if line == "CSInfo - Resumo Técnico do Dispositivo":
            continue
        elif line == "CSInfo by CEOsoftware":
            continue  # Rodapé
        elif line.startswith("Nome do computador"):
            continue  # Já temos no computer_name
        elif line.startswith("Tipo:"):
            machine_type = line.split(': ')[1] if ': ' in line else ""
            continue
        elif line.startswith("Relatório gerado em"):
            report_date = line
            continue
            
        # Detectar seções
        if line == "INFORMAÇÕES DO SISTEMA":
            current_section = "sistema"
            continue
        elif line == "INFORMAÇÕES DE HARDWARE":
            current_section = "hardware"
            continue
        elif line == "ADMINISTRADORES":
            current_section = "admin"
            continue
        elif line == "SOFTWARES INSTALADOS":
            current_section = "software"
            continue
            
        # Processar conteúdo baseado na seção
        if current_section == "sistema":
            if ":" in line:
                sistema_info[line.split(':')[0].strip()] = line
        elif current_section == "hardware":
            if ":" in line:  # Agora captura tanto linhas principais quanto indentadas
                hardware_info.append(line)
        elif current_section == "admin":
            if line.startswith("Administrador"):
                admin_name = line.split(': ')[1] if ': ' in line else line.replace("Administrador", "").strip()
                admin_info.append(admin_name)
        elif current_section == "software":
            if line and not line.startswith("Nenhum software"):
                # Remover numeração se existir
                clean_line = re.sub(r'^\s*\d+\.\s*', '', line)
                software_info.append(clean_line)
    
    # Ordenar listas
    admin_info.sort()
    software_info.sort()
    
    # Organizar hardware por grupos seguindo a nova estrutura
    hardware_groups = {
        'Memória': [],
        'Processador': [],
        'Disco': [],
        'Monitor': [],
        'Teclado': [],
        'Mouse': [],
        'Placa mãe': [],
        'Placa de Rede': [],
        'Placa de Vídeo': [],
        'Impressora': []
    }
    
    # Classificar hardware por grupos
    for info in hardware_info:
        assigned = False
        for group_name in hardware_groups.keys():
            if group_name.lower() in info.lower():
                hardware_groups[group_name].append(info)
                assigned = True
                break
        if not assigned:
            # Se não se encaixar em nenhum grupo específico, adicionar a 'Outros'
            if 'Outros' not in hardware_groups:
                hardware_groups['Outros'] = []
            hardware_groups['Outros'].append(info)
    
    # Ordenar dentro de cada grupo
    for group in hardware_groups.values():
        group.sort()
    
    return {
        'report_date': report_date,
        'computer_name': computer_name,
        'machine_type': machine_type,
        'sistema': sistema_info,
        'hardware': hardware_groups,
        'admins': admin_info,
        'software': software_info
    }

def write_pdf_report(path, lines, computer_name):
    """Gera um relatório em PDF com as informações coletadas - idêntico ao TXT"""
    try:
        from reportlab.platypus import Table, TableStyle, PageTemplate, Frame, Spacer, Paragraph
        from reportlab.lib.enums import TA_LEFT, TA_CENTER
        from reportlab.platypus import BaseDocTemplate
        from reportlab.lib.units import inch
        from reportlab.platypus.doctemplate import PageBreak
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        from reportlab.pdfgen import canvas
        from reportlab.platypus.flowables import HRFlowable
        
        # Remover duplicatas (igual ao TXT)
        filtered_lines = remove_duplicate_lines(lines)
        
        # Criar classe personalizada para rodapé
        class NumberedCanvas(canvas.Canvas):
            def __init__(self, *args, **kwargs):
                canvas.Canvas.__init__(self, *args, **kwargs)
                self._saved_page_states = []
                
            def showPage(self):
                self._saved_page_states.append(dict(self.__dict__))
                self._startPage()
                
            def save(self):
                """Adiciona informações de página em cada página"""
                num_pages = len(self._saved_page_states)
                for (page_num, state) in enumerate(self._saved_page_states):
                    self.__dict__.update(state)
                    self.draw_page_number(page_num + 1, num_pages)
                    canvas.Canvas.showPage(self)
                canvas.Canvas.save(self)
                
            def draw_page_number(self, page_num, total_pages):
                """Desenha o rodapé com numeração de página e texto centralizado"""
                # Desenhar linha fina cinza acima do rodapé
                self.setStrokeColor(colors.Color(0.7, 0.7, 0.7))  # Cinza claro
                self.setLineWidth(0.5)  # Linha fina
                self.line(0.5*inch, 0.5*inch, A4[0] - 0.5*inch, 0.5*inch)  # Linha horizontal
                
                # Configurar fonte para rodapé - tamanho 7, cinza 75%
                self.setFont("Helvetica", 7)
                self.setFillColor(colors.Color(0.25, 0.25, 0.25))  # Cinza 75% (25% preto)
                # Numeração de páginas (lado esquerdo)
                page_text = f"Página {page_num} de {total_pages}"
                self.drawString(0.5*inch, 0.3*inch, page_text)

                # Texto "CSInfo by CEOsoftware" centralizado (mantém apenas este texto + numeração)
                footer_text = "CSInfo by CEOsoftware"
                text_width = self.stringWidth(footer_text, "Helvetica", 7)
                page_width = A4[0]
                x_center = (page_width - text_width) / 2
                self.drawString(x_center, 0.3*inch, footer_text)
        
        # Criar documento PDF com template personalizado
        doc = BaseDocTemplate(path, pagesize=A4)
        
        # Definir frame com margens ajustadas para o rodapé
        frame = Frame(
            0.5*inch,  # x1 (margem esquerda)
            0.6*inch,  # y1 (margem inferior - espaço para rodapé)
            A4[0] - inch,  # width (largura da página - margens)
            A4[1] - 1.35*inch,  # altura - ajuste intermediário para o topo
            leftPadding=0,
            bottomPadding=0,
            rightPadding=0,
            topPadding=0.15*inch  # padding superior levemente maior
        )
        
        # Função para desenhar cabeçalho em cada página
        def draw_header(canvas, doc):
            canvas.saveState()
            # Ajusta a posição do cabeçalho para não sobrepor o conteúdo
            y_title = A4[1]-0.7*inch if doc.page == 1 else A4[1]-0.6*inch
            y_subtitle = A4[1]-0.9*inch if doc.page == 1 else A4[1]-0.8*inch
            y_version = A4[1]-1.05*inch if doc.page == 1 else A4[1]-0.95*inch
            y_line = A4[1]-0.95*inch if doc.page == 1 else A4[1]-0.85*inch
            # Novo cabeçalho alinhado ao formato do TXT
            try:
                import csinfo as _cs
                _ver = getattr(_cs, '__version__', None)
            except Exception:
                _ver = None
            # Tentar desenhar o ícone do app (assets/ico.png) no canto esquerdo do cabeçalho
            try:
                import sys
                # Resolve caminho ao arquivo assets mesmo quando empacotado pelo PyInstaller
                if getattr(sys, 'frozen', False):
                    base = sys._MEIPASS
                else:
                    base = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
                png_path = os.path.join(base, 'assets', 'ico.png')
                if os.path.exists(png_path):
                    # tamanho do ícone em polegadas
                    img_size = 0.5 * inch
                    # drawImage usa coordenada (x, y) do canto inferior esquerdo
                    x_img = 0.6 * inch
                    y_img = y_title - (img_size / 2)
                    try:
                        canvas.drawImage(png_path, x_img, y_img, width=img_size, height=img_size, preserveAspectRatio=True, mask='auto')
                    except Exception:
                        # se drawImage falhar (formatos), ignorar a imagem
                        pass
            except Exception:
                pass

            # Cabeçalho: apenas o título com versão (removido 'CEOsoftware Sistemas')
            ver_text = f"v{_ver}" if _ver else "vdesconhecida"
            canvas.setFont("Helvetica-Bold", 11)
            canvas.setFillColor(colors.navy)
            canvas.drawCentredString(A4[0]/2, y_title, f"CSInfo - Inventário de hardware e software - {ver_text}")
            canvas.setStrokeColor(colors.Color(0.7, 0.7, 0.7))
            canvas.setLineWidth(0.5)
            canvas.line(0.5*inch, y_line, A4[0]-0.5*inch, y_line)
            canvas.restoreState()
        # Criar template da página com cabeçalho
        template = PageTemplate(id='normal', frames=[frame], onPage=draw_header)
        doc.addPageTemplates([template])
        
        story = []
        
        # Estilos
        styles = getSampleStyleSheet()
        
        # Estilo para cabeçalho CSInfo
        header_style = ParagraphStyle(
            'CSInfoHeader',
            parent=styles['Normal'],
            fontSize=11,
            textColor=colors.navy,
            alignment=TA_LEFT,
            spaceAfter=6,
            leading=13,
            fontName='Helvetica-Bold'
        )
        
        # Estilos para títulos de seção personalizados
        section_title_styles = {
            "INFORMAÇÕES DO SISTEMA": ParagraphStyle(
                'SectionTitleSistema',
                parent=styles['Heading2'],
                fontSize=10,
                spaceAfter=6,
                spaceBefore=6,
                leading=12,
                textColor=colors.Color(0.5, 0, 0),  # Vermelho escuro
                fontName='Helvetica-Bold',
                alignment=TA_LEFT
            ),
            "INFORMAÇÕES DE HARDWARE": ParagraphStyle(
                'SectionTitleHardware',
                parent=styles['Heading2'],
                fontSize=10,
                spaceAfter=6,
                spaceBefore=6,
                leading=12,
                textColor=colors.Color(0, 0.35, 0),  # Verde escuro
                fontName='Helvetica-Bold',
                alignment=TA_LEFT
            ),
            "ADMINISTRADORES": ParagraphStyle(
                'SectionTitleAdmin',
                parent=styles['Heading2'],
                fontSize=10,
                spaceAfter=6,
                spaceBefore=6,
                leading=12,
                textColor=colors.Color(1, 0.5, 0),  # Laranja
                fontName='Helvetica-Bold',
                alignment=TA_LEFT
            ),
            "SOFTWARES INSTALADOS": ParagraphStyle(
                'SectionTitleSoft',
                parent=styles['Heading2'],
                fontSize=10,
                spaceAfter=6,
                spaceBefore=6,
                leading=12,
                textColor=colors.Color(0.7, 0.6, 0.1),  # Amarelo escuro
                fontName='Helvetica-Bold',
                alignment=TA_LEFT
            ),
            "INFORMAÇÕES DE REDE": ParagraphStyle(
                'SectionTitleNet',
                parent=styles['Heading2'],
                fontSize=10,
                spaceAfter=6,
                spaceBefore=6,
                leading=12,
                textColor=colors.Color(0.1, 0.3, 0.7),  # Azul escuro
                fontName='Helvetica-Bold',
                alignment=TA_LEFT
            ),
            "SEGURANÇA DO SISTEMA": ParagraphStyle(
                'SectionTitleSec',
                parent=styles['Heading2'],
                fontSize=10,
                spaceAfter=6,
                spaceBefore=6,
                leading=12,
                textColor=colors.Color(0.7, 0.1, 0.1),  # Vermelho escuro
                fontName='Helvetica-Bold',
                alignment=TA_LEFT
            ),
        }
        
        # Estilo para texto normal
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=2,
            leading=11,
            textColor=colors.black,
            fontName='Helvetica'
        )
        
        # Estilo para texto indentado (com 2 espaços)
        indented_style = ParagraphStyle(
            'IndentedNormal',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=2,
            leading=11,
            textColor=colors.black,
            fontName='Helvetica',
            leftIndent=12  # Indentação de 12 pontos
        )
        
        # Estilo para texto muito indentado (com 4 espaços)
        double_indented_style = ParagraphStyle(
            'DoubleIndentedNormal',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=2,
            leading=11,
            textColor=colors.black,
            fontName='Helvetica',
            leftIndent=24  # Indentação de 24 pontos
        )
        
        def clean_text(text):
            """Limpa texto para PDF: remove caracteres de controle (ex: \x00) e escapa marcações"""
            if text is None:
                return ""
            # Converte para str e remove caracteres de controle que geram quadrados pretos
            txt = str(text)
            # Remove caracteres de controle (0x00-0x1F, 0x7F-0x9F)
            txt = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', txt)
            # Substituições seguras para o mini-markup do ReportLab
            return (txt.replace('&', '&amp;')
                       .replace('<', '&lt;')
                       .replace('>', '&gt;')
                       .replace('"', '&quot;')
                       .replace("'", '&#39;'))
        
        # Adiciona espaçamento extra no início do story (aplica na primeira página)
        story.append(Spacer(1, 18))

        # Processar cada linha exatamente como no TXT
        for idx, line in enumerate(filtered_lines):
            line_stripped = line.strip()
            # Linha em branco
            if not line_stripped:
                story.append(Spacer(1, 6))
                continue
            # (Cabeçalho é desenhado pelo template; não há linhas de cabeçalho no corpo)
            # Títulos de seções com cor personalizada
            if line_stripped in section_title_styles:
                story.append(Paragraph(f"<b>{clean_text(line_stripped)}</b>", section_title_styles[line_stripped]))
                continue
            
            # Pular o rodapé do conteúdo (será adicionado automaticamente)
            if line_stripped == "CSInfo by CEOsoftware":
                continue
            
            # Determinar o estilo baseado na indentação
            if line.startswith("    "):  # 4 espaços - indentação dupla
                story.append(Paragraph(clean_text(line_stripped), double_indented_style))
            elif line.startswith("  "):   # 2 espaços - indentação simples
                story.append(Paragraph(clean_text(line_stripped), indented_style))
            else:  # Sem indentação
                story.append(Paragraph(clean_text(line_stripped), normal_style))
        
        # Gerar PDF com canvas personalizado
        doc.build(story, canvasmaker=NumberedCanvas)
        return True
        
    except Exception as e:
        print(f"Erro ao gerar PDF: {e}")
        return False

def check_remote_machine(computer_name):
    if not computer_name:
        return True
    import subprocess, os, socket, time
    try:
        # 1) Ping com algumas tentativas (sem abrir janela no Windows)
        creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000) if os.name == 'nt' else 0
        def try_ping():
            try:
                p = subprocess.run(["ping", "-n", "1", computer_name], capture_output=True, text=True, timeout=3, creationflags=creationflags)
                return p.returncode == 0
            except Exception:
                return False

        for _ in range(2):
            if try_ping():
                return True
            time.sleep(0.4)

        # 2) Tentar conectar em portas comuns (WinRM HTTP/HTTPS, SMB) - se qualquer uma responder, consideramos acessível
        ports = [5985, 5986, 445]
        for port in ports:
            try:
                sock = socket.create_connection((computer_name, port), timeout=2)
                sock.close()
                return True
            except Exception:
                pass

        # 3) Tentativa final via WinRM/PowerShell (pode falhar se WinRM não estiver configurado)
        try:
            out = run_powershell('Write-Output $env:COMPUTERNAME', computer_name=computer_name)
            if out and str(out).strip():
                return True
        except Exception:
            pass

        return False
    except Exception:
        return False

def is_remote_admin(computer_name=None):
    if not computer_name or computer_name.lower() in ("localhost", os.environ.get("COMPUTERNAME", "").lower()):
        return True
    ps = r'''
    try {
        $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $adminGroup = [ADSI]("WinNT://%s/Administrators")
        $isAdmin = $false
        foreach ($member in $adminGroup.Members()) {
            if ($member.GetType().InvokeMember("Name", 'GetProperty', $null, $member, $null) -eq $user.Split("\\")[1]) {
                $isAdmin = $true
                break
            }
        }
        $isAdmin
    } catch {
        $false
    }
    ''' % computer_name
    out = run_powershell(ps, computer_name=computer_name)
    return str(out).strip().lower() in ("true", "1")

def main(export_type=None, barra_callback=None, computer_name=None):
    # --- COLETA DE INFORMAÇÕES AVANÇADAS ---
    network_details = get_network_details(computer_name)
    firewall_status = get_firewall_status(computer_name)
    windows_update_status = get_windows_update_status(computer_name)
    running_processes = get_running_processes(computer_name)
    critical_services = get_critical_services(computer_name)
    firewall_controller = get_firewall_controller(computer_name)
    modo_gui = export_type is not None or barra_callback is not None
    # Inicializar flags de geração
    gerar_txt = False
    gerar_pdf = False
    # Se chamado pela GUI, não pedir input nem printar
    if not modo_gui:
        print("=== Coletor de Informações de Máquina ===")
        computer_input = input("Digite o nome da máquina remota ou pressione ENTER com o campo em branco para analisar a máquina local: ").strip()
        computer_name = None if not computer_input else computer_input
        if computer_name:
            print(f"Coletando informações da máquina remota: {computer_name}")
            print("Nota: Certifique-se de que você tem permissões administrativas na máquina remota")
            print("e que o WinRM está habilitado na máquina de destino.\n")
        else:
            print("Coletando informações da máquina local\n")
    else:
        if modo_gui:
            gerar_txt = export_type in ('txt', 'ambos')
            gerar_pdf = export_type in ('pdf', 'ambos')
        else:
            print("\nEscolha o formato de relatório para gerar:")
            print("1 - TXT")
            print("2 - PDF")
            print("3 - Ambos (TXT e PDF)")
            escolha = input("Digite o número da opção desejada [1/2/3]: ").strip()
            gerar_txt = escolha in ('1', '3')
            gerar_pdf = escolha in ('2', '3')

    etapas = [
        "Obtendo nome do computador",
        "Verificando tipo (Notebook/Desktop)",
        "Obtendo versão do sistema operacional",
        "Verificando ativação do Windows",
        "Verificando tipo de rede (Domínio/Workgroup)",
        "Obtendo informações do processador",
        "Obtendo quantidade de memória RAM",
        "Obtendo informações dos pentes de memória",
        "Obtendo informações dos discos rígidos",
        "Obtendo informações das unidades lógicas",
        "Obtendo versão do Office",
        "Verificando ativação do Office",
        "Obtendo informações da placa mãe",
        "Obtendo informações dos monitores",
        "Obtendo informações das placas de rede",
        "Obtendo informações das placas de vídeo",
        "Verificando teclado e mouse",
        "Obtendo informações do SQL Server",
        "Obtendo informações do antivírus",
        "Obtendo usuários administradores",
        "Obtendo informações das impressoras",
        "Obtendo lista de softwares instalados",
        "Salvando relatório"
    ]
    total = len(etapas)
    inicio = time.time()
    def barra_progresso(atual):
        largura = 30
        perc = int((atual/total)*100)
        preenchido = int(largura*atual/total)
        barra = '[' + '#' * preenchido + '-' * (largura-preenchido) + f'] {perc}%'
        tempo = int(time.time()-inicio)
        etapa_texto = etapas[atual-1] if atual-1 < len(etapas) else "Finalizado"
        if hasattr(barra_progresso, "callback") and callable(barra_progresso.callback):
            barra_progresso.callback(perc, etapa_texto)
        elif not modo_gui:
            print(f"\r{barra} | {etapa_texto} | Tempo: {tempo}s", end='', flush=True)

    # Se a função externa de callback foi passada, linka-a à barra_progresso
    if barra_callback and callable(barra_callback):
        barra_progresso.callback = barra_callback

    # Adiciona callback para cada linha apurada
    def add_line(line):
        lines.append(line)
        if barra_callback:
            barra_callback(None, line)

    machine = get_machine_name(computer_name); barra_progresso(1)
    import getpass
    if not computer_name or computer_name.lower() == machine.lower():
        usuario_logado = getpass.getuser()
    else:
        import importlib.util, sys
        mod_name = "network_discovery"
        mod_path = os.path.join(os.path.dirname(__file__), "network_discovery.py")
        spec = importlib.util.spec_from_file_location(mod_name, mod_path)
        network_discovery = importlib.util.module_from_spec(spec)
        sys.modules[mod_name] = network_discovery
        spec.loader.exec_module(network_discovery)
        usuario_logado = network_discovery.get_logged_user(machine)
    safe_name = safe_filename(machine)
    filename = f"info_maquina_{safe_name}.txt"
    path = os.path.join(os.getcwd(), filename)

    lines = []
    def padrao(valor):
        return valor if valor and str(valor).strip() else "NÃO OBTIDO"
    
    # LINHAS INICIAIS SOLICITADAS: Nome, Tipo, Gerado por
    add_line(f"Nome do computador: {machine}")
    add_line(f"Tipo: {'Notebook' if is_laptop(computer_name) else 'Desktop'}")
    add_line(f"Gerado por: {usuario_logado}")
    add_line("")
    # (Cabeçalho visual será desenhado pelo template do PDF; não adicionamos essas linhas ao corpo)
    # DATA DE GERAÇÃO
    add_line(f"Relatório gerado em: {datetime.now().strftime('%d/%m/%Y - %H:%M')}")
    add_line("")
    barra_progresso(2)
    
    # INFORMAÇÕES DO SISTEMA
    add_line("INFORMAÇÕES DO SISTEMA")
    add_line(f"Versão do sistema operacional: {get_os_version(computer_name)}"); barra_progresso(4)
    add_line(f"  Status do sistema operacional: {get_windows_activation_status(computer_name)}"); barra_progresso(5)
    add_line("")  # Linha em branco após Windows
    add_line(f"Versão do Office: {get_office_version(computer_name)}"); barra_progresso(6)
    add_line(f"  Status do Office: {get_office_activation_status(computer_name)}"); barra_progresso(7)
    add_line("")  # Linha em branco após Office
    
    # SQL Server
    sql_instances = get_sql_server_info(computer_name)
    if sql_instances:
        for idx, instance in enumerate(sql_instances, start=1):
            instance_name = padrao(instance.get('Instance', ''))
            version = padrao(instance.get('Version', ''))
            status = padrao(instance.get('Status', ''))
            add_line(f"SQL Server {idx}: Instância: {instance_name} | Versão: {version} | Status: {status}")
    else:
        add_line("SQL Server: NÃO INSTALADO")
    add_line("")  # Linha em branco após SQL Server
    barra_progresso(8)
    
    # Antivírus
    antivirus_list = get_antivirus_info(computer_name)
    if antivirus_list:
        for idx, av in enumerate(antivirus_list, start=1):
            name = padrao(av.get('Name', ''))
            enabled = padrao(av.get('Enabled', ''))
            add_line(f"Antivírus {idx}: {name} | Status: {enabled}")
    else:
        add_line("Antivírus: NÃO DETECTADO")
    add_line("")  # Linha em branco após Antivírus
    barra_progresso(9)
    
    add_line(f"Rede: {is_domain_computer(computer_name)}"); barra_progresso(10)
    add_line("")
    
    # INFORMAÇÕES DE HARDWARE
    add_line("INFORMAÇÕES DE HARDWARE")
    
    # MEMÓRIA RAM
    add_line(f"Memória RAM total: {get_memory_info(computer_name)}"); barra_progresso(11)
    
    # Pentes de memória
    memory_modules = get_memory_modules_info(computer_name)
    if memory_modules:
        for idx, module in enumerate(memory_modules, start=1):
            manufacturer = padrao(module.get('Manufacturer', ''))
            part_number = padrao(module.get('PartNumber', ''))
            capacity = padrao(module.get('Capacity', ''))
            speed = padrao(module.get('Speed', ''))
            mem_type = padrao(module.get('MemoryType', ''))
            form_factor = padrao(module.get('FormFactor', ''))
            location = padrao(module.get('Location', ''))
            add_line(f"  Pente de Memória {idx}: {capacity} | {mem_type} | {speed} | {form_factor}")
            add_line(f"    Fabricante: {manufacturer}")
    else:
        add_line("  Pentes de Memória: NÃO OBTIDO")
    add_line("")  # Linha em branco após Memória
    barra_progresso(12)
    
    # PROCESSADOR
    processors = get_processor_info(computer_name)
    if processors:
        for idx, cpu in enumerate(processors, start=1):
            name = padrao(cpu.get('Name', ''))
            manufacturer = padrao(cpu.get('Manufacturer', ''))
            cores = padrao(cpu.get('Cores', ''))
            logical = padrao(cpu.get('LogicalProcessors', ''))
            max_speed = padrao(cpu.get('MaxSpeed', ''))
            arch = padrao(cpu.get('Architecture', ''))
            l2_cache = padrao(cpu.get('L2Cache', ''))
            l3_cache = padrao(cpu.get('L3Cache', ''))
            add_line(f"Processador {idx}: {name}")
            add_line(f"  Cores: {cores} físicos | {logical} lógicos")
            add_line(f"  Cache: L2: {l2_cache} | L3: {l3_cache}")
            add_line(f"  Fabricante: {manufacturer}")
    else:
        add_line("Processador: NÃO OBTIDO")
    add_line("")  # Linha em branco após Processador
    barra_progresso(13)
    
    # DISCOS
    disks = get_disk_info(computer_name)
    if disks:
        for idx, (modelo, tamanho, espaco_usado, espaco_livre, particoes, tipo, interface) in enumerate(disks, start=1):
            add_line(f"Disco {idx}: {padrao(modelo)} | Tamanho: {padrao(tamanho)}")
            add_line(f"  Tipo: {padrao(tipo)} | Interface: {padrao(interface)}")
    else:
        add_line("Disco: NÃO OBTIDO")
    barra_progresso(14)

    # Unidades (após todos os discos)
    logical_drives = get_logical_drives_info(computer_name)
    if logical_drives:
        for drive in logical_drives:
            drive_letter = padrao(drive.get('Drive', ''))
            size = padrao(drive.get('Size', ''))
            used = padrao(drive.get('Used', ''))
            free = padrao(drive.get('Free', ''))
            file_system = padrao(drive.get('FileSystem', ''))
            label = padrao(drive.get('Label', ''))
            # Truncar espaço usado para 2 casas decimais
            try:
                used_float = float(used)
                used = f"{used_float:.2f}"
            except:
                pass
            try:
                size_float = float(size)
                size = f"{size_float:.2f}"
            except:
                pass
            try:
                free_float = float(free)
                free = f"{free_float:.2f}"
            except:
                pass
            lines.append(f"Unidade {drive_letter} ({label}) | Total: {size} GB | Usado: {used} GB | Livre: {free} GB | Sistema: {file_system}")
    else:
        lines.append("Unidades lógicas: NÃO OBTIDO")
    lines.append("")  # Linha em branco após unidades lógicas
    barra_progresso(15)

    # MONITORES
    monitors = get_monitor_infos(computer_name)
    if monitors:
        for idx, m in enumerate(monitors, start=1):
            lines.append(f"Monitor {idx}: {padrao(m.get('Fabricante',''))} | Modelo: {padrao(m.get('Modelo',''))} | Serial: {padrao(m.get('Serial',''))}")
    else:
        lines.append("Monitor 1: NÃO OBTIDO")
    lines.append("")  # Linha em branco após monitores
    barra_progresso(16)

    # TECLADO
    has_keyboard, has_mouse = get_keyboard_mouse_status(computer_name)
    lines.append(f"Teclado conectado: {'SIM' if has_keyboard else 'NÃO'}")
    
    # MOUSE
    lines.append(f"Mouse conectado: {'SIM' if has_mouse else 'NÃO'}")
    lines.append("")  # Linha em branco após teclado/mouse
    barra_progresso(17)

    # PLACA MÃE
    fabricante_mb, modelo_mb, serial_mb = get_motherboard_info(computer_name)
    lines.append(f"Placa mãe: {padrao(fabricante_mb)} | Modelo: {padrao(modelo_mb)} | Serial: {padrao(serial_mb)}")
    lines.append("")  # Linha em branco após placa mãe
    barra_progresso(18)

    # PLACA DE REDE
    network_adapters = get_network_adapters_info(computer_name)
    if network_adapters:
        for idx, adapter in enumerate(network_adapters, start=1):
            name = padrao(adapter.get('Name', ''))
            manufacturer = padrao(adapter.get('Manufacturer', ''))
            speed = padrao(adapter.get('Speed', ''))
            mac = padrao(adapter.get('MACAddress', ''))
            lines.append(f"Placa de Rede {idx}: {name} | Fabricante: {manufacturer} | Velocidade: {speed} | MAC: {mac}")
    else:
        lines.append("Placa de Rede: NÃO OBTIDO")
    lines.append("")  # Linha em branco após placa de rede
    barra_progresso(19)

    # PLACA DE VÍDEO
    video_cards = get_video_cards_info(computer_name)
    if video_cards:
        for idx, card in enumerate(video_cards, start=1):
            name = padrao(card.get('Name', ''))
            manufacturer = padrao(card.get('Manufacturer', ''))
           
            memory = padrao(card.get('Memory', ''))

            card_type = padrao(card.get('Type', ''))
            lines.append(f"Placa de Vídeo {idx}: {name} | Fabricante: {manufacturer} | Memória: {memory} | Tipo: {card_type}")
    else:
        lines.append("Placa de Vídeo: NÃO OBTIDO")
    lines.append("")  # Linha em branco após placa de vídeo
    barra_progresso(20)

    # IMPRESSORAS
    printers = get_printers(computer_name)
    if printers:
        for idx, (name, serial, fabricante, modelo) in enumerate(printers, start=1):
            lines.append(f"Impressora {idx}: {padrao(name)} | Serial/ID: {padrao(serial)} | Fabricante: {padrao(fabricante)} | Modelo: {padrao(modelo)}")
    else:
        lines.append("Impressora: NÃO OBTIDO")
    lines.append("")
    barra_progresso(21)

    # ADMINISTRADORES (mantém como está)
    lines.append("ADMINISTRADORES")
    admin_users = get_admin_users(computer_name)
    if admin_users:
        for idx, user in enumerate(admin_users, start=1):
            # Como agora get_admin_users retorna apenas strings (nomes)
            name = padrao(user if isinstance(user, str) else user.get('Name', ''))
            lines.append(f"Administrador {idx}: {name}")
    else:
        lines.append("Usuários Administradores: NÃO OBTIDO")
    barra_progresso(22)

    # SOFTWARES INSTALADOS (mantém como está)
    lines.append("")  # Linha em branco para separar
    lines.append("SOFTWARES INSTALADOS")
    installed_software = get_installed_software(computer_name)
    if installed_software:
        for idx, software in enumerate(installed_software, start=1):
            name = padrao(software.get('Name', ''))
            version = padrao(software.get('Version', ''))
            publisher = padrao(software.get('Publisher', ''))
            lines.append(f"{idx}. {name} | Versão: {version} | Editor: {publisher}")
    else:
        lines.append("Nenhum software detectado")
    barra_progresso(23)

    # --- NOVAS SEÇÕES ---
    lines.append("")
    lines.append("INFORMAÇÕES DE REDE")
    if network_details:
        for idx, net in enumerate(network_details, start=1):
            ip = padrao(net.get('IP', ''))
            gw = padrao(net.get('Gateway', ''))
            dns = padrao(net.get('DNS', ''))
            mac = padrao(net.get('MAC', ''))
            desc = padrao(net.get('Descricao', ''))
            lines.append(f"Adaptador {idx}: {desc}")
            lines.append(f"  IP: {ip} | Gateway: {gw} | DNS: {dns} | MAC: {mac}")
    else:
        lines.append("Informações de rede: NÃO OBTIDO")

    lines.append("")
    lines.append("SEGURANÇA DO SISTEMA")
    # Firewall
    lines.append("Firewall:")
    if firewall_status:
        for fw in firewall_status:
            perfil = padrao(fw.get('Perfil', ''))
            ativado = 'Ativado' if fw.get('Ativado', False) else 'Desativado'
            lines.append(f"  Perfil: {perfil} | Status: {ativado}")
    else:
        lines.append("  Status: NÃO OBTIDO")
    # Controle externo do firewall
    if firewall_controller and firewall_controller != "Windows Firewall (padrão)" and firewall_controller != "NÃO OBTIDO":
        lines.append(f"Firewall controlado por: {firewall_controller}")
    # Espaço de uma linha antes do Windows Update
    lines.append("")
    # Windows Update
    if windows_update_status:
        lines.append("Windows Update:")
        lines.append(f"  {windows_update_status}")
    else:
        lines.append("Windows Update: NÃO OBTIDO")
    
    # RODAPÉ
    lines.append("")
    lines.append("")
    lines.append("CSInfo by CEOsoftware")

    # Só gera TXT/PDF se export_type for passado explicitamente e for um dos valores esperados
    pdf_path = None
    print(f"DEBUG csinfo: export_type={repr(export_type)}, gerar_txt={gerar_txt}, gerar_pdf={gerar_pdf}")
    if export_type in ('txt', 'pdf', 'ambos'):
        # Escrita explícita conforme seleção
        if export_type in ('txt', 'ambos'):
            try:
                print(f"DEBUG csinfo: Gravando TXT em: {path}")
                write_report(path, lines)
                barra_progresso(23)
                print(f"\u2705 Arquivo TXT gerado: {path}")
            except Exception as e:
                print(f"\u274c Erro ao escrever TXT: {e}")
        if export_type in ('pdf', 'ambos'):
            pdf_path = path.replace('.txt', '.pdf')
            print("Gerando arquivo PDF...")
            try:
                ok = write_pdf_report(pdf_path, lines, machine)
                if ok:
                    print(f"\u2705 Arquivo PDF gerado com sucesso: {pdf_path}")
                else:
                    print("\u274c Erro ao gerar arquivo PDF.")
                    pdf_path = None
            except Exception as e:
                print(f"\u274c Exceção ao gerar PDF: {e}")

    # Perguntar se deseja abrir os arquivos gerados
    if not modo_gui:
        if gerar_txt:
            resposta = input(f"Deseja abrir o arquivo TXT gerado ({filename})? [s/N]: ").strip().lower()
            if resposta == 's':
                try:
                    os.startfile(path)
                except Exception as e:
                    print(f"Não foi possível abrir o arquivo TXT: {e}")
    
    # Retornar sempre os caminhos dos arquivos gerados (ou None) e as linhas coletadas
    resultado = {
        'txt': path if export_type in ('txt', 'ambos') and os.path.exists(path) else None,
        'pdf': pdf_path if pdf_path and os.path.exists(pdf_path) else None,
        'lines': lines,
        'machine': machine,
        'user': usuario_logado
    }
    return resultado
