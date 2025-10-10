"""Implementação interna do pacote csinfo.

Este arquivo contém a implementação completa que antes vivia em `csinfo.py` na raiz
do projeto. Foi movida para dentro do pacote para permitir importação limpa
(`import csinfo`) e facilitar o empacotamento com PyInstaller.
"""
# Conteúdo migrado de csinfo.py
import socket
import json
import re
import tempfile
import os

def get_debug_session_log():
    """Retorna o path do arquivo de sessão de debug atual, se existir.

    Retorna None se CSINFO_DEBUG não estiver habilitado ou se o log não existir.
    """
    try:
        path = os.environ.get('CSINFO_DEBUG_SESSION') or getattr(run_powershell, '_csinfo_session_log', None)
        if path and os.path.exists(path):
            return path
    except Exception:
        pass
    return None


# Credencial padrão usada por run_powershell quando não fornecida explicitamente
_CSINFO_DEFAULT_CREDENTIAL = None


def set_default_credential(user, password):
    """Define uma credencial padrão (usuário, senha) usada por run_powershell.

    A implementação armazena em uma variável global do módulo. A validação
    é mínima — preservamos os valores como strings.
    """
    global _CSINFO_DEFAULT_CREDENTIAL
    try:
        _CSINFO_DEFAULT_CREDENTIAL = (str(user), str(password))
    except Exception:
        _CSINFO_DEFAULT_CREDENTIAL = (user, password)
    return True


def clear_default_credential():
    """Limpa a credencial padrão previamente definida."""
    global _CSINFO_DEFAULT_CREDENTIAL
    _CSINFO_DEFAULT_CREDENTIAL = None
    return True

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
                if os.path.exists(tmp_path):
                    os.replace(tmp_path, path)
                # limpar sidecar se ainda existir (não deixamos arquivos temporários após sucesso)
                try:
                    # tentar várias localizações possíveis do sidecar, pois o path
                    # de saída pode ser absoluto/externo ao workspace
                    candidates = []
                    try:
                        candidates.append(annots_sidecar)
                    except Exception:
                        pass
                    try:
                        candidates.append(path + '.tmp.annots.jsonl')
                    except Exception:
                        pass
                    try:
                        candidates.append(path + '.annots.jsonl')
                    except Exception:
                        pass
                    try:
                        candidates.append(tmp_path + '.annots.jsonl')
                    except Exception:
                        pass
                    # deduplicate
                    seen = set()
                    for c in candidates:
                        if not c:
                            continue
                        if c in seen:
                            continue
                        seen.add(c)
                        try:
                            if os.path.exists(c):
                                os.remove(c)
                        except Exception:
                            # não falhar a exportação por conta do cleanup
                            pass
                except Exception:
                    pass
                # também limpar qualquer sidecar no diretório de exportação (padrão do usuário)
                try:
                    export_dir = os.path.dirname(os.path.abspath(path)) or os.getcwd()
                    base = os.path.splitext(os.path.basename(path))[0]
                    # padrão: arquivos que contenham o base do nome e terminem com .annots.jsonl
                    candidates = []
                    try:
                        for fn in os.listdir(export_dir):
                            if fn.endswith('.annots.jsonl') and base in fn:
                                candidates.append(os.path.join(export_dir, fn))
                    except Exception:
                        candidates = []

                    # tentar remover com algumas tentativas (em caso de locks temporários)
                    import time
                    for full in candidates:
                        for attempt in range(3):
                            try:
                                if os.path.exists(full):
                                    os.remove(full)
                                break
                            except Exception:
                                time.sleep(0.12)
                                continue
                except Exception:
                    pass
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
    out = run_powershell(ps, computer_name=computer_name, timeout=60, retries=3)
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

# --- Helpers para criação de títulos e índice com anchors/links internos ---
def criar_titulo_pdf(secao, styles=None):
    """Retorna um Paragraph que contém um anchor nomeado igual ao texto da seção.

    Uso:
        estilo = section_title_styles.get(secao, styles['Heading2'])
        p = criar_titulo_pdf('INFORMAÇÕES DO SISTEMA', styles)

    Isso cria internamente: <a name='INFORMAÇÕES DO SISTEMA'/>INFORMAÇÕES DO SISTEMA
    """
    try:
        if styles is None:
            styles = getSampleStyleSheet()
        estilo = None
        try:
            estilo = styles.get(secao)
        except Exception:
            estilo = None
        if estilo is None:
            # fallback genérico
            estilo = ParagraphStyle('SectionTitle', parent=styles['Heading2'], fontSize=10, leading=12, textColor=colors.whitesmoke, fontName='Helvetica-Bold')
        texto = f"<a name='{secao}'/>{secao}"
        return Paragraph(texto, estilo)
    except Exception:
        # Em caso de erro, retornar um Paragraph simples
        try:
            return Paragraph(secao, styles['Normal'])
        except Exception:
            return Paragraph(str(secao), getSampleStyleSheet()['Normal'])


def criar_indice_pdf(itens, styles=None, indice_style_name='IndiceStyle'):
    """Retorna uma lista de flowables (Paragraph + Spacer) representando o índice lateral.

    Cada item será renderizado como um link interno para o anchor com o mesmo nome.
    Exemplo de tag usado: <a href='#INFORMAÇÕES DO SISTEMA'>INFORMAÇÕES DO SISTEMA</a>
    """
    try:
        if styles is None:
            styles = getSampleStyleSheet()
        indice_estilo = ParagraphStyle(indice_style_name, parent=styles['Normal'], fontSize=9, textColor=colors.whitesmoke, leftIndent=4, leading=11)
        flow = []
        for item in itens:
            # criar link interno usando tag <a href='#...'> disponível no ReportLab
            # preservamos acentos e caixa do texto
            link_text = f"<a href='#{item}'><font color='white'>{item}</font></a>"
            flow.append(Paragraph(link_text, indice_estilo))
            flow.append(Spacer(1, 4))
        return flow
    except Exception:
        # fallback simples
        sheet = getSampleStyleSheet()
        return [Paragraph(i, sheet['Normal']) for i in itens]


def run_powershell(cmd, computer_name=None, timeout=20, retries=2, initial_backoff=0.3, credential=None):
    """Execute um comando PowerShell com retries/backoff e timeout configurável.

    - timeout: segundos por tentativa (pode ser sobrescrito pela variável de ambiente CSINFO_PS_TIMEOUT)
    - retries: número de tentativas totais (inclui a primeira)
    - initial_backoff: tempo inicial em segundos para backoff exponencial
    """
    # permitir override global via env
    try:
        env_to = int(os.environ.get('CSINFO_PS_TIMEOUT', str(timeout)))
        timeout = env_to
    except Exception:
        pass

    # compatibilidade: se caller passou computer_name como positional (antes era timeout), garantir que computer_name seja uma string ou None
    if isinstance(computer_name, (int, float)):
        # recebido um número como segundo argumento -> isso na verdade era timeout; ajustar
        timeout = computer_name
        computer_name = None

    # guardar comando original para possíveis re-execucoes
    original_cmd = cmd

    # se não foi passada credencial explicitamente, tentar usar a default definida por set_default_credential
    if credential is None:
        try:
            credential = globals().get('_CSINFO_DEFAULT_CREDENTIAL')
        except Exception:
            credential = None

    # Detectar se o alvo é local (nome corresponde ao hostname/local computername)
    is_local = False
    if computer_name:
        try:
            resolved_local = set()
            try:
                resolved_local.add(socket.gethostname().lower())
            except Exception:
                pass
            try:
                env_name = os.environ.get('COMPUTERNAME')
                if env_name:
                    resolved_local.add(str(env_name).lower())
            except Exception:
                pass
            resolved_local.update(('localhost', '127.0.0.1', '::1'))
            is_local = str(computer_name).lower() in {n for n in resolved_local if n}
        except Exception:
            is_local = False

    # Construir variantes de comando remoto (serão tentadas ciclicamente nas tentativas)
    remote_variants = None
    if computer_name and not is_local:
        remote_variants = []
        if credential and isinstance(credential, (list, tuple)) and len(credential) == 2:
            user, pwd = credential
            # escapar aspas simples para PowerShell (duplicar)
            user_esc = str(user).replace("'", "''")
            pwd_esc = str(pwd).replace("'", "''")
            # 1) Tentar Negotiate (NTLM/Kerberos negotiable)
            remote_variants.append(
                f"$sec = ConvertTo-SecureString '{pwd_esc}' -AsPlainText -Force; $cred = New-Object System.Management.Automation.PSCredential('{user_esc}',$sec); Invoke-Command -ComputerName {computer_name} -Credential $cred -Authentication Negotiate -ScriptBlock {{ {original_cmd} }} -ErrorAction Stop"
            )
            # 2) Tentar padrão com credencial (sem explicit Authentication)
            remote_variants.append(
                f"$sec = ConvertTo-SecureString '{pwd_esc}' -AsPlainText -Force; $cred = New-Object System.Management.Automation.PSCredential('{user_esc}',$sec); Invoke-Command -ComputerName {computer_name} -Credential $cred -ScriptBlock {{ {original_cmd} }} -ErrorAction Stop"
            )
            # 3) Tentar via SSL (se configurado)
            remote_variants.append(
                f"$sec = ConvertTo-SecureString '{pwd_esc}' -AsPlainText -Force; $cred = New-Object System.Management.Automation.PSCredential('{user_esc}',$sec); Invoke-Command -ComputerName {computer_name} -Credential $cred -Authentication Negotiate -UseSSL -ScriptBlock {{ {original_cmd} }} -ErrorAction Stop"
            )
        else:
            # Sem credenciais: tentar Negotiate e padrão
            remote_variants.append(f"Invoke-Command -ComputerName {computer_name} -ScriptBlock {{ {original_cmd} }} -Authentication Negotiate -ErrorAction Stop")
            remote_variants.append(f"Invoke-Command -ComputerName {computer_name} -ScriptBlock {{ {original_cmd} }} -ErrorAction Stop")

    # Forçar codificação UTF-8 no PowerShell
    # Note: se remote_variants estiver definido, iremos selecionar uma variante com base na tentativa atual

    # Usar credencial default do módulo se credential não foi fornecida
    try:
        if credential is None and globals().get('_CSINFO_DEFAULT_CREDENTIAL'):
            credential = globals().get('_CSINFO_DEFAULT_CREDENTIAL')
    except Exception:
        pass

    # preparar arquivo de sessão (único por execução) quando CSINFO_DEBUG habilitado
    debug_enabled = bool(os.environ.get('CSINFO_DEBUG'))
    session_log = None
    if debug_enabled:
        # se o usuário passou um caminho customizado para o log de sessão via env, usar
        session_log = os.environ.get('CSINFO_DEBUG_SESSION')
        if not session_log:
            # criar e reutilizar um arquivo na primeira chamada desta execução
            # armazenamos em uma variável no módulo para não recriar em chamadas subsequentes
            try:
                if not hasattr(run_powershell, '_csinfo_session_log') or not run_powershell._csinfo_session_log:
                    run_powershell._csinfo_session_log = os.path.join(tempfile.gettempdir(), f"csinfo_debug_session_{os.getpid()}_{int(time.time())}.log")
                session_log = run_powershell._csinfo_session_log
            except Exception:
                session_log = os.path.join(tempfile.gettempdir(), f"csinfo_debug_session_{os.getpid()}_{int(time.time())}.log")

    def _write_debug_entry(full_cmd, computer, to, dur=None, return_code=None, output=None, exc=None):
        """Escreve uma entrada de debug no arquivo de sessão (append) e opcionalmente cria um arquivo individual."""
        try:
            header = "--- CSInfo debug entry ---\n"
            ts = datetime.utcnow().isoformat() + 'Z'
            lines = [header, f"TIMESTAMP_UTC: {ts}\n", f"COMMAND: {full_cmd}\n", f"COMPUTER: {computer}\n", f"TIMEOUT: {to}\n"]
            if dur is not None:
                lines.append(f"DURATION_SECONDS: {dur}\n")
            if return_code is not None:
                lines.append(f"RETURN_CODE: {return_code}\n")
            if exc is not None:
                lines.append(f"EXCEPTION: {exc}\n")
            lines.append("OUTPUT:\n")
            if output:
                try:
                    lines.append(output)
                    if not output.endswith('\n'):
                        lines.append('\n')
                except Exception:
                    lines.append(str(output) + '\n')

            # anexar ao arquivo de sessão
            if session_log:
                try:
                    with open(session_log, 'a', encoding='utf-8', errors='replace') as sfh:
                        sfh.writelines(lines)
                except Exception:
                    pass

            # opcional: manter também arquivos individuais (útil para ferramentas que já esperam esse padrão)
            if os.environ.get('CSINFO_DEBUG_INDIVIDUAL') == '1':
                try:
                    indiv = os.path.join(tempfile.gettempdir(), f"csinfo_debug_{os.getpid()}_{int(time.time())}.log")
                    with open(indiv, 'w', encoding='utf-8') as fh:
                        fh.writelines(lines)
                except Exception:
                    pass
        except Exception:
            pass

    attempt = 0
    backoff = initial_backoff
    last_exc = None
    while attempt < retries:
        attempt += 1
        ts = time.time()
        try:
            # selecionar comando: se houver variantes remotas, usar uma variante baseada na tentativa atual
            if remote_variants:
                idx = min(len(remote_variants)-1, attempt-1)
                sel_cmd = remote_variants[idx]
            else:
                sel_cmd = cmd

            cmd_with_encoding = f"[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; {sel_cmd}"
            full = ['powershell', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', cmd_with_encoding]

            if os.name == 'nt':
                creationflags = getattr(subprocess, 'CREATE_NO_WINDOW', 0x08000000)
                raw = subprocess.check_output(full, stderr=subprocess.STDOUT, text=False, timeout=timeout, creationflags=creationflags)
            else:
                raw = subprocess.check_output(full, stderr=subprocess.STDOUT, text=False, timeout=timeout)
            try:
                out = raw.decode('utf-8', errors='replace')
            except Exception:
                try:
                    out = str(raw)
                except Exception:
                    out = ''
            duration = time.time() - ts
            # debug
            if debug_enabled:
                try:
                    _write_debug_entry(full, computer_name, timeout, dur=duration, return_code=0, output=out)
                except Exception:
                    pass
            return out.strip()
        except subprocess.CalledProcessError as cpe:
            last_exc = cpe
            # gravar debug e sair (erro do comando)
            if debug_enabled:
                try:
                    _write_debug_entry(full, computer_name, timeout, dur=(time.time()-ts), return_code=getattr(cpe, 'returncode', 'ERR'), output=getattr(cpe, 'output', ''))
                except Exception:
                    pass
            # não retryar em caso de erro específico de execução (mas permitiremos retry em timeout)
            break
        except Exception as exc:
            last_exc = exc
            # se não for a última tentativa, esperar e retryar
            if debug_enabled:
                try:
                    _write_debug_entry(full, computer_name, timeout, dur=(time.time()-ts), exc=exc)
                except Exception:
                    pass
            if attempt < retries:
                time.sleep(backoff)
                backoff *= 2
                continue
            break

    return ""

    # Se falhou e estamos executando remotamente sem credencial, oferecer prompt interativo (uma vez)
    # Observação: isso só ocorre se não passamos credential explicitamente
    if computer_name and not credential:
        should_prompt = os.environ.get('CSINFO_PROMPT_CREDS', '0') == '1' or sys.stdin.isatty()
        if should_prompt:
            try:
                import getpass
                prompt_user = input(f"Credenciais necessárias para {computer_name} (formato DOMAIN\\user): ")
                if prompt_user:
                    prompt_pwd = getpass.getpass(f"Senha para {prompt_user}: ")
                    if prompt_pwd is None:
                        return ""
                    # tentar novamente com credencial fornecida
                    return run_powershell(original_cmd, computer_name=computer_name, timeout=timeout, retries=max(2, retries), initial_backoff=initial_backoff, credential=(prompt_user, prompt_pwd))
            except Exception:
                pass
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
        cmd = r"""
        $apps = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue | Where-Object {$_.DisplayName -like "*Office*" -or $_.DisplayName -like "*Microsoft 365*"}
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
        # fallback: se a saída não for JSON, tentar dividir por linhas
        if out:
            for l in out.splitlines():
                if l.strip():
                    items.append((l, "", "", ""))
    return items or [("NÃO OBTIDO", "", "", "")]

# --- Wrappers / aliases para compatibilidade com teste/harness ---
def get_disks_short(computer_name=None):
    """Alias antigo/curto usado pelo test harness -> retorna drives lógicos compactos."""
    return get_logical_drives_info(computer_name=computer_name)

def get_monitors(computer_name=None):
    return get_monitor_infos(computer_name=computer_name)

def get_kbd_mouse(computer_name=None):
    return get_keyboard_mouse_status(computer_name=computer_name)

def get_nics(computer_name=None):
    return get_network_adapters_info(computer_name=computer_name)

def get_video(computer_name=None):
    return get_video_cards_info(computer_name=computer_name)

def get_installed(computer_name=None):
    return get_installed_software(computer_name=computer_name)

def get_winupdate(computer_name=None):
    return get_windows_update_status(computer_name=computer_name)
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
                # Para evitar falsos positivos em servidores/headless, só considerar portátil
                # se houver pelo menos um monitor detectado (notebooks normalmente reportam monitores)
                try:
                    mons = get_monitor_infos(computer_name)
                    # mons retorna lista de dicts ou lista vazia; considerar portátil apenas se houver monitores válidos
                    if mons and any(m and (m.get('Fabricante') or m.get('Modelo') or m.get('Serial')) for m in mons):
                        return True
                    # caso contrário, não assumir portátil apenas pelo chassi
                except Exception:
                    # se não for possível obter monitores, assumir conservadoramente que NÃO é laptop
                    pass
    except Exception:
        pass
    return False

def get_chassis_type_name(computer_name=None):
    """Retorna um nome legível baseado no(s) ChassisTypes retornado(s) pelo WMI.

    Mapeamento baseado na tabela comum de ChassisTypes:
      3 -> Desktop
      4 -> Low Profile Desktop
      5 -> Pizza Box
      8,9,10,14 -> Notebook / Laptop / Portable
      12 -> Docking Station
      16 -> Tablet
      17 -> Convertible
      23 -> Server
    Se vários valores forem retornados, retorna o primeiro que bater no mapeamento.
    Caso não seja possível determinar, retorna 'Desconhecido'.
    """
    try:
        cmd = "(Get-CimInstance Win32_SystemEnclosure | Select-Object -ExpandProperty ChassisTypes | ConvertTo-Json -Compress)"
        out = run_powershell(cmd, computer_name=computer_name)
        arr = json.loads(out) if out else []
        if isinstance(arr, int):
            arr = [arr]
    except Exception:
        arr = []

    mapping = {
        3: 'Desktop',
        4: 'Low Profile Desktop',
        5: 'Pizza Box',
        8: 'Notebook',
        9: 'Notebook',
        10: 'Notebook',
        12: 'Docking Station',
        14: 'Notebook',
        16: 'Tablet',
        17: 'Convertible',
        23: 'Server'
    }

    for v in arr:
        try:
            iv = int(v)
            if iv in mapping:
                return mapping[iv]
        except Exception:
            continue

    # fallback: se não houver chassi conhecido, tentar heurística antiga
    try:
        if is_laptop(computer_name):
            return 'Notebook'
    except Exception:
        pass
    return 'Desconhecido'

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

def write_report(path, lines, include_debug=False):
    # Remove duplicidades antes de escrever
    filtered_lines = remove_duplicate_lines(lines)
    # Sanitizar linhas: remover caracteres de controle como NUL
    def _sanitize_line(s):
        if s is None:
            return ""
        t = str(s)
        # remove control chars 0x00-0x1F and 0x7F-0x9F
        return re.sub(r'[\x00-\x1f\x7f-\x9f]', '', t)
    # Cabeçalho com versão e data
    try:
        import csinfo as _cs
        ver = getattr(_cs, '__version__', None)
    except Exception:
        ver = None
    # Novo cabeçalho conforme solicitado:
    # CEOsoftware Sistemas
    # CSInfo - Inventário de hardware e software - v{versão}
    import getpass
    user = getpass.getuser()
    # Evitar duplicar a linha "Gerado por:" — se já existe nas linhas coletadas, não adicionamos no cabeçalho
    try:
        has_gerado = any((str(l) or '').strip().lower().startswith('gerado por:') for l in filtered_lines)
    except Exception:
        has_gerado = False

    header = [
        "CEOsoftware Sistemas",
        f"CSInfo - Inventário de hardware e software - v{ver if ver else 'desconhecida'}",
    ]
    if not has_gerado:
        header.append(f"Gerado por: {user}")
    header.append("")
    # Sanitizar todas as linhas antes de escrever
    sanitized = [_sanitize_line(l) for l in filtered_lines]
    with open(path, 'w', encoding='utf-8-sig') as f:
        f.write("\n".join(header + sanitized))
        # Se solicitado, anexar log de sessão de debug ao final (marcado)
        if include_debug:
            try:
                sess = get_debug_session_log()
                if sess:
                    f.write("\n\n=== LOG DE DEBUG (CSINFO) ===\n")
                    with open(sess, 'r', encoding='utf-8', errors='replace') as lf:
                        for l in lf:
                            f.write(_sanitize_line(l))
            except Exception:
                pass

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
    """Gera um relatório em PDF com as informações coletadas - idêntico ao TXT.

    Esta versão adiciona um índice lateral (sidebar) com links para os agrupamentos
    principais e cria bookmarks para facilitar a navegação no PDF.
    """
    if not path:
        return False
    try:
        import os, sys, traceback
        from reportlab.platypus import Table, TableStyle, PageTemplate, Frame, Spacer, Paragraph, Flowable, FrameBreak, NextPageTemplate
        from reportlab.lib.enums import TA_LEFT
        from reportlab.platypus import BaseDocTemplate
        from reportlab.lib.units import inch
        from reportlab.platypus.doctemplate import PageBreak
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        from reportlab.pdfgen import canvas
        import re

        filtered_lines = remove_duplicate_lines(lines)

        # Canvas numerado com rodapé
        class NumberedCanvas(canvas.Canvas):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self._saved_page_states = []

            def showPage(self):
                self._saved_page_states.append(dict(self.__dict__))
                self._startPage()

            def save(self):
                num_pages = len(self._saved_page_states)
                for page_num, state in enumerate(self._saved_page_states):
                    self.__dict__.update(state)
                    self.draw_page_number(page_num + 1, num_pages)
                    super().showPage()
                super().save()

            def draw_page_number(self, page_num, total_pages):
                self.setStrokeColor(colors.Color(0.7, 0.7, 0.7))
                self.setLineWidth(0.5)
                self.line(0.5 * inch, 0.5 * inch, A4[0] - 0.5 * inch, 0.5 * inch)
                self.setFont("Helvetica", 7)
                self.setFillColor(colors.Color(0.25, 0.25, 0.25))
                page_text = f"Página {page_num} de {total_pages}"
                self.drawString(0.5 * inch, 0.3 * inch, page_text)
                footer_text = "CSInfo by CEOsoftware"
                text_width = self.stringWidth(footer_text, "Helvetica", 7)
                x_center = (A4[0] - text_width) / 2
                self.drawString(x_center, 0.3 * inch, footer_text)

        class MapCollectCanvas(canvas.Canvas):
            """Canvas usado apenas para coletar bookmarks/destinos criados durante a
            primeira passagem. Registra nome -> page_number em self.named_map.
            """
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self.named_map = {}
                # manter referência da última instância criada para posterior leitura
                try:
                    MapCollectCanvas._last_instance = self
                except Exception:
                    pass

            def bookmarkPage(self, name):
                try:
                    # registrar o número da página atual para este nome
                    pnum = getattr(self, '_pageNumber', None)
                    if pnum is None:
                        pnum = 1
                    self.named_map[name] = pnum
                    try:
                        # debug: log each bookmark observed during collection
                        print(f"MapCollectCanvas.bookmarkPage: name={name!r} page={pnum}", file=sys.stderr)
                    except Exception:
                        pass
                except Exception:
                    pass
                try:
                    return super().bookmarkPage(name)
                except Exception:
                    return None
            
            def addOutlineEntry(self, title, key, level=0, closed=None):
                """Intercepta a criação de entradas de outline (bookmarks) para
                registrar também o mapeamento nome->página.
                """
                try:
                    pnum = getattr(self, '_pageNumber', None)
                    if pnum is None:
                        pnum = 1
                    if key:
                        self.named_map[key] = pnum
                except Exception:
                    pass
                try:
                    return super().addOutlineEntry(title, key, level=level, closed=closed)
                except Exception:
                    return None

        # document and layout
        tmp_path = path + ".tmp"
        # Persistir sidecar em pasta temporária do sistema para evitar
        # vazamento de arquivos .annots.jsonl junto ao diretório de export
        try:
            import tempfile, uuid
            tmp_dir = tempfile.gettempdir()
            unique = uuid.uuid4().hex
            # nome legível baseado no nome do pdf final + sufixo único
            annots_sidecar = os.path.join(tmp_dir, f"{os.path.basename(path)}.tmp.{unique}.annots.jsonl")
        except Exception:
            # fallback para colocá-lo ao lado do tmp (comportamento anterior)
            annots_sidecar = tmp_path + '.annots.jsonl'
        try:
            # log curto para diagnóstico: onde o sidecar será escrito
            print(f"annots_sidecar path: {annots_sidecar}", file=sys.stderr)
        except Exception:
            pass
        doc = BaseDocTemplate(tmp_path, pagesize=A4)
        sidebar_width = 1.0 * inch
        gap = 8
        left_x = 0.5 * inch
        left_y = 0.6 * inch
        left_h = A4[1] - 1.35 * inch
        left_frame = Frame(
            left_x,
            left_y,
            sidebar_width - gap,
            left_h,
            leftPadding=6,
            bottomPadding=6,
            rightPadding=6,
            topPadding=6,
            id='left'
        )

        # main_frame alinhado à esquerda sob o logo (logo x ~= 0.6*inch)
        logo_x = 0.6 * inch
        main_x = logo_x
        # largura: margem esquerda+direita (0.5 in) subtraída do espaço do logo
        main_frame = Frame(
            main_x,
            0.6 * inch,
            A4[0] - main_x - 0.5 * inch,
            A4[1] - 1.35 * inch,
            leftPadding=0,
            bottomPadding=6,
            rightPadding=6,
            topPadding=6,
            id='main'
        )

    # coletor de seções para o índice
        toc_sections = []

        def add_toc_section(title, dest):
            """Adiciona uma entrada ao toc_sections garantindo unicidade por destino."""
            try:
                for t, d in toc_sections:
                    if d == dest:
                        return
                toc_sections.append((title, dest))
            except Exception:
                try:
                    if (title, dest) not in toc_sections:
                        toc_sections.append((title, dest))
                except Exception:
                    pass

        class SectionAnchor(Flowable):
            def __init__(self, name, title):
                super().__init__()
                self.name = name
                self.title = title

            def draw(self):
                try:
                    self.canv.bookmarkPage(self.name)
                    # tentar também registrar diretamente no coletor de mapa (caso bookmarkPage
                    # não seja interceptado pela MapCollectCanvas por algum motivo)
                    try:
                        coll = getattr(MapCollectCanvas, '_last_instance', None)
                        if coll is not None:
                            pnum = getattr(self.canv, 'getPageNumber', None)
                            try:
                                pnum = self.canv.getPageNumber()
                            except Exception:
                                pnum = getattr(self.canv, '_pageNumber', None) or 1
                            coll.named_map[self.name] = pnum
                            try:
                                print(f"SectionAnchor.registered directly: name={self.name!r} page={pnum}", file=sys.stderr)
                            except Exception:
                                pass
                    except Exception:
                        pass
                    try:
                        self.canv.addOutlineEntry(self.title, self.name, level=0, closed=False)
                    except Exception:
                        pass
                except Exception:
                    pass

        class LinkedParagraph(Flowable):
            """Paragraph-like flowable that also creates an internal PDF link rectangle
            to a named destination when drawn. The link area covers the paragraph box.
            """
            def __init__(self, text, style, destname=None):
                super().__init__()
                self.para = Paragraph(text, style)
                self.dest = destname

            def wrap(self, availWidth, availHeight):
                w, h = self.para.wrap(availWidth, availHeight)
                self._w = w
                self._h = h
                return w, h

            def draw(self):
                try:
                    # draw paragraph at origin
                    self.para.drawOn(self.canv, 0, 0)
                    if self.dest:
                        try:
                            # registrar a área do link para criação posterior
                            try:
                                rect = self.canv._absRect((0, 0, self._w, self._h), relative=1)
                            except Exception:
                                rect = (0, 0, self._w, self._h)
                            try:
                                pnum = getattr(self.canv, 'getPageNumber', None)
                                if pnum:
                                    pnum = self.canv.getPageNumber()
                                else:
                                    pnum = getattr(self.canv, '_pageNumber', 1)
                            except Exception:
                                pnum = getattr(self.canv, '_pageNumber', 1)
                            try:
                                # tentar resolver destino para número de página (segunda passagem)
                                tgt_page = None
                                try:
                                    named_map_local = None
                                    named_map_local = getattr(self.canv, '_named_map', None) or getattr(MapCollectCanvas, '_named_map', None)
                                    if named_map_local and isinstance(named_map_local, dict):
                                        if isinstance(self.dest, str) and self.dest in named_map_local:
                                            tgt_page = named_map_local[self.dest]
                                except Exception:
                                    tgt_page = None
                                annotations_to_create.append((pnum, rect, self.dest))
                                # persistir em sidecar para pós-processamento confiável
                                try:
                                    import json
                                    payload = {'page': pnum, 'rect': rect, 'dest': self.dest}
                                    if tgt_page:
                                        payload['tgt_page'] = int(tgt_page)
                                    with open(annots_sidecar, 'a', encoding='utf-8') as _af:
                                        _af.write(json.dumps(payload) + '\n')
                                except Exception:
                                    pass
                            except Exception:
                                pass
                            # tentar criar a anotação diretamente usando o nome do destino
                            try:
                                # relative=0 garante coordenadas em coordenadas de página
                                self.canv.linkRect('', self.dest, rect, relative=0)
                            except Exception:
                                pass
                        except Exception:
                            pass
                except Exception:
                    pass

        def draw_header(canvas_obj, doc_obj):
            canvas_obj.saveState()
            y_title = A4[1] - 0.7 * inch if doc_obj.page == 1 else A4[1] - 0.6 * inch
            # desenhar logotipo pequeno à esquerda, se disponível
            try:
                import os
                from reportlab.lib.utils import ImageReader
                logo_path = os.path.join(os.path.dirname(__file__), '..', 'assets', 'ico.png')
                logo_path = os.path.abspath(logo_path)
                if os.path.exists(logo_path):
                    try:
                        img = ImageReader(logo_path)
                        iw, ih = img.getSize()
                        desired_w = 0.38 * inch
                        scale = desired_w / float(iw) if iw else 1.0
                        desired_h = float(ih) * scale
                        x_img = 0.6 * inch
                        # subir levemente o logo para ficar alinhado com o título
                        logo_offset_up = 0.08 * inch
                        y_img = y_title - (desired_h / 2.0) + logo_offset_up
                        canvas_obj.drawImage(img, x_img, y_img, width=desired_w, height=desired_h, preserveAspectRatio=True, mask='auto')
                    except Exception:
                        pass
            except Exception:
                pass
            canvas_obj.setFillColor(colors.navy)
            canvas_obj.setFont("Helvetica-Bold", 11)
            canvas_obj.drawCentredString(A4[0] / 2, y_title, "CSInfo - Inventário de hardware e software")
            try:
                import csinfo as _cs
                _ver = getattr(_cs, '__version__', None)
            except Exception:
                _ver = None
            ver_text = f"v{_ver}" if _ver else "vdesconhecida"
            try:
                canvas_obj.setFont("Helvetica", 8)
                ver_x = A4[0] - (0.6 * inch)
                canvas_obj.drawRightString(ver_x, y_title, ver_text)
            except Exception:
                pass
            canvas_obj.setStrokeColor(colors.Color(0.7, 0.7, 0.7))
            canvas_obj.setLineWidth(0.5)
            canvas_obj.line(0.5 * inch, y_title - (0.12 * inch), A4[0] - 0.5 * inch, y_title - (0.12 * inch))
            canvas_obj.restoreState()

        def draw_sidebar_on_canvas(canvas_obj, doc_obj):
            # desenha o índice no lado esquerdo da primeira página diretamente no canvas
            try:
                if not toc_sections:
                    return
                try:
                    if not pdf_sidebar_enabled:
                        return
                except Exception:
                    pass
                # área do sidebar
                try:
                    x = left_x
                    w = sidebar_width - gap
                    y_top = A4[1] - 1.05 * inch
                    y = y_top
                    line_h = 12
                    # background do sidebar
                    canvas_obj.saveState()
                    canvas_obj.setFillColor(colors.Color(0.2, 0.2, 0.2))
                    canvas_obj.rect(x, 0.6 * inch, w, left_h, fill=1, stroke=0)
                    canvas_obj.setFont('Helvetica', 8)
                    canvas_obj.setFillColor(colors.whitesmoke)
                    # desenhar cada item com link para o destino
                    for title, dest in toc_sections:
                        if y - (line_h - 2) < 0.7 * inch:
                            break
                        # desenhar texto e tentar criar link (registrar para pós-processamento e
                        # também chamar canvas.linkRect como fallback imediato)
                        try:
                            canvas_obj.drawString(x + 6, y, clean_text(title))
                            txt_w = canvas_obj.stringWidth(clean_text(title), 'Helvetica', 8)
                            rect = (x + 6, y - 2, x + 6 + txt_w, y + 10)
                            pnum = getattr(canvas_obj, 'getPageNumber', None)
                            if pnum:
                                try:
                                    pnum = canvas_obj.getPageNumber()
                                except Exception:
                                    pnum = getattr(canvas_obj, '_pageNumber', 1)
                            else:
                                pnum = getattr(canvas_obj, '_pageNumber', 1)
                                try:
                                    # tentar resolver destino para número de página
                                    tgt_page = None
                                    try:
                                        named_map_local = getattr(canvas_obj, '_named_map', None) or getattr(MapCollectCanvas, '_named_map', None)
                                        if named_map_local and isinstance(named_map_local, dict) and isinstance(dest, str) and dest in named_map_local:
                                            tgt_page = named_map_local[dest]
                                    except Exception:
                                        tgt_page = None
                                    annotations_to_create.append((pnum, rect, dest))
                                    try:
                                        import json
                                        payload = {'page': pnum, 'rect': rect, 'dest': dest}
                                        if tgt_page:
                                            payload['tgt_page'] = int(tgt_page)
                                        with open(annots_sidecar, 'a', encoding='utf-8') as _af:
                                            _af.write(json.dumps(payload) + '\n')
                                    except Exception:
                                        pass
                                except Exception:
                                    pass
                            try:
                                canvas_obj.linkRect('', dest, rect, relative=0)
                            except Exception:
                                pass
                        except Exception:
                            # se falhar ao desenhar/registrar este item, ignore e prossiga
                            pass
                        y -= line_h
                    canvas_obj.restoreState()
                except Exception:
                    pass
            except Exception:
                pass

        def draw_page(canvas_obj, doc_obj):
            # Cabeçalho com título e logo
            draw_header(canvas_obj, doc_obj)
            # desenhar sidebar interativo apenas na primeira página
            try:
                if doc_obj.page == 1:
                    draw_sidebar_on_canvas(canvas_obj, doc_obj)
            except Exception:
                pass

        # coletará anotações para criar no pós-processamento: (page_num, rect, destname)
        story = []
        annotations_to_create = []
        styles = getSampleStyleSheet()
        # header para o topo e títulos das seções
        header_style = ParagraphStyle('CSInfoHeader', parent=styles['Normal'], fontSize=10, textColor=colors.navy, alignment=TA_LEFT, spaceAfter=6, leading=14, fontName='Helvetica-Bold')
        # padronizar espaçamento/leading e usar cores suaves de fundo por seção
        TITLE_LEADING = 14
        TITLE_SPACE_BEFORE = 1
        TITLE_SPACE_AFTER = 1
        section_title_styles = {
            "INFORMAÇÕES DO SISTEMA": ParagraphStyle('SectionTitleSistema', parent=styles['Heading2'], fontSize=10, spaceAfter=TITLE_SPACE_AFTER, spaceBefore=TITLE_SPACE_BEFORE, leading=TITLE_LEADING, textColor=colors.whitesmoke, fontName='Helvetica-Bold', alignment=TA_LEFT),
            "INFORMAÇÕES DE HARDWARE": ParagraphStyle('SectionTitleHardware', parent=styles['Heading2'], fontSize=10, spaceAfter=TITLE_SPACE_AFTER, spaceBefore=TITLE_SPACE_BEFORE, leading=TITLE_LEADING, textColor=colors.whitesmoke, fontName='Helvetica-Bold', alignment=TA_LEFT),
            "ADMINISTRADORES": ParagraphStyle('SectionTitleAdmin', parent=styles['Heading2'], fontSize=10, spaceAfter=TITLE_SPACE_AFTER, spaceBefore=TITLE_SPACE_BEFORE, leading=TITLE_LEADING, textColor=colors.whitesmoke, fontName='Helvetica-Bold', alignment=TA_LEFT),
            "SOFTWARES INSTALADOS": ParagraphStyle('SectionTitleSoft', parent=styles['Heading2'], fontSize=10, spaceAfter=TITLE_SPACE_AFTER, spaceBefore=TITLE_SPACE_BEFORE, leading=TITLE_LEADING, textColor=colors.whitesmoke, fontName='Helvetica-Bold', alignment=TA_LEFT),
            "INFORMAÇÕES DE REDE": ParagraphStyle('SectionTitleNet', parent=styles['Heading2'], fontSize=10, spaceAfter=TITLE_SPACE_AFTER, spaceBefore=TITLE_SPACE_BEFORE, leading=TITLE_LEADING, textColor=colors.whitesmoke, fontName='Helvetica-Bold', alignment=TA_LEFT),
            "SEGURANÇA DO SISTEMA": ParagraphStyle('SectionTitleSec', parent=styles['Heading2'], fontSize=10, spaceAfter=TITLE_SPACE_AFTER, spaceBefore=TITLE_SPACE_BEFORE, leading=TITLE_LEADING, textColor=colors.whitesmoke, fontName='Helvetica-Bold', alignment=TA_LEFT),
        }
        # cores mais saturadas/escurecidas para contraste com texto branco
        section_bg_colors = {
            "IDENTIFICAÇÃO": colors.Color(0.20, 0.30, 0.40),  
            "INFORMAÇÕES DO SISTEMA": colors.Color(0.10, 0.45, 0.70),  
            "INFORMAÇÕES DE HARDWARE": colors.Color(0.10, 0.60, 0.25),  
            "INFORMAÇÕES DE REDE": colors.Color(0.05, 0.55, 0.55),              
            "SEGURANÇA DO SISTEMA": colors.Color(0.15, 0.25, 0.25),  
            "ADMINISTRADORES": colors.Color(0.70, 0.45, 0.10),  
            "SOFTWARES INSTALADOS": colors.Color(0.50, 0.20, 0.10),  
        }
        # padronizar paddings das células de título
        TITLE_LEFTPAD = 6
        TITLE_RIGHTPAD = 6
        TITLE_TOPPAD = 8
        TITLE_BOTTOMPAD = 8
        # gap vertical padrão entre blocos de título
        TITLE_GAP = 8
        # Small gap specifically used after certain titles to tighten spacing
        # Reduced to 0 to remove extra spacing between certain consecutive sections
        SMALL_TITLE_GAP = 0
        normal_style = ParagraphStyle('CustomNormal', parent=styles['Normal'], fontSize=9, spaceAfter=2, leading=11, textColor=colors.black, fontName='Helvetica')
        indented_style = ParagraphStyle('IndentedNormal', parent=styles['Normal'], fontSize=9, spaceAfter=2, leading=11, textColor=colors.black, fontName='Helvetica', leftIndent=12)
        double_indented_style = ParagraphStyle('DoubleIndentedNormal', parent=styles['Normal'], fontSize=9, spaceAfter=2, leading=11, textColor=colors.black, fontName='Helvetica', leftIndent=24)

        def clean_text(text):
            if text is None:
                return ""
            txt = str(text)
            txt = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', txt)
            return (txt.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;').replace("'", '&#39;'))


        # IDENTIFICAÇÃO
        try:
            id_keys = {'Nome do computador': None, 'Tipo': None, 'Gerado por': None, 'Relatório gerado em': None}
            matched_indices = set()
            for idx, ln in enumerate(filtered_lines):
                if not ln:
                    continue
                for k in list(id_keys.keys()):
                    if ln.startswith(k):
                        try:
                            if ':' in ln:
                                id_keys[k] = ln.split(':', 1)[1].strip()
                            else:
                                id_keys[k] = ''
                        except Exception:
                            id_keys[k] = ''
                        matched_indices.add(idx)
                        break
            # ajustar Tipo se for linha de disco
            try:
                tipo_val = id_keys.get('Tipo')
                if tipo_val and re.search(r'\bHDD\b|\bSSD\b|Interface:|\|', tipo_val, flags=re.IGNORECASE):
                    tipo_chassi = None
                    try:
                        tipo_chassi = get_chassis_type_name(computer_name)
                    except Exception:
                        tipo_chassi = None
                    if tipo_chassi and tipo_chassi != 'Desconhecido':
                        id_keys['Tipo'] = tipo_chassi
                    else:
                        try:
                            id_keys['Tipo'] = 'Notebook' if is_laptop(computer_name) else 'Desktop'
                        except Exception:
                            id_keys['Tipo'] = 'Desconhecido'
            except Exception:
                pass

            id_title_style = ParagraphStyle('IdTitle', parent=styles['Normal'], fontSize=10, leading=12, alignment=TA_LEFT, fontName='Helvetica-Bold', textColor=colors.white)
            id_dest = 'sec_IDENTIFICACAO'
            add_toc_section('IDENTIFICAÇÃO', id_dest)
            # garantir anchor literal também: criar_titulo_pdf insere <a name='IDENTIFICAÇÃO'/>
            story.append(SectionAnchor(id_dest, 'IDENTIFICAÇÃO'))
            title_para = criar_titulo_pdf('IDENTIFICAÇÃO', styles=styles)
            table_width = A4[0] - inch
            tbl = Table([[title_para]], colWidths=[table_width])
            try:
                bg = section_bg_colors.get('IDENTIFICAÇÃO', colors.Color(0.8, 0.8, 0.8))
            except Exception:
                bg = colors.Color(0.8, 0.8, 0.8)
            tbl.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (0, 0), bg),
                ('LEFTPADDING', (0, 0), (0, 0), TITLE_LEFTPAD),
                ('RIGHTPADDING', (0, 0), (0, 0), TITLE_RIGHTPAD),
                ('TOPPADDING', (0, 0), (0, 0), TITLE_TOPPAD),
                ('BOTTOMPADDING', (0, 0), (0, 0), TITLE_BOTTOMPAD),
            ]))
            story.append(tbl)
            story.append(Spacer(1, TITLE_GAP))
            for k in ('Nome do computador', 'Tipo', 'Gerado por', 'Relatório gerado em'):
                v = id_keys.get(k)
                if v is None:
                    for ln in filtered_lines:
                        if ln.startswith(k) and ':' not in ln:
                            v = ln.strip()
                            break
                if v is None:
                    continue
                try:
                    story.append(Paragraph(f"{clean_text(k)}: <b>{clean_text(v)}</b>", normal_style))
                except Exception:
                    try:
                        story.append(Paragraph(f"{clean_text(k)}: {clean_text(v)}", normal_style))
                    except Exception:
                        pass
            story.append(Spacer(1, 8))
            try:
                body_lines = [ln for i, ln in enumerate(filtered_lines) if i not in matched_indices]
            except Exception:
                body_lines = list(filtered_lines)
        except Exception:
            body_lines = list(filtered_lines)

        # carregar cores
        try:
            import csinfo as _cs
            pdf_field_colors = getattr(_cs, '__pdf_field_colors__', {}) or {}
            pdf_sidebar_enabled = getattr(_cs, '__pdf_sidebar_enabled__', True)
        except Exception:
            pdf_field_colors = {}
            pdf_sidebar_enabled = True

        # Filtrar eventuais linhas que representam o próprio índice (várias variantes)
        def _normalize(s):
            try:
                import unicodedata
                s2 = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
                return s2.lower().strip()
            except Exception:
                return s.lower().strip()

        index_aliases = {'indice', 'índice', 'index'}
        cleaned_body = []
        for ln in body_lines:
            if not ln:
                cleaned_body.append(ln)
                continue
            norm = _normalize(ln.strip())
            # Filtrar qualquer linha que contenha a palavra 'indice' (com ou sem acento)
            if any(alias in norm for alias in index_aliases):
                # pular cabeçalhos explícitos do índice
                continue
            # pular linhas muito curtas que correspondam exatamente a títulos do toc
            if any(_normalize(ln.strip()) == _normalize(t) for t, _ in toc_sections):
                continue
            cleaned_body.append(ln)

        iter_lines = cleaned_body

        # Sumário removido: não inserir página de sumário no documento.
        # Os bookmarks (SectionAnchor) são mantidos para navegação via outline,
        # mas agora transformamos títulos de seção em agrupamentos: o corpo
        # mostra apenas o título (resumo) e um link "Ver detalhes"; os detalhes
        # completos de cada agrupamento são adicionados ao final do documento
        # em seções separadas. Isso simula um comportamento de expand/retract
        # sem Javascript, usando links/named destinations.

        # Preparar coleta de linhas por grupo
        group_contents = {}
        current_group = None
        collecting_group = False

        for idx, line in enumerate(iter_lines):
            line_stripped = line.strip()
            if not line_stripped:
                story.append(Spacer(1, 6))
                continue

            # Ignorar linhas que sejam o título do índice lateral ou que coincidam
            # exatamente com qualquer título do índice para evitar repetição no corpo
            try:
                if re.match(r'^\s*Índice\b', line_stripped, flags=re.IGNORECASE):
                    continue
            except Exception:
                pass
            try:
                if any(line_stripped == t for t, _ in toc_sections):
                    # já será desenhado na sidebar, não colocar no corpo
                    continue
            except Exception:
                pass

            # Só adicionar títulos de agrupamento reais no corpo (nunca o texto 'Índice' ou qualquer item do sidebar)
            # Ignorar títulos que contenham 'indice' inadvertidamente
            if any(alias in _normalize(line_stripped) for alias in index_aliases):
                continue
            if line_stripped in section_title_styles:
                # iniciar novo agrupamento: registrar destino e criar resumo
                destname = 'sec_' + re.sub(r'[^0-9a-zA-Z_]', '_', line_stripped).upper()
                detail_dest = 'detail_' + destname
                # âncora do título-resumo (bookmark)
                story.append(SectionAnchor(destname, line_stripped))
                try:
                    add_toc_section(line_stripped, destname)
                except Exception:
                    pass
                try:
                    # criar título com link para a seção de detalhes
                    link_text = f"{clean_text(line_stripped)} [ver detalhes]"
                    # usar LinkedParagraph para garantir área de link funcional
                    title_para = LinkedParagraph(link_text, section_title_styles.get(line_stripped, header_style), destname=detail_dest)
                    table_width = A4[0] - inch
                    tbl = Table([[title_para]], colWidths=[table_width])
                    try:
                        bg = section_bg_colors.get(line_stripped, colors.Color(0.5, 0.5, 0.5))
                    except Exception:
                        bg = colors.Color(0.5, 0.5, 0.5)
                    # ajustar paddings para títulos consecutivos específicos
                    try:
                        pad_top = TITLE_TOPPAD
                        pad_bottom = TITLE_BOTTOMPAD
                        # olhar próxima linha não-vazia para detectar título seguinte
                        next_title = None
                        for k in range(idx+1, len(iter_lines)):
                            nxt = iter_lines[k].strip() if iter_lines[k] else ''
                            if not nxt:
                                continue
                            if nxt in section_title_styles:
                                next_title = nxt
                            break
                        # reduzir padding quando for o par SISTEMA -> HARDWARE
                        if str(line_stripped).strip().upper() == 'INFORMAÇÕES DO SISTEMA' and next_title and str(next_title).strip().upper() == 'INFORMAÇÕES DE HARDWARE':
                            pad_bottom = 0
                        # também reduzir o top padding do HARDWARE se o anterior foi SISTEMA
                        if str(line_stripped).strip().upper() == 'INFORMAÇÕES DE HARDWARE':
                            # checar o título anterior simples (procura reversa)
                            prev_title = None
                            for k in range(idx-1, -1, -1):
                                prv = iter_lines[k].strip() if iter_lines[k] else ''
                                if not prv:
                                    continue
                                if prv in section_title_styles:
                                    prev_title = prv
                                break
                            if prev_title and str(prev_title).strip().upper() == 'INFORMAÇÕES DO SISTEMA':
                                pad_top = 0
                    except Exception:
                        pad_top = TITLE_TOPPAD
                        pad_bottom = TITLE_BOTTOMPAD
                    tbl.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (0, 0), bg),
                        ('LEFTPADDING', (0, 0), (0, 0), TITLE_LEFTPAD),
                        ('RIGHTPADDING', (0, 0), (0, 0), TITLE_RIGHTPAD),
                        ('TOPPADDING', (0, 0), (0, 0), pad_top),
                        ('BOTTOMPADDING', (0, 0), (0, 0), pad_bottom),
                    ]))
                    story.append(tbl)
                    try:
                        # Para o par INFORMAÇÕES DO SISTEMA -> INFORMAÇÕES DE HARDWARE,
                        # usar um Spacer negativo para sobrepor levemente os blocos
                        if str(line_stripped).strip().upper() == 'INFORMAÇÕES DO SISTEMA' and next_title and str(next_title).strip().upper() == 'INFORMAÇÕES DE HARDWARE':
                            # usar gap pequeno de 2 pontos entre os blocos
                            story.append(Spacer(1, 2))
                        else:
                            gap = SMALL_TITLE_GAP if str(line_stripped).strip().upper() == 'INFORMAÇÕES DO SISTEMA' else TITLE_GAP
                            story.append(Spacer(1, gap))
                    except Exception:
                        story.append(Spacer(1, TITLE_GAP))
                except Exception:
                    try:
                        story.append(Paragraph(clean_text(line_stripped), section_title_styles.get(line_stripped, header_style)))
                        try:
                            gap = SMALL_TITLE_GAP if str(line_stripped).strip().upper() == 'INFORMAÇÕES DO SISTEMA' else TITLE_GAP
                            story.append(Spacer(1, gap))
                        except Exception:
                            pass
                    except Exception:
                        pass
                # iniciar coleta de linhas do grupo (não inserir conteúdo no corpo)
                current_group = destname
                collecting_group = True
                group_contents[current_group] = []
                continue

            if line_stripped == "CSInfo by CEOsoftware":
                continue

            if collecting_group and current_group:
                # armazenar linha no grupo atual
                group_contents[current_group].append(line)
                continue

            if line.startswith("    "):
                story.append(Paragraph(clean_text(line_stripped), double_indented_style))
            elif line.startswith("  "):
                story.append(Paragraph(clean_text(line_stripped), indented_style))
            else:
                important_prefixes = ['Nome do computador', 'Tipo', 'Versão do sistema operacional', 'Antivírus', 'Memória RAM total', 'Processador']
                # Keys for which we want the VALUE (right-hand side) to be bolded
                value_bold_keys = {'Versão do sistema operacional', 'Antivírus', 'Memória RAM total', 'Processador'}
                if ':' in line_stripped:
                    left, right = line_stripped.split(':', 1)
                    left_clean = clean_text(left).strip()
                    right_clean = clean_text(right).strip()
                    left_check = re.sub(r"\s+\d+$", '', left_clean)
                    if any(left_check.startswith(p) for p in important_prefixes):
                        color = None
                        try:
                            color = pdf_field_colors.get(left_clean) or pdf_field_colors.get(left_check)
                        except Exception:
                            color = None
                        try:
                            # If this key is in the set that requires the VALUE to be bold,
                            # render as: Label: <b>Value</b> (with optional color on the value).
                            if left_check in value_bold_keys or left_clean in value_bold_keys:
                                if color:
                                    paragraph_text = f"{left_clean}: <font color=\"{color}\"><b>{right_clean}</b></font>"
                                else:
                                    paragraph_text = f"{left_clean}: <b>{right_clean}</b>"
                                story.append(Paragraph(paragraph_text, normal_style))
                                continue
                        except Exception:
                            pass
                        # Default behavior for other important prefixes: bold the label
                        if color:
                            paragraph_text = f"<font color=\"{color}\"><b>{left_clean}:</b> {right_clean}</font>"
                        else:
                            paragraph_text = f"<b>{left_clean}:</b> {right_clean}"
                        story.append(Paragraph(paragraph_text, normal_style))
                        continue
                else:
                    for p in important_prefixes:
                        if line_stripped.startswith(p) and len(line_stripped) <= len(p) + 10:
                            try:
                                color = pdf_field_colors.get(line_stripped) or pdf_field_colors.get(re.sub(r"\s+\d+$", '', line_stripped))
                            except Exception:
                                color = None
                            try:
                                special_bg_keys = {'Nome do computador', 'Tipo', 'Versão do sistema operacional', 'Antivírus', 'Memória RAM total', 'Processador'}
                                if re.sub(r"\s+\d+$", '', line_stripped) in special_bg_keys or line_stripped in special_bg_keys:
                                    if color:
                                        story.append(Paragraph(f"<font color=\"{color}\">{clean_text(line_stripped)}</font>", normal_style))
                                    else:
                                        story.append(Paragraph(clean_text(line_stripped), normal_style))
                                    break
                            except Exception:
                                if color:
                                    story.append(Paragraph(f"<font color=\"{color}\"><b>{clean_text(line_stripped)}</b></font>", normal_style))
                                else:
                                    story.append(Paragraph(f"<b>{clean_text(line_stripped)}</b>", normal_style))
                                break
                            if color:
                                story.append(Paragraph(f"<font color=\"{color}\"><b>{clean_text(line_stripped)}</b></font>", normal_style))
                            else:
                                story.append(Paragraph(f"<b>{clean_text(line_stripped)}</b>", normal_style))
                            break
                    else:
                        story.append(Paragraph(clean_text(line_stripped), normal_style))
                    continue

                story.append(Paragraph(clean_text(line_stripped), normal_style))

        try:
            sess = get_debug_session_log()
            if sess and os.path.exists(sess):
                story.append(PageBreak())
                story.append(Paragraph('<b>LOG DE DEBUG (CSINFO)</b>', section_title_styles.get('INFORMAÇÕES DO SISTEMA', header_style)))
                try:
                    with open(sess, 'r', encoding='utf-8', errors='replace') as lf:
                        for raw in lf:
                            txt = clean_text(raw.rstrip('\n'))
                            if txt:
                                story.append(Paragraph(txt, normal_style))
                except Exception:
                    pass
        except Exception:
            pass

        try:
            # Inserir o índice lateral (como links) no início do story
            try:
                # construir índice usando (titulo, dest) para que o href aponte
                # exatamente para os destinos (dest) já registrados pelo
                # SectionAnchor (ex.: 'sec_IDENTIFICACAO'). Isso evita erros
                # de "undefined destination target" quando o link apontar
                # para um nome que não existe como bookmark.
                if toc_sections:
                    indice_estilo = ParagraphStyle('IndiceInline', parent=styles['Normal'], fontSize=9, textColor=colors.whitesmoke, leftIndent=4, leading=11)
                    idx_flow = []
                    for title, dest in toc_sections:
                        try:
                            link_text = f"<a href='#{dest}'><font color='white'>{clean_text(title)}</font></a>"
                            idx_flow.append(Paragraph(link_text, indice_estilo))
                            idx_flow.append(Spacer(1, 4))
                        except Exception:
                            try:
                                idx_flow.append(Paragraph(clean_text(title), indice_estilo))
                                idx_flow.append(Spacer(1, 4))
                            except Exception:
                                pass
                    # O índice será desenhado diretamente no canvas; não
                    # inserir os flowables do índice no story para evitar
                    # layout em colunas.
                    # (idx_flow descartado)
                    pass
            except Exception:
                pass

            # preparar os templates: 'first' com main_frame deslocado (para não sobrepor o sidebar)
            # e 'later' com frame full-width para as páginas subsequentes
            try:
                main_frame_later = Frame(
                    left_x,
                    0.6 * inch,
                    A4[0] - inch,
                    A4[1] - 1.35 * inch,
                    leftPadding=0,
                    bottomPadding=6,
                    rightPadding=6,
                    topPadding=6,
                    id='main_later'
                )
                template_first = PageTemplate(id='first', frames=[main_frame], onPage=draw_page)
                template_later = PageTemplate(id='later', frames=[main_frame_later], onPage=draw_page)
                doc.addPageTemplates([template_first, template_later])
                # garantir que, após a primeira página, o template 'later' seja usado
                try:
                    # inserir NextPageTemplate apenas se a sidebar estiver habilitada
                    try:
                        if pdf_sidebar_enabled:
                            story = [NextPageTemplate('later')] + story
                    except Exception:
                        pass
                except Exception:
                    pass
            except Exception:
                pass

            # DEBUG: informações de layout para ajudar a diagnosticar páginas em branco / colunas
            try:
                print(f"PDF tmp: {tmp_path}", file=sys.stderr)
                print(f"toc_sections count: {len(toc_sections)}", file=sys.stderr)
                try:
                    print(f"left_frame: x={left_x}, width={sidebar_width - gap}, height={left_h}", file=sys.stderr)
                    print(f"main_frame: x={main_x}, width={A4[0] - inch - sidebar_width - gap}", file=sys.stderr)
                except Exception:
                    pass
                # listar os primeiros flowables do story para depuração
                try:
                    print(f"story length: {len(story)}", file=sys.stderr)
                    for i, f in enumerate(story[:20]):
                        try:
                            tname = type(f).__name__
                            preview = ''
                            if hasattr(f, 'getPlainText'):
                                preview = f.getPlainText()[:80].replace('\n', ' ')
                            elif hasattr(f, 'text'):
                                preview = str(f.text)[:80]
                            else:
                                preview = repr(f)[:80]
                            print(f"[{i}] {tname}: {preview}", file=sys.stderr)
                        except Exception:
                            print(f"[{i}] {type(f).__name__}", file=sys.stderr)
                except Exception:
                    pass
            except Exception:
                pass
            # Adicionar seções de detalhe para cada agrupamento coletado
            try:
                if group_contents:
                    story.append(PageBreak())
                    for sec_dest, lines in list(group_contents.items()):
                        try:
                            detail_dest = 'detail_' + sec_dest
                            # buscar o título legível a partir do toc (se disponível)
                            title_text = sec_dest
                            try:
                                for t, d in toc_sections:
                                    if d == sec_dest:
                                        title_text = t
                                        break
                            except Exception:
                                pass
                            # âncora para destino de detalhes: criar via SectionAnchor (bookmark)
                            try:
                                story.append(SectionAnchor(detail_dest, f"Detalhes: {title_text}"))
                            except Exception:
                                pass
                            # título de detalhe com link de volta (bloco colorido)
                            try:
                                # construir paragraph com texto branco
                                detail_text = f"Detalhes: {clean_text(title_text)} [voltar]"
                                # criar um estilo de detalhe baseado em header_style com texto branco
                                try:
                                    detail_style = ParagraphStyle(f"Detail_{title_text}", parent=header_style, textColor=colors.whitesmoke, leading=TITLE_LEADING)
                                except Exception:
                                    detail_style = ParagraphStyle('DetailFallback', parent=header_style, textColor=colors.whitesmoke)
                                # Use LinkedParagraph so the '[voltar]' clickable area is recorded
                                # and later turned into a GoTo annotation in post-processing.
                                title_para = LinkedParagraph(f"<b>{detail_text}</b>", detail_style, destname=sec_dest)
                                # selecionar a cor de fundo correspondente ao título principal (sem 'Detalhes: ' prefix)
                                base_title = title_text
                                try:
                                    bg = section_bg_colors.get(base_title, colors.Color(0.2, 0.2, 0.2))
                                except Exception:
                                    bg = colors.Color(0.2, 0.2, 0.2)
                                # criar tabela para fundo colorido e paddings uniformes
                                tblw = Table([[title_para]], colWidths=[A4[0] - inch])
                                tblw.setStyle(TableStyle([
                                    ('BACKGROUND', (0, 0), (0, 0), bg),
                                    ('LEFTPADDING', (0, 0), (0, 0), TITLE_LEFTPAD),
                                    ('RIGHTPADDING', (0, 0), (0, 0), TITLE_RIGHTPAD),
                                    ('TOPPADDING', (0, 0), (0, 0), TITLE_TOPPAD),
                                    ('BOTTOMPADDING', (0, 0), (0, 0), TITLE_BOTTOMPAD),
                                ]))
                                story.append(tblw)
                                # controlar gap (usar SMALL_TITLE_GAP se aplicável)
                                try:
                                    # para detalhes, se a seção base for INFORMAÇÕES DO SISTEMA,
                                    # também usar sobreposição leve para manter consistência visual
                                    if str(base_title).strip().upper() == 'INFORMAÇÕES DO SISTEMA':
                                        # usar gap pequeno de 2 pontos também no bloco de detalhes
                                        story.append(Spacer(1, 2))
                                    else:
                                        gap = SMALL_TITLE_GAP if str(base_title).strip().upper() == 'INFORMAÇÕES DO SISTEMA' else TITLE_GAP
                                        story.append(Spacer(1, gap))
                                except Exception:
                                    story.append(Spacer(1, TITLE_GAP))
                            except Exception:
                                try:
                                    story.append(Paragraph(f"Detalhes: {clean_text(title_text)}", header_style))
                                    story.append(Spacer(1, TITLE_GAP))
                                except Exception:
                                    pass
                            # conteúdo do grupo
                            for raw in lines:
                                try:
                                    ln = raw.rstrip('\n')
                                    if ln.startswith('    '):
                                        story.append(Paragraph(clean_text(ln.strip()), double_indented_style))
                                    elif ln.startswith('  '):
                                        story.append(Paragraph(clean_text(ln.strip()), indented_style))
                                    else:
                                        # Apply same key:value formatting as used in the body so
                                        # specific keys have their VALUES rendered in bold.
                                        try:
                                            line_strip = ln.strip()
                                            important_prefixes = ['Nome do computador', 'Tipo', 'Versão do sistema operacional', 'Antivírus', 'Memória RAM total', 'Processador']
                                            value_bold_keys = {'Versão do sistema operacional', 'Antivírus', 'Memória RAM total', 'Processador'}
                                            if ':' in line_strip:
                                                left, right = line_strip.split(':', 1)
                                                left_clean = clean_text(left).strip()
                                                right_clean = clean_text(right).strip()
                                                left_check = re.sub(r"\s+\d+$", '', left_clean)
                                                color = None
                                                try:
                                                    color = pdf_field_colors.get(left_clean) or pdf_field_colors.get(left_check)
                                                except Exception:
                                                    color = None
                                                try:
                                                    if left_check in value_bold_keys or left_clean in value_bold_keys:
                                                        if color:
                                                            paragraph_text = f"{left_clean}: <font color=\"{color}\"><b>{right_clean}</b></font>"
                                                        else:
                                                            paragraph_text = f"{left_clean}: <b>{right_clean}</b>"
                                                        story.append(Paragraph(paragraph_text, normal_style))
                                                        continue
                                                except Exception:
                                                    pass
                                                if left_check in important_prefixes or left_clean in important_prefixes:
                                                    if color:
                                                        paragraph_text = f"<font color=\"{color}\"><b>{left_clean}:</b> {right_clean}</font>"
                                                    else:
                                                        paragraph_text = f"<b>{left_clean}:</b> {right_clean}"
                                                    story.append(Paragraph(paragraph_text, normal_style))
                                                    continue
                                            # fallback: plain paragraph
                                            story.append(Paragraph(clean_text(line_strip), normal_style))
                                        except Exception:
                                            story.append(Paragraph(clean_text(ln.strip()), normal_style))
                                except Exception:
                                    pass
                            story.append(Spacer(1, 12))
                        except Exception:
                            pass
            except Exception:
                pass

            # primeira passagem: coletar mapa de nomes -> página
            try:
                try:
                    MapCollectCanvas._last_instance = None
                except Exception:
                    pass
                # limpar sidecar antes da primeira passagem
                try:
                    if os.path.exists(annots_sidecar):
                        os.remove(annots_sidecar)
                except Exception:
                    pass
                try:
                    doc.build(story, canvasmaker=MapCollectCanvas)
                except Exception:
                    pass
                try:
                    coll = getattr(MapCollectCanvas, '_last_instance', None)
                    named_map = coll.named_map if coll is not None else {}
                    try:
                        print('collected_named_map:', named_map, file=sys.stderr)
                    except Exception:
                        pass
                except Exception:
                    named_map = {}
            except Exception:
                named_map = {}

            # segunda passagem: build final com canvas numerado
            try:
                # anexar o mapa coletado às classes/canvas para que os flowables
                # possam criar anotações paginadas na segunda passagem
                try:
                    NumberedCanvas._named_map = named_map
                except Exception:
                    pass
                try:
                    MapCollectCanvas._named_map = named_map
                except Exception:
                    pass
                # limpar sidecar antes da segunda passagem para gravar posições reais
                try:
                    if os.path.exists(annots_sidecar):
                        os.remove(annots_sidecar)
                except Exception:
                    pass
                try:
                    doc.build(story, canvasmaker=NumberedCanvas)
                except Exception:
                    # fallback: tentar uma build simples
                    doc.build(story)
            except Exception:
                # fallback handled above
                pass
            # move temp file to final path atomically
            try:
                # tentar corrigir anotações internas pós-processando o PDF temporário
                def _fix_annotation_destinations(pdf_path, name_map=None, annotations_list=None):
                    try:
                        from pypdf import PdfReader, PdfWriter
                        from pypdf.generic import DictionaryObject, NameObject, ArrayObject, FloatObject, NumberObject
                    except Exception:
                        return False

                    try:
                        rdr = PdfReader(pdf_path)
                        # carregar sidecar se necessário
                        if not annotations_list:
                            annotations_list = []
                            try:
                                import json
                                if os.path.exists(annots_sidecar):
                                    with open(annots_sidecar, 'r', encoding='utf-8') as _af:
                                        for ln in _af:
                                            try:
                                                obj = json.loads(ln)
                                                annotations_list.append((obj.get('page'), tuple(obj.get('rect')), obj.get('dest')))
                                            except Exception:
                                                pass
                            except Exception:
                                annotations_list = annotations_list or []

                        nd = getattr(rdr, 'named_destinations', None) or {}

                        writer = PdfWriter()
                        # append all pages first so writer has its own page objects/refs
                        writer.append_pages_from_reader(rdr)

                        # collect existing annots from writer pages and preserve non-Link ones
                        num_pages = len(writer.pages)
                        page_new_annots = {i: [] for i in range(num_pages)}
                        for i in range(num_pages):
                            try:
                                existing = writer.pages[i].get('/Annots') or []
                            except Exception:
                                existing = []
                            kept = []
                            for a in existing:
                                try:
                                    obj = a.get_object()
                                    subtype = obj.get('/Subtype')
                                    if subtype != '/Link':
                                        kept.append(a)
                                except Exception:
                                    kept.append(a)
                            page_new_annots[i] = kept

                        # create named destinations in writer for every target used
                        dest_to_index = {}
                        for (_pnum, _rect, destname) in (annotations_list or []):
                            if destname in dest_to_index:
                                continue
                            # try to resolve the target page
                            tgt_page = None
                            if name_map and destname in name_map:
                                tgt_page = name_map[destname]
                            elif destname in nd:
                                try:
                                    tgt_page = rdr.get_destination_page_number(nd[destname]) + 1
                                except Exception:
                                    tgt_page = None
                            if not tgt_page:
                                continue
                            tgt_index = int(tgt_page) - 1
                            if tgt_index < 0 or tgt_index >= num_pages:
                                continue
                            dest_to_index[destname] = tgt_index

                        # register named destinations in writer
                        for destname, tgt_index in dest_to_index.items():
                            try:
                                # pypdf expects a name and a page (object or index)
                                writer.add_named_destination(destname, writer.pages[tgt_index])
                            except Exception:
                                try:
                                    writer.add_named_destination(destname, tgt_index)
                                except Exception:
                                    pass

                        # create new annotations and append to the appropriate writer page
                        for (pnum, rect, destname) in (annotations_list or []):
                            try:
                                page_index = int(pnum) - 1
                                if page_index < 0 or page_index >= num_pages:
                                    continue
                                # create an annotation that performs a GoTo to the named destination
                                if destname not in dest_to_index:
                                    continue
                                ann = DictionaryObject()
                                ann[NameObject('/Type')] = NameObject('/Annot')
                                ann[NameObject('/Subtype')] = NameObject('/Link')
                                x0, y0, x1, y1 = rect
                                ann[NameObject('/Rect')] = ArrayObject([FloatObject(x0), FloatObject(y0), FloatObject(x1), FloatObject(y1)])
                                ann[NameObject('/Border')] = ArrayObject([NumberObject(0), NumberObject(0), NumberObject(0)])
                                # action dictionary: /S /GoTo, /D <name>
                                action = DictionaryObject()
                                action[NameObject('/S')] = NameObject('/GoTo')
                                # Named destination should be a Name or string; use NameObject with leading slash
                                try:
                                    action[NameObject('/D')] = NameObject('/' + destname)
                                except Exception:
                                    try:
                                        from pypdf.generic import TextStringObject
                                        action[NameObject('/D')] = TextStringObject(destname)
                                    except Exception:
                                        action[NameObject('/D')] = NameObject('/' + destname)
                                ann[NameObject('/A')] = action
                                try:
                                    newref = writer._add_object(ann)
                                    page_new_annots[page_index].append(newref)
                                except Exception:
                                    page_new_annots[page_index].append(ann)
                            except Exception:
                                pass

                        # assign updated Annots arrays back to writer pages
                        from collections import defaultdict
                        for i in range(num_pages):
                            try:
                                writer.pages[i][NameObject('/Annots')] = ArrayObject(page_new_annots[i])
                            except Exception:
                                pass

                        outp = pdf_path + '.fixed'
                        with open(outp, 'wb') as f:
                            writer.write(f)
                        try:
                            os.replace(outp, pdf_path)
                        except Exception:
                            try:
                                os.remove(pdf_path)
                                os.replace(outp, pdf_path)
                            except Exception:
                                pass
                        return True
                    except Exception:
                        return False

                try:
                    _fix_annotation_destinations(tmp_path, name_map=named_map, annotations_list=annotations_to_create)
                except Exception:
                    pass
                if os.path.exists(tmp_path):
                    os.replace(tmp_path, path)
            except Exception:
                # best-effort; if rename fails, leave temp file for inspection
                pass
            return True
        except Exception:
            tb = traceback.format_exc()
            try:
                print("Erro ao gerar arquivo PDF:", tb, file=sys.stderr)
            except Exception:
                pass
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass
            return False
    except Exception:
        try:
            import traceback as _tb, sys as _sys
            _txt = _tb.format_exc()
            try:
                print("write_pdf_report top-level exception:", _txt, file=_sys.stderr)
            except Exception:
                pass
        except Exception:
            pass
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

def main(export_type=None, barra_callback=None, computer_name=None, include_debug_on_export=False, machine_alias=None):
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
        try:
            emitted_lines.append(line)
        except Exception:
            pass
        if barra_callback:
            try:
                barra_callback(None, line)
            except Exception:
                pass

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
    # Se houver apelido (machine_alias), use o padrão Info_maquina_<apelido>_<nomemaquina>.txt
    if machine_alias and str(machine_alias).strip():
        safe_alias = safe_filename(machine_alias)
        filename = f"Info_maquina_{safe_alias}_{safe_name}.txt"
    else:
        filename = f"Info_maquina_{safe_name}.txt"
    path = os.path.join(os.getcwd(), filename)

    lines = []
    emitted_lines = []
    def padrao(valor):
        return valor if valor and str(valor).strip() else "NÃO OBTIDO"
    
    # LINHAS INICIAIS SOLICITADAS: Nome, Tipo, Gerado por
    # Inserir título textual para exportações em TXT quando o PDF não for gerado
    try:
        add_line("IDENTIFICAÇÃO")
    except Exception:
        pass
    add_line(f"Nome do computador: {machine}")
    # Determinar tipo com base no ChassisTypes quando possível
    tipo_chassi = get_chassis_type_name(computer_name)
    if tipo_chassi and tipo_chassi != 'Desconhecido':
        add_line(f"Tipo: {tipo_chassi}")
    else:
        # fallback conservador: usar heurística de is_laptop
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
    # --- INFORMAÇÕES DE REDE (moveram para cá) ---
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

    # RODAPÉ (mantido mais adiante)
    lines.append("")
    lines.append("")
    lines.append("CSInfo by CEOsoftware")

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

    # Só gera TXT/PDF se export_type for passado explicitamente e for um dos valores esperados
    pdf_path = path.replace('.txt', '.pdf')
    print(f"DEBUG csinfo: export_type={repr(export_type)}, gerar_txt={gerar_txt}, gerar_pdf={gerar_pdf}")
    if export_type in ('txt', 'pdf', 'ambos'):
        if export_type in ('txt', 'ambos'):
            try:
                print(f"DEBUG csinfo: Gravando TXT em: {path}")
                write_report(path, lines, include_debug=include_debug_on_export)
                barra_progresso(23)
                print(f"Arquivo TXT gerado: {path}")
            except Exception as e:
                print('Erro ao gerar TXT:', e)
        if export_type in ('pdf', 'ambos'):
            try:
                print("Gerando arquivo PDF...")
                ok = write_pdf_report(pdf_path, lines, machine)
                if ok:
                    print(f"Arquivo PDF gerado com sucesso: {pdf_path}")
                else:
                    print("Erro ao gerar arquivo PDF.")
                    pdf_path = None
            except Exception as e:
                print('Exceção ao gerar PDF:', e)

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
    # Se em modo GUI e foi passado um barra_callback, enviar qualquer linha
    # que tenha sido adicionada diretamente a `lines` mas não foi encaminhada
    # via add_line durante a execução.
    if barra_callback:
        try:
            for l in lines:
                try:
                    if l not in emitted_lines:
                        barra_callback(None, l)
                except Exception:
                    try:
                        barra_callback(None, str(l))
                    except Exception:
                        pass
        except Exception:
            pass

    return resultado
