# coletar_info_maquina.py
import subprocess
import platform
import os
import re
import sys
import winreg
from datetime import datetime
import json
import time

def run_powershell(cmd, timeout=20):
    full = ['powershell', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', cmd]
    try:
        out = subprocess.check_output(full, stderr=subprocess.STDOUT, text=True, timeout=timeout)
        return out.strip()
    except subprocess.CalledProcessError:
        return ""
    except Exception:
        return ""

def safe_filename(s):
    # remove caracteres inválidos para nome de arquivo
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', s)

def get_machine_name():
    return platform.node() or "Desconhecido"

def get_os_version():
    cmd = "(Get-CimInstance Win32_OperatingSystem | Select-Object -Property Caption,Version | ConvertTo-Json -Compress)"
    out = run_powershell(cmd)
    m_caption = re.search(r'"Caption"\s*:\s*"([^"]+)"', out)
    m_version = re.search(r'"Version"\s*:\s*"([^"]+)"', out)
    caption = m_caption.group(1) if m_caption else ""
    version = m_version.group(1) if m_version else ""
    if caption:
        return f"{caption} (Version {version})" if version else caption
    return platform.platform()

def get_office_version():
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
    return results[0] if results else "Não encontrado"

def get_motherboard_info():
    cmd = "(Get-CimInstance Win32_BaseBoard | Select-Object Manufacturer,Product,SerialNumber | ConvertTo-Json -Compress)"
    out = run_powershell(cmd)
    try:
        info = json.loads(out) if out else {}
        if isinstance(info, list):
            info = info[0] if info else {}
        fabricante = info.get('Manufacturer', '')
        modelo = info.get('Product', '')
        serial = info.get('SerialNumber', '')
        return fabricante, modelo, serial
    except Exception:
        return "", "", "Não encontrado"

def get_monitor_infos():
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
        out = run_powershell(ps)
        try:
                parsed = json.loads(out) if out else []
                if isinstance(parsed, dict):
                        parsed = [parsed]
                return parsed if parsed else [{"Fabricante":"","Modelo":"","Serial":"Nenhum serial encontrado"}]
        except Exception:
                return [{"Fabricante":"","Modelo":"","Serial":"Nenhum serial encontrado"}]

def get_devices_by_class(devclass):
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
    out = run_powershell(ps.format(cls=devclass))
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
    return items or [("Nenhum dispositivo encontrado", "", "", "")]

def get_printers():
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
        out = run_powershell(ps)
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
        return items or [("Nenhuma impressora encontrada", "", "", "")]

def is_laptop():
    # 1) verificar Win32_Battery (se existir, provavelmente notebook)
    out = run_powershell("Get-CimInstance Win32_Battery | ConvertTo-Json -Compress")
    if out and out.strip() != "null":
        return True
    # 2) verificar ChassisTypes em Win32_SystemEnclosure (valores que indicam portátil: 8,9,10,14)
    cmd = "(Get-CimInstance Win32_SystemEnclosure | Select-Object -ExpandProperty ChassisTypes | ConvertTo-Json -Compress)"
    out2 = run_powershell(cmd)
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

def write_report(path, lines):
    with open(path, 'w', encoding='utf-8-sig') as f:
        f.write("\n".join(lines))

def main():
    etapas = [
        "Obtendo nome do computador",
        "Verificando tipo (Notebook/Desktop)",
        "Obtendo versão do sistema operacional",
        "Obtendo versão do Office",
        "Obtendo informações da placa mãe",
        "Obtendo informações dos monitores",
        "Obtendo informações dos teclados",
        "Obtendo informações dos mouses",
        "Obtendo informações das impressoras",
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
        print(f"\r{barra} | {etapas[atual-1]} | Tempo: {tempo}s", end='', flush=True)

    machine = get_machine_name(); barra_progresso(1)
    safe_name = safe_filename(machine)
    filename = f"info_maquina_{safe_name}.txt"
    path = os.path.join(os.getcwd(), filename)

    lines = []
    lines.append(f"Relatório gerado: {datetime.now().isoformat()}")
    lines.append(f"Nome do computador: {machine}"); barra_progresso(2)
    lines.append(f"Tipo: {'Notebook' if is_laptop() else 'Desktop'}"); barra_progresso(3)
    lines.append(f"Versão do sistema operacional: {get_os_version()}"); barra_progresso(4)
    lines.append(f"Versão do Office (se tiver): {get_office_version()}"); barra_progresso(5)
    fabricante_mb, modelo_mb, serial_mb = get_motherboard_info()
    def padrao(valor):
        return valor if valor and str(valor).strip() else "NÃO OBTIDO"
    lines.append(f"Placa mãe: {padrao(fabricante_mb)} | Modelo: {padrao(modelo_mb)} | Serial: {padrao(serial_mb)}"); barra_progresso(6)

    monitors = get_monitor_infos()
    if monitors:
        for idx, m in enumerate(monitors, start=1):
            lines.append(f"Monitor {idx}: {padrao(m.get('Fabricante',''))} | Modelo: {padrao(m.get('Modelo',''))} | Serial: {padrao(m.get('Serial',''))}")
    else:
        lines.append("Monitor 1: NÃO OBTIDO")
    barra_progresso(7)

    keyboards = get_devices_by_class("Keyboard")
    if keyboards:
        for idx, (name, serial, fabricante, modelo) in enumerate(keyboards, start=1):
            lines.append(f"Teclado {idx}: {padrao(name)} | Serial/ID: {padrao(serial)} | Fabricante: {padrao(fabricante)} | Modelo: {padrao(modelo)}")
    else:
        lines.append("Teclado: NÃO OBTIDO")
    barra_progresso(8)

    mice = get_devices_by_class("Mouse")
    if mice:
        for idx, (name, serial, fabricante, modelo) in enumerate(mice, start=1):
            lines.append(f"Mouse {idx}: {padrao(name)} | Serial/ID: {padrao(serial)} | Fabricante: {padrao(fabricante)} | Modelo: {padrao(modelo)}")
    else:
        lines.append("Mouse: NÃO OBTIDO")
    barra_progresso(9)

    printers = get_printers()
    if printers:
        for idx, (name, serial, fabricante, modelo) in enumerate(printers, start=1):
            lines.append(f"Impressora {idx}: {padrao(name)} | Serial/ID: {padrao(serial)} | Fabricante: {padrao(fabricante)} | Modelo: {padrao(modelo)}")
    else:
        lines.append("Impressora: NÃO OBTIDO")
    barra_progresso(10)

    write_report(path, lines)
    print(f"\nArquivo gerado: {path}")
    resposta = input(f"Deseja abrir o arquivo gerado ({filename})? [s/N]: ").strip().lower()
    if resposta == 's':
        try:
            os.startfile(path)
        except Exception as e:
            print(f"Não foi possível abrir o arquivo: {e}")

if __name__ == "__main__":
    main()
