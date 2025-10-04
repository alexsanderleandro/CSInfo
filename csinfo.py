# coletar_info_maquina.py
import subprocess
import platform
import os
import re
import sys
import winreg
from datetime import datetime
import json

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

def get_motherboard_serial():
    cmd = "(Get-CimInstance Win32_BaseBoard | Select-Object -ExpandProperty SerialNumber) -join ''"
    out = run_powershell(cmd)
    return out or "Não encontrado"

def get_monitor_serials():
    ps = r"""
    $s = @()
    try {
      $mons = Get-WmiObject -Namespace root\wmi -Class WmiMonitorID -ErrorAction SilentlyContinue
      foreach ($m in $mons) {
        $arr = $m.SerialNumberID
        if ($arr) {
          $serial = ($arr | ForEach-Object {[char]$_}) -join ''
          if ($serial) { $s += $serial }
        } else {
          if ($m.InstanceName) { $s += $m.InstanceName }
        }
      }
    } catch {}
    $s | ConvertTo-Json -Compress
    """
    out = run_powershell(ps)
    try:
        parsed = json.loads(out) if out else []
        if isinstance(parsed, str):
            parsed = [parsed]
        return parsed if parsed else ["Nenhum serial encontrado"]
    except Exception:
        return [l for l in out.splitlines() if l] or ["Nenhum serial encontrado"]

def get_devices_by_class(devclass):
    ps = r"""
    $out = @()
    try {
      $devs = Get-PnpDevice -Class {cls} -ErrorAction SilentlyContinue
      foreach ($d in $devs) {
        $id = $d.InstanceId
        $name = $d.FriendlyName
        if (-not $name) { $name = $d.Name }
        $serial = ""
        try {
          $prop = Get-PnpDeviceProperty -InstanceId $id -KeyName 'DEVPKEY_Device_SerialNumber' -ErrorAction SilentlyContinue
          if ($prop) { $serial = $prop.Data }
        } catch {}
        if (-not $serial) { $serial = $id }
        $out += [PSCustomObject]@{ Name = $name; Serial = $serial }
      }
    } catch {}
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
            items.append((name, serial))
    except Exception:
        for l in out.splitlines():
            if l.strip():
                items.append((l, ""))
    return items or [("Nenhum dispositivo encontrado", "")]

def get_printers():
    ps = r"""
    $o = @()
    try {
      $pr = Get-CimInstance Win32_Printer -ErrorAction SilentlyContinue
      foreach ($p in $pr) {
        $name = $p.Name
        $pnp = $p.PNPDeviceID
        $serial = ""
        if ($pnp) {
          try {
            $prop = Get-PnpDeviceProperty -InstanceId $pnp -KeyName 'DEVPKEY_Device_SerialNumber' -ErrorAction SilentlyContinue
            if ($prop) { $serial = $prop.Data }
          } catch {}
        }
        if (-not $serial) { $serial = $pnp }
        $o += [PSCustomObject]@{Name=$name; Serial=$serial}
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
            items.append((p.get('Name'), p.get('Serial')))
    except Exception:
        for l in out.splitlines():
            if l.strip():
                items.append((l, ""))
    return items or [("Nenhuma impressora encontrada", "")]

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
    with open(path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))

def main():
    machine = get_machine_name()
    safe_name = safe_filename(machine)
    filename = f"info_maquina_{safe_name}.txt"
    path = os.path.join(os.getcwd(), filename)

    lines = []
    lines.append(f"Relatório gerado: {datetime.now().isoformat()}")
    lines.append(f"Nome: {machine}")
    lines.append(f"Tipo: {'Notebook' if is_laptop() else 'Desktop'}")
    lines.append(f"Versão do sistema operacional: {get_os_version()}")
    lines.append(f"Versão do Office (se tiver): {get_office_version()}")
    lines.append(f"Serial da placa mãe: {get_motherboard_serial()}")

    # Monitores
    monitors = get_monitor_serials()
    if monitors:
        for idx, m in enumerate(monitors, start=1):
            lines.append(f"Monitor {idx} Serial: {m}")
    else:
        lines.append("Monitor 1 Serial: Nenhum serial encontrado")

    # Teclado
    keyboards = get_devices_by_class("Keyboard")
    if keyboards:
        for idx, (name, serial) in enumerate(keyboards, start=1):
            lines.append(f"Teclado {idx}: {name} | Serial/ID: {serial}")
    else:
        lines.append("Teclado: Nenhum encontrado")

    # Mouse
    mice = get_devices_by_class("Mouse")
    if mice:
        for idx, (name, serial) in enumerate(mice, start=1):
            lines.append(f"Mouse {idx}: {name} | Serial/ID: {serial}")
    else:
        lines.append("Mouse: Nenhum encontrado")

    # Impressoras
    printers = get_printers()
    if printers:
        for idx, (name, serial) in enumerate(printers, start=1):
            lines.append(f"Impressora {idx}: {name} | Serial/ID: {serial}")
    else:
        lines.append("Impressora: Nenhuma encontrada")

    write_report(path, lines)
    print(f"Arquivo gerado: {path}")

if __name__ == "__main__":
    main()
