# tmp_run_csinfo_monitor.ps1
# Inicia dist\CSInfo.exe, monitora processos powershell/pwsh por 12 segundos (1s amostras) e registra novos PIDs que surgirem
$baseline = Get-Process -Name powershell,pwsh -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id -ErrorAction SilentlyContinue
if (-not $baseline) { $baseline = @() }
Write-Output "Baseline PowerShell PIDs: $($baseline -join ',')"
$p = Start-Process -FilePath .\dist\CSInfo.exe -PassThru
Write-Output "Started CSInfo PID:$($p.Id)"
$events = @()
$duration = 12
for ($i=0; $i -lt $duration; $i++) {
    Start-Sleep -Seconds 1
    $now = Get-Date
    $current = Get-Process -Name powershell,pwsh -ErrorAction SilentlyContinue | Select-Object Id, ProcessName, StartTime
    if ($current) {
        foreach ($proc in $current) {
            if ($baseline -notcontains $proc.Id) {
                # novo PID detectado
                $events += [PSCustomObject]@{ Time = $now; Id = $proc.Id; Name = $proc.ProcessName; StartTime = $proc.StartTime }
            }
        }
    }
}
if ($events.Count -eq 0) {
    Write-Output "No new powershell/pwsh processes detected during monitoring."
} else {
    Write-Output "New powershell/pwsh processes detected:"
    $events | Format-Table -AutoSize
}
Write-Output 'Recent files modified in the repo (last 10 minutes):'
Get-ChildItem -Path . -Recurse -File | Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-10) } | Select-Object FullName, LastWriteTime | Format-Table -AutoSize
Write-Output 'Stopping GUI process'
try {
    Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    Write-Output "Stopped PID:$($p.Id)"
} catch {
    Write-Output "Failed to stop PID:$($p.Id): $_"
}
Write-Output 'Done.'
