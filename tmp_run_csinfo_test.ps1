# tmp_run_csinfo_test.ps1 - inicia dist\CSInfo.exe, aguarda, verifica processos powershell/pwsh, lista arquivos recentes e encerra o EXE
$p = Start-Process -FilePath .\dist\CSInfo.exe -PassThru
Write-Output "Started PID:$($p.Id)"
Start-Sleep -Seconds 4
$ps = Get-Process -Name powershell,pwsh -ErrorAction SilentlyContinue
if ($ps) {
    $ps | Select-Object Id, ProcessName, StartTime | Format-Table -AutoSize
} else {
    Write-Output 'No powershell/pwsh processes found.'
}
Write-Output 'Recent files modified in the repo (last 5 minutes):'
Get-ChildItem -Path . -Recurse -File | Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) } | Select-Object FullName, LastWriteTime | Format-Table -AutoSize
Write-Output 'Stopping GUI process'
try {
    Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    Write-Output "Stopped PID:$($p.Id)"
} catch {
    Write-Output "Failed to stop PID:$($p.Id): $_"
}
Write-Output 'Done.'
