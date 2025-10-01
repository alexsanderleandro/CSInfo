$exe = Join-Path -Path $PWD -ChildPath 'dist\CSInfo.exe'
if (-Not (Test-Path $exe)) { Write-Error "Executable not found: $exe"; exit 1 }
$p = Start-Process -FilePath $exe -PassThru
Write-Output "Started PID:$($p.Id)" 
Start-Sleep -Seconds 4
$py = Join-Path -Path $PWD -ChildPath 'tmp_auto_export.py'
if (-Not (Test-Path $py)) { Write-Error "Script not found: $py"; Stop-Process -Id $p.Id -Force; exit 1 }
# Run the python script
& python $py
# List generated files
Get-ChildItem -Path . -Filter 'auto_test_report.*' -File | Select-Object FullName, LastWriteTime | Format-Table -AutoSize
# Stop the EXE
try { Stop-Process -Id $p.Id -Force; Write-Output "Stopped PID:$($p.Id)" } catch { Write-Warning "Failed to stop PID:$($p.Id): $_" }
