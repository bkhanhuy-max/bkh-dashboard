@echo off
echo Dang kiem tra API Server...
powershell -Command "try { Invoke-WebRequest -Uri 'http://localhost:8001/api/status' -UseBasicParsing -TimeoutSec 2 -ErrorAction Stop | Out-Null; Write-Host 'API Server da chay roi!' } catch { Write-Host 'Dang khoi dong...'; wscript //nologo D:\BKH-AI\start_api_server.vbs; Start-Sleep 3; Write-Host 'API Server san sang!' }"

