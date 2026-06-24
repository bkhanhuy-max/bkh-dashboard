@echo off
echo Khoi dong cac dich vu BKH-AI (API Server & SSH Tunnel)...
powershell -ExecutionPolicy Bypass -File "%~dp0start_services.ps1"
pause
