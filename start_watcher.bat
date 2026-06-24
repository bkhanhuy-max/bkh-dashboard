@echo off
title BKH Excel Auto-Deploy Watcher
echo Dang khoi dong cong cu tu dong cap nhat khi file Excel thay doi...
powershell -WindowStyle Minimized -ExecutionPolicy Bypass -File "%~dp0watch_excel_changes.ps1"
