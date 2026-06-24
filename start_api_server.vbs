Set WShell = CreateObject("WScript.Shell")
WShell.Run "powershell -ExecutionPolicy Bypass -NonInteractive -File D:\BKH-AI\api_server.ps1", 0, False

