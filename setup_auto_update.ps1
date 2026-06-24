# setup_auto_update.ps1 — Đăng ký watcher chạy ngầm trên Windows Task Scheduler
# Script này giúp đăng ký tệp watch_excel_changes.ps1 tự động chạy ẩn mỗi khi bạn đăng nhập Windows.

$scriptPath = "d:\BKH-AI\watch_excel_changes.ps1"

Write-Output "--- ĐĂNG KÝ TỰ ĐỘNG CẬP NHẬT TRÊN WINDOWS ---"

# Kiểm tra quyền Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Warning "Bạn cần chạy PowerShell bằng quyền Administrator để đăng ký tác vụ chạy ngầm Windows."
    Write-Output "Cách chạy:"
    Write-Output "1. Nhấp chuột phải vào biểu tượng PowerShell trên Windows và chọn 'Run as Administrator'."
    Write-Output "2. Chạy lệnh sau để đăng ký:"
    Write-Output "   powershell -ExecutionPolicy Bypass -File d:\BKH-AI\setup_auto_update.ps1"
    exit 1
}

# Cấu hình Scheduled Task chạy ẩn
$taskName = "BKH_Excel_Auto_Deploy"
$trigger = New-ScheduledTaskTrigger -AtLogon
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

try {
    Register-ScheduledTask -TaskName $taskName -Trigger $trigger -Action $action -Settings $settings -Force | Out-Null
    Write-Output "✅ Đăng ký thành công tác vụ tự động chạy ngầm!"
    Write-Output "Tác vụ này sẽ chạy ẩn hoàn toàn (Hidden) ở chế độ nền khi khởi động máy."
    Write-Output "Để kích hoạt chạy thử ngay lập tức mà không cần khởi động lại máy, bạn chạy lệnh:"
    Write-Output "   Start-ScheduledTask -TaskName `"$taskName`""
} catch {
    Write-Error "Không thể đăng ký Scheduled Task: $_"
}
