# watch_excel_changes.ps1 — Tự động cập nhật số liệu khi phát hiện file Excel thay đổi
# Chạy ngầm bằng FileSystemWatcher để tự động chạy deploy.ps1 khi bạn lưu file Excel.

$watcherPath = "d:\BKH-AI\01. DU AN 2026"
$deployScript = "d:\BKH-AI\deploy.ps1"
$logPath = "d:\BKH-AI\watcher_log.txt"

function LogWatcher($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$ts] $msg" | Tee-Object -FilePath $logPath -Append
}

LogWatcher "===== KHỞI ĐỘNG WATCHER EXCEL ====="
LogWatcher "Đang theo dõi thư mục: $watcherPath"

# Khởi tạo FileSystemWatcher
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $watcherPath
$watcher.Filter = "*.xlsx"
$watcher.IncludeSubdirectories = $true
$watcher.EnableRaisingEvents = $true

# Hàm xử lý khi phát hiện thay đổi
$action = {
    $path = $Event.SourceEventArgs.FullPath
    $changeType = $Event.SourceEventArgs.ChangeType
    
    # Bỏ qua các file tạm thời của Excel bắt đầu bằng ~$
    if ($path -match "\\~\$") {
        return
    }
    
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Ghi log sự kiện
    "[$ts] Phát hiện thay đổi: $path ($changeType)" | Out-File -FilePath $logPath -Append
    
    # Chờ 5 giây để Excel lưu hoàn tất và giải phóng khóa file
    Start-Sleep -Seconds 5
    
    "[$ts] Bắt đầu chạy tự động deploy..." | Out-File -FilePath $logPath -Append
    try {
        # Thực thi deploy.ps1
        $output = & powershell -ExecutionPolicy Bypass -File $deployScript 2>&1
        $output | Out-File -FilePath $logPath -Append
        "[$ts] ✅ Tự động deploy thành công!" | Out-File -FilePath $logPath -Append
    } catch {
        "[$ts] ❌ Lỗi khi tự động deploy: $_" | Out-File -FilePath $logPath -Append
    }
}

# Đăng ký sự kiện thay đổi ghi (Changed) và tạo mới (Created)
Register-ObjectEvent $watcher "Changed" -Action $action | Out-Null
Register-ObjectEvent $watcher "Created" -Action $action | Out-Null

LogWatcher "Watcher đang chạy ngầm..."

# Giữ vòng lặp chạy liên tục để duy trì tiến trình
while ($true) {
    Start-Sleep -Seconds 1
}
