# start_services.ps1 — Khởi động API Server và SSH Tunnel và cập nhật URL
$ProjectDir = "d:\BKH-AI"
$UrlFile = "$ProjectDir\pinggy_url.json"
$LogFile = "$ProjectDir\services_startup_log.txt"

function Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$ts] $msg" | Out-File -FilePath $LogFile -Append
    Write-Host "[$ts] $msg"
}

Log "=== BẮT ĐẦU KHỞI ĐỘNG DỊCH VỤ BKH-AI ==="

# Bước 1: Dọn dẹp tiến trình cũ để tránh xung đột cổng
Log "Dọn dẹp các tiến trình cũ..."
try {
    wmic process where "CommandLine like '%api_server.ps1%'" call terminate | Out-Null
    wmic process where "CommandLine like '%free.pinggy.io%'" call terminate | Out-Null
} catch {}

Start-Sleep -Seconds 2

# Bước 2: Khởi động API Server chạy ngầm
Log "Khởi động API Server..."
if (Test-Path "$ProjectDir\start_api_server.vbs") {
    wscript //nologo "$ProjectDir\start_api_server.vbs"
} else {
    Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -NonInteractive -File `"$ProjectDir\api_server.ps1`"" -WindowStyle Hidden
}

# Bước 3: Khởi động SSH Tunnel chạy ngầm
Log "Khởi động SSH Tunnel (Pinggy)..."
$sshCmd = "ssh -o StrictHostKeyChecking=no -o UserKnownHostsFile=NUL -p 443 -R 80:127.0.0.1:8001 -L 4300:127.0.0.1:4300 free@free.pinggy.io u:Host:localhost"
Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$sshCmd`"" -WindowStyle Hidden

# Bước 4: Chờ kết nối và lấy URL công khai
Log "Đang đợi lấy URL công khai..."
$attempts = 0
$url = $null
while ($attempts -lt 6 -and -not $url) {
    Start-Sleep -Seconds 3
    $attempts++
    try {
        $res = Invoke-RestMethod -Uri "http://127.0.0.1:4300/urls" -ErrorAction SilentlyContinue
        if ($res -and $res.urls) {
            # Chọn URL HTTPS
            foreach ($u in $res.urls) {
                if ($u -like "https://*") {
                    $url = $u
                    break
                }
            }
        }
    } catch {
        Log "Thử lần ${attempts}: Chưa kết nối được API Pinggy..."
    }
}

if ($url) {
    Log "✅ Lấy thành công URL: $url"
    
    # Ghi file pinggy_url.json
    $jsonObj = [Ordered]@{ url = $url }
    $jsonObj | ConvertTo-Json -Compress | Out-File -FilePath $UrlFile -Encoding utf8
    Log "✅ Đã ghi $UrlFile"

    # Push lên GitHub
    Log "Đang push URL lên GitHub..."
    Set-Location $ProjectDir
    git add pinggy_url.json
    git commit -m "data: update tunnel url [skip ci]" 2>&1
    git push origin master:main 2>&1
    Log "✅ Đã đồng bộ URL lên GitHub Pages."
} else {
    Log "❌ Lỗi: Không thể lấy URL công khai từ Pinggy sau 18 giây."
}

Log "=== KẾT THÚC KHỞI ĐỘNG DỊCH VỤ ===`n"
