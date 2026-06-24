# api_server.ps1 — BKH-AI Local API Server (Robust version)
$Port = 8001
$ProjectDir = "D:\BKH-AI"
$LogFile = "$ProjectDir\api_server_log.txt"

function WriteLog($msg) {
    $ts = Get-Date -Format "HH:mm:ss"
    "[$ts] $msg" | Out-File -FilePath $LogFile -Append
    Write-Host "[$ts] $msg"
}

function SendJson($response, $json, $statusCode = 200) {
    $response.StatusCode = $statusCode
    $response.ContentType = "application/json; charset=utf-8"
    $response.Headers.Set("Access-Control-Allow-Origin", "*")
    $response.Headers.Set("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    $response.Headers.Set("Access-Control-Allow-Headers", "Content-Type")
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $response.ContentLength64 = $bytes.Length
    $response.OutputStream.Write($bytes, 0, $bytes.Length)
    $response.Close()
}

$IsUpdating = $false

WriteLog "=== BKH-AI API Server khoi dong port $Port ==="

# Vong lap tu khoi dong neu gap loi
while ($true) {
    $listener = $null
    try {
        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add("http://localhost:$Port/")
        $listener.Start()
        WriteLog "✅ Server san sang tai http://localhost:$Port"

        while ($listener.IsListening) {
            $context = $null
            try {
                $context = $listener.GetContext()
                $req = $context.Request
                $res = $context.Response
                $path = $req.Url.LocalPath
                WriteLog "<- $($req.HttpMethod) $path"

                # Xu ly CORS preflight
                if ($req.HttpMethod -eq "OPTIONS") {
                    SendJson $res '{"ok":true}' 200
                    continue
                }

                if ($path -eq "/api/status") {
                    $d = [math]::Round((Get-Item "$ProjectDir\dashboard_data.json" -ErrorAction SilentlyContinue).Length / 1KB, 1)
                    $a = [math]::Round((Get-Item "$ProjectDir\ai_data.json" -ErrorAction SilentlyContinue).Length / 1KB, 1)
                    $t = (Get-Item "$ProjectDir\ai_data.json" -ErrorAction SilentlyContinue).LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                    $json = "{`"server`":`"running`",`"dashboard_kb`":$d,`"ai_kb`":$a,`"updated`":`"$t`"}"
                    SendJson $res $json

                } elseif ($path -eq "/api/update") {
                    if ($IsUpdating) {
                        WriteLog "⚠️ Dang co tien trinh cap nhat dang chay. Tu choi request."
                        SendJson $res '{"success":false,"error":"Hệ thống đang tiến hành cập nhật, vui lòng đợi trong giây lát...","output":""}' 200
                        continue
                    }
                    $IsUpdating = $true
                    WriteLog "▶ Bat dau cap nhat..."
                    $start = Get-Date
                    $out = ""
                    $err = ""
                    try {
                        $out = (& powershell -ExecutionPolicy Bypass -NonInteractive -File "$ProjectDir\deploy.ps1" 2>&1) -join "`n"
                        WriteLog "✅ deploy.ps1 xong"
                    } catch {
                        $err = $_.ToString()
                    } finally {
                        $IsUpdating = $false
                    }

                    $elapsed = [math]::Round(((Get-Date) - $start).TotalSeconds, 1)
                    $success = ($err -eq "" -and $out -notmatch "❌")
                    $responseObj = [Ordered]@{
                        success = $success
                        elapsed = $elapsed
                        output  = $out.Trim()
                        error   = $err.Trim()
                    }
                    $json = ConvertTo-Json $responseObj -Compress
                    WriteLog "✅ Xong sau ${elapsed}s"
                    SendJson $res $json

                } else {
                    SendJson $res '{"error":"not found"}' 404
                }

            } catch {
                WriteLog "⚠️ Loi xu ly request: $_"
                try { $context.Response.Close() } catch {}
            }
        }
    } catch {
        WriteLog "❌ Server loi: $_ — Khoi dong lai sau 3s..."
        Start-Sleep -Seconds 3
    } finally {
        try { if ($listener) { $listener.Stop(); $listener.Close() } } catch {}
    }
}
