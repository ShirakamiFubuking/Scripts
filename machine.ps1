
# $OutputEncoding = [System.Text.Encoding]::UTF8
# [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# [Console]::InputEncoding = [System.Text.Encoding]::UTF8
# 每小時執行
# 基本檢查
# TODO

# 設定目標與路徑
$targetVersion = "1.3.4.103349"
$logFile = Join-Path $env:TEMP "pwb_update_machine.log"

# 定義日誌函式 (方便統一格式並加入時間)
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$timestamp] $Message" | Out-File -FilePath $logFile -Append -Encoding utf8
}

# 1. 請求本地服務
try {
    $response = Invoke-WebRequest -Uri "http://127.0.0.1:61161" -ErrorAction Stop
    
    # 2. 使用正規表示法過濾出版本號
    if ($response.Content -match 'version:(?<version>[\d\.]+)') {
        $currentVersion = $Matches['version']
        Write-Log "Current version detected: $currentVersion"
        
        if ($currentVersion -ne $targetVersion) {
            Write-Log "Not target version ($targetVersion). Starting update..."
            
            # --- 執行更新程式 ---
            $tempDir = $env:TEMP
            $zipPath = Join-Path $tempDir "hicos.zip"
            
            Write-Log "Downloading HiCOS update..."
            Invoke-WebRequest -Uri "https://api-hisecurecdn.cdn.hinet.net/HiCOS_Client.zip" -OutFile $zipPath
            
            Write-Log "Extracting files..."
            Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force
            
            Write-Log "Installing HiCOS (Silent Mode)..."
            $exePath = Join-Path $tempDir "HiCOS_Client.exe"
            if (Test-Path $exePath) {
                Start-Process -FilePath $exePath -ArgumentList "/install", "/quiet", "/norestart" -Wait
                Write-Log "Installation process finished."
            } else {
                Write-Log "Error: HiCOS_Client.exe not found after extraction."
            }
            
            # 清理暫存檔
            Remove-Item -Path $zipPath -ErrorAction SilentlyContinue
            Remove-Item -Path $exePath -ErrorAction SilentlyContinue
            Write-Log "Update sequence completed."
        } else {
            Write-Log "Already at target version: $targetVersion. No action needed."
        }
    }
} catch {
    Write-Log "Warning: Unable to connect to HiCOS service at http://127.0.0.1:61161. Ensure the service is running."
}