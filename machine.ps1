
# $OutputEncoding = [System.Text.Encoding]::UTF8
# [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# [Console]::InputEncoding = [System.Text.Encoding]::UTF8
# 每小時執行
# 基本檢查
# TODO

# ==================== 參數:log檔案 ====================

$logFile = Join-Path $env:TEMP "pwb_update_machine.log"

# ==================== 參數:Hicos更新 ====================

$targetVersion = "1.3.4.103349"
$setupFileUrl = "https://api-hisecurecdn.cdn.hinet.net/MOICA/HiCOS_Client.zip"

# ==================== 參數:hosts檔案 ====================

$hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
$sourceUrl = "https://raw.githubusercontent.com/ShirakamiFubuking/Scripts/refs/heads/main/hosts"

# ==================== 參數:登錄檔更新 ====================

$regUrl = "https://raw.githubusercontent.com/ShirakamiFubuking/Scripts/refs/heads/main/reg/local_machine.reg"

# ==================== 程式碼 ====================
# 定義日誌函式 (方便統一格式並加入時間)
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$timestamp] $Message" | Out-File -FilePath $logFile -Append -Encoding utf8
}

############# Hicos更新
# 1. 請求本地服務
try {
    $response = Invoke-WebRequest -Uri "http://127.0.0.1:61161" -ErrorAction Stop
    
    # 2. 使用正規表示法過濾出版本號
    if ($response.Content -match 'version:(?<version>[\d\.]+)') {
        # 將擷取的字串轉換為 [version] 物件
        $currentVersion = [version]$Matches['version']
        $targetVersionObj = [version]$targetVersion
        
        Write-Log "Current version detected: $currentVersion"

        # 使用 [version] 物件進行「小於 (-lt)」判斷
        if ($currentVersion -lt $targetVersionObj) {
            Write-Log "Current version ($currentVersion) is older than target ($targetVersionObj). Starting update..."
            
            # --- 執行更新程式 ---
            $tempDir = $env:TEMP
            $zipPath = Join-Path $tempDir "hicos.zip"
            
            Write-Log "Downloading HiCOS update..."
            Invoke-WebRequest -Uri $setupFileUrl -OutFile $zipPath
            
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
        }elseif ($currentVersion -eq $targetVersionObj) {
            Write-Log "Already at target version: $targetVersion. No action needed."
        }else {
            Write-Log "Current version ($currentVersion) is newer than target ($targetVersionObj). Skipping."
        }
    }
} catch {
    Write-Log "Warning: Unable to connect to HiCOS service at http://127.0.0.1:61161. Ensure the service is running."
}

############# hosts更新
# 3. 檢查權限
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $errorMsg = "權限不足：請以系統管理員身分執行此腳本。"
    Write-Error $errorMsg
    # 如果日誌路徑可寫，記錄錯誤
    if (Test-Path (Split-Path $logFile)) { "[(Get-Date)] $errorMsg" | Out-File $logFile -Append }
    break
}

Write-Log "開始更新流程..."

try {
    # 4. 下載最新 hosts
    Write-Log "正在從 GitHub 獲取內容: $sourceUrl"
    $newHostsContent = Invoke-RestMethod -Uri $sourceUrl -UseBasicParsing

    if ($null -eq $newHostsContent -or $newHostsContent.Length -lt 10) {
        throw "下載內容為空或長度異常，取消寫入。"
    }

    # 5. 直接覆寫 hosts 檔案 (不進行備份)
    # 使用 .NET 方法確保 UTF-8 無 BOM 格式，避免系統解析問題
    $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllLines($hostsPath, $newHostsContent, $Utf8NoBomEncoding)
    Write-Log "hosts 檔案已成功更新。"

    # 6. 重新整理 DNS 快取
    ipconfig /flushdns | Out-Null
    Write-Log "DNS 快取已清理，設定立即生效。"

} catch {
    Write-Log "錯誤發生：$($_.Exception.Message)"
}

Write-Log "更新流程結束。"

############# 登錄檔更新
$tempRegPath = "$env:TEMP\setting.reg"
try {
    # 3. 下載 .reg 檔案到暫存資料夾
    Write-Log "正在從 GitHub 下載登錄檔..."
    Invoke-WebRequest -Uri $regUrl -OutFile $tempRegPath -UseBasicParsing
    Write-Log "檔案已暫存至: $tempRegPath"

    # 4. 使用 reg import 進行匯入
    # /s 參數代表「安靜模式」，不會跳出確認視窗
    Write-Log "執行匯入作業..."
    $process = Start-Process -FilePath "reg.exe" -ArgumentList "import `"$tempRegPath`"" -Wait -PassThru -WindowStyle Hidden

    if ($process.ExitCode -eq 0) {
        Write-Log "登錄表設定已成功匯入。"
    } else {
        throw "reg.exe 回傳錯誤代碼: $($process.ExitCode)"
    }

    # 5. 清理暫存檔
    if (Test-Path $tempRegPath) {
        Remove-Item $tempRegPath -Force
        Write-Log "暫存檔已刪除。"
    }

} catch {
    Write-Log "匯入失敗：$($_.Exception.Message)"
}

# SIG # Begin signature block
# MIIFRgYJKoZIhvcNAQcCoIIFNzCCBTMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU/6L9Ak/S+lX0fE4J6CMlBo+4
# tiygggLuMIIC6jCCAdKgAwIBAgIQf/nbIZcJG6BDLhvkFK6gNzANBgkqhkiG9w0B
# AQsFADANMQswCQYDVQQDDAJHZTAeFw0yNjAyMTAwMjA2NTRaFw0zMTAyMTAwMjE2
# NTNaMA0xCzAJBgNVBAMMAkdlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
# AQEAo4PtDhHC0bm/MGf+ud5v7E0gD80T4anDq36e98xeUv+TzZ7VUtP5uATp5APe
# pUGRPEw7P8nkNwUKhym/V6ZjcMwege6lUO2+FjgfmE31gj+1d27c3Ier+5w6P3p4
# hxZz8p8tH6y59MS5EKeIYFktK0qi0S3sNTpOoZeIUsoUgtizmS7Yx0L/LgSTBGxG
# 4R//lQ36v7xLPIZkB7SM55Y2jgQcRAr9ca662wGYudOx6xD645Y5Q+G+YqOT1joY
# QRXkA7GLNjIQuSR+aEn3iJH6BA0SR9IytsOMeMbKyoUmjFeH++1X0/eRJlhhaAbQ
# iBQewXF+Nu8WvYms5CgZOPgB0QIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFHP8T/a7U3TwIjlf6HN0bKtM+enW
# MA0GCSqGSIb3DQEBCwUAA4IBAQAwfcBw64YZuVf5Eaun/7x/UTbZNfbvD+kjaj4l
# /D5mBGUewRvgduYc52Kvxfvisbj+nH3wG2VBNVHDjssdfkRVRavQ63E0CtxxSHS3
# Ahe0vSl8ztEBbYdfXskmCx2u9dYaiX0TeHBCtngDKsjaCbT2TAn51sQnWCnw2cpk
# kZcKLwfNj6V60U0yMgp/lc0DgSFinEGZE2gLbqtFEGhGTYpQncgMSe94Of8NaGfP
# YskaNSTqwG3WZ4zp5mgrXbRXrC4c57rZ7HIH3aWSTCvVeswfmDJz9pLSrCAP8L7r
# mQHC8v+eziWG9v6fesTjJaTyqkWM4zZaDCiZncR8PGkZi0jUMYIBwjCCAb4CAQEw
# ITANMQswCQYDVQQDDAJHZQIQf/nbIZcJG6BDLhvkFK6gNzAJBgUrDgMCGgUAoHgw
# GAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQx
# FgQUGEKo5jAKkgJ3SB8RecBPpMlYOjswDQYJKoZIhvcNAQEBBQAEggEAP9K7Fo1d
# b4/S2lnUFkxaHifV7KBFlsdzAaHaL7nAj/vedERHDZ4vQmioJcddlVKOpJi3iwnP
# LpVGLWo+qcg6+hnQu8Hx6Xf8spJvrup297dDRqTwx0tcNXMrGdmtGhOTBuW711gW
# TpwtG0Ge/kNgawAuI0aXYa+HGbGzFTtResiBRP1xkjVcECD46e0MbgS4JgIN2bcv
# gaL7c+Ex31HEfH02Kbbr9u8flsBWzh8Sx6UKYXpo58AEKpaP9Diqbd1x84lGZrhM
# EjeHJY06cN6JLVu+2EWy6ob9YxztBcRU4g0A4TCCWzfJiY31KIKqJXRdyztRiIJ2
# NXwCyO6+aDCFQQ==
# SIG # End signature block
