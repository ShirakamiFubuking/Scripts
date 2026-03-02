
# $OutputEncoding = [System.Text.Encoding]::UTF8
# [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# [Console]::InputEncoding = [System.Text.Encoding]::UTF8
# 每小時執行
# 基本檢查
# TODO

# 設定目標與路徑
$targetVersion = "1.3.4.103349"
$logFile = Join-Path $env:TEMP "pwb_update.log"

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
# Set-Location -Path $env:TEMP; Invoke-WebRequest -Uri "https://api-hisecurecdn.cdn.hinet.net/HiCOS_Client.zip" -OutFile "hicos.zip"; Expand-Archive -Path "hicos.zip" -DestinationPath "." -Force; Start-Process -FilePath "HiCOS_Client.exe" -ArgumentList "/install", "/quiet", "/norestart" -Wait; Remove-Item -Path "hicos.zip", "HiCOS_Client.exe" -ErrorAction SilentlyContinue
# SIG # Begin signature block
# MIIFRgYJKoZIhvcNAQcCoIIFNzCCBTMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUT30s00ErO0G2k5xuYOzJUNKH
# a/GgggLuMIIC6jCCAdKgAwIBAgIQf/nbIZcJG6BDLhvkFK6gNzANBgkqhkiG9w0B
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
# FgQUp0eqKEIQwAYs/8yKMoy94PeU09swDQYJKoZIhvcNAQEBBQAEggEAJIZ0ExaQ
# y2TEBWofL4tad8y1ZvDiwZ10Aucb3eSOAaOCp1+lHBOhF387iWRjNghELeWroTle
# gXkaJpc5mURKvCJPxrNl+vCSwu8wR4HTZj0aPnNFcYVEC8RiQPN4UR/QLGe4o8xB
# tVIlSolxlJ6WTVvg6i5e+woTxxLOrWRKGSQzyXgCLT+Gi1M47znChgdKLLWnYSrs
# PcJcwK3GV/0/O8qoYiNMZ+P6GNeh86vb/DflpmLUxZG450F6/7AoJJaB7hBqB940
# lgroo9LYXL5Sm48TamdLJkWTEidXCxUxf9cfSulM2ob/q0iUKYnYt84Wz5kHa6C0
# 31b/X7t+/SPVzQ==
# SIG # End signature block
