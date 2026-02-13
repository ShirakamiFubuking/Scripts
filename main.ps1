
# $OutputEncoding = [System.Text.Encoding]::UTF8
# [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# [Console]::InputEncoding = [System.Text.Encoding]::UTF8
# 每小時執行
# 基本檢查
# TODO

# Update Hicos
# 目標版本號
$targetVersion = "1.3.4.103349"
# 1. 請求本地服務
$response = Invoke-WebRequest -Uri "http://127.0.0.1:61161" -ErrorAction SilentlyContinue
if ($response) {
    # 2. 使用正規表示法過濾出版本號 (尋找 version: 後面的數字與點)
    if ($response.Content -match 'version:(?<version>[\d\.]+)') {
        $currentVersion = $Matches['version']
        Write-Host "now version: $currentVersion" -ForegroundColor Cyan
        
        if ($currentVersion -ne $targetVersion) {
            Write-Host "not target version: $targetVersion, updating..." -ForegroundColor Yellow
            
            # --- 執行更新程式 ---
            Set-Location -Path $env:TEMP
            Invoke-WebRequest -Uri "https://api-hisecurecdn.cdn.hinet.net/HiCOS_Client.zip" -OutFile "hicos.zip"
            Expand-Archive -Path "hicos.zip" -DestinationPath "." -Force
            
            Write-Host "installing HiCOS..." -ForegroundColor Green
            Start-Process -FilePath "HiCOS_Client.exe" -ArgumentList "/install", "/quiet", "/norestart" -Wait
            
            # 清理暫存檔
            Remove-Item -Path "hicos.zip", "HiCOS_Client.exe" -ErrorAction SilentlyContinue
            Write-Host "update completed." -ForegroundColor Green
        } else {
            Write-Host "is target version: $targetVersion" -ForegroundColor Green
        }
    }
} else {
    Write-Warning "check Hicos service at http://127.0.0.1:61161"
}
# Set-Location -Path $env:TEMP; Invoke-WebRequest -Uri "https://api-hisecurecdn.cdn.hinet.net/HiCOS_Client.zip" -OutFile "hicos.zip"; Expand-Archive -Path "hicos.zip" -DestinationPath "." -Force; Start-Process -FilePath "HiCOS_Client.exe" -ArgumentList "/install", "/quiet", "/norestart" -Wait; Remove-Item -Path "hicos.zip", "HiCOS_Client.exe" -ErrorAction SilentlyContinue
# SIG # Begin signature block
# MIIFRgYJKoZIhvcNAQcCoIIFNzCCBTMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUt7RiXFYt7ZTq/g2GGCzsQg3C
# +eigggLuMIIC6jCCAdKgAwIBAgIQf/nbIZcJG6BDLhvkFK6gNzANBgkqhkiG9w0B
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
# FgQUCfMOZpqUdGIs0p24zqRNFlSuoBQwDQYJKoZIhvcNAQEBBQAEggEAPQqPVBRp
# y1DZi0S7MAzvQbxJstsGF/PcHPIPyWRLEYJeQ3XVZbadPO809jnUeOtvXLMrkr9f
# NqOf2sMDIS1IZvlqRadOPtjk1OtlaYbCS1nEEdkwCOKW15JQhx+BPdAF7DyQewMX
# lAkmnHoavwL83VJEiTS99mB45Aio0s7wBK9TP6TYWmmlnmd3lWA9PhOGqVXzc+fs
# Ka8+ithBs2yu765SewfVXZuz4TpzzC8QOvqVwyOWtdrC8qt0mp83j8ILO20YKG1M
# AKDPyXe9CLbkWGtpdb8eKiiwxfnMnDULHtQl0YTUz3EUQ0+3zQ90A1hNPt7ISBhh
# ek/YNJk9P/PFRQ==
# SIG # End signature block
