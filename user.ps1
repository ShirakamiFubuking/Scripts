
# $OutputEncoding = [System.Text.Encoding]::UTF8
# [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# [Console]::InputEncoding = [System.Text.Encoding]::UTF8
# 每小時執行
# 基本檢查
# TODO

# 設定目標與路徑
$logFile = Join-Path $env:TEMP "pwb_update_user.log"

# 定義日誌函式 (方便統一格式並加入時間)
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "[$timestamp] $Message" | Out-File -FilePath $logFile -Append -Encoding utf8
}
$TaskName = "pwb_update_machine"
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Set-ScheduledTask -TaskName $TaskName -Principal $Principal

$os = Get-CimInstance Win32_OperatingSystem
$cpu = Get-CimInstance Win32_Processor
$computer_info = @{
    # 唯一識別 ID (UUID)
    uuid = (Get-CimInstance Win32_ComputerSystemProduct).UUID
    
    # 電腦名稱
    ComputerName = $env:COMPUTERNAME
    
    # 登入使用者
    LoginUser = whoami
    
    # 系統版本 (例如: Microsoft Windows 11 Pro)
    OS_Name = $os.Caption
    
    # Build 版本 (例如: 22631)
    OS_Version = $os.Version
    
    # 網卡資訊：獲取特定網段 (128.5.47.*) 的 IP 與 MAC
    # 如果要獲取「所有」非虛擬網卡，可將 Where 條件改為 { $_.InterfaceAlias -notlike "*Loopback*" }
    Network_Interfaces = @(
        Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.IPAddress -notlike "127.0.0.1" } | ForEach-Object {
            $adapter = Get-NetAdapter -InterfaceIndex $_.InterfaceIndex
            [PSCustomObject]@{
                IPAddress  = $_.IPAddress
                MACAddress = $adapter.MacAddress
                Interface  = $adapter.InterfaceAlias
            }
        }
    )
    
    # CPU 型號
    CPU_Model = $cpu.Name
}
$computer_info | Format-Table -AutoSize

# 強制使用 UTF8 編碼轉換 JSON
$jsonBody = $computer_info | ConvertTo-Json -Compress
# 在 Content-Type 中明確指定 charset=utf-8
Write-Log $jsonBody
try {
    Invoke-RestMethod -Uri "http://128.5.47.252:5000/report" `
                      -Method Post `
                      -Body ([System.Text.Encoding]::UTF8.GetBytes($jsonBody)) `
                      -ContentType "application/json; charset=utf-8"
    Write-Log "資料已成功以 UTF-8 編碼回傳！" -ForegroundColor Cyan
} catch {
    Write-Log "回傳失敗"
}
# SIG # Begin signature block
# MIIFRgYJKoZIhvcNAQcCoIIFNzCCBTMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUeu4BxtvCfbnQB6k3SBNhZ398
# bsKgggLuMIIC6jCCAdKgAwIBAgIQf/nbIZcJG6BDLhvkFK6gNzANBgkqhkiG9w0B
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
# FgQUEql8JTHWJrbOgrMFLpEKwpDnhE8wDQYJKoZIhvcNAQEBBQAEggEAdUwVaQVE
# NPnFYEDlJu6tBEXx0CcH/Dg3Dpghttu3A96Q1/BlC7MXsjMcXsBU+2ZkNpK69tBA
# Z/psKYzlx54n42VsNPdxEMK/69G0f+/9LJeSgOqVdNeQOkjH8+zh6T9hbcO4RPnm
# J2z6vM1WZnEjUbo6UiGvD+m3+pyk3EMthH+jiikM7pG3Leo6GbXhfiWWNcyo7315
# ueVpJDNDWfM/fF62wedky8KnHXYrAkxpzfFzP0C8D/Go5zC36XFkhF4fCZqm5MA+
# TDlDDwiwMoGv5tNWTpV2KZ07hDBpqtpMi5gHC9uf7+VqN+j0ICPdcmTvXtHjiW+C
# MwELx8VswmKcWA==
# SIG # End signature block
