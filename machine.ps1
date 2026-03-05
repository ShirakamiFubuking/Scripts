$TaskName = "pwb_update_machine"
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Set-ScheduledTask -TaskName $TaskName -Principal $Principal
# -------------------- Configuration --------------------
$Config = @{
    LogFile         = Join-Path $env:TEMP "pwb_update_machine.log"
    Hicos = @{
        TargetVersion = "1.3.4.103349"
        DownloadUrl   = "https://api-hisecurecdn.cdn.hinet.net/MOICA/HiCOS_Client.zip"
        ServiceUrl    = "http://127.0.0.1:61161"
    }
    SevenZip = @{
        TargetVersion = "25.01"
        Architecture  = "x64"  # "x64" or "x86"
        # 7-Zip 官網下載網址會隨版本變動，腳本會動態產生
    }
    Hosts = @{
        Path          = "$env:SystemRoot\System32\drivers\etc\hosts"
        SourceUrl     = "https://raw.githubusercontent.com/ShirakamiFubuking/Scripts/refs/heads/main/hosts"
    }
    BrowserPolicies = @{
        Chrome = "HKLM:\SOFTWARE\Policies\Google\Chrome\PopupsAllowedForUrls"
        Edge   = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\PopupsAllowedForUrls"
        Urls   = @{
            "1" = "https://ecpa.dgpa.gov.tw"
            "2" = "https://dm.kcg.gov.tw:443"
            "3" = "http://localhost:61161"
        }
    }
}
# -------------------- Helper Functions --------------------
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Write-Host $logEntry
    $logEntry | Out-File -FilePath $Config.LogFile -Append -Encoding utf8
}

# -------------------- Pre-Flight Checks --------------------
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log "Access Denied: Please run this script as Administrator."
    exit
}

# ==================== Hosts Update ====================
Write-Log "[Hosts] Updating hosts file..."
try {
    Write-Log "[Hosts] Fetching latest content from source..."
    $newHostsContent = Invoke-RestMethod -Uri $Config.Hosts.SourceUrl -UseBasicParsing

    if ($null -ne $newHostsContent -and $newHostsContent.Length -ge 10) {
        $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllLines($Config.Hosts.Path, $newHostsContent, $Utf8NoBomEncoding)
        Write-Log "[Hosts] Update successful."

        ipconfig /flushdns | Out-Null
        Write-Log "[Hosts] DNS cache flushed."
    } else {
        Write-Log "[Hosts] Error: Downloaded content is invalid or empty."
    }
} catch {
    Write-Log "[Hosts] Failed to update: $($_.Exception.Message)"
}

# ==================== Registry Update ====================
Write-Log "[Registry] Updating system registry..."
try {
    $PolicyPaths = @($Config.BrowserPolicies.Chrome, $Config.BrowserPolicies.Edge)

    foreach ($keyPath in $PolicyPaths) {
        if (-not (Test-Path $keyPath)) {
            New-Item -Path $keyPath -Force | Out-Null
            Write-Log "[Registry] Created key: $keyPath"
        }
        $Config.BrowserPolicies.Urls.GetEnumerator() | ForEach-Object {
            Set-ItemProperty -Path $keyPath -Name $_.Key -Value $_.Value -Type String
        }
    }
    Write-Log "[Registry] Browser policies applied successfully."
} catch {
    Write-Log "[Registry] Failed to apply settings: $($_.Exception.Message)"
}

# ==================== HiCOS Update ====================
Write-Log "[HiCOS] Checking version..."
try {
    $response = Invoke-WebRequest -Uri $Config.Hicos.ServiceUrl -TimeoutSec 5 -ErrorAction Stop
    
    if ($response.Content -match 'version:(?<version>[\d\.]+)') {
        $currentVersion = [version]$Matches['version']
        $targetVersionObj = [version]$Config.Hicos.TargetVersion
        
        Write-Log "[HiCOS] Current: $currentVersion | Target: $targetVersionObj"

        if ($currentVersion -ne $targetVersionObj) {
            Write-Log "[HiCOS] Newer version available. Starting update sequence..."
            
            $zipPath = Join-Path $env:TEMP "hicos.zip"
            $extractDir = Join-Path $env:TEMP "hicos_extracted"

            Write-Log "[HiCOS] Downloading package..."
            Invoke-WebRequest -Uri $Config.Hicos.DownloadUrl -OutFile $zipPath -ErrorAction Stop
            
            Write-Log "[HiCOS] Extracting files..."
            Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
            
            $exePath = Join-Path $extractDir "HiCOS_Client.exe"
            if (Test-Path $exePath) {
                Write-Log "[HiCOS] Executing silent installation..."
                $proc = Start-Process -FilePath $exePath -ArgumentList "/install", "/quiet", "/norestart" -Wait -PassThru
                Write-Log "[HiCOS] Installation finished (ExitCode: $($proc.ExitCode))."
            } else {
                Write-Log "[HiCOS] Error: HiCOS_Client.exe not found in package."
            }

            # Cleanup
            Remove-Item $zipPath -ErrorAction SilentlyContinue
            Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue
        } else {
            Write-Log "[HiCOS] System is up to date or version is newer. Skipping."
        }
    }
} catch {
    Write-Log "[HiCOS] Warning: Service unreachable at $($Config.Hicos.ServiceUrl). Check if HiCOS is installed."
}

# ==================== 7-Zip Update ====================
Write-Log "[7-Zip] Checking installation..."
try {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    # find installed 7-Zip infomation
    $installedApp = Get-ItemProperty $regPaths -ErrorAction SilentlyContinue | 
                    Where-Object { $_.DisplayName -like "*7-Zip*" } | Select-Object -First 1

    $current7zVer = if ($installedApp) { [version]($installedApp.DisplayVersion -replace '[^0-9.]', '') } else { [version]"0.0.0.0" }
    $target7zVerObj = [version]$Config.SevenZip.TargetVersion

    Write-Log "[7-Zip] Current: $current7zVer | Target: $target7zVerObj"

    if ($current7zVer -lt $target7zVerObj) {
        Write-Log "[7-Zip] Newer version available. Preparing download..."
        
        # 24.01 -> 7z2401-x64.exe
        $verClean = $Config.SevenZip.TargetVersion.Replace(".", "")
        $fileName = "7z$verClean-$($Config.SevenZip.Architecture).exe"
        $downloadUrl = "https://www.7-zip.org/a/$fileName"
        $exePath = Join-Path $env:TEMP $fileName

        Write-Log "[7-Zip] Downloading from $downloadUrl"
        Invoke-WebRequest -Uri $downloadUrl -OutFile $exePath -ErrorAction Stop

        Write-Log "[7-Zip] Executing silent installation..."
        $proc = Start-Process -FilePath $exePath -ArgumentList "/S" -Wait -PassThru
        
        if ($proc.ExitCode -eq 0) {
            Write-Log "[7-Zip] Installation successful."
        } else {
            Write-Log "[7-Zip] Installation failed with ExitCode: $($proc.ExitCode)"
        }

        # Cleanup
        Remove-Item $exePath -ErrorAction SilentlyContinue
    } else {
        Write-Log "[7-Zip] System is up to date."
    }
} catch {
    Write-Log "[7-Zip] Error during update: $($_.Exception.Message)"
}

# SIG # Begin signature block
# MIIFRgYJKoZIhvcNAQcCoIIFNzCCBTMCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUPB0TRkhBORNb45wmUbh17hVF
# W+egggLuMIIC6jCCAdKgAwIBAgIQf/nbIZcJG6BDLhvkFK6gNzANBgkqhkiG9w0B
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
# FgQUjbx6FY+mmKOC2Q318gdnuVNrGoowDQYJKoZIhvcNAQEBBQAEggEAhSZspkYU
# 0+nUVpk3m23uD8FWCfozzZdTxK6N4pZF/fCRtPts6IeRyW36Hnph5tQoVaCrzmFy
# 9VtPBqjQlNByvRrAWyz5SCRWwJ7wemFqWJhgfLhjglr9scHEbemGMGLQFY0Xs02F
# oSqJ9E2OMkvxvKex3xTDqVQHr3EXrYxml8OG9hKEE0EnLfbeEIDbQIWq8YghFpGE
# sfC9uE8Yv1b38LUMAO6EtoRq3mFR1hYQhrD7Mc5c5fvTd7JQzmPoaXRFjU/QFzIh
# zW+BwEznTIYuPlE2rEWGpHZtpdzd9a7rb89gzTy+67vliWo+dsA5nf+s18cbdDE6
# 1QzEh5Kz4koRAw==
# SIG # End signature block
