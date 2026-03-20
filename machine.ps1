# Write-Log -LogFile <filepath> -Message <log message> -Level <INFO(default), WARN, ERROR>
function Write-Log {
    param(
        [string]$LogFile = (Join-Path $env:TEMP "pwb_update.log"),

        [Parameter(Mandatory=$true)]
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")] [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    # write to console
    $color = switch($Level) { "ERROR" {"Red"} "WARN" {"Yellow"} Default {"Gray"} }
    Write-Host $logEntry -ForegroundColor $color
    # write to file
    $logEntry | Out-File -FilePath $LogFile -Append -Encoding utf8
}

function Get-Getool-Version {
    Write-Log -Message $ToolVersion
}

function Report {
    param(
        [string]$Uri = "http://128.5.47.252:5000/report"
    )
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
    Write-Log -Message $jsonBody
    try {
        Invoke-RestMethod -Uri $Uri `
                        -Method Post `
                        -Body ([System.Text.Encoding]::UTF8.GetBytes($jsonBody)) `
                        -ContentType "application/json; charset=utf-8"
        Write-Log -Message "report successed"
    } catch {
        Write-Log -Message "report failed"
    }
}
# 桌面顯示控制台等等
function Show-Control {
    # 定義路徑
    $path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"

    # 如果路徑不存在則建立（針對新使用者環境）
    if (-not (Test-Path $path)) { New-Item -Path $path -Force }

    # 設定圖示顯示（0 代表「顯示」，1 代表「隱藏」）
    Set-ItemProperty -Path $path -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Value 0  # 本機 (電腦)
    Set-ItemProperty -Path $path -Name "{21EC2020-3AEA-1069-A2DD-08002B30309D}" -Value 0  # 控制台
    Set-ItemProperty -Path $path -Name "{645FF040-5081-101B-9F08-00AA002F954E}" -Value 0  # 資源回收筒

    # 重新整理桌面以套用變更
    (New-Object -ComObject Shell.Application).Namespace(0).Self.InvokeVerb("Properties")
    Stop-Process -Name explorer -Force
}

function Set-Default-WSUS {
    # 定義目標 WSUS 伺服器網址
    $targetWSUS = "http://10.101.188.68:8530"

    # 定義註冊表路徑
    $wpPath = "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate"
    $auPath = "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU"

    # 檢查路徑是否存在並讀取目前的 WUServer 設定
    $currentWSUS = $null
    if (Test-Path $wpPath) {
        $currentWSUS = (Get-ItemProperty -Path $wpPath -Name "WUServer" -ErrorAction SilentlyContinue).WUServer
    }

    # --- 開始邏輯判斷 ---
    if ($currentWSUS -eq $targetWSUS) {
        Write-Log -Message "設定檢查：目前 WSUS 伺服器已正確設定為 $targetWSUS。無需變更。"
    } else {
        Write-Log -Message "設定檢查：偵測到設定不符（目前為：$currentWSUS），正在更新為 $targetWSUS..."

        # 1. 確保路徑存在
        if (-not (Test-Path $wpPath)) { New-Item -Path $wpPath -Force | Out-Null }
        if (-not (Test-Path $auPath)) { New-Item -Path $auPath -Force | Out-Null }

        # 2. 寫入 WindowsUpdate 核心設定
        $wpSettings = @{
            "WUServer" = $targetWSUS
            "WUStatusServer" = $targetWSUS
            "UpdateServiceUrlAlternate" = ""
            "DoNotEnforceEnterpriseTLSCertPinningForUpdateDetection" = 1
        }
        foreach ($name in $wpSettings.Keys) {
            Set-ItemProperty -Path $wpPath -Name $name -Value $wpSettings[$name] -Type String -ErrorAction SilentlyContinue
        }
        # 修正 DWord 型別
        Set-ItemProperty -Path $wpPath -Name "DoNotEnforceEnterpriseTLSCertPinningForUpdateDetection" -Value 1 -Type DWord

        # 3. 寫入 AU (自動更新) 設定
        $auSettings = @{
            "DetectionFrequencyEnabled" = 1
            "DetectionFrequency"        = 4
            "AutoInstallMinorUpdates"   = 1
            "AUOptions"                 = 4
            "ScheduledInstallTime"      = 12
            "ScheduledInstallEveryWeek" = 1
            "AllowMUUpdateService"      = 1
            "UseWUServer"               = 1
        }
        foreach ($name in $auSettings.Keys) {
            Set-ItemProperty -Path $auPath -Name $name -Value $auSettings[$name] -Type DWord
        }

        # 4. 刪除不需要的舊標記 (對應 **del.)
        $toDelete = @("FillEmptyContentUrls", "AutomaticMaintenanceEnabled", "ScheduledInstallFirstWeek", "ScheduledInstallSecondWeek", "ScheduledInstallThirdWeek", "ScheduledInstallFourthWeek")
        foreach ($item in $toDelete) {
            Remove-ItemProperty -Path $wpPath -Name $item -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path $auPath -Name $item -ErrorAction SilentlyContinue
        }

        # 5. 重啟服務使設定生效
        Write-Log -Message "正在重啟 Windows Update 服務以套用變更..."
        Restart-Service -Name wuauserv -Force
        
        Write-Log -Message "設定更新完成！"
    }
}

# Update-7zip -TargetVersion 25.01
function Update-7zip {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TargetVersion,
        [ValidateSet("x64", "x86", "arm64")]
        [string]$Architecture = "x64"
    )
    # 1. Load the external logging tool (Dot Sourcing)
    
    Write-Log -Message "[7-Zip] Checking installation status..."

    try {
        # Define Registry paths for both 64-bit and 32-bit software lists
        $regPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        # Locate installed 7-Zip using DisplayName
        $installedApp = Get-ItemProperty $regPaths -ErrorAction SilentlyContinue | 
                        Where-Object { $_.DisplayName -like "*7-Zip*" } | 
                        Select-Object -First 1

        # Version comparison logic
        $current7zVer = if ($installedApp) { 
            [version]($installedApp.DisplayVersion -replace '[^0-9.]', '') 
        } else { 
            [version]"0.0.0.0" 
        }
        
        $target7zVerObj = [version]$TargetVersion

        Write-Log -Message "[7-Zip] Current: $current7zVer | Target: $target7zVerObj"

        if ($current7zVer -lt $target7zVerObj) {
            Write-Log -Message "[7-Zip] Update required. Preparing download..."
            
            # Optimization: Format version for URL (e.g., 24.01 -> 2401)
            $verClean = $TargetVersion.Replace(".", "")
            $fileName = "7z$verClean-$Architecture.exe"
            $downloadUrl = "https://www.7-zip.org/a/$fileName"
            $exePath = Join-Path $env:TEMP $fileName

            # Download installer
            Write-Log -Message "[7-Zip] Downloading from $downloadUrl"
            Invoke-WebRequest -Uri $downloadUrl -OutFile $exePath -ErrorAction Stop

            # Silent Installation (/S)
            Write-Log -Message "[7-Zip] Executing silent installation..."
            $proc = Start-Process -FilePath $exePath -ArgumentList "/S" -Wait -PassThru
            
            if ($proc.ExitCode -eq 0) {
                Write-Log -Message "[7-Zip] Installation successful."
            } else {
                Write-Log -Message "[7-Zip] Installation failed with ExitCode: $($proc.ExitCode)"
            }

            # Cleanup temporary installer
            if (Test-Path $exePath) {
                Remove-Item $exePath -Force
            }
        } else {
            Write-Log -Message "[7-Zip] System is already up to date."
        }
    } catch {
        Write-Log -Message "[7-Zip] CRITICAL ERROR: $($_.Exception.Message)"
        exit 1
    }
}

# Update-Hicos -TargetVersion 1.3.4.103349
function Update-Hicos {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TargetVersion,
        [string]$DownloadUrl = "https://api-hisecurecdn.cdn.hinet.net/MOICA/HiCOS_Client.zip",
        [string]$ServiceUrl = "http://127.0.0.1:61161"
    )

    Write-Log -Message "[HiCOS] Checking version status..."

    try {
        # Attempt to get the current version from the local service
        $response = Invoke-WebRequest -Uri $ServiceUrl -TimeoutSec 5 -ErrorAction Stop
        
        # Extract version using Regex from the response content
        if ($response.Content -match 'version:(?<version>[\d\.]+)') {
            $currentVersion = [version]$Matches['version']
            $targetVersionObj = [version]$TargetVersion
            
            Write-Log -Message "[HiCOS] Current: $currentVersion | Target: $targetVersionObj"

            if ($currentVersion -ne $targetVersionObj) {
                Write-Log -Message "[HiCOS] Version mismatch. Starting update sequence..."
                
                $zipPath = Join-Path $env:TEMP "hicos.zip"
                $extractDir = Join-Path $env:TEMP "hicos_extracted"

                # Download package
                Write-Log -Message "[HiCOS] Downloading package from $DownloadUrl"
                Invoke-WebRequest -Uri $DownloadUrl -OutFile $zipPath -ErrorAction Stop
                
                # Extract package
                Write-Log -Message "[HiCOS] Extracting files..."
                if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
                Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
                
                # Locate and run installer
                $exePath = Join-Path $extractDir "HiCOS_Client.exe"
                if (Test-Path $exePath) {
                    Write-Log -Message "[HiCOS] Executing silent installation..."
                    # Parameters /install /quiet /norestart are standard for this installer
                    $proc = Start-Process -FilePath $exePath -ArgumentList "/install", "/quiet", "/norestart" -Wait -PassThru
                    Write-Log -Message "[HiCOS] Installation finished (ExitCode: $($proc.ExitCode))."
                } else {
                    Write-Log -Message "[HiCOS] Error: HiCOS_Client.exe not found in extracted package."
                    exit 1
                }

                # Cleanup temporary files
                Write-Log -Message "[HiCOS] Cleaning up temporary files..."
                if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
                if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }
                
            } else {
                Write-Log -Message "[HiCOS] System is up to date ($currentVersion). Skipping update."
            }
        } else {
            Write-Log -Message "[HiCOS] Error: Could not parse version from service response."
        }
    } catch {
        Write-Log -Message "[HiCOS] Warning: Service unreachable at $ServiceUrl. Ensure HiCOS is installed and running."
        # Optionally: Trigger a fresh install logic here if the service is missing
    }
}

# -------------------- hide user update popup window --------------------
$newAction = New-ScheduledTaskAction -Execute "mshta" -Argument "vbscript:Execute(""CreateObject(""""WScript.Shell"""").Run """"powershell -Command IEX (Invoke-RestMethod -Uri 'https://raw.githubusercontent.com/ShirakamiFubuking/Scripts/refs/heads/main/user.ps1')"""",0,True:close()"")"
Set-ScheduledTask -TaskName "pwb_update_user" -Action $newAction

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

# -------------------- Install psm --------------------
# $url = "http://128.5.47.252/Module.zip"
# $tempZip = Join-Path $env:TEMP "Module.zip"
# $destPath = "C:\Program Files\WindowsPowerShell\Modules"

# try {
#     # 2. 檢查目標目錄是否存在，不存在則建立
#     if (!(Test-Path $destPath)) {
#         Write-Log "Creating directory: $destPath" -ForegroundColor Cyan
#         New-Item -ItemType Directory -Path $destPath -Force | Out-Null
#     }

#     # 3. 下載檔案
#     Write-Log "Downloading $url ..." -ForegroundColor Cyan
#     Invoke-WebRequest -Uri $url -OutFile $tempZip -ErrorAction Stop

#     # 4. 解壓縮檔案
#     # -Force 參數確保若檔案已存在會直接覆蓋
#     Write-Log "Extracting to $destPath ..." -ForegroundColor Cyan
#     Expand-Archive -Path $tempZip -DestinationPath $destPath -Force -ErrorAction Stop

#     # 5. 清理暫存檔
#     Remove-Item $tempZip -Force
#     Write-Log "Successfully installed module to $destPath" -ForegroundColor Green
# }
# catch {
#     Write-Error "An error occurred: $($_.Exception.Message)"
# }
Remove-Item -Recurse "C:\Program Files\WindowsPowerShell\Modules\LogTool"
Remove-Item -Recurse "C:\Program Files\WindowsPowerShell\Modules\Report"
Remove-Item -Recurse "C:\Program Files\WindowsPowerShell\Modules\UpdateTools"
Remove-Item -Recurse "C:\Program Files\WindowsPowerShell\Modules\GeTools"
Report "http://128.5.47.252:5000/report"

# -------------------- Pre-Flight Checks --------------------
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Log -Message "Access Denied: Please run this script as Administrator."
    exit
}

Set-Default-WSUS

# ==================== Hosts Update ====================
Write-Log -Message "[Hosts] Updating hosts file..."
try {
    Write-Log -Message "[Hosts] Fetching latest content from source..."
    $newHostsContent = Invoke-RestMethod -Uri $Config.Hosts.SourceUrl -UseBasicParsing

    if ($null -ne $newHostsContent -and $newHostsContent.Length -ge 10) {
        $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllLines($Config.Hosts.Path, $newHostsContent, $Utf8NoBomEncoding)
        Write-Log -Message "[Hosts] Update successful."

        ipconfig /flushdns | Out-Null
        Write-Log -Message "[Hosts] DNS cache flushed."
    } else {
        Write-Log -Message "[Hosts] Error: Downloaded content is invalid or empty."
    }
} catch {
    Write-Log -Message "[Hosts] Failed to update: $($_.Exception.Message)"
}

# ==================== Registry Update ====================
Write-Log -Message "[Registry] Updating system registry..."
try {
    $PolicyPaths = @($Config.BrowserPolicies.Chrome, $Config.BrowserPolicies.Edge)

    foreach ($keyPath in $PolicyPaths) {
        if (-not (Test-Path $keyPath)) {
            New-Item -Path $keyPath -Force | Out-Null
            Write-Log -Message "[Registry] Created key: $keyPath"
        }
        $Config.BrowserPolicies.Urls.GetEnumerator() | ForEach-Object {
            Set-ItemProperty -Path $keyPath -Name $_.Key -Value $_.Value -Type String
        }
    }
    Write-Log -Message "[Registry] Browser policies applied successfully."
} catch {
    Write-Log -Message "[Registry] Failed to apply settings: $($_.Exception.Message)"
}

# ==================== HiCOS Update ====================
Write-Log -Message "[HiCOS] Checking version..."
Update-Hicos $Config.Hicos.TargetVersion

# ==================== 7-Zip Update ====================
Write-Log -Message "[7-Zip] Checking installation..."
Update-7zip $Config.SevenZip.TargetVersion