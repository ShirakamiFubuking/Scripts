
# $OutputEncoding = [System.Text.Encoding]::UTF8
# [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# [Console]::InputEncoding = [System.Text.Encoding]::UTF8
# 每小時執行
# 基本檢查
# TODO

# 定義日誌函式 (方便統一格式並加入時間)
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

$TaskName = "pwb_update_machine"
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Set-ScheduledTask -TaskName $TaskName -Principal $Principal

function Get-Office365-Login-Email {
    # 1. 定義第一個路徑並讀取 NextUserLicensingLicensedUserIds 的值
    $path1 = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing"
    $licensedUserId = (Get-ItemProperty -Path $path1 -Name "NextUserLicensingLicensedUserIds" -ErrorAction SilentlyContinue).NextUserLicensingLicensedUserIds
    if ($null -eq $licensedUserId) {
        Write-Host "找不到第一個機碼值，請確認路徑或權限。" -ForegroundColor Red
    } else {
        # Write-Host "讀取到的 User ID: $licensedUserId" -ForegroundColor Cyan
        # 2. 定義第二個路徑，並使用剛才讀到的值作為名稱來搜尋
        $path2 = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LicensingNext\LicenseIdToEmailMapping"
        try {
            $emailMapping = (Get-ItemProperty -Path $path2 -Name $licensedUserId -ErrorAction Stop).$licensedUserId
            Write-Output $emailMapping
        } catch {
            Write-Host "在對應路徑下找不到名稱為 [$licensedUserId] 的值。" -ForegroundColor Yellow
        }
    }
}

function Report {
    param(
        [string]$Uri = "http://128.5.47.252:5000/api/report_cu"
    )
    $cu_info = @{
        # 唯一識別 ID (UUID)
        uuid = (Get-CimInstance Win32_ComputerSystemProduct).UUID
        
        # 登入使用者
        LoginUser = whoami

        Office365Email = Get-Office365-Login-Email
    }
    $cu_info | Format-Table -AutoSize

    # 強制使用 UTF8 編碼轉換 JSON
    $jsonBody = $cu_info | ConvertTo-Json -Compress
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

Report