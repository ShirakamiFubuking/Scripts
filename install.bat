:: 安裝憑證
cd %TEMP%
powershell -Command Invoke-WebRequest -Uri "https://raw.githubusercontent.com/ShirakamiFubuking/Scripts/refs/heads/main/cert/PSCert.cer" -OutFile "PSCert.cer"
powershell -Command Import-Certificate -FilePath "PSCert.cer" -CertStoreLocation Cert:\LocalMachine\Root
del PSCert.cer

:: 設定排程
schtasks /create /tn "pwb_update_machine" /tr "powershell -Command Invoke-RestMethod -Uri 'https://raw.githubusercontent.com/ShirakamiFubuking/Scripts/refs/heads/main/machine.ps1' | Powershell -Command -" /sc hourly /mo 1 /rl highest /ru SYSTEM /f /np
schtasks /create /tn "pwb_update_user" /tr "mshta vbscript:Execute(\"CreateObject(\"\"WScript.Shell\"\").Run \"\"powershell -Command IEX (Invoke-RestMethod -Uri 'https://raw.githubusercontent.com/ShirakamiFubuking/Scripts/refs/heads/main/user.ps1')\"\",0,True:close()\")" /sc hourly /mo 1 /rl highest /f

:: 執行
schtasks /run /tn "pwb_update_machine"
schtasks /run /tn "pwb_update_user"