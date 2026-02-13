:: 安裝憑證
cd %TEMP%
powershell -Command Invoke-WebRequest -Uri "https://raw.githubusercontent.com/ShirakamiFubuking/Scripts/refs/heads/main/cert/PSCert.cer" -OutFile "PSCert.cer"
powershell -Command Import-Certificate -FilePath "PSCert.cer" -CertStoreLocation Cert:\LocalMachine\Root
del PSCert.cer

:: 設定排程
schtasks /create /tn "pwb_update" /tr "powershell -Command Invoke-RestMethod -Uri 'https://raw.githubusercontent.com/ShirakamiFubuking/Scripts/refs/heads/main/main.ps1' | Powershell -Command -" /sc hourly /mo 1 /rl highest /f /np