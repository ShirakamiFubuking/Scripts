# $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=Ge" -KeyUsage DigitalSignature -FriendlyName "Ge Code Signing Cert" -NotAfter (Get-Date).AddYears(5)
# 1. 取得憑證
$cert = Get-ChildItem Cert:\LocalMachine\My\ -CodeSigningCert | Select-Object -First 1

# 2. 先確保檔案是帶 BOM 的 UTF-8 格式 (防止簽署時轉換失敗)
$content = Get-Content -Path "machine.ps1" -Raw
$content | Out-File -FilePath "machine.ps1" -Encoding utf8
# 3. 對腳本進行簽署
Set-AuthenticodeSignature -FilePath "machine.ps1" -Certificate $cert


$content = Get-Content -Path "user.ps1" -Raw
$content | Out-File -FilePath "user.ps1" -Encoding utf8
Set-AuthenticodeSignature -FilePath "user.ps1" -Certificate $cert