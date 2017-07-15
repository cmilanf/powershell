Write-Host "Welcome to the DM Credentials installer"
Write-Host ""
Write-Host "This module allows saving yor PSCredential to disk in order to restore it whenever you want."
Write-Host "This is particulary useful to autoload your admin credentials at runspace opening."
Write-Host ""
Write-Host -NoNewLine "The script will set a Get-DMCredential at startup with "
Write-Host -ForegroundColor Yellow -NoNewline "pscred"
Write-Host " filename (alias)."
Write-Host "Make sure you use "
Write-Host -ForegroundColor Yellow "New-DMCredential -UserName <youradmhere@yourdomain.com> -Alias pscred"
Write-Host "Your credentials will be available at `$global:pscred"
Write-Host ""

$PSPROFILE = "C:\Users\$env:username\Documents\WindowsPowerShell"
$file = "DMCredential.ps1"
$file_profile = "DMCredential_profile.ps1"

if(!(Test-Path $PSPROFILE))
{ 
    Write-Host "'$PSPROFILE' not found, creating..."
    mkdir $PSPROFILE
}
Write-Host "Copying $file..."
copy $file $PSPROFILE
Write-Host "Appending aliases and functions to PowerShell user profile..."
Get-Content $file_profile | Out-File -FilePath "$PSPROFILE\Microsoft.PowerShell_profile.ps1" -Append
Write-Host "Installation finished"