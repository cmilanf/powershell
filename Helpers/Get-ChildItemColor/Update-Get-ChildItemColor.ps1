Write-Host -NoNewline "Welcome to the "
Write-Host -ForegroundColor Green -NoNewLine "C"
Write-Host -ForegroundColor Red -NoNewLine "O"
Write-Host -ForegroundColor Yellow -NoNewLine "L"
Write-Host -ForegroundColor Cyan -NoNewLine "O"
Write-Host -ForegroundColor Magenta -NoNewLine "R"
Write-Host " Get-ChildItem installer"
Write-Host ""
$PSPROFILE = "C:\Users\$env:username\Documents\WindowsPowerShell"
$file = "Get-ChildItemColor.ps1"
$file_profile = "Get-ChildItemColor_profile.ps1"

if(!(Test-Path $PSPROFILE))
{ 
    Write-Host "'$PSPROFILE' not found, creating..."
    mkdir $PSPROFILE
}
Write-Host "Copying $file..."
copy $file $PSPROFILE
Write-Host "Update finished"