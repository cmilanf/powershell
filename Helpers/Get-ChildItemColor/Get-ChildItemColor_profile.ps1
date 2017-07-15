# BEGIN Get-ChildItemColor
Write-Host -NoNewline "Loading Get-ChildItemColor module..."
$psprofile=Split-Path $profile.CurrentUserCurrentHost
. "$psprofile\Get-ChildItemColor.ps1" # read the colourized ls
set-alias ls Get-ChildItemColor -force -option allscope
function Get-ChildItem-Force { ls -Force }
set-alias la Get-ChildItem-Force -option allscope
Write-Host -ForegroundColor Green "[OK]"
# END Get-ChildItemColor 