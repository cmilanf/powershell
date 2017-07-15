# BEGIN Autoload Credential

# Function Get-CurrentUserPrincipalName credit:
# https://digitaljive.wordpress.com/2013/03/27/powershell-get-current-user-principle-name-upn/
Function Get-CurrentUserPrincipalName
{
  $strFilter = “(&(objectCategory=User)(SAMAccountName=$Env:USERNAME))”
  $objDomain = New-Object System.DirectoryServices.DirectoryEntry
  $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
  $objSearcher.SearchRoot = $objDomain
  $objSearcher.PageSize = 1
  $objSearcher.Filter = $strFilter
  $objSearcher.SearchScope = “Subtree”
  $objSearcher.PropertiesToLoad.Add(“userprincipalname”) | Out-Null
  $colResults = $objSearcher.FindAll()

  $UPN = $colResults[0].Properties.userprincipalname
  return $UPN
}

$upn = Get-CurrentUserPrincipalName
Write-Host -NoNewline "Loading administrator credentials module..."
. "C:\Users\$env:username\Documents\WindowsPowerShell\DMCredential.ps1"
if(Test-Path -Path "C:\Users\$env:username\pscred")
{
    Get-DMCredential -UserName $upn -Alias pscred
    Write-Host -ForegroundColor Green "[OK]"
} else {
    Write-Host -ForegroundColor Red "[Credentials not found]"
}
# END Autoload Credential
