# Copyright David Moravec, PowerShell Magazine
# http://www.powershellmagazine.com/2012/10/30/pstip-storing-of-credentials/

function New-DMCredential
{
   [CmdletBinding()]
 
   param(
      [Parameter(Mandatory = $true, Position = 0)]
      [string]$UserName,
	  
      [Parameter( Mandatory = $true, Position = 1)]
      [string]$Alias,
	  
	  [Parameter(Mandatory = $false, Position = 2)]
      [string]$Path
   )
 
   # Where the credentials will be stored
   if([string]::isNullorEmpty($Path))
   {
		$Path = $env:userprofile
   }
	
   # get credentials for given username
   $cred = Get-Credential $UserName
 
   # and save encrypted text to a file
   $cred.Password | ConvertFrom-SecureString | Set-Content -Path ("$path\$Alias")
}
 
function Get-DMCredential
{
   [CmdletBinding()]
 
   param(
      [Parameter(Mandatory = $true, Position = 0)]
      [string]$UserName,
	  
      [Parameter(Mandatory = $true, Position = 1)]
      [string]$Alias,
	  
	  [Parameter(Mandatory = $false, Position = 2)]
      [string]$Path
   )
 
   # where to load credentials from
   if([string]::isNullorEmpty($Path))
   {
		$Path = $env:userprofile
   }
 
   # receive cred as a PSCredential object
   $pwd = Get-Content -Path ("$path\$Alias") | ConvertTo-SecureString
   $cred = New-Object System.Management.Automation.PSCredential $UserName, $pwd
 
   # assign a cred to a global variable based on input
   Invoke-Expression "`$Global:$($Alias) = `$cred"
   Remove-Variable -Name cred
   Remove-Variable -Name pwd
}
 
function Show-DMCredential
{
   param($cred)
 
   Write-Host "Password is about to be shown in clear text. Press a key to continue o CTRL+C to cancel..."
   $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
   # Just to see password in clear text
   [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.Password))
}