<#
.SYNOPSIS
    Generates a RDCMan connection file with IP addresses from Azure Virtual Machines.
.DESCRIPTION
    This PowerShell script will use your active Azure Resource Manager subscription and get all the virtual machines IP addresses from a resource group. After that it will build a RDCMan compatible file with the retrieved information, easying the connection to the resources.
    
    You can download RDCMan from https://www.microsoft.com/en-us/download/details.aspx?id=44989
.PARAMETER ResourceGroupName
    The resource group name where we are going to search for Virtual Machines. If you use * in this parameter it will retrieve the virtual machines from every single resource group in the subscription.
.PARAMETER Name
    The name our RDCMan tree will have. If empty it will take the ResourceGroupName value.
    Provide a PARAMETER section for each parameter that your script or function accepts.
.PARAMETER IncludeLinux
    Usually, you won't connect through RDP to Linux virtual machines, so they won't be of any use to RDCMan, but if you want to include them, turn on this switch.
.PARAMETER PreferPublicIP
    If the script finds a virtual machines with both, public and private IP addresses, prefer the public one.
.PARAMETER OutputXmlFile
    The file name used to save the generated RDCMan config file. Path will be current script execution.
.EXAMPLE
    Generate-AzureRmRdcFile.ps1 -ResourceGroupName * -Name "HOME Machines" -PreferPublicIP -OutputXmlFile home.rdg
.LINK
    https://calnus.com
.LINK
    RDCMan download: https://www.microsoft.com/en-us/download/details.aspx?id=44989
.NOTES
    Carlos Milán Figueredo
    https://www.calnus.com

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
    INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
    PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
    FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
    OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
    DEALINGS IN THE SOFTWARE.
#>

param(
    [Parameter(Mandatory=$true)]
    [String]$ResourceGroupName,
    [Parameter(Mandatory=$true)]
    [String]$Name=$ResourceGroupName,
    [Switch]$IncludeLinux,
    [Switch]$PreferPublicIP,
    [Parameter(Mandatory=$true)]
    [String]$OutputXmlFile
    )

$rdgXmlFile = [XML]'<?xml version="1.0" encoding="utf-8"?>
<RDCMan programVersion="2.7" schemaVersion="3">
  <file>
    <properties>
      <expanded>True</expanded>
      <name></name>
    </properties>
  </file>
  <connected />
  <favorites />
  <recentlyUsed />
</RDCMan>'

Write-Output "Get-AzureRMmRdcFile.ps1 - The Azure Resource Manager RDCMan generator"
Write-Output "by Carlos Milán Figueredo - https://www.calnus.com"

$rdgXmlFile.RDCMan.file.properties.name = $Name

if($ResourceGroupName -eq '*')
{
    $VMs = Get-AzureRmVm
} else {
    $VMs = Get-AzureRmVm -ResourceGroupName $ResourceGroupName
}
Write-Output "Found $($VMs.Count) virtual machines..."
foreach($vm in $VMs)
{
    Write-Output ""
    Write-Output "========== $($vm.name) is based on $($vm.StorageProfile.OsDisk.OsType) OS"
    Write-Output "Found $($vm.NetworkInterfaceIDs.Count) network interface(s)"
    if(($vm.StorageProfile.OsDisk.OsType -eq "Windows") -or $IncludeLinux)
    {
        foreach($nicId in $vm.NetworkInterfaceIDs)
        {
            $nic = Get-AzureRmResource -ResourceId $nicId
            Write-Output "Got $($nic.Name) interface."
            Write-Output "* Private IP: $($nic.Properties.ipConfigurations.Properties.privateIPAddress)"
            Write-Output "* Private IP Allocation Method: $($nic.Properties.ipConfigurations.Properties.privateIPAllocationMethod)"
            $ip=$nic.Properties.ipConfigurations.Properties.privateIPAddress
            try
            {
                $pip = Get-AzureRmResource -ResourceId $nic.Properties.ipConfigurations.Properties.PublicIpAddress.Id
                Write-Output "* Public IP: $($pip.Properties.ipAddress)"
                Write-Output "* Public IP Allocation Method: $($pip.Properties.publicIPAllocationMethod)"
                if($PreferPublicIp)
                {
                    $ip=$pip.Properties.ipAddress
                }
            } catch {
                Write-Output "* Public IP: not found"
            }
            if($vm.NetworkInterfaceIDs.Count -le 1)
            {
                $xmlDisplayNameTextNode=$rdgXmlFile.CreateTextNode($vm.Name)
            } else {
                $xmlDisplayNameTextNode=$rdgXmlFile.CreateTextNode($nic.Name)
            }
            $xmlNameTextNode=$rdgXmlFile.CreateTextNode($ip)
            $xmlServerElement=$rdgXmlFile.CreateElement("server")
            $xmlPropElement=$rdgXmlFile.CreateElement("properties")
            $xmlNameElement=$rdgXmlFile.CreateElement("name")
            $xmlDisplayNameElement=$rdgXmlFile.CreateElement("displayName")
            [void]$xmlDisplayNameElement.AppendChild($xmlDisplayNameTextNode)
            [void]$xmlNameElement.AppendChild($xmlNameTextNode)
            [void]$xmlPropElement.AppendChild($xmlDisplayNameElement)
            [void]$xmlPropElement.AppendChild($xmlNameElement)
            [void]$xmlServerElement.AppendChild($xmlPropElement)
            [void]$rdgXmlFile.RDCMan.file.AppendChild($xmlServerElement)
        }
    }
}
Write-Output "`r`nWritting file to $PSScriptRoot\$OutputXmlFile"
$rdgXmlFile.Save("$PSScriptRoot\$OutputXmlFile")
Write-Output "Done!"
