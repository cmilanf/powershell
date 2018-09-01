<#
.SYNOPSIS
    Creates a Volume Shadow Copy in order to backup and iSCSI server VHD volume that is currently in use, upload the file to an Storage Account and change it to Archive tier. This script must be run with Administrator privilege to access VSS.
.DESCRIPTION
    This script leverages the Azure Storage ARCHIVE tier for storing high volumes of data, such as VHD files served through iSCSI. As this could be troublesome due to the files being in use, this script takes the approach of using Volume Shadow Copy for allowing the operation without actual downtime of the disk.
.PARAMETER LogFile
    Path and filename of the operation log. It can be a UNC path.
    Example: "\\NAS\logs\iSCSI-backup.log"
.PARAMETER BlobContainerName
    Blob container name in the Azure Storage account.
.PARAMETER StorageAccountName
    Azure Storage account name.
.PARAMETER StorageAccountKey
    Key for accessing the storage account.
.PARAMETER BackupFile
    Path and filename of the the file to backup.
    Example: "E:\DATA.VHDX"
.PARAMETER DaysToRemoveOldObject
    Objects found in storage account older that this value will be removed
.PARAMETER ShadowDriveLetter
    Drive letter that should be use for Shadow Copy. It must not be in use by the system.
.EXAMPLE
    BackupISCSI-ToAzure.ps1 -BackupFile "E:\mymassivefile.vhdx" -StorageAccountName "backups" -StorageAccountKey "fh328wtjgewace098j3d39DSVMVCHNA9P" -BlobContainerName "iscsi" -LogFile "\\NAS\iSCSI-backup.log"
.LINK
    https://calnus.com
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
    [Parameter(Mandatory=$true)][string]$BackupFile,
    [Parameter(Mandatory=$true)][string]$StorageAccountName,
    [Parameter(Mandatory=$true)][string]$StorageAccountKey,
    [Parameter(Mandatory=$true)][string]$BlobContainerName,
    [string]$LogFile,
    [int]$DaysToRemoveOldObject=181,
    [string]$ShadowDriveLetter="S:",
    [string]$TempFolder="C:\Windows\TEMP"
)

$ErrorActionPreference = "Stop"
if(Get-Module -ListAvailable -Name "AzureRM.Storage") {
    Import-Module -Name AzureRM.Storage
} else {
    throw "AzureRM.Storage module not found"
}

$date = (Get-Date).ToString("yyyyMMdd")
$context = New-AzureStorageContext -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey
$isOldDate = [DateTime]::UtcNow.AddDays(-$DaysToRemoveOldObject)
# $backupFileNameExtensionLess is just the filename without extension. Example: "myfile"
$backupFileNameExtensionLess = [System.IO.Path]::GetFileNameWithoutExtension($BackupFile)
# $backupFileNameExtension is the extension of the filename WITH dot. Example: ".vhdx"
$backupFileNameExtension = [System.IO.Path]::GetExtension($BackupFile)
# $backupFileNameDriveLetter is the driver letter followed by colon. Example: "E:"
$backupFileNameDriveLetter = (Get-Item $BackupFile) | Split-Path -Qualifier

# ##### DISKSHADOW SCRIPT #####
if($backupFileNameDriveLetter -eq $ShadowDriveLetter) {
    throw "Backup file and shadow file cannot share the same drive letter"
}
$DiskShadowScriptStart = "set context persistent nowriters
set verbose on
add volume $backupFilenameDriveLetter alias vssBackup
create
expose %vssBackup% $ShadowDriveLetter"
$DiskShadowScriptStop = "set verbose on
delete shadows exposed $ShadowDriveLetter"

if (Test-Path $(Split-Path -Path $LogFile -Parent) -PathType Container) {
    Start-Transcript -Path $LogFile -Force
}
Write-Output "File: $BackupFile"
Write-Output "Azure Storage Account: $StorageAccountName"
Write-Output "Blob Container: $BlobContainerName"
Write-Output "Days to remove old objects: $DaysToRemoveOldObject"
Write-Output ""

# ##### VOLUME SHADOW COPY CREATION #####
$DiskShadowScriptStart | Out-File -FilePath "$TempFolder\diskshadow-start.txt" -Encoding ascii
cmd.exe /c "diskshadow /s $TempFolder\diskshadow-start.txt"
Remove-Item -Path "$TempFolder\diskshadow-start.txt"
$BackupFileShadow = $BackupFile.Remove(0,2).Insert(0,$ShadowDriveLetter)

# ##### AZURE STORAGE UPLOAD & TIER SET #####
Set-AzureStorageBlobContent -Blob "$backupFileNameExtensionLess-$date$backupFilenameExtension" -Container $BlobContainerName -File $BackupFileShadow -Context $context -Force -Verbose
if($?) {
    $blob=Get-AzureStorageBlob -Blob "$backupFileNameExtensionLess-$date$backupFilenameExtension" -Container $BlobContainerName -Context $context -Verbose
    Write-Output "Setting storage tier to ARCHIVE..."
    $blob.ICloudBlob.SetStandardBlobTier("Archive")
    if($?) {
        Write-Output "Cleaning old files..."
        Get-AzureStorageBlob -Context $context -Container $BlobContainerName | Where-Object { $_.LastModified.UtcDateTime -lt $isOldDate -and $_.BlobType -eq "BlockBlob" } | Remove-AzureStorageBlob -Verbose
    }
}
# ##### VOLUME SHADOW COPY DELETION #####
$DiskShadowScriptStop | Out-File -FilePath "$TempFolder\diskshadow-stop.txt" -Encoding ascii
cmd.exe /c "diskshadow /s $TempFolder\diskshadow-stop.txt"
Remove-Item -Path "$TempFolder\diskshadow-stop.txt"
if (Test-Path $(Split-Path -Path $LogFile -Parent) -PathType Container) {
    Stop-Transcript
}