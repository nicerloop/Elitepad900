
$ProgressPreference = "SilentlyContinue"

# Define work folders location
$ScriptFolder = $PSScriptRoot
$DownloadsFolder = (Join-Path -Path $HOME -ChildPath "Downloads")
$WorkDirectory = (Join-Path -Path $HOME -ChildPath "ELITEPAD900")
$DriversFolder = (Join-Path -Path $WorkDirectory -ChildPath "drivers")
$BackupFolder = (Join-Path -Path $ScriptFolder -ChildPath "Downloads")



function Copy-Files {
    param (
        $SourceFolder,
        $DestinationFolder,
        $FileName
    )
    Write-Host "Copy $FileName from $SourceFolder to $DestinationFolder"
    # https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
    Robocopy.exe $SourceFolder $DestinationFolder $FileName /E /Z /A-:R /IM /NDL /NFL /NJH /NJS | Out-Null
}

function Start-Curl {
    param (
        $File,
        $Url
    )
    $UserAgent = "Mozilla/5.0 (X11; Linux i586; rv:100.0) Gecko/20100101 Firefox/100.0"
    # https://curl.se/docs/manpage.html
    curl.exe -o $File -# -C - -A $UserAgent $Url
}

# Download if absent
function Get-File {
    param (
        $FileName,
        $DownloadFolder,
        $BackupFolder,
        $Url,
        $Description
    )
    $DownloadFile = (Join-Path -Path $DownloadFolder -ChildPath $FileName)
    if (Test-Path -Path $DownloadFile -PathType Leaf) {
        Write-Host "Found $Description as $DownloadFile"
    }
    else {
        $BackupFile = (Join-Path -Path $BackupFolder -ChildPath $FileName)
        if (Test-Path -Path $BackupFile -PathType Leaf) {
            Write-Host "Found $Description as $BackupFile"
            Copy-Files -SourceFolder $BackupFolder -DestinationFolder $DownloadFolder -FileName $FileName
        }
        else {
            Write-Host "Download $Description from $Url"
            Start-Curl -File $DownloadFile -Url $Url
            Copy-Files -SourceFolder $DownloadFolder -DestinationFolder $BackupFolder -FileName $FileName
        }
    }
}

# Download HP Elitepad900 drivers for Windows 8.1
# https://support.hp.com/us-en/drivers/selfservice/hp-elitepad-900-g1-tablet/5298028

$HpSoftPacks = @(
    (71504, "HP Elitepad 900 System BIOS/Firmware and Driver Update 1.0.2.3 Rev.A Jun 3, 2015"),
    (65877, "Qualcomm Atheros AR600x 802.11a/b/g/n Wireless LAN Driver 3.7 C Apr 18, 2014"),
    (65376, "Qualcomm Atheros AR3002 Bluetooth 4.0+HS Driver for Microsoft Windows 2.2 Rev.A Feb 12, 2014"),
    (64673, "NXP Semiconductors Near Field Proximity (NFP) Driver 1.4.7.2 Dec 2, 2013"),
    (64682, "Broadcom GPS Driver 19.17 Dec 6, 2013")
)
For ($SoftPackIndex = 0; $SoftPackIndex -lt $HpSoftPacks.Length; $SoftPackIndex++) {
    $SoftPackNumber, $SoftPackDescription = $HpSoftPacks[$SoftPackIndex]
    $SoftPackFile = "sp$SoftPackNumber.exe"
    $M = [int][Math]::Floor(($SoftPackNumber - 1) / 1000)
    $D = [int][Math]::Floor((($SoftPackNumber - 1) % 1000) / 500)
    $Min = $M * 1000 + $D * 500 + 1
    $Max = $M * 1000 + ($D + 1) * 500
    $SoftPackUrl = "https://ftp.hp.com/pub/softpaq/sp$Min-$Max/$SoftPackFile"
    Get-File -File $SoftPackFile -DownloadFolder $DownloadsFolder -BackupFolder $BackupFolder -Url $SoftPackUrl -Description "HP SoftPack $SoftPackNumber $SoftPackDescription"
}
For ($SoftPackIndex = 0; $SoftPackIndex -lt $HpSoftPacks.Length; $SoftPackIndex++) {
    $SoftPackNumber, $SoftPackDescription = $HpSoftPacks[$SoftPackIndex]
    $SoftPackFile = "sp$SoftPackNumber.exe"
    $SoftPackFolder = (Join-Path -Path $DriversFolder -ChildPath "sp$SoftPackNumber")
    Write-Host "Unpack HP SoftPack $SoftPackNumber $SoftPackDescription to $SoftPackFolder"
    Start-Process -FilePath (Join-Path -Path $DownloadsFolder -ChildPath $SoftPackFile) -ArgumentList "/s /e /f `"$SoftPackFolder`"" -Wait 
}

# Broadcom GPS Driver installer from Dell
# HP SoftPack provided installer does not run on Windows 10
# Dell provided on runs on Windows 10
# and can be expanded and streamlined from the command line
# https://www.dell.com/support/home/en-us/drivers/driversdetails?driverid=p0p15

$ZipPackFileBaseName = "GPS_BCM4751_W8_A02-P0P15_ZPE"
$ZipPackFileName = "$ZipPackFileBaseName.exe"
$ZipPackUrl = "https://dl.dell.com/FOLDER00998748M/3/GPS_BCM4751_W8_A02-P0P15_ZPE.exe"
$ZipPackDescription = "Dell ZipPack P0P15 BCM47511 Standalone GPS Solution 19.14.6362.4, A02 30 Nov 2012"
Get-File -File $ZipPackFileName -DownloadFolder $DownloadsFolder -BackupFolder $BackupFolder -Url $ZipPackUrl -Description $ZipPackDescription
$ZipPackFolder = (Join-Path -Path $DriversFolder -ChildPath $ZipPackFileBaseName)
Write-Host "Unpack $ZipPackDescription to $ZipPackFolder"
$ZipPackFile = (Join-Path -Path $DownloadsFolder -ChildPath $ZipPackFileName)
$ZipPackArchive = "$ZipPackFileBaseName.zip"
Copy-Item -Path $ZipPackFile -Destination $ZipPackArchive
New-Item -ItemType Directory -Path $ZipPackFolder -Force | Out-Null
Remove-Item -Recurse -Force -Path $ZipPackFolder
# Prevent issues with CLI non-interactive execution
# https://github.com/PowerShell/Microsoft.PowerShell.Archive/issues/77#issuecomment-601947496
$global:ProgressPreference = "SilentlyContinue"
Expand-Archive -Path $ZipPackArchive -DestinationPath $ZipPackFolder
$global:ProgressPreference = "Continue"
Remove-Item -Path $ZipPackArchive

### List drivers
# Get-ChildItem -Recurse -Filter "*.inf" -Path $DriversFolder | Select-Object -ExpandProperty FullName

# Download Windows 10 Pro 22H2 en_US x32 ISO
# https://www.microsoft.com/en-us/software-download/windows10ISO

$ImageFileName = "Win10_22H2_EnglishInternational_x32v1.iso"
$ImagePath = (Join-Path -Path $DownloadsFolder -ChildPath $ImageFileName)
$ImageDescription = "Windows 10 Pro 22H2 en_US x86 ISO"

if (Test-Path $ImagePath -PathType leaf) {
    Write-Host "Found $ImageDescription as $ImagePath"
}
else {
    Write-Host "Get $ImageDescription download URL"
    # Windows ISO Downloader
    # https://github.com/pbatard/Fido
    $FidoFileName = "Fido.ps1"
    Get-File -FileName $FidoFileName -DownloadFolder $DownloadsFolder -BackupFolder $BackupFolder -Url "https://raw.githubusercontent.com/pbatard/Fido/master/Fido.ps1" -Description "Windows ISO downloader"
    $FidoFile = (Join-Path -Path $DownloadsFolder -ChildPath $FidoFileName)
    $ImageUrl = (& $FidoFile -Win 10 -Arch x86 -GetUrl)
    Get-File -FileName $ImageFileName -DownloadFolder $DownloadsFolder -BackupFolder $BackupFolder -Url $ImageUrl -Description $ImageDescription
}

# Get install files from ISO

Write-Host "Mount $ImageDescription"
$DriveLetter = (Mount-DiskImage -ImagePath $ImagePath -PassThru | Get-Volume).DriveLetter
Write-Host "Mounted $ImageDescription as $DriveLetter`:"
Copy-Files -SourceFolder "$DriveLetter`:" -DestinationFolder $WorkDirectory -FileName "*.*"
Write-Host "Unmount $ImageDescription"
Dismount-DiskImage -ImagePath $ImagePath | Out-Null

# Keep only Windows 10 Pro variant

$SourcesFolder = (Join-Path $WorkDirectory "sources")
$InstallWim = (Join-Path $SourcesFolder "install.wim")
$InstallWimSlim = (Join-Path $SourcesFolder "install.slim.wim")

Write-Host "Keep only Windows 10 Pro variant"
$ImageName = Get-WindowsImage -ImagePath $InstallWim | Select-Object -ExpandProperty ImageName | Select-String -Pattern 'Pro' | Select-Object -ExpandProperty Line | Select-String -NotMatch -Pattern ' N', 'Edu', 'Work' | Select-Object -ExpandProperty Line
$SourceIndex = Get-WindowsImage -ImagePath $InstallWim -Name $ImageName | Select-Object -ExpandProperty ImageIndex
Export-WindowsImage -SourceImagePath $InstallWim -SourceIndex $SourceIndex -DestinationImagePath $InstallWimSlim | Out-Null
Remove-Item -Path $InstallWim
Rename-Item -Path $InstallWimSlim -NewName $InstallWim

# Prepare Boot and Install Windows Images
# Create mount point

$BootWim = (Join-Path $SourcesFolder "boot.wim")
$WimMount = "C:\wim_mount"
New-Item -Path $WimMount -ItemType directory -Force | Out-Null

# Add drivers to WinPE image

Write-Host "Add drivers to WinPE image"
Mount-WindowsImage -ImagePath $BootWim -Index 1 -Path $WimMount | Out-Null
Add-WindowsDriver -Path $WimMount -Driver $DriversFolder -Recurse | Out-Null
Dismount-WindowsImage -Path $WimMount -Save | Out-Null
Clear-WindowsCorruptMountPoint | Out-Null

# Add drivers to Setup image

Write-Host "Add drivers to Setup image"
Mount-WindowsImage -ImagePath $BootWim -Index 2 -Path $WimMount | Out-Null
Add-WindowsDriver -Path $WimMount -Driver $DriversFolder -Recurse | Out-Null
Dismount-WindowsImage -Path $WimMount -Save | Out-Null
Clear-WindowsCorruptMountPoint | Out-Null

# Add drivers to Install image

Write-Host "Add drivers to Install image"
Mount-WindowsImage -ImagePath $InstallWim -Index 1 -Path $WimMount | Out-Null
Add-WindowsDriver -Path $WimMount -Driver $DriversFolder -Recurse | Out-Null

Write-Host "Enable SMB1 Client"
Enable-WindowsOptionalFeature -Path $WimMount -FeatureName "smb1protocol-client" -All | Out-Null

Dismount-WindowsImage -Path $WimMount -Save | Out-Null
Clear-WindowsCorruptMountPoint | Out-Null

# remove mount point
Remove-Item $WimMount | Out-Null

$TargetDirectory = (Join-Path -Path $ScriptFolder -ChildPath "ELITEPAD900")
Copy-Files -SourceFolder $WorkDirectory -DestinationFolder $TargetDirectory -FileName "*.*"

$ProgressPreference = "Continue"
