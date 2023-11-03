#Requires -Version 5.1
$ErrorActionPreference = "Stop"
# Set-StrictMode -Version 3
# Set-PSDebug -Trace 2
$ProgressPreference = "SilentlyContinue"
# $GlobalProgressPreference = $global:ProgressPreference
$global:ProgressPreference = "SilentlyContinue"

function Copy-Files {
    param (
        [Parameter(Mandatory)] $Path,
        [Parameter(Mandatory)] $DestinationPath,
        [Parameter(Mandatory)] $FileNamePattern
    )
    Write-Host "Copy $FileNamePattern from $Path to $DestinationPath"
    Robocopy.exe $Path $DestinationPath $FileNamePattern /E /Z /A-:R /IM /NDL /NFL /NJH /NJS | Out-Null
}

function Get-Url {
    param (
        [Parameter(Mandatory)] $Url,
        [Parameter(Mandatory)] $DestinationPath
    )
    $UserAgent = "Mozilla/5.0 (X11; Linux i586; rv:100.0) Gecko/20100101 Firefox/100.0"
    curl.exe --output-dir $DestinationPath -O -C - -A $UserAgent -L -# $Url
}

function Get-File {
    param (
        [Parameter(Mandatory)] $FileName,
        [Parameter(Mandatory)] $DestinationFolder,
        [Parameter(Mandatory)] $BackupFolder,
        $Url,
        $ScriptBlock,
        [Parameter(Mandatory)] $Description
    )
    $DestinationFile = (Join-Path -Path $DestinationFolder -ChildPath $FileName)
    if (Test-Path -Path $DestinationFile -PathType Leaf) {
        Write-Host "Found $Description as $DestinationFile"
    }
    else {
        $BackupFile = (Join-Path -Path $BackupFolder -ChildPath $FileName)
        if (Test-Path -Path $BackupFile -PathType Leaf) {
            Write-Host "Found $Description as $BackupFile"
            Copy-Files -Path $BackupFolder -DestinationPath $DestinationFolder -FileNamePattern $FileName
        }
        else {
            if ($ScriptBlock) {
                Write-Host "Build URL"
                $Url = (& $ScriptBlock)
            }
            Write-Host "Download $Description from $Url"
            Get-Url -DestinationPath $DestinationFolder -Url $Url
            Copy-Files -Path $DestinationFolder -DestinationPath $BackupFolder -FileNamePattern $FileName
        }
    }
}

function Get-HpSoftPack {
    param (
        [Parameter(Mandatory)] $Number,
        [Parameter(Mandatory)] $Description,
        [Parameter(Mandatory)] $DestinationFolder,
        [Parameter(Mandatory)] $BackupFolder
    )
    $File = "sp$Number.exe"
    $M = [int][Math]::Floor(($Number - 1) / 1000)
    $D = [int][Math]::Floor((($Number - 1) % 1000) / 500)
    $Min = $M * 1000 + $D * 500 + 1
    $Max = $M * 1000 + ($D + 1) * 500
    $Url = "https://ftp.hp.com/pub/softpaq/sp$Min-$Max/$File"
    Get-File -File $File -DestinationFolder $DestinationFolder -BackupFolder $BackupFolder -Url $Url -Description "HP SoftPack $Number $Description"
}

function Expand-HpSoftPack {
    param (
        [Parameter(Mandatory)] $Number,
        [Parameter(Mandatory)] $Description,
        [Parameter(Mandatory)] $SourceFolder,
        [Parameter(Mandatory)] $DestinationFolder
    )
    $File = "sp$Number.exe"
    $FilePath = (Join-Path -Path $SourceFolder -ChildPath $File)
    $Folder = (Join-Path -Path $DestinationFolder -ChildPath "sp$Number")
    Write-Host "Expand HP SoftPack $Number $Description to $Folder"
    Start-Process -FilePath $FilePath -ArgumentList "/s /e /f `"$Folder`"" -Wait
}

function Expand-Archive-Silent {
    param (
        [Parameter(Mandatory)] $Path,
        [Parameter(Mandatory)] $DestinationPath
    )
    # Prevent issues with CLI non-interactive execution
    # https://github.com/PowerShell/Microsoft.PowerShell.Archive/issues/77#issuecomment-601947496
    # $GlobalProgressPreference = $global:ProgressPreference
    # $global:ProgressPreference = "SilentlyContinue"
    Expand-Archive -Path $Path -DestinationPath $DestinationPath
    # $global:ProgressPreference = $GlobalProgressPreference
}

# function Expand-DellZipPack {
#     param (
#         $ZipPackFileBaseName,
#         $ZipPackDescription,
#         $DownloadsFolder,
#         $DriversFolder
#     )
#     $ZipPackFolder = (Join-Path -Path $DriversFolder -ChildPath $ZipPackFileBaseName)
#     Write-Host "Unpack $ZipPackDescription to $ZipPackFolder"
#     $ZipPackFile = (Join-Path -Path $DownloadsFolder -ChildPath "$ZipPackFileBaseName.exe")
#     $ZipPackArchive = (Join-Path -Path $DownloadsFolder -ChildPath "$ZipPackFileBaseName.zip")
#     Copy-Item -Path $ZipPackFile -Destination $ZipPackArchive
#     New-Item -ItemType Directory -Path $ZipPackFolder -Force | Out-Null
#     Remove-Item -Recurse -Force -Path $ZipPackFolder
#     # # Prevent issues with CLI non-interactive execution
#     # # https://github.com/PowerShell/Microsoft.PowerShell.Archive/issues/77#issuecomment-601947496
#     # $GlobalProgressPreference = $global:ProgressPreference
#     # $global:ProgressPreference = "SilentlyContinue"
#     # Expand-Archive -Path $ZipPackArchive -DestinationPath $ZipPackFolder
#     # $global:ProgressPreference = $GlobalProgressPreference
#     Expand-Archive-Silent -Path $ZipPackArchive -DestinationPath $ZipPackFolder
#     Remove-Item -Path $ZipPackArchive
# }

function Expand-InnoInstaller {
    param (
        [Parameter(Mandatory)] $Path,
        [Parameter(Mandatory)] $DestinationPath,
        [Parameter(Mandatory)] $DownloadsFolder,
        [Parameter(Mandatory)] $BackupFolder,
        [Parameter(Mandatory)] $WorkFolder
    )
    $InnoExtractVersion = "1.9"
    $InnoExtractBaseName = "innoextract-$InnoExtractVersion-windows"
    $InnoExtractFileName = "$InnoExtractBaseName.zip"
    $InnoExtractUrl = "https://github.com/dscharrer/innoextract/releases/download/$InnoExtractVersion/$InnoExtractFileName"
    $InnoExtractDescription = "Inno Setup installer unpacker"
    Get-File -File $InnoExtractFileName -DestinationFolder $DownloadsFolder -BackupFolder $BackupFolder -Url $InnoExtractUrl -Description $InnoExtractDescription
    $InnoExtractFolder = (Join-Path -Path $WorkFolder -ChildPath $InnoExtractBaseName)
    $InnoExtractPath = (Join-Path -Path $InnoExtractFolder -ChildPath "innoextract.exe")
    if (-Not (Test-Path -Path $InnoExtractPath -PathType Leaf)) {
        $InnoExtractArchive = (Join-Path -Path $DownloadsFolder -ChildPath $InnoExtractFileName)
        Expand-Archive-Silent -Path $InnoExtractArchive -DestinationPath $InnoExtractFolder
    }
    Write-Host "Expand $Path to $DestinationPath"
    Start-Process -FilePath $InnoExtractPath -ArgumentList "--extract --output-dir `"$DestinationPath`" --silent `"$Path`"" -Wait
}

Write-Host "Disable progress bars for CLI non-interactive execution"
$ProgressPreference = "SilentlyContinue"

Write-Host "Define work folders location"
$ScriptFolder = $PSScriptRoot
$DownloadsFolder = (Join-Path -Path $HOME -ChildPath "Downloads")
$WorkFolder = (Join-Path -Path $HOME -ChildPath "ELITEPAD900")
$DriversFolder = (Join-Path -Path $WorkFolder -ChildPath "drivers")
$BackupFolder = (Join-Path -Path $ScriptFolder -ChildPath "Downloads")

Write-Host "Download HP Elitepad900 drivers for Windows 8.1"
# https://support.hp.com/us-en/drivers/selfservice/hp-elitepad-900-g1-tablet/5298028
$HpSoftPacks = @(
    (71504, "HP Elitepad 900 System BIOS/Firmware and Driver Update 1.0.2.3 Rev.A Jun 3, 2015"),
    (65877, "Qualcomm Atheros AR600x 802.11a/b/g/n Wireless LAN Driver 3.7 C Apr 18, 2014"),
    (65376, "Qualcomm Atheros AR3002 Bluetooth 4.0+HS Driver for Microsoft Windows 2.2 Rev.A Feb 12, 2014"),
    (64673, "NXP Semiconductors Near Field Proximity (NFP) Driver 1.4.7.2 Dec 2, 2013"),
    (64682, "Broadcom GPS Driver 19.17 Dec 6, 2013")
)
For ($Index = 0; $Index -lt $HpSoftPacks.Length; $Index++) {
    $Number, $Description = $HpSoftPacks[$Index]
    Get-HpSoftPack -Number $Number -Description $Description -DestinationFolder $DownloadsFolder -BackupFolder $BackupFolder
}
Write-Host "Expand HP Elitepad900 drivers for Windows 8.1"
For ($Index = 0; $Index -lt $HpSoftPacks.Length; $Index++) {
    $Number, $Description = $HpSoftPacks[$Index]
    Expand-HpSoftPack -Number $Number -Description $Description -SourceFolder $DownloadsFolder -DestinationFolder $DriversFolder
}

Write-Host "Remove DASL/Win8 from HP SoftPack 71504"
Remove-Item -Path (Join-Path -Path $DriversFolder -ChildPath "sp71504/DASL/DASL/Win8") -Recurse

Write-Host "Expand HP SoftPack 64673 installer"
# $HpSoftPack64673 = (Join-Path -Path $DriversFolder -ChildPath "sp64673/NXP_NFC_Driver_Package_Setup.exe")
$HpSoftPack64673Folder = (Join-Path -Path $DriversFolder -ChildPath "sp64673")
$HpSoftPack64673 = (Join-Path -Path $HpSoftPack64673Folder -ChildPath "*.exe" -Resolve)
Expand-InnoInstaller -Path $HpSoftPack64673 -DestinationPath $HpSoftPack64673Folder -DownloadsFolder $DownloadsFolder -BackupFolder $BackupFolder -WorkFolder $WorkFolder

# # Download Dell Broadcom GPS Driver for Windows 8.1
# # https://www.dell.com/support/home/en-us/drivers/driversdetails?driverid=p0p15
# # HP SoftPack provided installer does not run on Windows 10
# # Dell provided one runs on Windows 10
# # and can be expanded and streamlined from the command line

# $ZipPackFileBaseName = "GPS_BCM4751_W8_A02-P0P15_ZPE"
# # $ZipPackFileName = "$ZipPackFileBaseName.exe"
# $ZipPackUrl = "https://dl.dell.com/FOLDER00998748M/3/GPS_BCM4751_W8_A02-P0P15_ZPE.exe"
# $ZipPackDescription = "Dell ZipPack P0P15 BCM47511 Standalone GPS Solution 19.14.6362.4, A02 30 Nov 2012"
# # Get-File -File $ZipPackFileName -DownloadFolder $DownloadsFolder -BackupFolder $BackupFolder -Url $ZipPackUrl -Description $ZipPackDescription
# Get-File -File "$ZipPackFileBaseName.exe" -DownloadFolder $DownloadsFolder -BackupFolder $BackupFolder -Url $ZipPackUrl -Description $ZipPackDescription
# # $ZipPackFolder = (Join-Path -Path $DriversFolder -ChildPath $ZipPackFileBaseName)
# # Write-Host "Unpack $ZipPackDescription to $ZipPackFolder"
# # $ZipPackFile = (Join-Path -Path $DownloadsFolder -ChildPath $ZipPackFileName)
# # $ZipPackArchive = "$ZipPackFileBaseName.zip"
# # Copy-Item -Path $ZipPackFile -Destination $ZipPackArchive
# # New-Item -ItemType Directory -Path $ZipPackFolder -Force | Out-Null
# # Remove-Item -Recurse -Force -Path $ZipPackFolder
# # # Prevent issues with CLI non-interactive execution
# # # https://github.com/PowerShell/Microsoft.PowerShell.Archive/issues/77#issuecomment-601947496
# # $global:ProgressPreference = "SilentlyContinue"
# # Expand-Archive -Path $ZipPackArchive -DestinationPath $ZipPackFolder
# # $global:ProgressPreference = "Continue"
# # Remove-Item -Path $ZipPackArchive
# Expand-DellZipPack -ZipPackFileBaseName $ZipPackFileBaseName -ZipPackDescription $ZipPackDescription -DownloadsFolder $DownloadsFolder -DriversFolder $DriversFolder

Write-Host "Available drivers:"
$DriverInfFiles = (Get-ChildItem -Recurse -Filter "*.inf" -Path $DriversFolder)
$DriverInfFiles | Select-Object -ExpandProperty FullName
$DriverInfFilesCount = ($DriverInfFiles).count
Write-Host "Available drivers count: $DriverInfFilesCount"

# Download Windows 10 Pro x32 ISO
# https://www.microsoft.com/en-us/software-download/windows10ISO

$ImageRelease = "1703"
$ImageSuffix = ""
# $ImageRelease = "22H2"
# $ImageSuffix = "v1"
$ImageLanguage = "English"
# $ImageLanguage = "French"
$ImageFileName = "Win10_${ImageRelease}_${ImageLanguage}_x32${ImageSuffix}.iso"
$ImagePath = (Join-Path -Path $DownloadsFolder -ChildPath $ImageFileName)
$ImageDescription = "Windows 10 Pro x86 ISO release $ImageRelease with $ImageLanguage language"

# if (Test-Path $ImagePath -PathType leaf) {
#     Write-Host "Found $ImageDescription as $ImagePath"
# }
# else {
#     Write-Host "Get $ImageDescription download URL"
#     # Windows ISO Downloader
#     # https://github.com/pbatard/Fido
#     $FidoFileName = "Fido.ps1"
#     Get-File -FileName $FidoFileName -DestinationFolder $DownloadsFolder -BackupFolder $BackupFolder -Url "https://raw.githubusercontent.com/pbatard/Fido/master/Fido.ps1" -Description "Windows ISO downloader"
#     $FidoFile = (Join-Path -Path $DownloadsFolder -ChildPath $FidoFileName)
#     $ImageUrl = (& $FidoFile -Win 10 -Rel $ImageRelease -Ed Pro -Lang $ImageLanguage -Arch x86 -GetUrl)
#     Get-File -FileName $ImageFileName -DestinationFolder $DownloadsFolder -BackupFolder $BackupFolder -Url $ImageUrl -Description $ImageDescription
# }

Get-File -FileName $ImageFileName -DestinationFolder $DownloadsFolder -BackupFolder $BackupFolder -Description $ImageDescription -ScriptBlock {
    Write-Host "Get $ImageDescription download URL"
    # Windows ISO Downloader
    # https://github.com/pbatard/Fido
    $FidoFileName = "Fido.ps1"
    Get-File -FileName $FidoFileName -DestinationFolder $DownloadsFolder -BackupFolder $BackupFolder -Url "https://raw.githubusercontent.com/pbatard/Fido/master/Fido.ps1" -Description "Windows ISO downloader"
    $FidoFile = (Join-Path -Path $DownloadsFolder -ChildPath $FidoFileName)
    (& $FidoFile -Win 10 -Rel $ImageRelease -Ed Pro -Lang $ImageLanguage -Arch x86 -GetUrl)
}

Write-Host "Get install files from ISO"

Write-Host "Mount $ImageDescription"
$DriveLetter = (Mount-DiskImage -ImagePath $ImagePath -PassThru | Get-Volume).DriveLetter
Write-Host "Mounted $ImageDescription as $DriveLetter`:"
Copy-Files -Path "$DriveLetter`:" -DestinationPath $WorkFolder -FileNamePattern "*.*"
Write-Host "Unmount $ImageDescription"
Dismount-DiskImage -ImagePath $ImagePath | Out-Null

$SourcesFolder = (Join-Path $WorkFolder "sources")
$BootWim = (Join-Path $SourcesFolder "boot.wim")
$InstallWim = (Join-Path $SourcesFolder "install.wim")
$InstallWimSlim = (Join-Path $SourcesFolder "install.slim.wim")

Write-Host "Find image index for Windows 10 Pro variant"
$ImageName = Get-WindowsImage -ImagePath $InstallWim | Select-Object -ExpandProperty ImageName | Select-String -Pattern 'Pro' | Select-Object -ExpandProperty Line -First 1
Write-Host "Windows 10 Pro found with name $ImageName"
$ImageIndex = Get-WindowsImage -ImagePath $InstallWim -Name $ImageName | Select-Object -ExpandProperty ImageIndex
Write-Host "Windows 10 Pro found with index $ImageIndex"

Write-Host "Create mount point"
$WimMount = "C:\wim_mount"
New-Item -Path $WimMount -ItemType directory -Force | Out-Null

Write-Host "Prepare WinPE image"
Mount-WindowsImage -ImagePath $BootWim -Index 1 -Path $WimMount | Out-Null
Write-Host "Add drivers"
Add-WindowsDriver -Path $WimMount -Driver $DriversFolder -Recurse | Out-Null
Dismount-WindowsImage -Path $WimMount -Save | Out-Null
Clear-WindowsCorruptMountPoint | Out-Null

Write-Host "Prepare Setup image"
Mount-WindowsImage -ImagePath $BootWim -Index 2 -Path $WimMount | Out-Null
Write-Host "Add drivers"
Add-WindowsDriver -Path $WimMount -Driver $DriversFolder -Recurse | Out-Null
Dismount-WindowsImage -Path $WimMount -Save | Out-Null
Clear-WindowsCorruptMountPoint | Out-Null

Write-Host "Prepare Install image"
Mount-WindowsImage -ImagePath $InstallWim -Index $ImageIndex -Path $WimMount | Out-Null
Write-Host "Add drivers"
Add-WindowsDriver -Path $WimMount -Driver $DriversFolder -Recurse | Out-Null
# Write-Host "Enable SMB1 Client"
# Enable-WindowsOptionalFeature -Path $WimMount -FeatureName "smb1protocol-client" -All | Out-Null
Dismount-WindowsImage -Path $WimMount -Save | Out-Null
Clear-WindowsCorruptMountPoint | Out-Null

Write-Host "Keep only Windows 10 Pro variant"
Export-WindowsImage -SourceImagePath $InstallWim -SourceIndex $ImageIndex -DestinationImagePath $InstallWimSlim -CompressionType Max | Out-Null
Remove-Item -Path $InstallWim
Rename-Item -Path $InstallWimSlim -NewName $InstallWim

Write-Host "Remove mount point"
Remove-Item $WimMount | Out-Null

# Write-Host "Add unattended install configuration"
# Copy-Files -Path $ScriptFolder -DestinationPath $WorkFolder -FileNamePattern "autounattended.xml"

$TargetDirectory = (Join-Path -Path $ScriptFolder -ChildPath "ELITEPAD900")
Copy-Files -Path $WorkFolder -DestinationPath $TargetDirectory -FileNamePattern "*.*"

# $ProgressPreference = "Continue"
