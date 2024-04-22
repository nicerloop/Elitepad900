#Requires -Version 5.1
$ErrorActionPreference = "Stop"

Set-StrictMode -Version 3

# Prevent issues with CLI non-interactive execution
# https://github.com/PowerShell/Microsoft.PowerShell.Archive/issues/77#issuecomment-601947496
Write-Host "Disable progress bars for CLI non-interactive execution"
$ProgressPreference = "SilentlyContinue"
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

function Get-HpSoftPaq {
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
    Get-File -File $File -DestinationFolder $DestinationFolder -BackupFolder $BackupFolder -Url $Url -Description "HP SoftPaq $Number $Description"
}

function Expand-HpSoftPaq {
    param (
        [Parameter(Mandatory)] $Number,
        [Parameter(Mandatory)] $Description,
        [Parameter(Mandatory)] $SourceFolder,
        [Parameter(Mandatory)] $DestinationFolder
    )
    $File = "sp$Number.exe"
    $FilePath = (Join-Path -Path $SourceFolder -ChildPath $File)
    $Folder = (Join-Path -Path $DestinationFolder -ChildPath "sp$Number")
    Write-Host "Expand HP SoftPaq $Number $Description to $Folder"
    Start-Process -FilePath $FilePath -ArgumentList "/s /e /f `"$Folder`"" -Wait
}

function Expand-DellZipPack {
    param (
        $ZipPackFileBaseName,
        $ZipPackDescription,
        $DownloadsFolder,
        $DriversFolder
    )
    $ZipPackFolder = (Join-Path -Path $DriversFolder -ChildPath $ZipPackFileBaseName)
    Write-Host "Unpack $ZipPackDescription to $ZipPackFolder"
    $ZipPackFile = (Join-Path -Path $DownloadsFolder -ChildPath "$ZipPackFileBaseName.exe")
    $ZipPackArchive = (Join-Path -Path $DownloadsFolder -ChildPath "$ZipPackFileBaseName.zip")
    Copy-Item -Path $ZipPackFile -Destination $ZipPackArchive
    New-Item -ItemType Directory -Path $ZipPackFolder -Force | Out-Null
    Remove-Item -Recurse -Force -Path $ZipPackFolder
    Expand-Archive -Path $ZipPackArchive -DestinationPath $ZipPackFolder
    Remove-Item -Path $ZipPackArchive
}

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
        Expand-Archive -Path $InnoExtractArchive -DestinationPath $InnoExtractFolder
    }
    Write-Host "Expand $Path to $DestinationPath"
    Start-Process -FilePath $InnoExtractPath -ArgumentList "--extract --output-dir `"$DestinationPath`" --silent `"$Path`"" -Wait
}

function Get-WindowsIsoUrl {
    param (
        [Parameter(Mandatory)] $DownloadsFolder,
        [Parameter(Mandatory)] $BackupFolder,
        [Parameter(Mandatory)] $ImageVersion,
        [Parameter(Mandatory)] $ImageEdition,
        [Parameter(Mandatory)] $ImageLanguage
    )
    # Windows ISO Downloader
    # https://github.com/pbatard/Fido
    $FidoFileName = "Fido.ps1"
    Get-File -FileName $FidoFileName -DestinationFolder $DownloadsFolder -BackupFolder $BackupFolder -Url "https://raw.githubusercontent.com/pbatard/Fido/master/Fido.ps1" -Description "Windows ISO downloader"
    $FidoFile = (Join-Path -Path $DownloadsFolder -ChildPath $FidoFileName)
    # StrictMode makes Fido fail
    Set-StrictMode -Off
    (& $FidoFile -Win $ImageVersion -Ed $ImageEdition -Lang $ImageLanguage -Arch x86 -GetUrl)
}

# Download Windows x32 ISO
# https://www.microsoft.com/en-us/software-download/windows8ISO
# https://www.microsoft.com/en-us/software-download/windows10ISO
function Get-WindowsIso {
    param (
        [Parameter(Mandatory)] $DownloadsFolder,
        [Parameter(Mandatory)] $BackupFolder,
        [Parameter(Mandatory)] $ImageVersion,
        [Parameter(Mandatory)] $ImageLanguage
    )
    if ($ImageVersion -eq "10") {
        $ImageEdition = "Pro"
        $ImageRelease = "22H2"
        $ImageSuffix = "v1"
        $ImageFileName = "Win${ImageVersion}_${ImageRelease}_${ImageLanguage}_x32${ImageSuffix}.iso"
    } elseif ($ImageVersion -eq "8" -or $ImageVersion -eq "8.1") {
        $ImageVersion = "8.1"
        $ImageEdition = "Standard"
        $ImageFileName = "Win${ImageVersion}_${ImageLanguage}_x32.iso"
    } else {
        throw "Unsupported Windows version $ImageVersion"
    }
    $ImageDescription = "Windows $ImageVersion $ImageEdition x86 ISO with $ImageLanguage language"
    $ImagePath = (Join-Path -Path $DownloadsFolder -ChildPath $ImageFileName)
    Get-File -FileName $ImageFileName -DestinationFolder $DownloadsFolder -BackupFolder $BackupFolder -Description $ImageDescription -ScriptBlock {
        Write-Host "Get $ImageDescription download URL"
        Get-WindowsIsoUrl -DownloadsFolder $DownloadsFolder -BackupFolder $BackupFolder -ImageVersion $ImageVersion -ImageEdition $ImageEdition -ImageLanguage $ImageLanguage
    }
    return $ImagePath, $ImageDescription
}

function Add-Drivers {
    param (
        [Parameter(Mandatory)] $ImageDescription,
        [Parameter(Mandatory)] $ImagePath,
        [Parameter(Mandatory)] $ImageIndex,
        [Parameter(Mandatory)] $MountPath,
        [Parameter(Mandatory)] $DriversFolder
    )
    Write-Host "Prepare $ImageDescription image"
    Mount-WindowsImage -ImagePath $ImagePath -Index $ImageIndex -Path $MountPath | Out-Null
    Write-Host "Add drivers"
    Add-WindowsDriver -Path $WimMount -Driver $DriversFolder -Recurse | Out-Null
    Dismount-WindowsImage -Path $MountPath -Save | Out-Null
    Clear-WindowsCorruptMountPoint | Out-Null
}

Write-Host "Define work folders location"
$ScriptFolder = $PSScriptRoot
$DownloadsFolder = (Join-Path -Path $HOME -ChildPath "Downloads")
$WorkFolder = (Join-Path -Path $HOME -ChildPath "ELITEPAD900")
$DriversFolder = (Join-Path -Path $WorkFolder -ChildPath "drivers")
$BackupFolder = (Join-Path -Path $ScriptFolder -ChildPath "Downloads")

Write-Host "Download HP Elitepad900 drivers for Windows 8.1"
# https://support.hp.com/us-en/drivers/selfservice/hp-elitepad-900-g1-tablet/5298028
$HpSoftPaqs = @(
    (71504, "HP Elitepad 900 System BIOS/Firmware and Driver Update 1.0.2.3 Rev.A Jun 3, 2015"),
    (65877, "Qualcomm Atheros AR600x 802.11a/b/g/n Wireless LAN Driver 3.7 C Apr 18, 2014"),
    (65376, "Qualcomm Atheros AR3002 Bluetooth 4.0+HS Driver for Microsoft Windows 2.2 Rev.A Feb 12, 2014"),
    (64673, "NXP Semiconductors Near Field Proximity (NFP) Driver 1.4.7.2 Dec 2, 2013"),
    (64682, "Broadcom GPS Driver 19.17 Dec 6, 2013")
)
For ($Index = 0; $Index -lt $HpSoftPaqs.Length; $Index++) {
    $Number, $Description = $HpSoftPaqs[$Index]
    Get-HpSoftPaq -Number $Number -Description $Description -DestinationFolder $DownloadsFolder -BackupFolder $BackupFolder
}
Write-Host "Expand HP Elitepad900 drivers for Windows 8.1"
For ($Index = 0; $Index -lt $HpSoftPaqs.Length; $Index++) {
    $Number, $Description = $HpSoftPaqs[$Index]
    Expand-HpSoftPaq -Number $Number -Description $Description -SourceFolder $DownloadsFolder -DestinationFolder $DriversFolder
}

# sp71504 contains video drivers problematic since Windows 10 1903
$AlterSP71504 = $true;
if ($AlterSP71504 -And (Test-Path -Path (Join-Path -Path $DriversFolder -ChildPath "sp71504"))) {
    Write-Host "Remove Video drivers from HP SoftPaq 71504"
    Remove-Item -Path (Join-Path -Path $DriversFolder -ChildPath "sp71504/Package/Drivers/Graphics") -Recurse -Force
}

# sp64673 must be expanded to access the drivers
if (Test-Path -Path (Join-Path -Path $DriversFolder -ChildPath "sp64673")) {
    Write-Host "Expand HP SoftPaq 64673 installer"
    $HpSoftPaq64673Folder = (Join-Path -Path $DriversFolder -ChildPath "sp64673")
    $HpSoftPaq64673 = (Join-Path -Path $HpSoftPaq64673Folder -ChildPath "*.exe" -Resolve)
    Expand-InnoInstaller -Path $HpSoftPaq64673 -DestinationPath $HpSoftPaq64673Folder -DownloadsFolder $DownloadsFolder -BackupFolder $BackupFolder -WorkFolder $WorkFolder
}

# sp64682 does not run on Windows 10
# Download Dell Broadcom GPS Driver for Windows 8.1
# https://www.dell.com/support/home/en-us/drivers/driversdetails?driverid=p0p15
$ZipPackFileBaseName = "GPS_BCM4751_W8_A02-P0P15_ZPE"
$ZipPackUrl = "https://dl.dell.com/FOLDER00998748M/3/GPS_BCM4751_W8_A02-P0P15_ZPE.exe"
$ZipPackDescription = "Dell ZipPack P0P15 BCM47511 Standalone GPS Solution 19.14.6362.4, A02 30 Nov 2012"
Get-File -File "$ZipPackFileBaseName.exe" -DestinationFolder $DownloadsFolder -BackupFolder $BackupFolder -Url $ZipPackUrl -Description $ZipPackDescription
# Expand-DellZipPack -ZipPackFileBaseName $ZipPackFileBaseName -ZipPackDescription $ZipPackDescription -DownloadsFolder $DownloadsFolder -DriversFolder $DriversFolder

Write-Host "Available drivers:"
$DriverInfFiles = (Get-ChildItem -Recurse -Filter "*.inf" -Path $DriversFolder)
$DriverInfFiles | Select-Object -ExpandProperty FullName
$DriverInfFilesCount = ($DriverInfFiles).count
Write-Host "Available drivers count: $DriverInfFilesCount"

$WindowsInstaller = $true
if ($WindowsInstaller) {
    $ImageVersion = "10"
    $ImageLanguage = "French"
$ImagePath, $ImageDescription = Get-WindowsIso -DownloadsFolder $DownloadsFolder -BackupFolder $BackupFolder -ImageVersion $ImageVersion -ImageLanguage $ImageLanguage

Write-Host "Get install files from ISO"

Write-Host "Mount $ImageDescription"
$DriveLetter = (Mount-DiskImage -ImagePath $ImagePath -PassThru | Get-Volume).DriveLetter
Write-Host "Mounted $ImageDescription as $DriveLetter`:"
Copy-Files -Path "$DriveLetter`:" -DestinationPath $WorkFolder -FileNamePattern "*.*"
Write-Host "Unmount $ImageDescription"
Dismount-DiskImage -ImagePath $ImagePath | Out-Null
}

$SlipstreamDrivers = $true
If ($SlipstreamDrivers -and $WindowsInstaller) {
    $SourcesFolder = (Join-Path $WorkFolder "sources")
    $BootWim = (Join-Path $SourcesFolder "boot.wim")
    $InstallWim = (Join-Path $SourcesFolder "install.wim")
    $InstallWimSlim = (Join-Path $SourcesFolder "install.slim.wim")

    Write-Host "Find image index for Windows $ImageVersion Professional variant"
    $ImageName = Get-WindowsImage -ImagePath $InstallWim | Select-Object -ExpandProperty ImageName | Select-String -Pattern 'Pro' | Select-Object -ExpandProperty Line -First 1
    Write-Host "Windows $ImageVersion Professional found with name $ImageName"
    $ImageIndex = Get-WindowsImage -ImagePath $InstallWim -Name $ImageName | Select-Object -ExpandProperty ImageIndex
    Write-Host "Windows $ImageVersion Professional found with index $ImageIndex"

    Write-Host "Create mount point"
    $WimMount = "C:\wim_mount"
    New-Item -Path $WimMount -ItemType directory -Force | Out-Null

    Add-Drivers -ImageDescription "WinPE" -ImagePath $BootWim -ImageIndex 1 -MountPath $WimMount -DriversFolder $DriversFolder
    Add-Drivers -ImageDescription "Setup" -ImagePath $BootWim -ImageIndex 2 -MountPath $WimMount -DriversFolder $DriversFolder
    Add-Drivers -ImageDescription "Install" -ImagePath $InstallWim -ImageIndex $ImageIndex -MountPath $WimMount -DriversFolder $DriversFolder

    Write-Host "Keep only Windows 10 Pro variant"
    Export-WindowsImage -SourceImagePath $InstallWim -SourceIndex $ImageIndex -DestinationImagePath $InstallWimSlim -CompressionType Max | Out-Null
    Remove-Item -Path $InstallWim
    Rename-Item -Path $InstallWimSlim -NewName $InstallWim

    Write-Host "Remove mount point"
    Remove-Item $WimMount | Out-Null

} else {
    Write-Host "Copy drivers installers"
    Copy-Files -Path $DownloadsFolder -DestinationPath $DriversFolder -FileNamePattern "*.exe"
}

Write-Host "Add drivers add and export scripts"
$DriversScriptsFolder = (Join-Path -Path $ScriptFolder -ChildPath "drivers-scripts")
Copy-Files -Path $DriversScriptsFolder -DestinationPath $DriversFolder -FileNamePattern "drivers-*.bat"

Write-Host "Copy installation files to origin location"
$TargetDirectory = (Join-Path -Path $ScriptFolder -ChildPath "ELITEPAD900")
if (Test-Path -Path $TargetDirectory) {
    Remove-Item -Path $TargetDirectory -Recurse -Force | Out-Null
}
Copy-Files -Path $WorkFolder -DestinationPath $TargetDirectory -FileNamePattern "*.*"
