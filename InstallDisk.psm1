# This function is enclosed in a module so that $deviceMethods isn't leaked
New-Module -ScriptBlock {
    $deviceMethods = Add-Type -Name NativeMethods -Namespace Kernel32 -PassThru -MemberDefinition @'
[DllImport("Kernel32.dll", EntryPoint = "QueryDosDeviceA", CharSet = CharSet.Ansi, SetLastError=true)]
public static extern int QueryDosDevice(string lpDeviceName, System.Text.StringBuilder lpTargetPath, int ucchMax);
'@
    function ConvertTo-DosDevice {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [String]
            $DeviceName
        )
        $sb = New-Object System.Text.StringBuilder(30)
        if (-not $deviceMethods::QueryDosDevice($DeviceName, $sb, 30)) {
            throw "no device mapping for $Devicename"
        }
        Write-Output $sb.ToString()
    }
} | Out-Null

# For this one, a simple HashTable will do the trick
# See links from https://msdn.microsoft.com/en-us/library/windows/desktop/aa362687%28v=vs.85%29.aspx
$elementTypes = @{
    # BootMgr
    "DisplayOrder" = 0x24000001
    "Timeout" = 0x25000004
    "DisplayBootMenu" = 0x26000020
    # DeviceObject
    # Library
    "ApplicationDevice" = 0x11000001
    "Description" = 0x12000004
    # MemDiag
    # OSLoader
    "OSDevice" = 0x21000001
}

# BCD functions

function Open-BcdStore {
    [CmdletBinding()]
    param (
        [String]$Path
    )
    $cimStore = Invoke-CimMethod -Namespace ROOT\wmi -ClassName BcdStore -MethodName OpenStore -Arguments @{
        "File" = $path
    }
    if (-not $cimStore.ReturnValue) {
        throw "No BCD store at $path"
    }
    Write-Output $cimStore.Store
}

function Open-BcdBootManager {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $BcdStore
    )
    $cimBootMgr = Invoke-CimMethod -InputObject $BcdStore -MethodName EnumerateObjects -Arguments @{
        "Type" = 0x10100002
    }
    if (-not $cimBootMgr.ReturnValue) {
        throw "No boot manager in BCD store $($BcdStore.FilePath)"
    }
    if (-not $cimBootMgr.Objects.Count) {
        throw "No boot manager in BCD store $($BcdStore.FilePath)"
    }
    Write-Output $cimBootMgr.Objects[0]
}

function Set-BcdBootManagerTimeout {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $BcdBootManager,

        [Int]
        $Timeout = 10
    )
    $ret = Invoke-CimMethod -InputObject $BcdBootManager -MethodName SetIntegerElement -Arguments @{
        "Type" = $elementTypes["Timeout"]
        "Integer" = $Timeout
    }
    if (-not $ret.ReturnValue) {
        throw "Unable to set boot timeout on BCD store $($BcdBootManager.StoreFilePath)"
    }
}

function Set-BcdBootManagerMenu {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $BcdBootManager,

        [Parameter(Mandatory)]
        [Boolean]
        $Enabled
    )
    $ret = Invoke-CimMethod -InputObject $BcdBootManager -MethodName SetBooleanElement -Arguments @{
        "Type" = $elementTypes["DisplayBootMenu"]
        "Boolean" = $Enabled
    }
    if (-not $ret.ReturnValue) {
        throw "Unable to set display boot menu on BCD store $($BcdBootManager.StoreFilePath)"
    }
}

function Add-BcdBootManagerMenuEntry {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $BcdBootManager,

        [Parameter(Mandatory)]
        [GUID]
        $ID
    )
    $displayOrder = Get-BcdElement -BcdObject $BcdBootManager -TypeName "DisplayOrder"
    $ret = Invoke-CimMethod -InputObject $BcdBootManager -MethodName SetObjectListElement -Arguments @{
        "Type" = $elementTypes["DisplayOrder"]
        "Ids" = $displayOrder.Ids + $ID.ToString("B")
    }
    if (-not $ret.ReturnValue) {
        throw "Unable to set display order"
    }
}

function Open-BcdObject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $BcdStore,

        [GUID]
        $ID = "{7619dcc9-fafe-11d9-b411-000476eba25f}" # Default bootloader
    )
    $cimBootLdr = Invoke-CimMethod -InputObject $BcdStore -MethodName OpenObject -Arguments @{
        "Id" = $ID.ToString("B")
    }
    if (-not $cimBootLdr.ReturnValue) {
        throw "Unable to open object"
    }
    Write-Output $cimBootLdr.Object
}

function Copy-BcdObject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $BcdStore,

        [Parameter(Mandatory)]
        [GUID]
        $ID
    )
    # I can copy objects only from another store, so create a copy to use as source
    $tempStorePath = [System.IO.Path]::GetTempFileName()
    Copy-Item -Path $BcdStore.FilePath.replace('\??\','') -Destination $tempStorePath | Out-Null
    $object = Invoke-CimMethod -InputObject $BcdStore -MethodName CopyObject -Arguments @{
        "SourceStoreFile" = $tempStorePath
        "SourceId" = $ID.ToString("B")
        "Flags" = 1
    }
    Remove-Item -Path $tempStorePath
    if (-not $object.ReturnValue -or $object.Object.ID.Count -ne 1) {
        throw "Error adding boot entry"
    }
    # Lots of parentheses but PowerShell has a strange sense of precedence.
    Write-Output ([GUID]($object.Object.Id))
}

function Get-BcdElement {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $BcdObject,

        [Parameter(Mandatory)]
        [String]
        $TypeName
    )
    if (-not $elementTypes.ContainsKey($TypeName)) {
        throw "Unknown element type: $TypeName"
    }
    $element = Invoke-CimMethod -InputObject $BcdObject -MethodName GetElement -Arguments @{
            "Type" = $elementTypes[$TypeName]
    }
    if (-not $element.ReturnValue) {
        throw "No device with type $type"
    }
    Write-Output $element.Element
}

function Set-BcdOSLoaderDevice {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $BcdOSLoader,

        [Parameter(Mandatory)]
        [String]
        $ImagePath
    )
    $parentPath = ConvertTo-DosDevice (Split-Path -Path $ImagePath -Qualifier)
    $path = Split-PAth -Path $ImagePath -NoQualifier
    foreach ($type in "ApplicationDevice","OSDevice") {
        $currentDevice = Get-BcdElement -BcdObject $BcdOSLoader -TypeName $type
        $device = Invoke-CimMethod -InputObject $BcdOSLoader -MethodName SetFileDeviceElement -Arguments @{
            "Type" = $elementTypes[$type]
            "DeviceType" = 4
            "Path" = $path
            "ParentDeviceType" = 2
            "ParentAdditionalOptions" = ""
            "ParentPath" = $parentPath
            "AdditionalOptions" = $currentDevice.Device.AdditionalOptions
        }
    }
}

function Set-BcdOSLoaderDescription {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $BcdOSLoader,

        [Parameter(Mandatory)]
        [String]
        $Description
    )
    $ret = Invoke-CimMethod -InputObject $BcdOSLoader -MethodName SetStringElement -Arguments @{
        "Type" = $elementTypes["Description"]
        "String" = $Description
    }
    if (-not $ret.ReturnValue) {
        throw "Can't set description"
    }
}

# Constructor for install disk handles
function New-InstallDisk {
    [CmdletBinding()]
    param (
         [Parameter(Mandatory)]
         [ValidatePattern('^[A-Z]:?$')]
         [String]
         $BootVolume,

         [Parameter(Mandatory)]
         [ValidatePattern('^[A-Z]:?$')]
         [String]
         $InstallerVolume
    )
    if ($BootVolume -match '^([a-z]):?') {
        $BootVolume = $Matches[1]
    }
    if ($InstallerVolume -match '^([a-z]):?') {
        $InstallerVolume = $Matches[1]
    }
    $new = New-Object -TypeName PSObject
    $new | Add-Member -MemberType NoteProperty -Name BootVolume -Value $BootVolume
    $new | Add-Member -MemberType NoteProperty -Name InstallerVolume -Value $InstallerVolume
    $new | Add-Member -MemberType ScriptProperty -Name BcdStores -Value {
        Write-Output "$($this.BootVolume):\boot\bcd"
        Write-Output "$($this.BootVolume):\efi\microsoft\boot\bcd"
    }
    Write-Output $new
}

function Get-InstallDisk {
    [CmdletBinding()]
    param (
        [Switch]
        $OnlyUSB
    )
    $partitions = Get-Partition
    Get-Disk | ForEach-Object {
        $disk = $_
        $parts = $partitions |
                Where-Object {$_.DiskNumber -eq $disk.Number} |
                Sort-Object -Property PartitionNumber
        if (
            ($_.BusType -eq "USB" -or -not $OnlyUSB) -and
            $_.PartitionStyle -eq "MBR" -and
            ($parts | Measure-Object).Count -ge 2 -and
            $parts[0].DriveLetter -and
            (Get-Volume -DriveLetter $parts[0].DriveLetter).FileSystem -eq "FAT32" -and
            $parts[1].DriveLetter -and
            (Get-Volume -DriveLetter $parts[1].DriveLetter).FileSystem -eq "NTFS"
        ) {
            New-InstallDisk -BootVolume $parts[0].DriveLetter -InstallerVolume $parts[1].DriveLetter
        }
    }
}

function Get-USBDisk {
    $disks = Get-Disk | Where-Object {$_.BusType -eq "USB"}
    if (-not $disks) {
        throw "No USB disk found"
    } elseif (($disks | Measure-Object).Count -ne 1) {
        throw "More than 1 USB disk found"
    }
    Write-Output $disks
}

function Initialize-InstallDisk {
    [CmdletBinding(ConfirmImpact = 'High',SupportsShouldProcess)]
    param (
        [String]
        $CDSource = "D:\",

        [Parameter(ValueFromPipeline,Mandatory)]
        [CimInstance]$Disk,

        [Switch]
        $PassThru
    )
    if (-not $PSCmdlet.ShouldProcess("Should initialize disk $($Disk.FriendlyName)",
            "Are you sure you want to initialize the disk?`nThis will erase all data on $($Disk.FriendlyName)",
            "Confirm?")) {
        return
    }
    Write-Verbose "Creating partitions"
    Clear-Disk -InputObject $Disk -RemoveData -Confirm:$false
    Initialize-Disk -InputObject $Disk -PartitionStyle MBR
    $bootPart = $Disk | New-Partition -AssignDriveLetter -Size 200MB -IsActive
    $filePart = $Disk | New-Partition -AssignDriveLetter -UseMaximumSize
    Format-Volume -Partition $bootPart -FileSystem FAT32 -NewFileSystemLabel BOOT | Out-Null
    Format-Volume -Partition $filePart -FileSystem NTFS -NewFileSystemLabel INSTALLERS | Out-Null

    Write-Verbose "Setting up MBR Boot sector"
    $bootsect = Join-Path -Path $CDSource -ChildPath "boot\bootsect.exe"
    $bootDrive = "$($bootPart.DriveLetter):"
    & "$bootsect" /nt60 $bootDrive /mbr | Out-Null

    Write-Verbose "Copying boot files"
    Copy-Item -Path (Join-Path -Path $CDSource -ChildPath "boot") -Destination $bootDrive -Recurse
    Copy-Item -Path (Join-Path -Path $CDSource -ChildPath "efi") -Destination $bootDrive -Recurse
    Copy-Item -Path (Join-Path -Path $CDSource -ChildPath "bootmgr") -Destination $bootDrive
    if (Test-Path (Join-Path -Path $CDSource -ChildPath "bootmgr.efi")) {
        Copy-Item -Path (Join-Path -Path $CDSource -ChildPath "bootmgr.efi") -Destination $bootDrive
    } else {
        Write-Warning "Couldn't find \bootmgr.efi from install CD files; this disk might not boot from UEFI."
    }

    Write-Verbose "Copying installation files"
    $fileDrive = "$($filePart.DriveLetter):"
    Copy-Item -Path (Join-Path -Path $CDSource -ChildPath "Sources") -Destination $fileDrive -Recurse

    Write-Verbose "Updating BCD stores"
    foreach ($path in "$bootDrive\boot\bcd","$bootDrive\efi\microsoft\boot\bcd") {
        $store = Open-BcdStore -Path $path
        $bootMgr = Open-BcdBootManager -BcdStore $store
        Set-BcdBootManagerTimeout -BcdBootManager $bootMgr -Timeout 10
        Set-BcdBootManagerMenu -BcdBootManager $bootMgr -Enabled $true
        # The GUID below is the default bootloader
        $bootLdr = Open-BcdObject -BcdStore $store -Id "{7619dcc9-fafe-11d9-b411-000476eba25f}"
        Set-BcdOSLoaderDevice -BcdOSLoader $bootLdr -ImagePath "$fileDrive\sources\boot.wim"
    }
    if ($PassThru) {
        New-InstallDisk -BootVolume $bootDrive -InstallerVolume $fileDrive
    }
}

function Set-WdsBootImage {
    param (
        [Parameter(Mandatory)][String]$Path,
        [String]$Server
    )
    # Needed to mount image
    Set-ItemProperty -Path $Path -Name IsReadOnly -Value $false
    $mnt = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
    New-Item -ItemType Directory -Path $mnt | Out-Null
    Mount-WindowsImage -ImagePath $Path -Index 2 -Path $mnt | Out-Null
    $cfg = "[LaunchApps]`n%SYSTEMDRIVE%\sources\setup.exe, /wds /wdsdiscover"
    if ($Server) {
        $cfg += " /wdsserver:$Server"
    }
    $cfg | Set-Content -Path "$mnt\Windows\system32\winpeshl.ini"
    Dismount-WindowsImage -Path $mnt -Save | Out-Null
}

function Set-CustomBootImage {
    # $SourceDirectory is the directory without the drive letter (which is dynamic at boot)
    param (
        [Parameter(Mandatory)][String]$Path,
        [Parameter(Mandatory)][String]$SourceDirectory
    )
    # Needed to mount image
    Set-ItemProperty -Path $Path -Name IsReadOnly -Value $false
    $mnt = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.IO.Path]::GetRandomFileName())
    # Ignore the drive: it's not the same in WinPE and in running system
    New-Item -ItemType Directory -Path $mnt | Out-Null
    Mount-WindowsImage -ImagePath $Path -Index 2 -Path $mnt | Out-Null
    @"
@echo off
:while
for %%d in ( d e f g h i j k l m n o p q r q r s t u v w x y z ) do (
	if exist "%%d:$SourceDirectory\install.wim" (
		set installers=%%d
	)
)
if not defined installers (
	goto while
)
%SYSTEMDRIVE%\sources\setup.exe /installfrom:%installers%:$SourceDirectory\install.wim /m:%installers%:$SourceDirectory
"@  | Set-Content "$mnt\setup.cmd"
    "[LaunchApps]`n%SYSTEMDRIVE%\setup.cmd" | Set-Content -Path "$mnt\Windows\system32\winpeshl.ini"
    Dismount-WindowsImage -Path $mnt -Save | Out-Null
}

function Add-BootEntry {
    param (
        [Parameter(Mandatory)][String[]]$BcdPath,
        [Parameter(Mandatory)][String]$Description,
        [Parameter(Mandatory)][ValidatePattern('[a-zA-Z]:.*')][String]$ImagePath
    )
    $fileDrive = Split-Path -Path $ImagePath -Qualifier
    $filePath = Split-Path -Path $ImagePath -NoQualifier
    foreach ($path in $BcdPath) {
        $store = Open-BcdStore -Path $path
        $id = Copy-BcdObject -BcdStore $store -ID "{7619dcc9-fafe-11d9-b411-000476eba25f}"
        $bootLdr = Open-BcdObject -BcdStore $store -ID $id
        Set-BcdOSLoaderDevice -BcdOSLoader $bootLdr -ImagePath $ImagePath
        Set-BcdOSLoaderDescription -BcdOSLoader $bootLdr -Description $Description
        $bootMgr = Open-BcdBootManager -BcdStore $store
        Add-BcdBootManagerMenuEntry -BcdBootManager $bootMgr -ID $id
    }
}

function Add-Installer {
    [CmdletBinding()]
    param (
        [String]$CDSource = "D:\",
        [Parameter(Mandatory)]$InstallDisk,
        [Parameter(Mandatory)][String]$Key,
        [Parameter(Mandatory)][String]$Description
    )
    Write-Verbose "Copying installation files"
    $targetDir = $InstallDisk.InstallerVolume + ":\$Key"
    if (Test-Path -Path $targetDir) {
        throw "Directory $targetDir already exists"
    }
    Copy-Item -Path (Join-Path -Path $CDSource -ChildPath "Sources") -Destination $targetDir -Recurse
    Set-CustomBootImage -Path "$targetDir\boot.wim" -SourceDirectory "\$key"
    $InstallDisk.BcdStores | ForEach-Object {
        Add-BootEntry -BcdPath $_ -Description $Description -ImagePath "$targetDir\boot.wim"
    }
}

function Add-WdsInstaller {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][String]$ImageSource,
        [Parameter(Mandatory)]$InstallDisk,
        [Parameter(Mandatory)][String]$Key,
        [Parameter(Mandatory)][String]$Description,
        [String]$Server
    )
    $targetDir = $InstallDisk.InstallerVolume + ":\$key"
    if (Test-Path -Path $targetDir) {
        throw "Directory $targetDir already exists"
    }
    New-Item -ItemType Directory -Path $targetDir | Out-Null
    $image = Copy-Item -Path $ImageSource -Destination "$targetDir\boot.wim" -PassThru
    Set-WdsBootImage -Path $image.FullName -Server $Server
    $InstallDisk.BcdStores | ForEach-Object {
        Add-BootEntry -BcdPath $_ -Description $Description -ImagePath $image.FullName
    }
}
