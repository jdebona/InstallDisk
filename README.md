InstallDisk
===========

This is a PowerShell module for creating USB installation media for Windows (Vista and newer).  Unlike many other solutions, it has all of the following features:

* It creates disks bootable on both BIOS and UEFI.
* It accepts files larger than 4 GB.
* It is a pure Windows procedure (no special drivers or bootloaders - the same results can be achieved with standard Windows tools)
* It can boot WDS installations.
* It's open source - you can review the code, especially as it runs with elevated privileges.  Furthermore, the code is not compiled (1OO% PowerShell).

Requirements
------------

* Windows 10 or newer (might work with any Windows version with PowerShell 4.0, but not tested)
* a USB disk (in Windows terminology, i. e. Windows must be able to partition it)
* administrator privileges (required for partitioning and installing boot sector)
* files from a Windows installation media (physical CD, mounted ISO or files extracted from ISO) for each installer you want to put on the disk

Installation
------------

Copy files in your %PSMODULEPATH% and make sure your execution policy and permissions allow you to run scripts.

Provided functions
------------------

### Common parameters

In the various functions, parameters are named consistenly with the following definitions:

#### CDSource

A path to the Windows installation files (the root of a CD).  If the path is at the root of a drive, the `\`directory must be included explicitly, e.g. `X:\`.  Default is `D:\`.

#### Disk

A raw disk object as returned by `Get-Disk`.

#### InstallDisk

An install disk object, representing a disk that has been initialized with at least one installer, as returned by `Initialize-InstallDisk -PassThru` or `Get-InstallDisk`.

#### Key

A unique identifier for an installer; used as a directory name to contain its files on the install disk.

#### Description

A description for the installer; printed in the boot menu of the install disk.

### Functions

#### Get-USBDisk

A wrapper around Get-Disk, for conveniently finding your blank media.

#### Initialize-InstallDisk

Prepares a disk with a first installer (WARNING: this wipes the disk):

1. Creates 2 partitions: a small FAT partition for boot files and a large NTFS partition for installation files
2. Configures bootloaders (UEFI and BIOS)
3. Copies installation files to the NTFS partition.

Parameters:

* CDSource
* Disk
* PassThru: switch: return the created install disk (see InstallDisk)

Note: the boot loaders are configured based on the CDSource parameter, so if your initial installer can't boot on UEFI, your install disk won't be able to boot on UEFI.

#### Get-InstallDisk

Returns all install disks (initialized with `Initialize-InstallDisk`) currently connected.

Parameters:

* OnlyUSB: switch: limit output to USB disks.

#### Add-Installer

Adds an installer to an existing install disk:

1. Copies installation files to the NTFS partition of the install disk
2. Edits the boot.wim image so that it can find installation files from the non-standard location
3. Adds entries to the install disk's bootloader

Parameters:

* CDSource
* InstallDisk
* Key
* Description

A good idea is, after initializing an install disk with the first installer, to add the same installer again.  This way,

* the second installer has a label showing what it is,
* files from another installer can be dropped in the `Sources` directory if you need to quickly add one and can't use this module.

#### Add-WdsInstaller

Adds a WDS boot image to an existing install disk.

Parameters:

* ImageSource: path to the boot.wim image to use (download it from a WDS server)
* InstallDisk
* Key
* Description
* Server: if set, force the boot image to use the specified WDS server; otherwise, rely on DNS to find one.

### Examples

Initialize a disk with a first installer (e.g. Windows 10) and put a second installer (e.g. Windows 2012 R2), assuming only one USB disk is connected, and your DVD drive is `E:`.

```powershell
$disk = Get-USBDisk
# Load Windows 10 DVD in drive E:
$installDisk = Initialize-InstallDisk -CDSource E:\ -Disk $disk -PassThru
# Load Windows 2012 R2 DVD in drive E:
Add-Installer -CDSource E:\ -InstallDisk $installDisk -Key win2012r2 -Description "Windows Server 2012 R2"
```

Add WDS installers to a previously created install disk (assuming only one install disk is connected). One installer will use DHCP to find a WDS server, the other one will use wds2.example.org.

```powershell
$installDisk = Get-InstallDisk
Add-WdsInstaller -ImageSource C:\path\to\boot.wim -InstallDisk $installDisk -Key "wds" -Description "WDS install"
Add-WdsInstaller -ImageSource C:\path\to\boot.wim -InstallDisk $installDisk -Key "wds2" -Description "WDS install via wds2.example.org" -Server wds2.example.org

```
