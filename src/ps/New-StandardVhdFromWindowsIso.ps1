<#
    .SYNOPSIS

    Creates a new 40Gb VHDX with an installation of Windows as given by the windows install ISO file.


    .DESCRIPTION

    The New-StandardVhdFromWindowsIso script takes all the actions to create a new VHDX virtual hard drive with a windows install.


    .PARAMETER osIsoFile

    The full path to the ISO file that contains the windows installation.


    .PARAMETER osEdition

    The SKU or edition of the operating system that should be taken from the ISO and applied to the disk.


    .PARAMETER configPath

    The full path to the directory that contains the unattended file that contains the parameters for an unattended setup
    and any necessary script files which will be used during the configuration of the operating system.


    .PARAMETER machineName

    The name of the machine that will be created.


    .PARAMETER localAdminCredential

    The credential for the local administrator on the new machine.


    .PARAMETER vhdPath

    The full path to where the VHDX file should be output.


    .PARAMETER hypervHost

    The name of the Hyper-V host machine on which a temporary VM can be created.


    .PARAMETER staticMacAddress

    An optional static MAC address that is applied to the VM so that it can be given a consistent IP address.


    .PARAMETER wsusServer

    The name of the WSUS server that can be used to download updates from.


    .PARAMETER wsusTargetGroup

    The name of the WSUS computer target group that should be used to determine which updates should be installed.


    .PARAMETER scriptPath

    The full path to the directory that contains the Convert-WindowsImage and the Apply-WindowsUpdate scripts.


    .PARAMETER logPath

    The full path to the directory in which output log files can be stored.


    .PARAMETER tempPath

    The full path to the directory in which temporary files can be stored.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $osIsoFile = $(throw 'Please specify the full path of the windows install ISO file.'),

    [Parameter(Mandatory = $false)]
    [string] $osEdition = '',

    [Parameter(Mandatory = $true)]
    [string] $configPath,

    [Parameter(Mandatory = $true)]
    [string] $machineName,

    [Parameter(Mandatory = $true)]
    [PSCredential] $localAdminCredential,

    [Parameter(Mandatory = $true)]
    [string] $vhdPath,

    [Parameter(Mandatory = $true)]
    [string] $hypervHost,

    [Parameter(Mandatory = $false)]
    [string] $staticMacAddress,

    [Parameter(Mandatory = $true)]
    [string] $wsusServer,

    [Parameter(Mandatory = $true)]
    [string] $wsusTargetGroup,

    [Parameter(Mandatory = $false)]
    [string] $scriptPath = $PSScriptRoot,

    [Parameter(Mandatory = $true)]
    [string] $logPath = $(Join-Path $env:Temp ([System.Guid]::NewGuid.ToString())),

    [Parameter(Mandatory = $true)]
    [string] $tempPath = $(Join-Path $env:Temp ([System.Guid]::NewGuid.ToString()))
)

Write-Verbose "New-StandardVhdFromWindowsIso - osIsoFile = $osIsoFile"
Write-Verbose "New-StandardVhdFromWindowsIso - osEdition = $osEdition"
Write-Verbose "New-StandardVhdFromWindowsIso - configPath = $configPath"
Write-Verbose "New-StandardVhdFromWindowsIso - machineName = $machineName"
Write-Verbose "New-StandardVhdFromWindowsIso - vhdPath = $vhdPath"
Write-Verbose "New-StandardVhdFromWindowsIso - hypervHost = $hypervHost"
Write-Verbose "New-StandardVhdFromWindowsIso - wsusServer = $wsusServer"
Write-Verbose "New-StandardVhdFromWindowsIso - wsusTargetGroup = $wsusTargetGroup"
Write-Verbose "New-StandardVhdFromWindowsIso - scriptPath = $scriptPath"
Write-Verbose "New-StandardVhdFromWindowsIso - tempPath = $tempPath"

$ErrorActionPreference = 'Stop'

$commonParameterSwitches =
    @{
        Verbose = $PSBoundParameters.ContainsKey('Verbose');
        Debug = $false;
        ErrorAction = 'Stop'
    }

. (Join-Path $PSScriptRoot hyperv.ps1)
. (Join-Path $PSScriptRoot sessions.ps1)
. (Join-Path $PSScriptRoot Windows.ps1)
. (Join-Path $PSScriptRoot WinRM.ps1)


# -------------------------- Script functions --------------------------------

function New-VhdFromIso
{
    [CmdletBinding()]
    param(
        [string] $osIsoFile,
        [string] $osEdition,
        [string] $vhdPath,
        [string] $configPath,
        [string] $scriptPath
    )

    Write-Verbose "New-VhdFromIso - osIsoFile = $osIsoFile"
    Write-Verbose "New-VhdFromIso - osEdition = $osEdition"
    Write-Verbose "New-VhdFromIso - vhdPath = $vhdPath"
    Write-Verbose "New-VhdFromIso - configPath = $configPath"
    Write-Verbose "New-VhdFromIso - scriptPath = $scriptPath"

    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    $convertWindowsImageUrl = 'https://gallery.technet.microsoft.com/scriptcenter/Convert-WindowsImageps1-0fe23a8f/file/59237/7/Convert-WindowsImage.ps1'
    $convertWindowsImagePath = Join-Path $scriptPath 'Convert-WindowsImage.ps1'
    if (-not (Test-Path $convertWindowsImagePath))
    {
        Invoke-WebRequest `
            -Uri $convertWindowsImageUrl `
            -UseBasicParsing `
            -Method Get `
            -OutFile $convertWindowsImagePath `
            @commonParameterSwitches
    }

    $unattendPath = Join-Path $configPath 'unattend.xml'

    . $convertWindowsImagePath
    Convert-WindowsImage `
        -SourcePath $osIsoFile `
        -Edition $osEdition `
        -VHDPath $vhdPath `
        -SizeBytes 40GB `
        -VHDFormat 'VHDX' `
        -VHDType 'Dynamic' `
        -VHDPartitionStyle 'GPT' `
        -BCDinVHD 'VirtualMachine' `
        -UnattendPath $unattendPath `
        @commonParameterSwitches

    # Copy the additional script files to the drive
    $driveLetter = Mount-Vhdx -vhdPath $vhdPath @commonParameterSwitches
    try
    {
        # Copy the remaining configuration scripts
        $unattendScriptsDirectory = "$($driveLetter):\UnattendResources"
        if (-not (Test-Path $unattendScriptsDirectory))
        {
            New-Item -Path $unattendScriptsDirectory -ItemType Directory | Out-Null
        }

        Copy-Item -Path "$configPath\*" -Destination $unattendScriptsDirectory @commonParameterSwitches
    }
    finally
    {
        Dismount-Vhdx -vhdPath $vhdPath @commonParameterSwitches
    }
}

function New-VmFromVhdAndWaitForBoot
{
    [CmdletBinding()]
    param(
        [string] $vhdPath,
        [string] $machineName,
        [string] $hypervHost,
        [string] $staticMacAddress,
        [string] $bootWaitTimeout
    )

    Write-Verbose "New-VmFromVhd - vhdPath = $vhdPath"
    Write-Verbose "New-VmFromVhd - machineName = $machineName"
    Write-Verbose "New-VmFromVhd - hypervHost = $hypervHost"
    Write-Verbose "New-VmFromVhd - staticMacAddress = $staticMacAddress"
    Write-Verbose "New-VmFromVhd - bootWaitTimeout = $bootWaitTimeout"

    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    # Create a new Hyper-V virtual machine based on a VHDX Os disk
    if ((Get-VM -ComputerName $hypervHost | Where-Object { $_.Name -eq $machineName}).Count -gt 0)
    {
        Stop-VM $machineName -ComputerName $hypervHost -TurnOff -Confirm:$false -Passthru | Remove-VM -ComputerName $hypervHost -Force -Confirm:$false
    }

    $vm = New-HypervVm `
        -hypervHost $hypervHost `
        -vmName $machineName `
        -osVhdPath $vhdPath `
        @commonParameterSwitches

    if ($staticMacAddress -ne '')
    {
        # Ensure that the VM has a specific Mac address so that it will get a known IP address
        # That IP address will be added to the trustedhosts list so that we can remote into
        # the machine without having it be attached to the domain.
        $vm | Get-VMNetworkAdapter | Set-VMNetworkAdapter -StaticMacAddress $staticMacAddress @commonParameterSwitches
    }

    Start-VM -Name $machineName -ComputerName $hypervHost @commonParameterSwitches
    $waitResult = Wait-VmGuestOS `
        -vmName $machineName `
        -hypervHost $hypervHost `
        -timeOutInSeconds $bootWaitTimeout `
        @commonParameterSwitches

    if (-not $waitResult)
    {
        throw "Waiting for $machineName to start past the given timeout of $bootWaitTimeout"
    }
}

function Update-VhdWithWindowsPatches
{
    [CmdletBinding()]
    param(
        [string] $vhdPath,
        [string] $wsusServer,
        [string] $wsusTargetGroup,
        [string] $scriptPath,
        [string] $tempPath
    )

    Write-Verbose "Update-VhdWithWindowsPatches - vhdPath = $vhdPath"
    Write-Verbose "Update-VhdWithWindowsPatches - wsusServer = $wsusServer"
    Write-Verbose "Update-VhdWithWindowsPatches - wsusTargetGroup = $wsusTargetGroup"
    Write-Verbose "Update-VhdWithWindowsPatches - scriptPath = $scriptPath"
    Write-Verbose "Update-VhdWithWindowsPatches - tempPath = $tempPath"

    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    $applyWindowsUpdateUrl = 'https://gallery.technet.microsoft.com/Offline-Servicing-of-VHDs-df776bda/file/104350/1/Apply-WindowsUpdate.ps1'
    $applyWindowsUpdatePath = Join-Path $scriptPath 'Apply-WindowsUpdate.ps1'
    if (-not (Test-Path $applyWindowsUpdatePath))
    {
        Invoke-WebRequest `
            -Uri $applyWindowsUpdateUrl `
            -UseBasicParsing `
            -Method Get `
            -OutFile $(Join-Path $scriptPath 'Apply-WindowsUpdate.ps1') `
            @commonParameterSwitches
    }

    # Grab all the update packages for the given OS
    $mountPath = Join-Path $tempPath 'VhdMount'
    if (-not (Test-Path $mountPath))
    {
        New-Item -Path $mountPath -ItemType Directory | Out-Null
    }

    & $applyWindowsUpdatePath `
        -VhdPath $vhdPath `
        -MountDir $mountPath `
        -WsusServerName $wsusServer `
        -WsusServerPort 8530 `
        -WsusTargetGroupName $wsusTargetGroup `
        -WsusContentPath "\\$($wsusServer)\WsusContent" `
        @commonParameterSwitches

    Get-ChildItem -Path (Split-Path $vhdPath -Parent) -Filter *.log |
        Foreach-Object {
            Copy-Item -Path $_.FullName -Destination (Join-Path $logPath "$([System.IO.Path]::GetFileNameWithoutExtension($_.FullName))-ApplyPatches.log") @commonParameterSwitches
        }
}

# -------------------------- Script start --------------------------------

if (-not (Test-Path $tempPath))
{
    New-Item -Path $tempPath -ItemType Directory | Out-Null
}

New-VhdFromIso `
    -osIsoFile $osIsoFile `
    -osEdition $osEdition `
    -vhdPath $vhdPath `
    -configPath $configPath `
    -scriptPath $scriptPath `
    @commonParameterSwitches

Update-VhdWithWindowsPatches `
    -vhdPath $vhdPath `
    -wsusServer $wsusServer `
    -wsusTargetGroup $wsusTargetGroup `
    -scriptPath $scriptPath `
    -tempPath $tempPath `
    @commonParameterSwitches

$timeOutInSeconds = 900
New-VmFromVhdAndWaitForBoot `
    -vhdPath $vhdPath `
    -machineName $machineName `
    -hypervHost $hypervHost `
    -staticMacAddress $staticMacAddress `
    -bootWaitTimeout $timeOutInSeconds `
    @commonParameterSwitches

$connection = Get-ConnectionInformationForVm `
    -machineName $machineName `
    -hypervHost $hypervHost `
    -localAdminCredential $localAdminCredential `
    -timeOutInSeconds $timeOutInSeconds `
    @commonParameterSwitches

Restart-Machine `
    -connection $connection `
    -localAdminCredential $localAdminCredential `
    -timeOutInSeconds $timeOutInSeconds `
    @commonParameterSwitches

    # Apply missing updates
    # while ($hasPatches)
    #{
    #    Start-VM
    #    apply patch
    #}

# reboot the machine to make sure everything is ready

New-HypervVhdxTemplateFromVm `
    -vmName $machineName `
    -vhdPath $vhdPath `
    -hypervHost $hypervHost `
    -localAdminCredential $localAdminCredential `
    -logPath $logPath `
    -timeOutInSeconds $timeOutInSeconds `
    @commonParameterSwitches