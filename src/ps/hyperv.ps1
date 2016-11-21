<#
    .SYNOPSIS

    Dismounts the VHDX drive from the operating system.


    .DESCRIPTION

    The Dismount-Vhdx function dismounts the VHDX drive from the operating system.


    .PARAMETER vhdPath

    The full path to the VHDX file that has been mounted.
#>
function Dismount-Vhdx
{
    [CmdletBinding()]
    param(
        [string] $vhdPath
    )

    Write-Verbose "Dismount-Vhdx - vhdPath = $vhdPath"

    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    Dismount-DiskImage -ImagePath $vhdPath @commonParameterSwitches
}

<#
    .SYNOPSIS

    Gets the connection information used to connect to a given Hyper-V VM.


    .DESCRIPTION

    The Get-ConnectionInformationForVm function gets the connection information used to connect to a given Hyper-V VM.


    .PARAMETER machineName

    The name of the VM.


    .PARAMETER hypervHost

    The name of the Hyper-V host machine.


    .PARAMETER localAdminCredential

    The credentials for the local administrator account.


    .PARAMETER timeOutInSeconds

    The amount of time that the function will wait for at the individual stages for a connection.


    .OUTPUTS

    A custom object containing the connection information for the VM. Available properties are:

        Name             The machine name of the VM
        IPAddress        The IP address of the VM
        Session          A powershell remoting session
#>
function Get-ConnectionInformationForVm
{
    [CmdletBinding()]
    param(
        [string] $machineName,
        [string] $hypervHost,
        [pscredential] $localAdminCredential,
        [int] $timeOutInSeconds
    )

    Write-Verbose "Get-ConnectionInformationForVm - machineName = $machineName"
    Write-Verbose "Get-ConnectionInformationForVm - hypervHost = $hypervHost"
    Write-Verbose "Get-ConnectionInformationForVm - localAdminCredential = $localAdminCredential"
    Write-Verbose "Get-ConnectionInformationForVm - timeOutInSeconds = $timeOutInSeconds"

    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    # Just because we have a positive connection does not mean that connection will stay there
    # During initialization the machine will reboot a number of times so it is possible
    # that we get caught out due to one of these reboots. In that case we'll just try again.
    $maxRetry = 10
    $count = 0

    $result = New-Object psObject
    Add-Member -InputObject $result -MemberType NoteProperty -Name Name -Value $machineName
    Add-Member -InputObject $result -MemberType NoteProperty -Name IPAddress -Value $null
    Add-Member -InputObject $result -MemberType NoteProperty -Name Session -Value $null

    [System.Management.Automation.Runspaces.PSSession]$vmSession = $null
    while (($vmSession -eq $null) -and ($count -lt $maxRetry))
    {
        $count = $count + 1

        try
        {
            $ipAddress = Wait-VmIPAddress `
                -vmName $machineName `
                -hypervHost $hypervHost `
                -timeOutInSeconds $timeOutInSeconds `
                @commonParameterSwitches
            if (($ipAddress -eq $null) -or ($ipAddress -eq ''))
            {
                throw "Failed to obtain an IP address for $machineName within the specified timeout of $timeOutInSeconds seconds."
            }

            Write-Verbose "IP address for $machineName is: $($ipAddress). Trying to connect via WinRM ..."

            # The guest OS may be up and running, but that doesn't mean we can connect to the
            # machine through powershell remoting, so ...
            $waitResult = Wait-WinRM `
                -ipAddress $ipAddress `
                -credential $localAdminCredential `
                -timeOutInSeconds $timeOutInSeconds `
                @commonParameterSwitches
            if (-not $waitResult)
            {
                throw "Waiting for $machineName to be ready for remote connections has timed out with timeout of $timeOutInSeconds"
            }

            Write-Verbose "Wait-WinRM completed successfully, making connection to machine $ipAddress ..."
            $vmSession = New-PSSession `
                -computerName $ipAddress `
                -credential $localAdminCredential `
                @commonParameterSwitches

            $result.IPAddress = $ipAddress
            $result.Session = $vmSession
        }
        catch
        {
            Write-Verbose "Failed to connect to the VM. Most likely due to a VM reboot. Trying another $($maxRetry - $count) times. Error was: $($_.Exception.ToString())"
        }
    }

    return $result
}

<#
    .SYNOPSIS

    Gets the drive letter for the drive with the given drive number


    .DESCRIPTION

    The Get-DriveLetter function returns the drive letter for the drive with the given drive number


    .PARAMETER driveNumber

    The number of the drive.


    .OUTPUT

    The letter of the drive.
#>
function Get-DriveLetter
{
    [CmdletBinding()]
    [OutputType([char])]
    param(
        [int] $driveNumber
    )

    # The first drive is C which is ASCII 67
    return [char]($driveNumber + 67)
}

<#
    .SYNOPSIS

    Gets the IP address for a given hyper-V VM.


    .DESCRIPTION

    The Get-IPAddressForVm function gets the IP address for a given VM.


    .PARAMETER vmName

    The name of the VM.


    .PARAMETER hypervHost

    The name of the machine which is the Hyper-V host for the domain.


    .OUTPUT

    The letter of the drive.
#>
function Get-IPAddressForVm
{
    [CmdletBinding()]
    param(
        [string] $vmName,
        [string] $hypervHost
    )

    Write-Verbose "Get-IPAddressForVm - vmName = $vmName"
    Write-Verbose "Get-IPAddressForVm - hypervHost = $hypervHost"

    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    # Get the IPv4 address for the VM
    $ipAddress = Get-VM -Name $vmName -ComputerName $hypervHost |
        Select-Object -ExpandProperty NetworkAdapters |
        Select-Object -ExpandProperty IPAddresses |
        Select-Object -First 1

    return $ipAddress
}

<#
    .SYNOPSIS

    Invokes sysprep on a Hyper-V VM and waits for the machine to shut down.


    .DESCRIPTION

    The Invoke-SysprepOnVmAndWaitShutdown function invokes sysprep on a Hyper-V VM and waits for the machine to shut down


    .PARAMETER machineName

    The name of the VM.


    .PARAMETER hypervHost

    The name of the machine which is the Hyper-V host for the domain.


    .PARAMETER localAdminCredential

    The credentials for the local administrator on the VM.


    .PARAMETER timeOutInSeconds

    The amount of time in seconds the function should wait for the guest OS to be started.
#>
function Invoke-SysprepOnVmAndWaitShutdown
{
    [CmdletBinding()]
    param(
        [string] $machineName,
        [string] $hypervHost,
        [pscredential] $localAdminCredential,
        [int] $timeOutInSeconds
    )

    Write-Verbose "Invoke-SysprepOnVmAndWaitShutdown - machineName = $machineName"
    Write-Verbose "Invoke-SysprepOnVmAndWaitShutdown - hypervHost = $hypervHost"
    Write-Verbose "Invoke-SysprepOnVmAndWaitShutdown - localAdminCredential = $localAdminCredential"
    Write-Verbose "Invoke-SysprepOnVmAndWaitShutdown - timeOutInSeconds = $timeOutInSeconds"

    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    . (Join-Path $PSScriptRoot 'windows.ps1')

    $result = Get-ConnectionInformationForVm `
        -machineName $machineName `
        -hypervHost $hypervHost `
        -localAdminCredential $localAdminCredential `
        -timeOutInSeconds $timeOutInSeconds `
        @commonParameterSwitches
    if ($result.Session -eq $null)
    {
        throw "Failed to connect to $machineName"
    }

    Invoke-Sysprep `
        -connectionInformation $result `
        -timeOutInSeconds $timeOutInSeconds `
        @commonParameterSwitches

    # Wait till machine is stopped
    $waitResult = Wait-VmStopped `
        -vmName $machineName `
        -hypervHost $hypervHost `
        -timeOutInSeconds $timeOutInSeconds `
        @commonParameterSwitches

    if (-not $waitResult)
    {
        throw "VM $machineName failed to shut down within $timeOutInSeconds seconds."
    }
}

<#
    .SYNOPSIS

    Mounts the VHDX drive in the operating system and returns the drive letter for the new drive.


    .DESCRIPTION

    The Mount-Vhdx function mounts the VHDX drive in the operating system and returns the
    drive letter for the new drive.


    .PARAMETER vhdPath

    The full path to the VHDX file that has been mounted.


    .OUTPUTS

    The drive letter for the newly mounted drive.
#>
function Mount-Vhdx
{
    [CmdletBinding()]
    param(
        [string] $vhdPath
    )

    Write-Verbose "Mount-Vhdx - vhdPath = $vhdPath"

    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    # store all the known drive letters because we can't directly get the drive letter
    # from the mounting operation so we have to compare the before and after pictures.
    $before = (Get-Volume).DriveLetter

    # Mounting the drive using Mount-DiskImage instead of Mount-Vhd because for the latter we need Hyper-V to be installed
    # which we can't do on a VM
    Mount-DiskImage -ImagePath $vhdPath -StorageType VHDX | Out-Null

    # Get all the current drive letters. The only new one should be the drive we just mounted
    $after = (Get-Volume).DriveLetter
    $driveLetter = compare $before $after -Passthru

    return $driveLetter
}

<#
    .SYNOPSIS

    Runs sysprep on the given VM, waits for it to turn off and then deletes the VM and turns the VM VHDX into a template.


    .DESCRIPTION

    The New-HyperVVhdxTemplateFromVm function runs sysprep on the given VM, waits for it to turn off and then
    deletes the VM and turns the VM VHDX into a template.


    .PARAMETER vmName

    The name of the VM.


    .PARAMETER vhdPath

    The full path to where the VHDX file should be output.


    .PARAMETER vhdxTemplatePath

    The full path to the VHDX file that will contain the template once the function completes.


    .PARAMETER hypervHost

    The name of the machine which is the Hyper-V host for the domain.


    .PARAMETER localAdminCredential

    The credential for the local administrator on the new machine.


    .PARAMETER logPath

    The full path to the directory into which the log files should be copied.


    .PARAMETER timeOutInSeconds

    The maximum amount of time in seconds that this function will wait for VM to enter the off state.
#>
function New-HyperVVhdxTemplateFromVm
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $vmName,

        [Parameter(Mandatory = $true)]
        [string] $vhdPath,

        [Parameter(Mandatory = $true)]
        [string] $vhdxTemplatePath,

        [Parameter(Mandatory = $true)]
        [string] $hypervHost,

        [Parameter(Mandatory = $true)]
        [PSCredential] $localAdminCredential,

        [Parameter(Mandatory = $true)]
        [string] $logPath,

        [Parameter()]
        [ValidateScript({$_ -ge 1 -and $_ -le [system.int64]::maxvalue})]
        [int] $timeOutInSeconds = 900
    )

    Write-Verbose "New-HyperVVhdxTemplateFromVm - vmName: $vmName"
    Write-Verbose "New-HyperVVhdxTemplateFromVm - vhdPath: $vhdPath"
    Write-Verbose "New-HyperVVhdxTemplateFromVm - hypervHost: $hypervHost"
    Write-Verbose "New-HyperVVhdxTemplateFromVm - logPath: $logPath"
    Write-Verbose "New-HyperVVhdxTemplateFromVm - timeOutInSeconds: $timeOutInSeconds"

    # Stop everything if there are errors
    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    Invoke-SysprepOnVmAndWaitShutdown `
        -machineName $vmName `
        -hypervHost $hypervHost `
        -localAdminCredential $localAdminCredential `
        -timeOutInSeconds $timeOutInSeconds `
        @commonParameterSwitches

    # Delete VM
    Remove-VM `
        -computerName $hypervHost `
        -Name $vmName `
        -Force `
        @commonParameterSwitches

    # Optimize the VHDX
    #
    # Mounting the drive using Mount-DiskImage instead of Mount-Vhd because for the latter we need Hyper-V to be installed
    # which we can't do on a VM
    $driveLetter = Mount-Vhdx -vhdPath $vhdPath @commonParameterSwitches
    try
    {
        # Copy the log files
        Get-ChildItem -Path "$($driveLetter):\windows\Panther" -Filter *.log -recurse |
            Foreach-Object {
                $directoryName = [System.IO.Path]::GetFileName((Split-Path $_.FullName -Parent))
                $fileName = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)
                Copy-Item -Path $_.FullName -Destination (Join-Path $logPath "$($fileName)-$($directoryName).log") @commonParameterSwitches
            }

        # Remove root level files we don't need anymore
        attrib -s -h "$($driveLetter):\pagefile.sys"
        Remove-Item -Path "$($driveLetter):\pagefile.sys" -Force -Verbose

        # Clean up all the user profiles except for the default one
        $userProfileDirectories = Get-ChildItem -Path "$($driveLetter):\Users\*" -Directory -Exclude 'Default', 'Public'
        foreach($userProfileDirectory in $userProfileDirectories)
        {
            Remove-Item -Path $userProfileDirectory.FullName -Recurse -Force @commonParameterSwitches
        }

        # Clean up the event logs
        $eventLogFiles = Get-ChildItem -Path "$($driveLetter):\windows\System32\Winevt\Logs\*" -File -Recurse
        foreach($eventLogFile in $eventLogFiles)
        {
            Remove-Item -Path $eventLogFile.FullName -Force @commonParameterSwitches
        }

        # Clean up the WinSXS store, and remove any superceded components. Updates will no longer be able to be uninstalled,
        # but saves a considerable amount of disk space.
        dism.exe /image:$($driveLetter):\ /Cleanup-Image /StartComponentCleanup /ResetBase

        $pathsToRemove = @(
            "$env:localappdata\Nuget",
            "$env:localappdata\temp\*",
            "$($driveLetter):\windows\logs",
            "$($driveLetter):\windows\panther",
            "$($driveLetter):\windows\temp\*",
            "$($driveLetter):\windows\winsxs\manifestcache")
        foreach($path in $pathsToRemove)
        {
            if (Test-Path $path)
            {
                try
                {
                    Remove-Item $path -Recurse -Force @commonParameterSwitches
                }
                catch
                {
                    # ignore it
                }
            }
        }

        Get-ChildItem -Path (Split-Path $vhdPath -Parent) -Filter *.log |
            Foreach-Object {
                $fileName = [System.IO.Path]::GetFileNameWithoutExtension($_.FullName)
                Copy-Item -Path $_.FullName -Destination (Join-Path $logPath "$($fileName)-cleanimage.log") @commonParameterSwitches
            }

        Remove-Item -Path "$($driveLetter):\*.log" -Force @commonParameterSwitches

        Write-Verbose "defragging ..."
        if (Test-Command -commandName 'Optimize-Volume')
        {
            Optimize-Volume -DriveLetter $driveLetter -Defrag @commonParameterSwitches
        }
            else
        {
            Defrag.exe $driveLetter /H
        }
    }
    finally
    {
        Dismount-Vhdx -vhdPath $vhdPath @commonParameterSwitches
    }

    Copy-Item -Path $vhdPath -Destination $vhdxTemplatePath @commonParameterSwitches
}

<#
    .SYNOPSIS

    Creates a new Hyper-V virtual machine with the given properties.


    .DESCRIPTION

    The New-HypervVm function creates a new Hyper-V virtual machine with the provided properties.


    .PARAMETER vmName

    The name of the VM.


    .PARAMETER osVhdPath

    The full path of the VHD that contains the pre-installed OS.


    .PARAMETER vmAdditionalDiskSizesInGb

    An array containing the sizes, in Gb, of any additional VHDs that should be attached to the virtual machine.


    .PARAMETER hypervHost

    The name of the machine which is the Hyper-V host for the domain.


    .PARAMETER vmStoragePath

    The full path of the directory where the virtual machine files should be stored.


    .PARAMETER vhdxStoragePath

    The full path of the directory where the virtual hard drive files should be stored.


    .OUTPUTS

    The VM object.
#>
function New-HypervVm
{
    [CmdletBinding()]
    [OutputType([void])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $vmName,

        [Parameter(Mandatory = $true)]
        [string] $osVhdPath,

        [Parameter(Mandatory = $false)]
        [int[]] $vmAdditionalDiskSizesInGb,

        [Parameter(Mandatory = $false)]
        [string] $hypervHost = $env:COMPUTERNAME,

        [Parameter(Mandatory = $false)]
        [string] $vhdxStoragePath,

        [Parameter(Mandatory = $false)]
        [string] $vmStoragePath
    )

    Write-Verbose "New-HypervVm - vmName: $vmName"
    Write-Verbose "New-HypervVm - osVhdPath: $osVhdPath"
    Write-Verbose "New-HypervVm - vmAdditionalDiskSizesInGb: $vmAdditionalDiskSizesInGb"
    Write-Verbose "New-HypervVm - hypervHost: $hypervHost"
    Write-Verbose "New-HypervVm - vhdxStoragePath: $vhdxStoragePath"
    Write-Verbose "New-HypervVm - vmStoragePath: $vmStoragePath"

    # Stop everything if there are errors
    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    # Make sure we have a local path to the VHD file
    $osVhdLocalPath = $osVhdPath
    if ($osVhdLocalPath.StartsWith("$([System.IO.Path]::DirectorySeparatorChar)$([System.IO.Path]::DirectorySeparatorChar)"))
    {
        $uncServerPath = "\\$($hypervHost)\"
        $shareRoot = $osVhdLocalPath.SubString($uncServerPath.Length, $osVhdLocalPath.IndexOf('\', $uncServerPath.Length) - $uncServerPath.Length)

        $shareList = Get-WmiObject -Class Win32_Share -ComputerName $hypervHost @commonParameterSwitches
        $localShareRoot = $shareList | Where-Object { $_.Name -eq $shareRoot} | Select-Object -ExpandProperty Path

        $osVhdLocalPath = $osVhdLocalPath.Replace((Join-Path $uncServerPath $shareRoot), $localShareRoot)
    }

    $vmSwitch = Get-VMSwitch -ComputerName $hypervHost @commonParameterSwitches | Select-Object -First 1

    $vmMemoryInBytes = 2 * 1024 * 1024 * 1024
    if (($vmStoragePath -ne $null) -and ($vmStoragePath -ne ''))
    {
        $vm = New-Vm `
            -Name $vmName `
            -Path $vmStoragePath `
            -VHDPath $osVhdLocalPath `
            -MemoryStartupBytes $vmMemoryInBytes `
            -SwitchName $vmSwitch.Name `
            -Generation 2 `
            -BootDevice 'VHD' `
            -ComputerName $hypervHost `
            -Confirm:$false `
            @commonParameterSwitches
    }
    else
    {
        $vm = New-Vm `
            -Name $vmName `
            -VHDPath $osVhdLocalPath `
            -MemoryStartupBytes $vmMemoryInBytes `
            -SwitchName $vmSwitch.Name `
            -Generation 2 `
            -BootDevice 'VHD' `
            -ComputerName $hypervHost `
            -Confirm:$false `
            @commonParameterSwitches
    }

     $vm = $vm |
        Set-Vm `
            -ProcessorCount 1 `
            -Confirm:$false `
            -Passthru `
            @commonParameterSwitches

    if ($vmAdditionalDiskSizesInGb -eq $null)
    {
        $vmAdditionalDiskSizesInGb = [int[]](@())
    }

    for ($i = 0; $i -lt $vmAdditionalDiskSizesInGb.Length; $i++)
    {
        $diskSize = $vmAdditionalDiskSizesInGb[$i]

        $driveLetter = Get-DriveLetter -driveNumber ($i + 1)
        $path = Join-Path $vhdxStoragePath "$($vmName)_$($driveLetter).vhdx"
        New-Vhd `
            -Path $path `
            -SizeBytes "$($diskSize)GB" `
            -VHDFormat 'VHDX'
            -Dynamic `
            @commonParameterSwitches
        Add-VMHardDiskDrive `
            -Path $path `
            -VM $vm `
            @commonParameterSwitches
    }

    return $vm
}


<#
    .SYNOPSIS

    Creates a new Hyper-V virtual machine from the given base template with the given properties.


    .DESCRIPTION

    The New-HypervVmFromBaseImage function creates a new Hyper-V virtual machine from the given base template with the provided properties.


    .PARAMETER vmName

    The name of the VM.


    .PARAMETER baseVhdx

    The full path of the template VHDx that contains the pre-installed OS.


    .PARAMETER vmAdditionalDiskSizesInGb

    An array containing the sizes, in Gb, of any additional VHDs that should be attached to the virtual machine.


    .PARAMETER configPath

    The full path to the directory that contains the unattended file that contains the parameters for an unattended setup
    and any necessary script files which will be used during the configuration of the operating system.


    .PARAMETER hypervHost

    The name of the machine which is the Hyper-V host for the domain.


    .PARAMETER vhdxStoragePath

    The full path of the directory where the virtual hard drive files should be stored.


    .OUTPUTS

    The VM object.
#>
function New-HypervVmFromBaseImage
{
    [CmdletBinding()]
    [OutputType([void])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '')]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $vmName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $baseVhdx,

        [Parameter(Mandatory = $false)]
        [int[]] $vmAdditionalDiskSizesInGb,

        [Parameter(Mandatory = $true)]
        [string] $configPath,

        [Parameter(Mandatory = $false)]
        [string] $hypervHost = $env:COMPUTERNAME,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $vhdxStoragePath
    )

    Write-Verbose "New-HypervVmFromBaseImage - vmName: $vmName"
    Write-Verbose "New-HypervVmFromBaseImage - baseVhdx: $baseVhdx"
    Write-Verbose "New-HypervVmFromBaseImage - vmAdditionalDiskSizesInGb: $vmAdditionalDiskSizesInGb"
    Write-Verbose "New-HypervVmFromBaseImage - configPath: $configPath"
    Write-Verbose "New-HypervVmFromBaseImage - hypervHost: $hypervHost"
    Write-Verbose "New-HypervVmFromBaseImage - vhdxStoragePath: $vhdxStoragePath"

    # Stop everything if there are errors
    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    # Create a copy of the VHDX file and then mount it
    $vhdxPath = Join-Path $vhdxStoragePath "$($vmName.ToLower()).vhdx"
    Copy-Item -Path $baseVhdx -Destination $vhdxPath -Verbose
    if (Get-ItemProperty -Path $vhdxPath -Name IsReadOnly)
    {
        Set-ItemProperty -Path $vhdxPath -Name IsReadOnly -Value $false
    }

    try
    {
        $driveLetter = Mount-Vhdx -vhdPath $vhdxPath @commonParameterSwitches

        # Copy the remaining configuration scripts
        $unattendScriptsDirectory = "$($driveLetter):\UnattendResources"
        if (-not (Test-Path $unattendScriptsDirectory))
        {
            New-Item -Path $unattendScriptsDirectory -ItemType Directory | Out-Null
        }

        Copy-Item -Path "$configPath\unattend.xml" -Destination "$($driveLetter):\unattend.xml" @commonParameterSwitches
        Copy-Item -Path "$configPath\*" -Exclude "$configPath\unattend.xml" -Destination $unattendScriptsDirectory @commonParameterSwitches
    }
    finally
    {
        Dismount-Vhdx -vhdPath $vhdxPath @commonParameterSwitches
    }

    $vm = New-HypervVm `
        -vmName $vmName `
        -osVhdPath $vhdxPath `
        -vmAdditionalDiskSizesInGb $vmAdditionalDiskSizesInGb `
        -hypervHost $hypervHost `
        -vhdxStoragePath '' `
        -vmStoragePath '' `
        @commonParameterSwitches

    return $vm
}

<#
    .SYNOPSIS

    Waits for the guest operating system to be started.


    .DESCRIPTION

    The Wait-VmGuestOS function waits for the guest operating system on a given VM to be started.


    .PARAMETER vmName

    The name of the VM.


    .PARAMETER hypervHost

    The name of the VM host machine.


    .PARAMETER timeOutInSeconds

    The amount of time in seconds the function should wait for the guest OS to be started.


    .OUTPUTS

    Returns $true if the guest OS was started within the timeout period or $false if the guest OS was not
    started within the timeout period.
#>
function Wait-VmGuestOS
{
    [CmdLetBinding()]
    param(
        [string] $vmName,

        [string] $hypervHost,

        [Parameter()]
        [ValidateScript({$_ -ge 1 -and $_ -le [system.int64]::maxvalue})]
        [int] $timeOutInSeconds = 900 #seconds
    )

    Write-Verbose "Wait-VmGuestOS - vmName = $vmName"
    Write-Verbose "Wait-VmGuestOS - hypervHost = $hypervHost"
    Write-Verbose "Wait-VmGuestOS - timeOutInSeconds = $timeOutInSeconds"

    # Stop everything if there are errors
    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    $startTime = Get-Date
    $endTime = $startTime + (New-TimeSpan -Seconds $timeOutInSeconds)
    do
    {
        if ((Get-Date) -ge $endTime)
        {
            Write-Verbose "The VM $vmName failed to shut down in the alotted time of $timeOutInSeconds"
            return $false
        }

        Write-Verbose "Waiting for VM $vmName to be ready for use [total wait time so far: $((Get-Date) - $startTime)] ..."
        Start-Sleep -seconds 5
    }
    until ((Get-VMIntegrationService -VMName $vmName -ComputerName $hypervHost @commonParameterSwitches | Where-Object { $_.name -eq "Heartbeat" }).PrimaryStatusDescription -eq "OK")

    return $true
}

<#
    .SYNOPSIS

    Waits for the guest operating system on a VM to be provided with an IP address.


    .DESCRIPTION

    The Wait-VmIPAddress function waits for the guest operating system on a VM to be provided with an IP address.


    .PARAMETER vmName

    The name of the VM.


    .PARAMETER hypervHost

    The name of the VM host machine.


    .PARAMETER timeOutInSeconds

    The amount of time in seconds the function should wait for the guest OS to be assigned an IP address.


    .OUTPUTS

    Returns the IP address of the VM or $null if no IP address could be obtained within the timeout period.
#>
function Wait-VmIPAddress
{
    [CmdletBinding()]
    param(
        [string] $vmName,
        [string] $hypervHost,

        [Parameter()]
        [ValidateScript({$_ -ge 1 -and $_ -le [system.int64]::maxvalue})]
        [int] $timeOutInSeconds = 900 #seconds
    )

    Write-Verbose "Wait-VmIPAddress - vmName = $vmName"
    Write-Verbose "Wait-VmIPAddress - hypervHost = $hypervHost"
    Write-Verbose "Wait-VmIPAddress - timeOutInSeconds = $timeOutInSeconds"

    # Stop everything if there are errors
    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    $startTime = Get-Date
    $endTime = $startTime + (New-TimeSpan -Seconds $timeOutInSeconds)
    while ((Get-Date) -le $endTime)
    {
        $ipAddress = Get-IPAddressForVm -vmName $vmName -hypervHost $hypervHost @commonParameterSwitches
        if (($ipAddress -ne $null) -and ($ipAddress -ne ''))
        {
            return $ipAddress
        }

        Write-Verbose "Waiting for VM $vmName to be given an IP address [total wait time so far: $((Get-Date) - $startTime)] ..."
        Start-Sleep -seconds 5
    }

    return $null
}

<#
    .SYNOPSIS

    Waits for a Hyper-V VM to be in the off state.


    .DESCRIPTION

    The Wait-VmStopped function waits for a Hyper-V VM to enter the off state.


    .PARAMETER vmName

    The name of the VM.


    .PARAMETER hypervHost

    The name of the VM host machine.


    .PARAMETER timeOutInSeconds

    The maximum amount of time in seconds that this function will wait for VM to enter
    the off state.
#>
function Wait-VmStopped
{
    [CmdletBinding()]
    param(
        [string] $vmName,

        [string] $hypervHost,

        [Parameter()]
        [ValidateScript({$_ -ge 1 -and $_ -le [system.int64]::maxvalue})]
        [int] $timeOutInSeconds = 900 #seconds
    )

    Write-Verbose "Wait-VmStopped - vmName = $vmName"
    Write-Verbose "Wait-VmStopped - hypervHost = $hypervHost"
    Write-Verbose "Wait-VmStopped - timeOutInSeconds = $timeOutInSeconds"

    # Stop everything if there are errors
    $ErrorActionPreference = 'Stop'

    $commonParameterSwitches =
        @{
            Verbose = $PSBoundParameters.ContainsKey('Verbose');
            Debug = $false;
            ErrorAction = 'Stop'
        }

    $startTime = Get-Date
    $endTime = $startTime + (New-TimeSpan -Seconds $timeOutInSeconds)
    Write-Verbose "Waiting till: $endTime"

    while ($true)
    {
        Write-Verbose "Start of the while loop ..."
        if ((Get-Date) -ge $endTime)
        {
            Write-Verbose "The VM $vmName failed to shut down in the alotted time of $timeOutInSeconds"
            return $false
        }

        Write-Verbose "Waiting for VM $vmName to shut down [total wait time so far: $((Get-Date) - $startTime)] ..."
        try
        {
            Write-Verbose "Getting VM state ..."
            $integrationServices = Get-VM -Name $vmName -ComputerName $hypervHost @commonParameterSwitches | Get-VMIntegrationService

            $offCount = 0
            foreach($service in $integrationServices)
            {
                Write-Verbose "vm $vmName integration service $($service.Name) is at state $($service.PrimaryStatusDescription)"
                if (($service.PrimaryStatusDescription -eq $null) -or ($service.PrimaryStatusDescription -eq ''))
                {
                    $offCount = $offCount + 1
                }
            }

            if ($offCount -eq $integrationServices.Length)
            {
                Write-Verbose "VM $vmName has turned off"
                return $true
            }

        }
        catch
        {
            Write-Verbose "Could not connect to $vmName. Error was $($_.Exception.Message)"
        }

        Write-Verbose "Waiting for 5 seconds ..."
        Start-Sleep -seconds 5
    }

    Write-Verbose "Waiting for VM $vmName to stop failed outside the normal failure paths."
    return $false
}