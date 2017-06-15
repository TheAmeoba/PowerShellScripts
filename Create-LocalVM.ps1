<#
    .SYNOPSIS
        Creates a new VM on the local machine
    
    .DESCRIPTION
        Created to make spinning up a VM on my laptop quick and easy using a differenced disk. 
        Then I've been adding features to make it more versitile.

    .PARAMETER Path
        Location of the VM's Virtual Hard Disks and configuration files

    .PARAMETER ReferenceVHD
        Defines a VHD to either copy or use in creating a differencing disk

    .PARAMETER VHDMethod
        For use with ReferenceVHD
        Valid inputs are Difference, Copy
        Defines what type of VHD is to be created from the reference

    .PARAMETER FixedMemory
        Used to fix the ammound of memory instead of dynamicly adjusting

    .PARAMETER ISO
        Iso file to be loaded on VMs DVD drive for use as boot device

    .PARAMETER VMSwitch
        Name of the Hyper-V switch to connect this VM to.

    .PARAMETER start
        Defines whether VM should start straight away after creation

    .NOTES
        Author: Simon Baker
        Version: 2.2
        Date Created: 2015-09-29
        Last Modified: 2017-06-02

        ToDo:
            - Data Drive?
#>
[CmdletBinding(SupportsShouldProcess=$True)]
Param(
    [Parameter(Mandatory=$True,Position=1)][string]$VMName,
    [Parameter(Mandatory=$True,Position=2)][string]$Path,
    [string]$ISO,
    [Parameter(ParameterSetName="TemplateNew", Mandatory=$True)][Parameter(ParameterSetName="TemplateVHD", Mandatory=$True)][ValidateSet('LinuxLegacy','LinuxClient','LinuxServer','WindowsClient','WindowsServer')][string]$Template,
    [Parameter(ParameterSetName="TemplateVHD", Mandatory=$True)][Parameter(ParameterSetName="CustomVHD", Mandatory=$True)][Parameter(ParameterSetName="CustomVHDDyn", Mandatory=$True)][string]$ReferenceVHD,
    [Parameter(ParameterSetName="TemplateVHD", Mandatory=$True)][Parameter(ParameterSetName="CustomVHD", Mandatory=$True)][Parameter(ParameterSetName="CustomVHDDyn", Mandatory=$True)][ValidateSet('Difference','Copy')][string]$VHDMethod,
    [Parameter(ParameterSetName="CustomVHD")][Parameter(ParameterSetName="CustomNew")][Parameter(ParameterSetName="CustomVHDDyn")][Parameter(ParameterSetName="CustomNewDyn")][Parameter(ParameterSetName="TemplateNew", Mandatory=$True)][Parameter(ParameterSetName="TemplateVHD", Mandatory=$True)][string]$VMSwitch,
    [Parameter(ParameterSetName="CustomVHD")][Parameter(ParameterSetName="CustomNew")][Parameter(ParameterSetName="CustomVHDDyn")][Parameter(ParameterSetName="CustomNewDyn")][System.Int64]$OSDriveSize=60GB,
    [Parameter(ParameterSetName="CustomNew")][Parameter(ParameterSetName="CustomNewDyn")][switch]$ThickProvision,
    [Parameter(ParameterSetName="CustomVHD")][Parameter(ParameterSetName="CustomNew", Mandatory=$true)][switch]$FixedMemory,
    [Parameter(ParameterSetName="CustomVHD")][Parameter(ParameterSetName="CustomNew")][Parameter(ParameterSetName="CustomVHDDyn")][Parameter(ParameterSetName="CustomNewDyn")][System.Int64]$Memory=1GB,
    [Parameter(ParameterSetName="CustomVHD")][Parameter(ParameterSetName="CustomNew")][Parameter(ParameterSetName="CustomVHDDyn")][Parameter(ParameterSetName="CustomNewDyn")][System.Int64]$CPUCount=1,
    [Parameter(ParameterSetName="CustomVHDDyn")][Parameter(ParameterSetName="CustomNewDyn")][System.Int64]$MaxMemory=2GB,
    [Parameter(ParameterSetName="CustomVHDDyn")][Parameter(ParameterSetName="CustomNewDyn")][System.Int64]$MinMemory=256MB,
    [Parameter(ParameterSetName="CustomVHD")][Parameter(ParameterSetName="CustomNew")][Parameter(ParameterSetName="CustomVHDDyn")][Parameter(ParameterSetName="CustomNewDyn")][ValidateSet(1,2)][int]$generation = 2,
    [Parameter(ParameterSetName="CustomVHD")][Parameter(ParameterSetName="CustomNew")][Parameter(ParameterSetName="CustomVHDDyn")][Parameter(ParameterSetName="CustomNewDyn")][validateSet('on','off')][string]$SecureBoot='on',
    [switch]$Start
)

#region Test actions
    Write-Host "Script Start" -ForegroundColor Green
    Write-Host "Performing Parameter Checks"
    # test if admin
    Write-Verbose "Testing for admin credentials"
    If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
        [Security.Principal.WindowsBuiltInRole] "Administrator")){
        Write-Error "Insufficient permissions" -Category PermissionDenied
        Break
    }else{
        Write-Debug "Admin Credentials confirmed"
    }

    # Test for exisiting VMs
    Write-Verbose "Test for unique VM name: $VMName"
    if(get-vm -Name $VMName -ErrorAction SilentlyContinue){
        write-Error "VM name '$VMName' not unique" -Category ResourceExists
        Break
    }else{
        Write-Debug "VM name $VMName confirmed unique"
    }

    # test for valid switch configuration
    if($VMSwitch){
        Write-Verbose "Testing for valid network switch: $VMSwitch"
        if(get-vmswitch $VMSwitch -ErrorAction SilentlyContinue){
            Write-Debug "Valid switch $VMSwitch confirmed"
        }else{
            Write-Error "Invalid VM switch name specified '$VMSwitch'" -Category ResourceUnavailable
            Break
        }
    }
#endregion

#region Templates
    if($PSCmdlet.ParameterSetName -in 'TemplateVHD','TemplateNew'){
        if($Template -eq 'WindowsClient'){
            $OSDriveSize = 120GB
            $ThickProvision = $False
            $FixedMemory = $False
            $Memory = 1GB
            $MinMemory = 256MB
            $MaxMemory = 2GB
            $generation = 2
            $SecureBoot = 'on'
        }
        if($Template -eq 'WindowsServer'){
            $OSDriveSize = 60GB
            $ThickProvision = $False
            $FixedMemory = $False
            $Memory = 1GB
            $MinMemory = 512MB
            $MaxMemory = 1GB
            $generation = 2
            $SecureBoot = 'on'
        }
        if($Template -eq 'LinuxClient'){
            $OSDriveSize = 60GB
            $ThickProvision = $False
            $FixedMemory = $True
            $Memory = 1GB
            $generation = 2
            $SecureBoot = 'off'
        }
        if($Template -eq 'LinuxServer'){
            $OSDriveSize = 30GB
            $ThickProvision = $False
            $FixedMemory = $True
            $Memory = 512MB
            $generation = 2
            $SecureBoot = 'off'
        }
        if($Template -eq 'LinuxLegacy'){
            $OSDriveSize = 30GB
            $ThickProvision = $False
            $FixedMemory = $True
            $Memory = 512MB
            $generation = 1
            $SecureBoot = 'off'
        }
    }
#endregion

#region Create VHD
    # Create or confirm use of existing VHD
    Write-Host "Creating Resources"
    $newVHD = "$Path\$VMName\Virtual Hard Disks\$VMName-OS.VHDX"
    Write-Verbose "Testing VHD unique: $newVHD"
    if(!(test-path $newVHD)){
        Write-debug "Confirmed VHD does not exist"
        if($ReferenceVHD -and ($VHDMethod -eq 'Difference')){
            Write-Verbose "Creating Differencing VHD $newVHD"
            New-VHD –Path $newVHD –ParentPath “$ReferenceVHD” –Differencing -SizeBytes $OSDriveSize
        }elseif($ReferenceVHD -and ($VHDMethod -eq 'Copy')){
            Write-Verbose "Copying $ReferenceVHD to $newVHD"
            Copy-Item -Path $ReferenceVHD -Destination $newVHD -ErrorAction Stop
        }elseif($ThickProvision){
            Write-Verbose "Creating thick provisioned VHD $newVHD"
            New-VHD -Path $newVHD -SizeBytes $OSDriveSize -Fixed
        }else{
            Write-Verbose "Creating Dynamic VHD $newVHD"
            New-VHD -Path $newVHD -SizeBytes $OSDriveSize -Dynamic
        }
        Write-Debug "New VHD created: $NewVHD"
    }else{
        Write-Warning "VHD already exists, skipping creation and using existing VHD" -WarningAction Inquire
    }
#endregion

#region Create VM
    Write-Host "Creating VM"
    Write-Verbose "Creating VM $VMName"
    if($VMSwitch){
        Write-Verbose "Including connection to switch: $VMSwitch"
        Write-Debug "New-VM -Name $VMName -Generation $generation -MemoryStartupBytes $Memory -VHDPath $newVHD -SwitchName $VMSwitch -Path $Path"
        New-VM -Name $VMName -Generation $generation -MemoryStartupBytes $Memory -VHDPath $newVHD -SwitchName $VMSwitch -Path $Path
    }else{
        Write-Debug "New-VM -Name $VMName -Generation $generation -MemoryStartupBytes $Memory -VHDPath $newVHD -Path $Path"
        New-VM -Name $VMName -Generation $generation -MemoryStartupBytes $Memory -VHDPath $newVHD -Path $Path
    }
#endregion

#region Edit VM Properties
    Write-Host "Configuring VM"
    if($FixedMemory){
        Write-Verbose "Assigning fixed memory limits"
        Set-VMMemory -VMName $VMName -DynamicMemoryEnabled $false
    }else{
        Write-Verbose "Assigning dynamic memory limits, $MinMemory - $MaxMemory"
        Set-VMMemory -VMName $VMName -DynamicMemoryEnabled $true -MaximumBytes $MaxMemory -MinimumBytes $MinMemory
    }
    Write-Debug "Memory limits set"

    # Configure CPU count
    Write-Verbose "Setting CPU Count"
    if($CPUCount -ne 1){
        Set-VMProcessor -VMName $VMName -Count $CPUCount
    }
    Write-Debug "CPU Count set to $CPUCount"

    # Add dvd drive for use with mounting ISO media
    if($ISO){
        Write-Verbose "Adding DVD drive and mounting ISO"
        if($generation -eq 2){
            Write-Debug "Add-VMDvdDrive -VMName $VMName -Path $ISO"
            Add-VMDvdDrive -VMName $VMName -Path $ISO

            Write-Verbose "Adjusting boot order to boot from DVD"
            $dvd=Get-VMDvdDrive -VMName $VMName
            Set-VMFirmware -VMNAME $VMName -FirstBootDevice $dvd
        }elseif($generation -eq 1){
            Write-Debug "Set-VMDvdDrive -VMDvdDrive (Get-VMDvdDrive -VMName $VMName) -Path $ISO"
            Set-VMDvdDrive -VMDvdDrive (Get-VMDvdDrive -VMName $VMName) -Path $ISO
        }
    }

    # Setting secure boot
    if($generation -eq 2){
        Write-Verbose "Setting secure boot option"
        Write-Debug "Setting secureboot $SecureBoot"
        Get-VMFirmware $VMName | Set-VMFirmware –EnableSecureBoot $SecureBoot
    }
#endregion

#region Start VM and open management
    # Open a connection to the VM
    Write-Verbose "Opening connection to the server for management"
    vmconnect localhost $VMName
    
    # Start VM
    if($start){
        Write-host "Starting $VMName"
        Write-Debug "Start-VM -Name $VMName"
        Start-VM -Name $VMName
    }
#endregion

#region Closing arguments
    Write-Host "End of script" -ForegroundColor Green
#endregion
