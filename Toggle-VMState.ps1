<#
    .SYNOPSIS
        Saves running VMs and starts them up again
#>

[CmdletBinding()]
Param(
    [switch]$save,
    [switch]$restore
)


####test if admin####
Write-Verbose "Test   - Admin credentials"
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")){
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    Break
}else{
    Write-Verbose "Result - Admin Credentials confirmed"
}

#####Functions#####

Write-Verbose "Test   - Determine exisitance of CSV"
$CSVExistance = Test-Path $env:SystemDrive\temp\RunningVMs.csv
Write-Verbose "Result - CSV exists: $CSVExistance"

Write-Verbose "Action - Define Functions"
Function RestoreVMs {
    Write-Verbose "Action - Restoring VMs"
    Import-Csv $env:systemdrive\temp\RunningVMs.csv | ForEach-Object {Start-VM -Name $_.VMName}
    Write-Verbose "Action - VMs restored, cleaning up"
    Remove-Item -Path $env:systemdrive\temp\RunningVMs.csv
}

Function SaveVMs {
    Write-Verbose "Action - Gathering list of VMs"
    $RunningVMs = Get-VM | Where{$_.State -eq 'Running'}
    Write-Verbose "Test   - Either creating new or appending to CSV for use in restore"
    if($CSVExistance){
        Write-Debug "Result - Add-Content to CSV"
        $RunningVMs | Add-Content -Path $env:SystemDrive\temp\RunningVMs.csv
    }else{
        Write-Debug "Result - Export-CSV to create CSV"
        $RunningVMs | Export-CSV -Path $env:SystemDrive\temp\RunningVMs.csv
    }
    Write-Verbose "Action - Run Save-VM for each object"
    $RunningVMs | ForEach-Object {save-VM -Name $_.VMName}
}

#####Start of Script#####

Write-Verbose "Test   - Check input switches"
if($save -and $restore){
    Write-debut "Result - Both switches set, throwing error and stopping script"
    Write-Error "Invalid input parametters"
    Return
}

if(!($save -or $restore)){
    Write-Verbose "Result - Neither switch found"
    Write-Verbose "Test   - Determining state of VMs"
    If ($CSVExistance){
        Write-Verbose "Result - Exisiting list present"
        Write-Debug "Calling function RestoreVms"
        RestoreVMs
    }else{
        Write-Verbose "Result - No list presnent, saving running VMs"
        Write-Debug "Calling function SaveVMs"
        SaveVMs
    }
}


if($save -and !$restore){
    Write-Verbose "Result - Save param set"
    Write-Debug "Calling SaveVMs"
    SaveVMs
}

if($restore -and !$save){
    Write-Verbose "Result - Restore Param set"
    Write-Debug "Calling RestoreVMs"
    RestoreVMs
}


Write-Host "   _"
Write-Host " ( (("
Write-Host "  \ =\"
Write-Host " __\_ `-\ "
Write-Host "(____))(  \----"
Write-Host "(____)) _  "
Write-Host "(____))"
Write-Host "(____))____/---- "
Write-Host ""