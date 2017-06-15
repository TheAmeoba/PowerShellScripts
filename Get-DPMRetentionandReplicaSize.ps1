<#
    .SYNOPSIS
        Gathers size and retention info from DPM Protection Groups

    .DESCRIPTION
        Simple report to get size and retention information from DPM Protection Groups.

        Will need to run from DPM management console or in a powershell console with correct modules loaded. 
        No checks have been implemented as yet to ensure this is the case.

    .LINK
        
    .PARAMETER DPMServer
        DNS name or IP of the DPM server to connect to. Optional is server is localhost

    .NOTES
        Version 0.1
        Author: Simon Baker
        Date: 2017-04-28
#>

[CmdletBinding(SupportsShouldProcess=$True)]
Param(
    [parameter(Mandatory=$false)]$DPMServer
)

Write-Warning "Please ensure script running from DPM console, script will not be auto checking" -WarningAction Inquire

#Initialize totals tally
$DatasourceTotals = New-Object PSObject -Property @{DatasourceSize=0;ReplicaSize=0;ShadowCopyAreaSize=0;ReplicaUsedSpace=0;ShadowCopyUsedSpace=0}

#Get all protection group info
if(!($DPMServer)){
    $GroupList=Get-ProtectionGroup 
}else{
    $GroupList=GET-PROTECTIONGROUP -DPMSERVERNAME $DPMServer
}

#Loop through all the protection groups gathering the retention info
foreach ( $Group in $GroupList ) {
    Write-Host "----------------------------------------------"
    Write-Host "Protection Group: "$Group.Name
    Write-Host "Disk Retention:"
    Get-DPMPolicyObjective -ProtectionGroup $Group -ShortTerm | fl
    Write-Host "Virtual Tape Retention:"
    Get-DPMPolicyObjective -ProtectionGroup $Group -LongTerm Tape | select RecoveryRange,Frequency,Schedules
    #get datasrouce sizes
    Write-Host "Data Source Size:"
    $Datasources=Get-DPMDatasource -ProtectionGroup $Group | Select Computer,Name,DatasourceSize,ReplicaSize,ShadowCopyAreaSize,ReplicaUsedSpace,ShadowCopyUsedSpace
    $Datasources | ft
    #add datasource sizes to final total
    foreach($Datasource in $Datasources){
        $DatasourceTotals.DatasourceSize += $Datasource.DatasourceSize
        $DatasourceTotals.ReplicaSize += $Datasource.ReplicaSize
        $DatasourceTotals.ShadowCopyAreaSize += $Datasource.ShadowCopyAreaSize
        $DatasourceTotals.ReplicaUsedSpace += $Datasource.ReplicaUsedSpace
        $DatasourceTotals.ShadowCopyUsedSpace += $Datasource.ShadowCopyUsedSpace
    }
}
Write-Host "----------------------------------------------"
Write-Host "----------------------------------------------"
Write-Host "Total Values (in GB):"
#Divite totals to display in GB
$DatasourceTotals.DatasourceSize = $DatasourceTotals.DatasourceSize /1gb
$DatasourceTotals.ReplicaSize += $DatasourceTotals.ReplicaSize /1gb
$DatasourceTotals.ShadowCopyAreaSize += $DatasourceTotals.ShadowCopyAreaSize /1gb
$DatasourceTotals.ReplicaUsedSpace += $DatasourceTotals.ReplicaUsedSpace /1gb
$DatasourceTotals.ShadowCopyUsedSpace += $DatasourceTotals.ShadowCopyUsedSpace /1gb
$DatasourceTotals | ft
