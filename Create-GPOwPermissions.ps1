<#
    .SYNOPSIS
        Script to create GPOs, link to OUs and assign required permissions

    .DESCRIPTION
        This script aims to create GPOs from a list and apply common settings to all of them. 
        Linking GPOs to OUs and applying edit and filtering permissions as requried.

        inputs are to be arrays eg:
        "OUName1","OUName2"

    .PARAMETER filtergroupou
        To create groups for filtering, provide the OU in which to create the groups

    .NOTES
        Created By: Simon Baker
        Created On: 2016-07-06
        Modified On: 2016-07-19
        Version: 4

        Script initially developed for client specific requirements and adapted to this generic version
#>

[CmdletBinding(SupportsShouldProcess=$True)]
Param(
    [Parameter(Mandatory=$true)][string]$domain,
    [Parameter(Mandatory=$false)][array]$mgmtusers,
    [Parameter(Mandatory=$false)][array]$linkOUs,
    [Parameter(Mandatory=$true)][array]$newGPOs,
    [Parameter(Mandatory=$false)][array]$filtergroups,
    [Parameter(Mandatory=$false)][string]$filtergroupou
)

#Loop through all GPOs
foreach($newGPO in $newGPOs){
    Write-Verbose "Creating $newGPO"
    $GPO = New-GPO -Name $newGPO -Domain $domain
    #Create permissions for management users
    if($mgmtusers){
        foreach($user in $mgmtusers){
            Set-GPPermissions -Guid $GPO.id -PermissionLevel GpoEdit -TargetName $user -TargetType User
        }
    }
    #Link GPO to OUs
    if($linkOUs){
        foreach($OU in $linkOUs){
            New-GPLink -Guid $GPO.id -Target $OU -LinkEnabled Yes
        }
    }
    #Set filtering groups, creating the group if requested
    if($filtergroups){
        foreach($group in $filtergroups){
            Set-GPPermissions -Guid $GPO.id -PermissionLevel None -TargetName "Authenticated Users" -TargetType Group
            if(Get-ADGroup -Filter {SamAccountName -eq $group}){
                Set-GPPermissions -Guid $GPO.id -PermissionLevel GpoApply -TargetName $group -TargetType Group
            }elseif($creategroups){
                new-adgroup -Name $group -GroupScope Global -GroupCategory Security -Path $filtergroupou
                Set-GPPermissions -Guid $GPO.id -PermissionLevel GpoApply -TargetName $group -TargetType Group
            }else{
                Write-Warning "Unable to add filter permissions for group $group as the group does not exist"
                Write-Verbose "Group $group does not exist, to create the group run script with -filtergroupou `"OU PATH`" specified"
            }
        }
    }
}
Write-Verbose "Script Complete"