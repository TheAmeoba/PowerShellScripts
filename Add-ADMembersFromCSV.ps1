<#
    .SYNOPSIS
        Adds users listed in a CSV file to an AD Group

    .DESCRIPTION
        Build to allow for bulk adding of users to an Active Directory group.
        Original code copied and modified for scripted use from http://powershell.com/cs/forums/t/17142.aspx

    .LINK
        http://powershell.com/cs/forums/t/17142.aspx

    .EXAMPLE
        Add-ADMembersFromCSV.ps1 -ADGroup Citrix_Users -UsersCSV C:\folder\users.csv -Domain mydomain.local
        
        This command adds users listed in users.csv to the group Citrix_Users in mydomain.local
    
    .PARAMETER UsersCSV
        Formating of the CSV should include coloumn headings on the first line and users listed on following lines.
        The heading "SAMAccountName" is required to be present in  CSV.
        Simple example with 3 users to add:

        SAMAccountName
        user1
        user2
        user3

    .NOTES
        Author: Simon Baker
        Date Created: 2015-12-04
        Last Modified by: Simon Baker
        Last Modified: 2015-12-04

        __Dev Notes__
        *Error with detection of existing users in powershell v2, users are added to fail log instead of existing log
#>

[CmdletBinding(SupportsShouldProcess=$True)]
Param(
   [Parameter(Mandatory=$True,Position=1)] [string]$ADGroup,
   [Parameter(Mandatory=$True,Position=2)] [string]$UsersCSV,
   [Parameter(Mandatory=$True,Position=3)] [string]$Domain,
   [string]$LogDir = ".\Add-ADMembersFromCSV"
)

Write-Verbose "Importing Active Directory Module"
Import-Module ActiveDirectory

$FailLog = "$LogDir\Faillog.csv"
$AddedLog = "$LogDir\AddedLog.csv"
$ExistingLog = "$LogDir\ExistingLog.csv"
$ErrorLog = "$LogDir\ErrorLog.csv"

Write-Verbose "Get existing group members"
$Members = Get-ADGroupMember -Identity $ADGroup -server $domain
Write-Verbose "Import list of users from CSV"
$Users = Import-Csv $UsersCSV

$Failures = @()
$Added = @()
$Existing = @()

Write-Verbose "Begining to add users"
foreach ($User in $Users){
    if ($Members.SAMAccountName -contains $User.SAMAccountName ){
        Write-Verbose "User $User already exists in group"
        $Existing += $User
    } else {
        try {
            Add-ADGroupMember -Identity $ADGroup -Members $User.SAMAccountName -server $domain
            Write-Verbose "Added $User to $ADGroup"
            $Added += $User
        } catch { 
            Write-Verbose "Failed to add $User to $ADGroup"
            $Failures += $User 
        }
    }
}

if ($Failures){
    Write-Verbose "Creating Failure Logs"
    $Failures | Export-Csv $FailLog -NoTypeInformation
    $Error | export-csv $ErrorLog -NoTypeInformation
}

if ($Added){
    Write-Verbose "Creating added users log"
    $Added | Export-Csv $AddedLog -NoTypeInformation
}

if ($Existing){
    Write-Verbose "Creating Existing users log"
    $Existing | Export-Csv $ExistingLog -NoTypeInformation
}

Write-Verbose "Script completed"