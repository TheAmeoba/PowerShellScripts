<#
    .SYNOPSIS
        Removes Policy Assignments and custom Definitions scoped to a resource group

    .DESCRIPTION
        Requires azure login before running script (e.g. Login-AzureRMAccount)
        Will find policy assignments scoped to a reource group and remove them, 
            along with any custom definitions used by those assignments

        Resource locks should be removed before running script

    .PARAMETER ResourceGroup
        Name of the resource group that policies are scoped to

    .PARAMETER subscriptionID
        The subscription ID that contains the resource group the assignments are scoped to.
        Can be left blank if current context is suitable

    .PARAMETER auto
        Skips confirmation. required if running headless

    .EXAMPLE
        remove-rgpolicyassignment.ps1 -ResourceGroup "testrgname"

    .NOTES
        Author: Simon Baker
        Created: 2018-11-02
        Modified: 
        Version: 1.0
#>

[CmdletBinding(SupportsShouldProcess=$True)]
Param(
    [parameter(Mandatory=$true,Position=1)][string]$ResourceGroup,
    [parameter()][string]$subscriptionID,
    [parameter()][switch]$auto
)

if(!($subscriptionID)){
    Write-Verbose "No subcription id provided, using current context..."
    $subscriptionID = (Get-AzureRmContext).Subscription
    Write-Verbose "Subscription ID: $($subscriptionID)"
}

$scopeID = "/subscriptions/" + $subscriptionID + "/resourceGroups/" + $ResourceGroup
[array]$definitions=@() 

Write-Verbose "Getting policy assignments..."
$policyassignments = Get-AzureRmPolicyAssignment -scope $scopeID
Write-Host "$($policyassignments.Count) policies found"
if(!$auto){pause}
$policyassignments | ForEach-Object {
    Write-Verbose "Removing Assignment: $($_.Name)"
    $definitions += $_.Properties.policyDefinitionId
    Remove-AzureRmPolicyAssignment -Id $_.PolicyAssignmentId
}

Write-Verbose "Getting Custom Definitions..."
$customdefinitions = Get-AzureRmPolicyDefinition | Where-Object {($_.Properties.policyType -like "Custom") -and ($definitions -contains $_.ResourceId)}
Write-Host "$($customdefinitions.count) custom definitions found"
if(!$auto){pause}
$customdefinitions | foreach-object {
    Write-Verbose "Removing Definition: $($_.Properties.displayName)"
    Remove-AzureRmPolicyDefinition -Id $_.PolicyDefinitionId -Force
}
