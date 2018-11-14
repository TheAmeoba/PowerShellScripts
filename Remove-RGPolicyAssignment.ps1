<#
    .SYNOPSIS
        Removes Policy Assignments and custom Definitions scoped to a resource group

    .DESCRIPTION
        Will find policy assignments scoped to a reource group and remove them, 
            along with any custom definitions used by those assignments

        Resource locks should be removed before running script
        Requires azure login before running script (e.g. Login-AzureRMAccount)

    .PARAMETER ResourceGroup
        Name of the resource group that policies are scoped to

    .PARAMETER subscriptionID
        The subscription ID that contains the resource group the assignments are scoped to.
        Can be left blank if current context is suitable

    .PARAMETER auto
        Skips confirmation. required if running headless

    .PARAMETER removeallcustdef
        If specified, script will search for all custom definitions and try to removed them. If they are still in use they will cause an error
        Useful if you've deleted the assignment and want to clean up a test environement

    .EXAMPLE
        remove-rgpolicyassignment.ps1 -ResourceGroup "testrgname" -Verbose

    .NOTES
        Author: Simon Baker
        Created: 2018-11-02
        Modified: 2018-11-08
        Version: 1.3

        Change Log:
            1.1 - Added Try-catch to address mid operation failures
            1.2 - Added function to remove all custom definitions
            1.3 - Bug fix: implemented switch to prompt for removal of all custom policy definitions
#>

[CmdletBinding(SupportsShouldProcess=$True)]
Param(
    [parameter(Mandatory=$true,Position=1)][string]$ResourceGroup,
    [parameter()][string]$subscriptionID,
    [parameter()][switch]$auto,
    [parameter()][switch]$removeallcustdef
)

# Variable definitions
[array]$definitions=@()
if(!($subscriptionID)){
    Write-Verbose "No subcription id provided, using current context..."
    $subscriptionID = (Get-AzureRmContext).Subscription
    Write-Verbose "Subscription ID: $($subscriptionID)"
}
$scopeID = "/subscriptions/" + $subscriptionID + "/resourceGroups/" + $ResourceGroup 

# Main script start
# Removing Policy Assignments
Write-Verbose "Getting policy assignments..."
$policyassignments = Get-AzureRmPolicyAssignment -scope $scopeID
Write-Host "$($policyassignments.Count) policies found"
if(!$auto){pause}
try{
    $policyassignments | ForEach-Object {
        Write-Verbose "Removing Assignment: $($_.Name)"
        Remove-AzureRmPolicyAssignment -Id $_.PolicyAssignmentId
        $definitions += $_.Properties.policyDefinitionId
    }
}catch{
    Write-Error $_
    if($definitions){
        Write-Error "Policy Definitions may be left behind due to script failing before definition step. Definitions below may need manual removal"
        Write-Error $definitions
    }
    Exit 10
}

# Removing Policy Definitions used by the removed policy assignments
Write-Verbose "Getting Custom Definitions used in removed assignments..."
$customdefinitions = Get-AzureRmPolicyDefinition | Where-Object {($_.Properties.policyType -like "Custom") -and ($definitions -contains $_.ResourceId)}
Write-Host "$($customdefinitions.count) custom definitions found"
if(!$auto){pause}
try{
    $customdefinitions | foreach-object {
        Write-Verbose "Removing Definition: $($_.Properties.displayName)"
        Remove-AzureRmPolicyDefinition -Id $_.PolicyDefinitionId -Force
    }
}catch{
    Write-Error $_
    Write-Error "Error during definition removal, not all defintions may have been removed."
    Write-Error "Manual removal may be required for the following definitions:"
    $customdefinitions.PolicyDefinitionId | Write-Error
    Exit 20
}

# attempt to remove all custom definitions
# no try catch here so script will report error for each definition that fails
Write-Verbose "Checking for remaining definitions..."
$remainingDefs = Get-AzureRmPolicyDefinition | Where-Object {($_.Properties.policyType -like "Custom")}
if($remainingDefs -and $removeallcustdef){
    Write-Host "$($remainingDefs.count) remaining custom definitions"
    if(!$auto){$continue = Read-Host "Continue with removal? [yes/no]"}else{$continue = "yes"}
    if($continue -eq "yes"){
        $remainingDefs | ForEach-Object {
            Write-Verbose "Removing Definition: $($_.Properties.displayName)"
            Remove-AzureRmPolicyDefinition -Id $_.PolicyDefinitionId -Force
        }
    }else{
        Write-Verbose "Cancel removal..."
    }
}

Write-Verbose "Script Completed"