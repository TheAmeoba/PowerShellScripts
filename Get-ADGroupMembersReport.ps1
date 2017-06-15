<#
    .SYNOPSIS
        Produces report on all AD group members

    .DESCRIPTION
        This script was developed to allow for easy viewing of group members from AD. 

        Results can be output to CSV, Excel table, Straight to the console or in Powershell's outgridview

        Requires AD Powershell tools to be installed in the context this script is run.

    .EXAMPLE
        Get-ADGroupMembersReport.ps1 -SearchBase "OU=groups,DC=testdomain,DC=com"

    .EXAMPLE
        Get-ADGroupMembersReport.ps1 -SearchBase "OU=groups,DC=testdomain,DC=com" -output excel -outputfile "C:\temp\output.xlsx"

    .NOTES
        Version: 1.1
        Author: Simon Baker
        Date: 2017-05-31
#>
[CmdletBinding(SupportsShouldProcess=$True)]
Param(
    [parameter(Mandatory=$true)][string]$SearchBase,
    [parameter(Mandatory=$false)][string]$outputfile,
    [parameter(Mandatory=$false)][ValidateSet('excel','csv','console','ogv')][string]$output="ogv"
)
#region Setup
    # Set up variables
    [array]$GroupMembership = @()
    # Check for valid inputs
    if(($output -in "csv","excel")){
        if(!($outputfile)){
            Write-Error "No output file defined" -Category InvalidArgument
        }
    }
    # Check for correct cmdlets
    Write-Verbose "Performing system checks"
    if(!(Get-Command Get-ADGroup -CommandType Cmdlet -errorAction SilentlyContinue)){
        Write-Verbose "The command Get-ADGroup and Get-ADGroupMember are required for the function of this script"
        Write-Error "AD Cmdlets not found" -Category NotInstalled 
        Return
    }
    # Gather group data
    Write-Verbose "Getting group data"
    $Groups = Get-ADGroup -Properties * -Filter * -SearchBase $SearchBase 
    Write-debug "Groups gathered"
#endregion

#region Data Loop
    Write-Verbose "Getting member data"
    # loop through each group
    Foreach($G In $Groups)
    {
        # Get details of group members
        $Members = Get-ADGroupMember $G
        Write-debug "Members gathered from group $($G.name)"
        foreach($M in $Members){
            $GroupMembership += New-Object psobject -Property @{"GroupName"=$G.Name;"User"=$M.Name;"LoginName"=$M.SamAccountName;"DistinguishedName"=$M.distinguishedName}
            Write-Debug "Writen object for member $($M.Name)"
        }
    }
#endregion

#region Output
If($output -eq "csv"){
    #region Export to csv
        Write-Verbose "Exporting to CSV"
        $GroupMembership | Export-Csv -Path $outputfile
    #endregion
}elseif($output -eq "excel"){
    #region spreadsheet check
        #Check we can create the spreadsheet
        Write-Verbose "Checking for Excel"
        $excelapp = $null
        try{
            $excelapp = new-object -comobject Excel.Application
        }
        catch{
            Write-Error "Excel not installed on this system" -Category NotInstalled
            return
        }
    #endregion

    #region Exporting to XLSX
        Write-Verbose "Crating XLSX"
        Write-debug "Break Point: About to write excel"
        # Asociate data with sheet names
        $ExcelSheets = @()
        $ExcelSheets += New-Object psobject -Property @{"SheetName"="Group Membership";"TableData"=$GroupMembership}

        # Create excel sheets
        $excelapp.sheetsInNewWorkbook = $ExcelSheets.count
        $xlsx = $excelapp.Workbooks.Add()
        # Set start point for writing to excel
        $sheet=1
        $row=1
        $column=1
        # Enter data into excel
        foreach($table in $ExcelSheets){
            $worksheet = $xlsx.Worksheets.Item($sheet)
            # Name the sheet
            $worksheet.Name = $table.SheetName
            # Create sheet headings
            foreach($item in ($table.TableData|Get-Member -MemberType 'NoteProperty'|Select-Object -ExpandProperty 'Name')){
                $worksheet.Cells.Item($row,$column)=$item
                $column++
            }
            $row++
            $column=1
            # Enter row data
            foreach($object in $table.TableData){
                foreach($item in ($table.TableData|Get-Member -MemberType 'NoteProperty'|Select-Object -ExpandProperty 'Name')){
                    $worksheet.Cells.Item($row,$column)=$object.$item
                    $column++
                }
                $row++
                $column=1
            }
            $sheet++
            $row=1
        }
        # Save the xlsx
        $xlsx.SaveAs($outputxlsx)
        $excelapp.quit()
    #endregion
}elseif($output -eq "console"){
    Return $GroupMembership
}elseif($output -eq "ogv"){
    $GroupMembership | ogv
}
#endregion

Write-Verbose "End of script"