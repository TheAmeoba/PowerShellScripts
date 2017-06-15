<#
    .SYNOPSIS
        Generates MS Excel report on the current GPO settings for a domain

    .DESCRIPTION
        The Get-GPOExcelReport script gathers GPO settings and compiles the useful information into an excel spread sheet for analysis and project planning.

        Settings can be gathered directly from the domain if the management Cmdlets are installed or they can be provided in XML format previously exporting using Get-GPOReport.
        MS Excel is required on the machine running this script in order to write the excel file for export.

    .LINK
        
    .PARAMETER outputxlsx
        Location and name to save the generated Excel report

    .PARAMETER outputxml
        Location and name of the XML file to save domain settings scanned from domain

    .PARAMETER inputxml
        Location and name of XML file containing settings previously generated from Get-GPOReport

    .PARAMETER domain
        MS AD Domain to query for GPO settings

    .PARAMETER moredetail
        Includes extra details that are normally not required for an overview of a GPOs functions. E.g. Every parameter for a folder redirection, no limited to just the path and security group applied

    .EXAMPLE
        Get-GPOExcelReport.ps1 -domain contoso.com -outputxml C:\temp\GPOSettings.xml -outputxlsx C:\temp\GPOSEttings.xlsx

        This command gets GPO settings from contoso.com and saves them to the provided xml file. XML is then formatted and saved to the XLSX file.

    .NOTES
        Version 3.0
        Author: Simon Baker
        Date: 2017-05-16

        ToDo List:
            - Software restriction trusted publishers
            - Rule collections - Review
            - System Services
#>
[CmdletBinding(SupportsShouldProcess=$True)]
Param(
    [parameter(ParameterSetName='ProvidedXML',Mandatory=$true)]$inputxml,
    [parameter(Mandatory=$true)]$outputxlsx,
    [parameter(ParameterSetName='QueryDomain',Mandatory=$true)]$domain,
    [parameter(ParameterSetName='QueryDomain',Mandatory=$true)]$outputxml,
    [parameter(Mandatory=$false)]$moredetail
)
#region Import Data
    #Check for parameter set and either query domain or import XML
    if($PSCmdlet.ParameterSetName -eq "QueryDomain"){
        Write-Verbose "Quering domain for group policy settings"
        $oldPreference = $ErrorActionPreference
        $ErrorActionPreference = ‘stop’
        try {if(Get-Command Get-GPOReport){
            Get-GPOReport -all -ReportType Xml -Path $outputxml
        }}catch{
            Write-Warning "Group policy management tools not found on this system."
            write-verbose "Please run from a machine with Group Policy tools or otherwise aquire an XML report (Get-GPOReport) and provide it to this script."
            Exit
        }Finally{$ErrorActionPreference=$oldPreference}
        [xml]$xmlreport = Get-Content -Path $outputxml
    }elseif($PSCmdlet.ParameterSetName -eq "ProvidedXML"){
        Write-Verbose "Importing XML"
        [xml]$xmlreport = Get-Content -Path $inputxml
    }
#endregion 
#region spreadsheet check
    #Check we can create the spreadsheet before going through all the calculations
    Write-Verbose "Checking for Excel"
    $excelapp = $null
    try{
        $excelapp = new-object -comobject Excel.Application
    }
    catch{
        Write-Warning "Excel not installed on this system. Please take the generated XML and run this script from a machine with Excel installed."
        exit
    }
#endregion
#region Looping through the GPOs
    #create array variables for storing data before writing to excel
    [array]$GPOLinks = @()
    [array]$GPOPolicies = @()
    [array]$GPOPermissions = @()
    #Begin collecting and formatting data into arrays
    Write-Verbose "Reading GPO Settings"
    write-debug "XML imported to xmlreport variable"
    foreach($gpo in $xmlreport.report.gpo){
        Write-Verbose "GPO: $($gpo.Name)"
        #Create list of GPO and the corresponding Links
        If($gpo.LinksTo.count -eq 0){
            $GPOLinks += New-Object psobject -Property @{"GPO"=$gpo.Name;"LinksToPath"="NO_LINKS_FOUND"}
        }else{
            foreach($link in $gpo.LinksTo){
                $GPOLinks += New-Object psobject -Property @{"GPO"=$gpo.Name;"LinksToPath"=$link.SOMPath}
            }
        }
        write-debug "gpo loop"
        #Loop through Computer settings then Users settings
        $gposcopes = $gpo | gm -MemberType Property | Where{$_.Name -in "Computer","User"}
        foreach($scope in $gposcopes.name){
            Write-debug "inside member loop"
            foreach($ExtData in $gpo.$($scope).ExtensionData){
                #Create list of GPO Policy Settings
                foreach($policy in $ExtData.Extension.policy){
                    if($policy.Name){
                        if($policy.edittext){
                            foreach($setting in $policy.edittext){
                                $GPOPolicies += New-Object psobject -Property @{"GPO"=$gpo.Name;"Scope"=$scope;"PolicyName"=$policy.Name;"State"=$policy.State;"Category"=$policy.Category;`
                                    "Setting"=$setting.Name;"SettingType"="edittext";"SettingState"=$setting.state;"SettingValue"=$setting.value
                                }
                            }
                        }
                        if($policy.DropDownList){
                            foreach($setting in $policy.DropDownList){
                                $GPOPolicies += New-Object psobject -Property @{"GPO"=$gpo.Name;"Scope"=$scope;"PolicyName"=$policy.Name;"State"=$policy.State;"Category"=$policy.Category;`
                                    "Setting"=$setting.Name;"SettingType"="DropDownList";"SettingState"=$setting.state;"SettingValue"=$setting.value.Name
                                }
                            }
                        }
                        if($policy.CheckBox){
                            foreach($setting in $policy.CheckBox){
                                $GPOPolicies += New-Object psobject -Property @{"GPO"=$gpo.Name;"Scope"=$scope;"PolicyName"=$policy.Name;"State"=$policy.State;"Category"=$policy.Category;`
                                    "Setting"=$setting.Name;"SettingType"="CheckBox";"SettingState"=$setting.state;"SettingValue"=""
                                }
                            }
                        }
                        if(!($policy.EditText -or $policy.DropDownList -or $policy.CheckBox)){
                        
                            $GPOPolicies += New-Object psobject -Property @{`
                                "GPO"=$gpo.Name;`
                                "Scope"=$scope;`
                                "PolicyName"=$policy.Name;`
                                "State"=$policy.State;`
                                "Category"=$policy.Category;`
                                "Setting"="";`
                                "SettingType"="";`
                                "SettingState"="";`
                                "SettingValue"=""
                            }
                        }
                    }
                }
        
                #Create list of GPO registry changes
                foreach($registrysetting in $ExtData.Extension.RegistrySettings){
                    foreach($registry in $registrysetting.Registry){
                        $GPOPolicies += New-Object psobject -Property @{`
                            "GPO"=$gpo.Name;`
                            "Scope"=$scope;`
                            "PolicyName"="Registry Settings";`
                            "State"="Enabled";`
                            "Category"="Registry";`
                            "Setting"="$($registry.Properties.hive)\$($registry.Properties.key)\$($registry.Properties.name)";`
                            "SettingType"=$registry.Properties.type;`
                            "SettingState"=$registry.Properties.action;`
                            "SettingValue"=$registry.Properties.value
                        }
                    }
                }

                #Create list of GPO Software restriction policies
                foreach($generalsoftwarerestrictionsetting in $ExtData.Extension.General){
                    if($generalsoftwarerestrictionsetting.ApplicableBinaries){
                        $GPOPolicies += New-Object psobject -Property @{`
                            "GPO"=$gpo.Name;`
                            "Scope"=$scope;`
                            "PolicyName"="Software Restriction Settings";`
                            "State"="Enabled";`
                            "Category"="Software Restiction";`
                            "Setting"="Applicable Binaries";`
                            "SettingType"="string";`
                            "SettingState"="Enabled";`
                            "SettingValue"=$generalsoftwarerestrictionsetting.ApplicableBinaries
                        }
                        $GPOPolicies += New-Object psobject -Property @{`
                            "GPO"=$gpo.Name;`
                            "Scope"=$scope;`
                            "PolicyName"="Software Restriction Settings";`
                            "State"="Enabled";`
                            "Category"="Software Restiction";`
                            "Setting"="Applicable Users";`
                            "SettingType"="string";`
                            "SettingState"="Enabled";`
                            "SettingValue"=$generalsoftwarerestrictionsetting.ApplicableUsers
                        }
                        $GPOPolicies += New-Object psobject -Property @{`
                            "GPO"=$gpo.Name;`
                            "Scope"=$scope;`
                            "PolicyName"="Software Restriction Settings";`
                            "State"="Enabled";`
                            "Category"="Software Restiction";`
                            "Setting"="Certificate Rules Enabled";`
                            "SettingType"="string";`
                            "SettingState"="Enabled";`
                            "SettingValue"=$generalsoftwarerestrictionsetting.CertificateRulesEnabled
                        }
                        [string]$FileTypes = ""
                        foreach($file in $generalsoftwarerestrictionsetting.ExecutableFiles.FileType){
                            $FileTypes += $file
                            $FileTypes += ", "
                        }
                        $GPOPolicies += New-Object psobject -Property @{`
                            "GPO"=$gpo.Name;`
                            "Scope"=$scope;`
                            "PolicyName"="Software Restriction Settings";`
                            "State"="Enabled";`
                            "Category"="Software Restiction";`
                            "Setting"="Executable File Types";`
                            "SettingType"="list";`
                            "SettingState"="Enabled";`
                            "SettingValue"=$FileTypes
                        }
                    }
                }
                foreach($pathrule in $ExtData.Extension.PathRule){
                    if($pathrule.SecurityLevel){
                        $GPOPolicies += New-Object psobject -Property @{`
                            "GPO"=$gpo.Name;`
                            "Scope"=$scope;`
                            "PolicyName"="Software Restriction Path Rule";`
                            "State"="Enabled";`
                            "Category"="Software Restiction";`
                            "Setting"=$pathrule.Path;`
                            "SettingType"="PathRule";`
                            "SettingState"="Enabled";`
                            "SettingValue"=$pathrule.SecurityLevel
                        }
                    }
                }
        
                #Create list of GPO Windows Settings\Account settings
                foreach($account in $ExtData.Extension.Account){
                    foreach($setting in $account){
                        $settingvalue = $setting | Get-member -MemberType Property | Where{$_.Name -like "Setting*"}
                        $GPOPolicies += New-Object psobject -Property @{"GPO"=$gpo.Name;"Scope"=$scope;"PolicyName"=$setting.Name;`
                            "State"="Enabled";"Category"=$setting.Type;`
                            "Setting"=$setting.Name;"SettingType"=$settingvalue.Name;"SettingState"="Enabled";`
                            "SettingValue"=$setting.($settingvalue.Name)
                        }
                    }
                }
            
                #Create list of GPO Folder redirection Policies
                foreach($folder in $ExtData.Extension.Folder){
                    if($folder.Location.DestinationPath){
                        write-debug "folder debug point"
                        $GPOPolicies += New-Object psobject -Property @{`
                            "GPO"=$gpo.Name;`
                            "Scope"=$scope;`
                            "PolicyName"=$folder.id;`
                            "State"="Enabled";`
                            "Category"="Folder Redirection";`
                            "Setting"="Destination Path";`
                            "SettingType"="Share Path";`
                            "SettingState"="Enabled";`
                            "SettingValue"=$folder.Location.DestinationPath
                        }
                        $GPOPolicies += New-Object psobject -Property @{`
                            "GPO"=$gpo.Name;`
                            "Scope"=$scope;`
                            "PolicyName"=$folder.id;`
                            "State"="Enabled";`
                            "Category"="Folder Redirection";`
                            "Setting"="Security Group";`
                            "SettingType"="Secuirty Group";`
                            "SettingState"="Enabled";`
                            "SettingValue"=$folder.Location.SecurityGroup.Name.'#text'
                        }
                        if($moredetail){
                            $folderSettings = $folder | gm -MemberType Property | Where{$_.Name -in "GrantExclusiveRights","MoveContents","FollowParent","ApplyToDownLevel","DoNotCare","RedirectToLocal","PolicyRemovalBehavior","ConfigurationControl","PrimaryComputerEvaluation"}
                            foreach($setting in $folderSettings.Name){
                                $GPOPolicies += New-Object psobject -Property @{`
                                    "GPO"=$gpo.Name;`
                                    "Scope"=$scope;`
                                    "PolicyName"=$folder.id;`
                                    "State"="Enabled";`
                                    "Category"="Folder Redirection";`
                                    "Setting"=$setting;`
                                    "SettingType"="string";`
                                    "SettingState"="";`
                                    "SettingValue"=$folder.$($setting)
                                }
                            }
                        }
                    }
                }

                #Create list of GPO Rule Collections
                foreach($RuleCollection in $ExtData.Extension.RuleCollection){
                    write-debug "Rule Collection Loop"
                    $rules = $RuleCollection | gm -MemberType Property | where{$_.Name -in "FilePublisherRule","FilePathRule"}
                    foreach($rule in $RuleCollection.$($rules.Name)){
                        $GPOPolicies += New-Object psobject -Property @{`
                            "GPO"=$gpo.Name;`
                            "Scope"=$scope;`
                            "PolicyName"=$rule.Name;`
                            "State"="Enabled";`
                            "Category"="App Control: $($RuleCollection.Type)";`
                            "Setting"="Action";`
                            "SettingType"="string";`
                            "SettingState"="";`
                            "SettingValue"=$rule.Action
                        }
                        $GPOPolicies += New-Object psobject -Property @{`
                            "GPO"=$gpo.Name;`
                            "Scope"=$scope;`
                            "PolicyName"=$rule.Name;`
                            "State"="Enabled";`
                            "Category"="App Control: $($RuleCollection.Type)";`
                            "Setting"="UserOrGroupSID";`
                            "SettingType"="string";`
                            "SettingState"="";`
                            "SettingValue"=$rule.UserOrGroupSid
                        }
                        foreach($condition in $rule.Conditions.FilePublisher){
                            $GPOPolicies += New-Object psobject -Property @{`
                                "GPO"=$gpo.Name;`
                                "Scope"=$scope;`
                                "PolicyName"=$rule.Name;`
                                "State"="Enabled";`
                                "Category"="App Control: $($RuleCollection.Type)";`
                                "Setting"="Publisher Name";`
                                "SettingType"="string";`
                                "SettingState"="";`
                                "SettingValue"=$condition.PublisherName
                            }
                        }
                        foreach($condition in $rule.Conditions.FilePath){
                            $GPOPolicies += New-Object psobject -Property @{`
                                "GPO"=$gpo.Name;`
                                "Scope"=$scope;`
                                "PolicyName"=$rule.Name;`
                                "State"="Enabled";`
                                "Category"="App Control: $($RuleCollection.Type)";`
                                "Setting"="File Path";`
                                "SettingType"="string";`
                                "SettingState"="";`
                                "SettingValue"=$condition.Path
                            }
                        }
                    }
                }

                #Create list of GPO System services
                foreach($SysService in $ExtData.Extension.SystemServices){
                    $GPOPolicies += New-Object psobject -Property @{`
                        "GPO"=$gpo.Name;`
                        "Scope"=$scope;`
                        "PolicyName"=$SysService.Name;`
                        "State"="Enabled";`
                        "Category"="System Services";`
                        "Setting"="Startup Mode";`
                        "SettingType"="string";`
                        "SettingState"="";`
                        "SettingValue"=$SysService.StartupMode
                    }
                }
            }
        }   
        #Create list of GPO Permissions 
        foreach($trustee in $gpo.SecurityDescriptor.Permissions.TrusteePermissions){
            $GPOPermissions += New-Object psobject -Property @{"GPO"=$gpo.Name;"Owner"=$Gpo.SecurityDescriptor.Owner.Name.'#text';"Group"=$Gpo.SecurityDescriptor.group.Name.'#text';`
            "Trustee"=$trustee.Trustee.Name.'#text';"type"=$trustee.Type.PermissionType;"Access"=$trustee.Standard.GPOGroupedAccessEnum}
        }
    }
#endregion
#region Exporting to XLSX
    Write-Verbose "Crating XLSX"
    Write-debug "Break Point: About to write excel"
    #Asociate data with sheet names
    $ExcelSheets = @()
    $ExcelSheets += New-Object psobject -Property @{"SheetName"="GPO Permissions";"TableData"=$GPOPermissions}
    $ExcelSheets += New-Object psobject -Property @{"SheetName"="GPO Policies";"TableData"=$GPOPolicies}
    $ExcelSheets += New-Object psobject -Property @{"SheetName"="GPO Links";"TableData"=$GPOLinks}
    #Create excel sheets
    $excelapp.sheetsInNewWorkbook = $ExcelSheets.count
    $xlsx = $excelapp.Workbooks.Add()
    #Set start point for writing to excel
    $sheet=1
    $row=1
    $column=1
    #Enter data into excel
    foreach($table in $ExcelSheets){
        $worksheet = $xlsx.Worksheets.Item($sheet)
        #Name the sheet
        $worksheet.Name = $table.SheetName
        #Create sheet headings
        foreach($item in ($table.TableData|Get-Member -MemberType 'NoteProperty'|Select-Object -ExpandProperty 'Name')){
            $worksheet.Cells.Item($row,$column)=$item
            $column++
        }
        $row++
        $column=1
        #Enter row data
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
    #Save the xlsx
    $xlsx.SaveAs($outputxlsx)
    $excelapp.quit()
    Write-Verbose "End of script"
#endregion