<#
    .SYNOPSIS
        Pings servers from a list and reports reply status

    .DESCRIPTION
        Tests a list of servers using Test-NetConnection and reports if the host is online, offline or unresolved by DNS.

        Can handle a mixed list of IPs and Hostnames

    .PARAMETER ServersTXT
        Path to a text file containing a list of devices to test connection to.
        List can include both IP addresses and Hostnames.
        Each entry should be on a new line.

    .PARAMETER LogFile
        File to write output to.
        Default: .\Test-Connections.csv

    .EXAMPLE
        Test-Connections.ps1 -ServersTXT C:\temp\serverlist.txt -LogFile C:\temp\ServerListStatus.csv

    .NOTES
        
#>
[CmdletBinding()]
Param(
   [Parameter(Mandatory=$True,Position=1)] [string]$ServersTXT,
   [Parameter()][string]$LogFile = ".\Test-Connections.csv"
)

Function Write-Log{
    Param([string]$logstring)
    $logline = $(Get-Date -Format FileDateTime) + "," + $logstring
    Add-content $LogFile -value $logline
    Write-Verbose $logstring
}

Write-Verbose "Getting content from Text file"
$servers = (Get-Content $ServersTXT)

Write-Verbose "Begining Test-NetConnection"
foreach ($server in $servers){
    Write-Debug "Testing server: $server"
    $result = Test-NetConnection -ComputerName $server -WarningAction Ignore -ErrorAction SilentlyContinue
    if($result.PingSucceeded -eq "True"){
        Write-Log "$($result.ComputerName),Success"
    }Else{
        if($result.RemoteAddress -eq $null){
            Write-Log "$($result.ComputerName),Unresolved"
        }else{
            Write-Log "$($result.ComputerName),Offline"
        }
    }
}

Write-Verbose "Script completed"