<#
    .SYNOPSIS
        Pings servers from a list and reports reply status

    .NOTES
        Legacy - superseeded by script Test-Connections.ps1
#>
[CmdletBinding()]
Param(
   [Parameter(Mandatory=$True,Position=1)] [string]$ServersTXT,
   [string]$LogDir = ".\Test-PingResponse-Logs"
)

$servers = (Get-Content $ServersTXT)
$collection = $()
foreach ($server in $servers)
{
    $status = @{ "ServerName" = $server}
    if (Test-Connection $server -Count 1 -ea 0 -Quiet)
    {
        $status["Results"] = "Up"
    }
    else
    {
        $status["Results"] = "Down"
    }
    New-Object -TypeName PSObject -Property $status -OutVariable serverStatus
    $collection += $serverStatus

}
$collection | Export-Csv -LiteralPath $LogDir\ServerStatus.csv -NoTypeInformation
