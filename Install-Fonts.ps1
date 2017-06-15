<#
    .SYNOPSIS
        Imports .ttf files into Windows Fonts collection

    .DESCRIPTION
        Creates a new shell application to correctly import font files into the windows font collection from either the local directory or a specified Path

        Window will be displayed showing fonts installing. Not completely transparent when run in the background

    .NOTES
        Author: Simon Baker
        Created: 2016-03-16
#>

[CmdletBinding(SupportsShouldProcess=$True)]
Param(
    [string]$Path = "."
)

$sa =  new-object -comobject shell.application
$Fonts =  $sa.NameSpace(0x14)

Get-ChildItem  "$Path\*.ttf" | %{$fonts.CopyHere($_.FullName)}