<#
    .SYNOPSIS
        Imports .ttf files into Windows Fonts collection

    .DESCRIPTION
        Creates a new shell application to correctly import font files into the windows font collection from either the local directory or a specified Path

        Window will be displayed showing fonts installing. Not completely transparent when run in the background

    .NOTES
        Author: Simon Baker

        Modified: 2017-09-20
        -Added support for otf files
        -Added comments to code

        Created: 2016-03-16
#>

[CmdletBinding(SupportsShouldProcess=$True)]
Param(
    [string]$Path = "."
)

# Create Shell Application
$sa =  new-object -comobject shell.application

# Define font namespace
$Fonts =  $sa.NameSpace(0x14)

# Get list of fonts to install
$fontlist = Get-ChildItem  "$Path\*.ttf"
$fontlist += Get-ChildItem  "$Path\*.otf"
Write-Verbose $fontlist
write-debug "Font list gathered, ready for install"

# Install fonts
$fontlist | %{$fonts.CopyHere($_.FullName)}