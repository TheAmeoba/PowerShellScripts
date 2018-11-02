<#
    .SYNOPSIS
        Starts IE in Full Screen Mode

    .DESCRIPTION
        Starts IE, sets window to full screen and loads specified URL instead of homepage

    .PARAMETER URL
        Webpage to load on start
        Default: www.google.com

    .NOTES
        Created By: Simon Baker
#>

[CmdletBinding(SupportsShouldProcess=$True)]
Param(
    [parameter(Mandatory=$false)][string]$URL="www.google.com"
)

$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible = $true
$ie.Navigate($URL)

$sw = @'
[DllImport("user32.dll")]
public static extern int ShowWindow(int hwnd, int nCmdShow);
'@

$type = Add-Type -Name ShowWindow2 -MemberDefinition $sw -Language CSharpVersion3 -Namespace Utils -PassThru
$type::ShowWindow($ie.hwnd, 3) # 3 = maximize 