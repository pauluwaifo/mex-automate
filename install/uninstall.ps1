<#
.SYNOPSIS
    Removes MEx Automate from Excel and PowerPoint on Windows.

.DESCRIPTION
    Undoes exactly what install.ps1 did: it removes the registry values that
    point Office at the add-ins, and deletes the downloaded manifests from
    %LOCALAPPDATA%\MEx Automate.

    Nothing in your files is touched. Sheets, dashboards, formulas and slides
    that MEx Automate wrote are ordinary Excel and PowerPoint content and stay
    exactly as they are.

.EXAMPLE
    irm https://mex-automate.vercel.app/uninstall.ps1 | iex
#>
[CmdletBinding()]
param(
    [switch] $Quiet
)

$ErrorActionPreference = "Stop"

$AddInIds = @(
    "fcf30730-1db1-4555-979a-87d048d8c363",  # Excel
    "6295c215-900a-4867-85a9-2a9293d10f98"   # PowerPoint
)
$DeveloperKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$InstallDir = Join-Path $env:LOCALAPPDATA "MEx Automate"

function Say([string] $Message) {
    if (-not $Quiet) { Write-Host $Message }
}

$removedRegistrations = 0
$removedFiles = $false

if (Test-Path $DeveloperKey) {
    $existing = (Get-Item $DeveloperKey).GetValueNames()
    foreach ($id in $AddInIds) {
        if ($existing -contains $id) {
            Remove-ItemProperty -Path $DeveloperKey -Name $id
            $removedRegistrations += 1
        }
    }
}

if (Test-Path $InstallDir) {
    Remove-Item -LiteralPath $InstallDir -Recurse -Force
    $removedFiles = $true
}

Say ""
if ($removedRegistrations -gt 0 -or $removedFiles) {
    Say "MEx Automate has been removed."
    if ((Get-Process EXCEL -ErrorAction SilentlyContinue) -or (Get-Process POWERPNT -ErrorAction SilentlyContinue)) {
        Say "Close and reopen Excel and PowerPoint for the buttons to disappear."
    }
} else {
    Say "MEx Automate was not installed for this user, so there was nothing to remove."
}
Say ""
Say "Your sheets, dashboards, formulas and slides are untouched."
Say ""
