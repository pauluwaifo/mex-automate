<#
.SYNOPSIS
    Removes MEx Automate from Excel on Windows.

.DESCRIPTION
    Undoes exactly what install.ps1 did: it removes the registry value that
    points Excel at the add-in, and deletes the downloaded manifest from
    %LOCALAPPDATA%\MEx Automate.

    Nothing in your workbooks is touched. Sheets, dashboards and formulas that
    MEx Automate wrote are ordinary Excel content and stay exactly as they are.

.EXAMPLE
    irm https://mex-automate.vercel.app/uninstall.ps1 | iex
#>
[CmdletBinding()]
param(
    [switch] $Quiet
)

$ErrorActionPreference = "Stop"

$AddInId = "fcf30730-1db1-4555-979a-87d048d8c363"
$DeveloperKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$InstallDir = Join-Path $env:LOCALAPPDATA "MEx Automate"

function Say([string] $Message) {
    if (-not $Quiet) { Write-Host $Message }
}

$removedRegistration = $false
$removedFiles = $false

if (Test-Path $DeveloperKey) {
    $existing = (Get-Item $DeveloperKey).GetValueNames()
    if ($existing -contains $AddInId) {
        Remove-ItemProperty -Path $DeveloperKey -Name $AddInId
        $removedRegistration = $true
    }
}

if (Test-Path $InstallDir) {
    Remove-Item -LiteralPath $InstallDir -Recurse -Force
    $removedFiles = $true
}

Say ""
if ($removedRegistration -or $removedFiles) {
    Say "MEx Automate has been removed."
    if ($null -ne (Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        Say "Close and reopen Excel for the Automate button to disappear."
    }
} else {
    Say "MEx Automate was not installed for this user, so there was nothing to remove."
}
Say ""
Say "Your sheets, dashboards and formulas are untouched."
Say ""
