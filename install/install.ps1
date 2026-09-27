<#
.SYNOPSIS
    Installs MEx Automate into Excel and PowerPoint on Windows.

.DESCRIPTION
    Downloads the add-ins' manifests to your user profile and registers them, so
    an "Automate" button appears on Excel's Home tab and a "Build deck" button
    on PowerPoint's.

    Nothing is installed machine-wide and no administrator rights are needed:
    the only things this script touches are

        %LOCALAPPDATA%\MEx Automate\*.xml
        HKCU:\Software\Microsoft\Office\16.0\WEF\Developer   (one value each)

    The add-in's own code is loaded by Office from the website each time the
    pane opens, so updates arrive without reinstalling.

    Run uninstall.ps1 to undo everything this script does.

.PARAMETER Base
    The site to install from. Defaults to the public deployment. Use this to
    install a preview build instead.

.PARAMETER Apps
    Which add-ins to install: Excel, PowerPoint, or both (the default).

.PARAMETER Quiet
    Suppress the explanatory output and print only errors.

.EXAMPLE
    irm https://mex-automate.vercel.app/install.ps1 | iex

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\install.ps1 -Apps Excel
#>
[CmdletBinding()]
param(
    [string] $Base = "https://mex-automate.vercel.app",
    [ValidateSet("Excel", "PowerPoint", "Both")]
    [string] $Apps = "Both",
    [switch] $Quiet
)

$ErrorActionPreference = "Stop"

# The ids in the hosted manifests. Registering under these names means a
# reinstall replaces the previous entry instead of adding a second one.
$AddIns = @(
    [pscustomobject]@{
        App      = "Excel"
        Id       = "fcf30730-1db1-4555-979a-87d048d8c363"
        File     = "manifest.xml"
        Button   = "Home > Automate"
    },
    [pscustomobject]@{
        App      = "PowerPoint"
        Id       = "6295c215-900a-4867-85a9-2a9293d10f98"
        File     = "manifest-powerpoint.xml"
        Button   = "Home > Build deck"
    }
)

$DeveloperKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$InstallDir = Join-Path $env:LOCALAPPDATA "MEx Automate"

function Say([string] $Message) {
    if (-not $Quiet) { Write-Host $Message }
}

function Fail([string] $Message) {
    Write-Host ""
    Write-Host "Install failed: $Message" -ForegroundColor Red
    Write-Host "Ask for help at https://github.com/pauluwaifo/mex-automate/issues"
    exit 1
}

if ($env:OS -ne "Windows_NT") {
    Fail "this installer is for Windows. On a Mac, or in Office on the web, follow the manual steps at $Base/#install"
}

$site = $Base.TrimEnd("/")
$wanted = if ($Apps -eq "Both") { $AddIns } else { $AddIns | Where-Object { $_.App -eq $Apps } }

Say ""
Say "MEx Automate"
Say "Installing for $($wanted.App -join ' and ') on Windows, for your user account only."
Say ""

# ---------------------------------------------------------------------------
# 1. Is Office even here? A missing key means the registration would be written
#    somewhere Office never reads.
# ---------------------------------------------------------------------------
if (-not (Test-Path "HKCU:\Software\Microsoft\Office\16.0")) {
    Fail "Microsoft 365 / Office 2016 or newer was not found for your account. Web add-ins need Office 2016 or later."
}

if (-not (Test-Path $InstallDir)) { New-Item -ItemType Directory -Path $InstallDir | Out-Null }
if (-not (Test-Path $DeveloperKey)) { New-Item -Path $DeveloperKey -Force | Out-Null }

$installed = @()

foreach ($addIn in $wanted) {
    $manifestUrl = "$site/$($addIn.File)"
    $manifestPath = Join-Path $InstallDir $addIn.File

    # Downloaded to a temporary file first, so a failed download cannot leave a
    # half-written manifest registered.
    $temp = Join-Path ([IO.Path]::GetTempPath()) ("mex-manifest-" + [Guid]::NewGuid().ToString("N") + ".xml")
    try {
        Say "Downloading the $($addIn.App) manifest..."
        Invoke-WebRequest -Uri $manifestUrl -OutFile $temp -UseBasicParsing
    } catch {
        Fail "could not download $manifestUrl ($($_.Exception.Message))"
    }

    try {
        [xml] $xml = Get-Content -LiteralPath $temp -Raw
    } catch {
        Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
        Fail "the file at $manifestUrl is not valid XML, so it is not the manifest we expected."
    }

    $foundId = $xml.OfficeApp.Id
    if ($foundId -ne $addIn.Id) {
        Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
        Fail "the manifest at $manifestUrl is for a different add-in (id $foundId). Refusing to register it."
    }

    try {
        Move-Item -LiteralPath $temp -Destination $manifestPath -Force
        New-ItemProperty -Path $DeveloperKey -Name $addIn.Id -Value $manifestPath -PropertyType String -Force | Out-Null
    } catch {
        Fail $_.Exception.Message
    }

    $registered = (Get-ItemProperty -Path $DeveloperKey -Name $addIn.Id).$($addIn.Id)
    if ($registered -ne $manifestPath) {
        Fail "the $($addIn.App) registration did not stick. Office would not find the add-in."
    }

    $installed += [pscustomobject]@{
        App     = $addIn.App
        Version = $xml.OfficeApp.Version
        Button  = $addIn.Button
    }
}

# ---------------------------------------------------------------------------
# 2. Tell the user what to do next, and be honest about restarting.
# ---------------------------------------------------------------------------
$running = @()
foreach ($process in @("EXCEL", "POWERPNT")) {
    if (Get-Process $process -ErrorAction SilentlyContinue) { $running += $process }
}

Say ""
Say "Installed MEx Automate $($installed[0].Version)."
Say ""
if ($running.Count -gt 0) {
    Say "Office only looks for new add-ins at startup, and you have something open:"
    Say "  1. Save your work and close Excel and PowerPoint completely."
    Say "  2. Open them again."
} else {
    Say "Next:"
}
foreach ($entry in $installed) {
    Say "  Open $($entry.App), then $($entry.Button)"
}
Say ""
Say "The first time a pane opens it may take a few seconds while Office"
Say "downloads the add-in. After that it is cached."
Say ""
Say "To remove:  irm $site/uninstall.ps1 | iex"
Say ""
