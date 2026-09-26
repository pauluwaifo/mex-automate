<#
.SYNOPSIS
    Installs MEx Automate into Excel on Windows.

.DESCRIPTION
    Downloads the add-in's manifest to your user profile and registers it with
    Excel, so an "Automate" button appears on the Home tab of every workbook.

    Nothing is installed machine-wide and no administrator rights are needed:
    the only things this script touches are

        %LOCALAPPDATA%\MEx Automate\manifest.xml
        HKCU:\Software\Microsoft\Office\16.0\WEF\Developer   (one value)

    The add-in's own code is loaded by Excel from the website each time the pane
    opens, so updates arrive without reinstalling.

    Run uninstall.ps1 to undo everything this script does.

.PARAMETER Base
    The site to install from. Defaults to the public deployment. Use this to
    install a preview build instead.

.PARAMETER Quiet
    Suppress the explanatory output and print only errors.

.EXAMPLE
    irm https://mex-automate.vercel.app/install.ps1 | iex

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\install.ps1
#>
[CmdletBinding()]
param(
    [string] $Base = "https://mex-automate.vercel.app",
    [switch] $Quiet
)

$ErrorActionPreference = "Stop"

# The id in the hosted manifest. Registering under this name means a reinstall
# replaces the previous entry instead of adding a second one.
$AddInId = "fcf30730-1db1-4555-979a-87d048d8c363"
$DeveloperKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$InstallDir = Join-Path $env:LOCALAPPDATA "MEx Automate"
$ManifestPath = Join-Path $InstallDir "manifest.xml"

function Say([string] $Message) {
    if (-not $Quiet) { Write-Host $Message }
}

function Fail([string] $Message) {
    Write-Host ""
    Write-Host "Install failed: $Message" -ForegroundColor Red
    Write-Host "Nothing was changed. Ask for help at https://github.com/pauluwaifo/mex-automate/issues"
    exit 1
}

if ($env:OS -ne "Windows_NT") {
    Fail "this installer is for Windows. On a Mac or in Excel for the web, follow the manual steps at $Base/#install"
}

$site = $Base.TrimEnd("/")
$manifestUrl = "$site/manifest.xml"

Say ""
Say "MEx Automate"
Say "Installing for Excel on Windows, for your user account only."
Say ""

# ---------------------------------------------------------------------------
# 1. Is Excel even here? A missing Office key means the registration would be
#    written somewhere Excel never reads.
# ---------------------------------------------------------------------------
if (-not (Test-Path "HKCU:\Software\Microsoft\Office\16.0")) {
    Fail "Microsoft 365 / Office 2016 or newer was not found for your account. Web add-ins need Office 2016 or later."
}

# ---------------------------------------------------------------------------
# 2. Fetch the manifest. It is downloaded to a temporary file first so a failed
#    download cannot leave a half-written manifest registered.
# ---------------------------------------------------------------------------
$temp = Join-Path ([IO.Path]::GetTempPath()) ("mex-manifest-" + [Guid]::NewGuid().ToString("N") + ".xml")
try {
    Say "Downloading the add-in manifest..."
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
if ($foundId -ne $AddInId) {
    Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
    Fail "the manifest at $manifestUrl is for a different add-in (id $foundId). Refusing to register it."
}
$version = $xml.OfficeApp.Version

# ---------------------------------------------------------------------------
# 3. Put it somewhere stable and point Excel at it.
# ---------------------------------------------------------------------------
try {
    if (-not (Test-Path $InstallDir)) { New-Item -ItemType Directory -Path $InstallDir | Out-Null }
    Move-Item -LiteralPath $temp -Destination $ManifestPath -Force

    if (-not (Test-Path $DeveloperKey)) { New-Item -Path $DeveloperKey -Force | Out-Null }
    New-ItemProperty -Path $DeveloperKey -Name $AddInId -Value $ManifestPath -PropertyType String -Force | Out-Null
} catch {
    Fail $_.Exception.Message
}

$registered = (Get-ItemProperty -Path $DeveloperKey -Name $AddInId).$AddInId
if ($registered -ne $ManifestPath) {
    Fail "the registration did not stick. Excel would not find the add-in."
}

# ---------------------------------------------------------------------------
# 4. Tell the user what to do next, and be honest about restarting Excel.
# ---------------------------------------------------------------------------
$excelRunning = $null -ne (Get-Process EXCEL -ErrorAction SilentlyContinue)

Say ""
Say "Installed MEx Automate $version."
Say ""
if ($excelRunning) {
    Say "Excel is open, and it only looks for new add-ins at startup:"
    Say "  1. Save your work and close Excel completely."
    Say "  2. Open Excel again."
    Say "  3. On the Home tab, click Automate."
} else {
    Say "Next:"
    Say "  1. Open Excel."
    Say "  2. On the Home tab, click Automate."
}
Say ""
Say "The first time the pane opens it may take a few seconds while Excel"
Say "downloads the add-in. After that it is cached."
Say ""
Say "To remove it later:  irm $site/uninstall.ps1 | iex"
Say ""
