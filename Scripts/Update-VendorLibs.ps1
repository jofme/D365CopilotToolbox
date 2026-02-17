<#
.SYNOPSIS
    Downloads vendor JavaScript libraries from npm and places them into the
    D365 AxResource Scripts folder.

.DESCRIPTION
    Reads Scripts/vendor-libs.json for the list of npm packages, versions, and
    source file paths. For each library it uses "npm pack" to download the
    package tarball, extracts the specified file, and copies it to the output
    directory.

    Prerequisites: Node.js / npm must be on PATH.

.PARAMETER CheckForUpdates
    Queries the npm registry for the latest version of each package and reports
    whether an update is available. Without -UpdateManifest, does NOT modify files.

.PARAMETER UpdateManifest
    When combined with -CheckForUpdates, writes the latest versions back into
    vendor-libs.json so a subsequent download fetches the new versions.

.PARAMETER Force
    Re-download files even if they already exist.

.EXAMPLE
    .\Update-VendorLibs.ps1

.EXAMPLE
    .\Update-VendorLibs.ps1 -CheckForUpdates

.EXAMPLE
    .\Update-VendorLibs.ps1 -Force
#>

[CmdletBinding()]
param(
    [switch]$CheckForUpdates,
    [switch]$UpdateManifest,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Verify npm is available
# ---------------------------------------------------------------------------

if (-not (Get-Command npm -ErrorAction SilentlyContinue)) {
    Write-Error "npm is not installed or not on PATH. Install Node.js from https://nodejs.org"
    exit 1
}

# ---------------------------------------------------------------------------
# Resolve paths
# ---------------------------------------------------------------------------

$ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ManifestPath = Join-Path $ScriptDir 'vendor-libs.json'

if (-not (Test-Path $ManifestPath)) {
    Write-Error "Manifest not found: $ManifestPath"
    exit 1
}

$Manifest  = Get-Content $ManifestPath -Raw | ConvertFrom-Json
$OutputDir = Join-Path $ScriptDir $Manifest.outputDir

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# ---------------------------------------------------------------------------
# Helper: Download and extract a single file from an npm package
# ---------------------------------------------------------------------------

function Get-NpmPackageFile {
    [CmdletBinding()]
    param(
        [string]$Package,
        [string]$Version,
        [string]$SourceFile,
        [string]$OutputPath
    )

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "vendor-libs-$([guid]::NewGuid().ToString('N'))"
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    try {
        # npm pack downloads the tarball into the current directory
        Push-Location $tempDir
        $tgzFile = npm pack "$Package@$Version" --silent 2>&1
        Pop-Location

        $tgzPath = Join-Path $tempDir $tgzFile.Trim()
        if (-not (Test-Path $tgzPath)) {
            throw "npm pack did not produce expected file: $tgzFile"
        }

        Write-Host "  Downloaded $Package@$Version"

        # Extract the tarball (npm tarballs are .tgz = gzip'd tar)
        # The target file is at package/<sourceFile> inside the tar
        $extractDir = Join-Path $tempDir 'extract'
        New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
        tar -xzf $tgzPath -C $extractDir 2>&1 | Out-Null

        $sourceFilePath = Join-Path $extractDir "package/$SourceFile"
        if (-not (Test-Path $sourceFilePath)) {
            throw "File '$SourceFile' not found in $Package@$Version package"
        }

        Copy-Item -Path $sourceFilePath -Destination $OutputPath -Force

        $hash = (Get-FileHash -Path $OutputPath -Algorithm SHA256).Hash
        Write-Host "  SHA256: $hash"
        Write-Host "  Saved:  $OutputPath"
    }
    finally {
        Pop-Location -ErrorAction SilentlyContinue
        if (Test-Path $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

$updatesAvailable = $false

foreach ($lib in $Manifest.libraries) {
    $outputPath = Join-Path $OutputDir $lib.outputFile

    Write-Host "`n[$($lib.name)] $($lib.package)@$($lib.version)" -ForegroundColor Cyan

    if ($CheckForUpdates) {
        $latest = (npm view "$($lib.package)" version --silent 2>$null)
        if ($latest -and $latest.Trim() -ne $lib.version) {
            Write-Host "  UPDATE AVAILABLE: $($lib.version) -> $($latest.Trim())" -ForegroundColor Yellow
            $updatesAvailable = $true
            if ($UpdateManifest) {
                $lib.version = $latest.Trim()
            }
        }
        elseif ($latest) {
            Write-Host "  Up to date ($($latest.Trim()))" -ForegroundColor Green
        }
        else {
            Write-Warning "  Could not query latest version"
        }
    }
    else {
        if ((Test-Path $outputPath) -and -not $Force) {
            Write-Host "  Already exists, skipping. Use -Force to re-download." -ForegroundColor DarkGray
            continue
        }

        Get-NpmPackageFile `
            -Package    $lib.package `
            -Version    $lib.version `
            -SourceFile $lib.sourceFile `
            -OutputPath $outputPath
    }
}

if ($CheckForUpdates) {
    if ($updatesAvailable -and $UpdateManifest) {
        $Manifest | ConvertTo-Json -Depth 10 | Set-Content -Path $ManifestPath -Encoding UTF8
        Write-Host "`nUpdated vendor-libs.json with new versions. Run without -CheckForUpdates to download." -ForegroundColor Yellow
        exit 1
    }
    elseif ($updatesAvailable) {
        Write-Host "`nUpdates available. Re-run with -UpdateManifest to write new versions, or edit Scripts/vendor-libs.json manually." -ForegroundColor Yellow
        exit 1
    }
    else {
        Write-Host "`nAll vendor libraries are up to date." -ForegroundColor Green
        exit 0
    }
}

Write-Host "`nDone." -ForegroundColor Green
