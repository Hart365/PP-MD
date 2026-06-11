param(
  [switch]$SkipBuild,
  [switch]$KeepTemp,
  [switch]$MirrorUnpacked
)

$ErrorActionPreference = 'Stop'

Set-Location (Join-Path $PSScriptRoot '..')

function Invoke-RobocopyMirror {
  param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Destination,
    [string[]]$ExtraArgs = @()
  )

  & robocopy $Source $Destination /MIR /R:1 /W:1 @ExtraArgs | Out-Null
  $code = $LASTEXITCODE

  # Robocopy uses bit flags. Codes 0-7 are successful copies/sync operations.
  if ($code -ge 8) {
    throw "Robocopy failed ($code): '$Source' -> '$Destination'"
  }
}

function Sync-LatestArtifacts {
  param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Destination,
    [Parameter(Mandatory = $true)][string[]]$Patterns
  )

  if (-not (Test-Path $Destination)) {
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
  }

  foreach ($pattern in $Patterns) {
    Get-ChildItem -Path $Destination -Filter $pattern -File -ErrorAction SilentlyContinue |
      Remove-Item -Force -ErrorAction SilentlyContinue
  }

  $copied = 0
  foreach ($pattern in $Patterns) {
    $matches = Get-ChildItem -Path $Source -Filter $pattern -File -ErrorAction SilentlyContinue
    foreach ($match in $matches) {
      Copy-Item -Path $match.FullName -Destination $Destination -Force
      $copied++
    }
  }

  if ($copied -eq 0) {
    throw "No artifacts were copied from '$Source' to '$Destination'"
  }
}

$packageJson = Get-Content '.\package.json' -Raw | ConvertFrom-Json
$appVersion = $packageJson.version

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$portableOut = "release-temp-portable-$stamp"
$installerOut = "release-temp-installer-$stamp"

if (-not $SkipBuild) {
  Write-Host 'Running npm build...'
  npm run build
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

Write-Host "Packaging portable to $portableOut ..."
$env:PPMD_FLAVOR = 'portable'
$env:CSC_IDENTITY_AUTO_DISCOVERY = 'false'
npx electron-builder --win portable "--config.directories.output=$portableOut"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host "Packaging installer to $installerOut ..."
$env:PPMD_FLAVOR = 'installer'
$env:CSC_IDENTITY_AUTO_DISCOVERY = 'false'
npx electron-builder --win nsis "--config.directories.output=$installerOut"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host 'Mirroring outputs to release_latest_* folders...'
# Copy only top-level latest artifacts (exe, metadata, blockmap) to avoid
# transient lock failures from full-folder robocopy mirroring.
Sync-LatestArtifacts -Source ".\\$portableOut" -Destination '.\release_latest_portable' -Patterns @(
  'PP-MD-*-portable.exe',
  'builder-debug.yml',
  'builder-effective-config.yaml'
)
Sync-LatestArtifacts -Source ".\\$installerOut" -Destination '.\release_latest_installer' -Patterns @(
  'PP-MD-*-installer.exe',
  'PP-MD-*-installer.exe.blockmap',
  'latest.yml',
  'builder-debug.yml',
  'builder-effective-config.yaml'
)

if ($MirrorUnpacked) {
  Write-Host 'Mirroring unpacked runtime to .\\release_latest ...'
  Invoke-RobocopyMirror -Source ".\\$portableOut\\win-unpacked" -Destination '.\\release_latest'
} else {
  Write-Host 'Skipping unpacked runtime mirror (use -MirrorUnpacked to include).'
}

if (-not $KeepTemp) {
  Write-Host 'Cleaning temporary output folders...'
  try {
    Remove-Item ".\\$portableOut" -Recurse -Force
  } catch {
    Write-Warning "Could not remove ${portableOut}: $($_.Exception.Message)"
  }
  try {
    Remove-Item ".\\$installerOut" -Recurse -Force
  } catch {
    Write-Warning "Could not remove ${installerOut}: $($_.Exception.Message)"
  }
}

Write-Host ''
Write-Host "Latest portable: .\\release_latest_portable\\PP-MD-$appVersion-x64-portable.exe"
Write-Host "Latest installer: .\\release_latest_installer\\PP-MD-$appVersion-x64-installer.exe"
if ($MirrorUnpacked) {
  Write-Host "Latest unpacked app: .\\release_latest\\PP-MD.exe"
} else {
  Write-Host 'Latest unpacked app mirror skipped.'
}

# Ensure successful script runs return zero even when robocopy returns informational non-zero codes.
$global:LASTEXITCODE = 0
