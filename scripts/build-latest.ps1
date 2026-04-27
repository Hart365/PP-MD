param(
  [switch]$SkipBuild,
  [switch]$KeepTemp,
  [switch]$MirrorUnpacked
)

$ErrorActionPreference = 'Stop'

Set-Location (Join-Path $PSScriptRoot '..')

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
# Mirror installer/portable outputs but exclude unpacked runtime files to avoid
# duplicating ~hundreds of MB in latest artifact folders.
if (Test-Path '.\release_latest_portable\win-unpacked') {
  Remove-Item '.\release_latest_portable\win-unpacked' -Recurse -Force
}
if (Test-Path '.\release_latest_installer\win-unpacked') {
  Remove-Item '.\release_latest_installer\win-unpacked' -Recurse -Force
}
robocopy ".\\$portableOut" ".\\release_latest_portable" /MIR /R:1 /W:1 /XD "win-unpacked" | Out-Null
robocopy ".\\$installerOut" ".\\release_latest_installer" /MIR /R:1 /W:1 /XD "win-unpacked" | Out-Null

if ($MirrorUnpacked) {
  Write-Host 'Mirroring unpacked runtime to .\\release_latest ...'
  robocopy ".\\$portableOut\\win-unpacked" ".\\release_latest" /MIR /R:1 /W:1 | Out-Null
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
Write-Host "Latest portable: .\\release_latest_portable\\PP-MD-1.0.0-x64-portable.exe"
Write-Host "Latest installer: .\\release_latest_installer\\PP-MD-1.0.0-x64-installer.exe"
if ($MirrorUnpacked) {
  Write-Host "Latest unpacked app: .\\release_latest\\PP-MD.exe"
} else {
  Write-Host 'Latest unpacked app mirror skipped.'
}
