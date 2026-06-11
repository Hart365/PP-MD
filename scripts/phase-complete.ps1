param(
  [switch]$SkipVersionBump,
  [switch]$SkipQualityGate,
  [switch]$MirrorUnpacked,
  [switch]$KeepTemp
)

$ErrorActionPreference = 'Stop'

Set-Location (Join-Path $PSScriptRoot '..')

Write-Host '== PP-MD Phase Close-Out ==' -ForegroundColor Cyan

if (-not $SkipVersionBump) {
  Write-Host 'Bumping package version (semver patch) for phase completion...'
  npm version patch --no-git-tag-version
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
} else {
  Write-Host 'Skipping version bump as requested.'
}

$packageJson = Get-Content '.\package.json' -Raw | ConvertFrom-Json
$currentVersion = $packageJson.version
Write-Host "Current package version: $currentVersion"
Write-Host 'Note: npm package versions use SemVer (major.minor.patch). This is the closest equivalent to a +0.0.0.1 phase increment.'

if (-not $SkipQualityGate) {
  Write-Host 'Running full quality gate (lint, regression, smoke, build, WCAG)...'
  npm run quality:gate
  if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
} else {
  Write-Host 'Skipping quality gate as requested.'
}

Write-Host 'Building fresh release executables (portable + installer)...'
& (Join-Path $PSScriptRoot 'build-latest.ps1') -MirrorUnpacked:$MirrorUnpacked -KeepTemp:$KeepTemp
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host ''
Write-Host 'Phase close-out completed successfully.' -ForegroundColor Green
Write-Host "Version ready for this phase: $currentVersion"
