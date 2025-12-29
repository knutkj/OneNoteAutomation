#
# .SYNOPSIS
# Updates the OneNoteAutomation module manifest with version and dynamic
# function exports.
# 
# .DESCRIPTION
# This script automatically updates the module manifest with:
#
# - ModuleVersion from release tag (if available)
# - Prerelease suffix (if present in tag)
# - FunctionsToExport discovered from Public/*.ps1 files
# 
# It follows the convention: one function per .ps1 file, where function name
# matches filename.
# 
# .NOTES
# This script is used both locally for testing and by the CI/CD pipeline.
#
[CmdletBinding()]
param(
    # The version tag (e.g., "v1.2.3" or "v1.2.3-preview1"). If not provided,
    # tries GITHUB_REF_NAME environment variable.
    [string]$VersionTag = $env:GITHUB_REF_NAME,
    
    # The release notes text. If not provided, tries to extract from GitHub
    # event.
    [string]$ReleaseNotes
)

# Version tag is required - fail early
if (-not $VersionTag) {
    throw "Version tag is required. Provide -VersionTag parameter or set GITHUB_REF_NAME environment variable."
}

# Resolve paths relative to script location
$scriptRoot = Split-Path -Parent $PSCommandPath
$repoRoot = Split-Path -Parent $scriptRoot
$manifestPath = Join-Path $repoRoot 'OneNoteAutomation\OneNoteAutomation.psd1'
$publicPath = Join-Path $repoRoot 'OneNoteAutomation\Public'

Write-Host "Repository root: $repoRoot"
Write-Host "Manifest path: $manifestPath"
Write-Host "Public path: $publicPath"

# Verify manifest exists
if (-not (Test-Path $manifestPath)) {
    throw "Module manifest not found: $manifestPath"
}

# Dynamically build FunctionsToExport from Public/*.ps1 files
# Convention: one function per file, function name matches filename
$functions = @()
if (Test-Path $publicPath) {
    $functions = Get-ChildItem -Path $publicPath -Filter '*.ps1' -File | 
        ForEach-Object { $_.BaseName }
}

if ($functions.Count -eq 0) {
    throw "No PowerShell files found in $publicPath. Cannot publish module without exported functions."
} else {
    Write-Host "Found functions to export: $($functions -join ', ')"
}

# Extract release notes from GitHub event if not provided
if (-not $ReleaseNotes -and $env:GITHUB_EVENT_PATH) {
    try {
        $githubEvent = Get-Content $env:GITHUB_EVENT_PATH | ConvertFrom-Json
        $ReleaseNotes = $githubEvent.release.body
        if ($ReleaseNotes) {
            Write-Host "Extracted release notes from GitHub event"
        }
    } catch {
        Write-Warning "Could not extract release notes from GitHub event: $_"
    }
}

Write-Host "Processing version tag: $VersionTag"

# Parse version tag (e.g., "v1.2.3" or "v1.2.3-preview1")
$versionTag = $VersionTag.TrimStart('v')
$versionParts = $versionTag -split '-', 2
$moduleVersion = $versionParts[0]
$prerelease = if ($versionParts.Count -gt 1) { $versionParts[1] } else { $null }

Write-Host "Setting module version to $moduleVersion"

# Prepare Update-ModuleManifest parameters
$manifestParams = @{
    Path = $manifestPath
    ModuleVersion = $moduleVersion
    FunctionsToExport = $functions
}

if ($prerelease) {
    Write-Host "Setting prerelease label to $prerelease"
    $manifestParams.Prerelease = $prerelease
}

if ($ReleaseNotes) {
    Write-Host "Setting release notes ($(($ReleaseNotes -split '\n').Count) lines)"
    $manifestParams.ReleaseNotes = $ReleaseNotes
}

# Update manifest with all parameters
Update-ModuleManifest @manifestParams

Write-Host "Module manifest updated successfully"