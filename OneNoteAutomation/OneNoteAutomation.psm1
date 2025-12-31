# Module entry point:
# - Dot-source functions from the Private folder first for internal utilities
# - Dot-source functions from the Public folder so each cmdlet can live in its
#   own file. This follows the convention of one function per .ps1 file.
# - Keep exports explicit (via .psd1 FunctionsToExport) for predictable public
#   surface area. Dynamic discovery via Get-ChildItem ensures all cmdlets in
#   Public/ are loaded without hardcoding names.
# - Assert required dependencies (Public folder, Private folder) at module load
#   time. Fail fast if module structure is invalid instead of deferring errors
#   to cmdlet invocation.

# Assert required dependencies exist - fail early if module is invalid.
$publicPath = Join-Path -Path $PSScriptRoot -ChildPath 'Public'
if (-not (Test-Path -Path $publicPath -PathType Container)) {
    throw "Module structure invalid: Public folder not found at $publicPath."
}

$privatePath = Join-Path -Path $PSScriptRoot -ChildPath 'Private'
if (-not (Test-Path -Path $privatePath -PathType Container)) {
    throw "Module structure invalid: Private folder not found at $privatePath."
}

# Dot-source Private functions first - required by Public cmdlets.
Get-ChildItem -Path $privatePath -Filter '*.ps1' -File | ForEach-Object {
    # Dot-source to load internal functions into the module scope.
    . $_.FullName
}

# Dot-source functions from the Public folder
Get-ChildItem -Path $publicPath -Filter '*.ps1' -File | ForEach-Object {
    # Dot-source to load functions into the module scope at import time.
    . $_.FullName
}
