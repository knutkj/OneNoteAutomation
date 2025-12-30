# Module entry point:
# - Dot-source functions from the Public folder so each cmdlet can live in its
#   own file. This follows the convention of one function per .ps1 file.
# - Keep exports explicit (via .psd1 FunctionsToExport) for predictable public
#   surface area. Dynamic discovery via Get-ChildItem ensures all cmdlets in
#   Public/ are loaded without hardcoding names.
# - Assert required dependencies (Public folder, ArgumentCompleters.ps1) at
#   module load time. Fail fast if module structure is invalid instead of
#   deferring errors to cmdlet invocation.

# Assert required dependencies exist - fail early if module is invalid.
$publicPath = Join-Path -Path $PSScriptRoot -ChildPath 'Public'
if (-not (Test-Path -Path $publicPath -PathType Container)) {
    throw "Module structure invalid: Public folder not found at $publicPath"
}

$completersPath = Join-Path -Path $PSScriptRoot -ChildPath 'ArgumentCompleters.ps1'
if (-not (Test-Path -Path $completersPath -PathType Leaf)) {
    throw "Module structure invalid: ArgumentCompleters.ps1 not found at $completersPath"
}

# Dot-source argument completers first - required by cmdlet parameters.
. $completersPath

# Dot-source functions from the Public folder
Get-ChildItem -Path $publicPath -Filter '*.ps1' -File | ForEach-Object {
    # Dot-source to load functions into the module scope at import time.
    . $_.FullName
}
