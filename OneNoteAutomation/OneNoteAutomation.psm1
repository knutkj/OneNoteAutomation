# Module entry point:
# - dot-source functions from the Public folder so each cmdlet can live in its own file.
# - keep exports explicit (via .psd1 FunctionsToExport) for predictable public surface area.
$publicPath = Join-Path -Path $PSScriptRoot -ChildPath 'Public'
if (Test-Path -Path $publicPath) {
    Get-ChildItem -Path $publicPath -Filter '*.ps1' -File | ForEach-Object {
        # Dot-source to load functions into the module scope at import time.
        . $_.FullName
    }
}
