function Initialize-OneNoteInterop {
    [CmdletBinding()]
    param()

    if (-not ('OneNoteAutomation.Interop.OneNoteApplication' -as [type])) {
        $interopPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Interop'
        Add-Type -Path @(
            (Join-Path $interopPath 'OneNoteApplication.cs')
            (Join-Path $interopPath 'OneNoteInterop.Generated.cs')
        ) -ErrorAction Stop
    }
}