function Assert-OneNoteApplication {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object]$Application
    )

    if ($PSEdition -eq 'Core' -and $null -ne $Application -and
        [System.Runtime.InteropServices.Marshal]::IsComObject($Application)) {
        throw 'Raw OneNote COM objects are not supported by -App in PowerShell 7. Create the application with New-OneNoteApplication instead.'
    }
}