#
# .SYNOPSIS
# Creates a OneNote application object for direct automation in PowerShell.
#
# .DESCRIPTION
# Creates an object for direct OneNote API calls or reuse through the -App
# parameter of module cmdlets. Uses raw COM in Windows PowerShell 5.1 and
# a managed wrapper in PowerShell 7 on Windows. The caller owns the object
# and must release it with Remove-ComObject when finished.
#
# .EXAMPLE
# # Call the OneNote API directly, using the same syntax in PS5 and PS7.
# $application = New-OneNoteApplication
# try {
#     [string]$hierarchyXml = ''
#     $application.GetHierarchy('', 2, [ref]$hierarchyXml)
#     [xml]$hierarchy = $hierarchyXml
#     $hierarchy.Notebooks.Notebook
# }
# finally {
#     Remove-ComObject -ComObject $application
# }
#
# .EXAMPLE
# $application = New-OneNoteApplication
# try {
#     Get-OneNoteNotebook -App $application
# }
# finally {
#     Remove-ComObject -ComObject $application
# }
#
function New-OneNoteApplication {
    [CmdletBinding()]
    param()

    if ($PSEdition -eq 'Core') {
        if (-not $IsWindows) {
            throw 'OneNote automation requires Windows and the OneNote desktop application.'
        }
        Initialize-OneNoteInterop
        New-Object -TypeName OneNoteAutomation.Interop.OneNoteApplication
    }
    else {
        New-Object -ComObject OneNote.Application
    }
}