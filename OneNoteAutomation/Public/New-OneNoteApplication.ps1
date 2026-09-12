#
# .SYNOPSIS
# Creates a OneNote application COM object.
#
# .DESCRIPTION
# Creates an application object that can be reused with the -App parameter
# of OneNote cmdlets. The caller owns the object and must release it with
# Remove-ComObject when finished.
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

    New-Object -ComObject OneNote.Application
}