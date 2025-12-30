# .SYNOPSIS
# Safely creates, uses, and releases a COM object in PowerShell scripts.
#
# .DESCRIPTION
# Use-ComObject provides a C#-like 'using' pattern for COM objects. It creates
# a COM object for the given ProgId, passes it to a script block, and ensures
# the object is released and cleaned up afterward, even if an error occurs.
# This prevents memory leaks and orphaned processes when automating Office or
# other COM-based applications.
#
#
# .EXAMPLE
# Use-ComObject -ProgId OneNote.Application -Script {
#     param($onenote)
#     # Use $onenote COM object here.
# }
#
function Use-ComObject {
    [CmdletBinding()]
    param(
        # The programmatic identifier (ProgId) of the COM object to create
        # (for example, 'OneNote.Application').
        [Parameter(Mandatory)]
        [string]$ProgId,

        # A script block that receives the created COM object as its first
        # argument.
        [Parameter(Mandatory)]
        [scriptblock]$Script
    )

    $comObject = $null
    try {
        $comObject = New-Object -ComObject $ProgId
        & $Script $comObject
    }
    finally {
        Remove-ComObject -ComObject $comObject
    }
}

