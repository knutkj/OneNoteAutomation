#
# .SYNOPSIS
# Releases and cleans up a COM object to prevent memory leaks and orphaned
# processes.
#
# .DESCRIPTION
# Remove-ComObject safely releases a COM object by calling ReleaseComObject,
# removing the variable, and forcing garbage collection. This is useful when
# automating Office or other COM-based applications to ensure resources are
# properly freed.
#
# .PARAMETER ComObject
# The COM object to release and clean up.
#
# .EXAMPLE
# # Releases the $onenote COM object and performs cleanup.
# Remove-ComObject -ComObject $onenote
#
# .OUTPUTS
# None. This function does not return any objects.
#
# .NOTES
# This function should be called to clean up COM objects created with
# New-Object -ComObject. Use-ComObject automatically calls this function,
# so manual cleanup is only needed when creating COM objects directly.
#
function Remove-ComObject {
    [CmdletBinding()]
    param(
        # The COM object to release and clean up.
        [Parameter(Mandatory)]
        [object]$ComObject
    )

    if ($null -ne $ComObject) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
        Remove-Variable ComObject -ErrorAction SilentlyContinue
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
