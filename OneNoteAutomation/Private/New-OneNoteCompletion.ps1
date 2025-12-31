#
# .SYNOPSIS
# Creates a PowerShell completion result for OneNote objects.
#
# .DESCRIPTION
# Internal helper function that creates properly formatted completion results
# for OneNote objects (notebooks, sections, pages) with ID information in tooltips.
#
filter New-OneNoteCompletion {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Id,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Name
    )

    [System.Management.Automation.CompletionResult]::new(
        "'$Name'",
        $Name,
        [System.Management.Automation.CompletionResultType]::ParameterValue,
        "ID: '$Id'."
    )
}