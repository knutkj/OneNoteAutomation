#
# .SYNOPSIS
# Retrieves the full page content XML element for a OneNote page.
#
# .DESCRIPTION
# Fetches the complete page XML content by page ID and returns the Page element.
# This is an internal helper for cmdlets that need full page content.
#
# .OUTPUTS
# System.Xml.XmlElement. The full page XML element.
#
function Get-OneNotePageContent {
    [CmdletBinding()]
    param(
        # The ID of the page to retrieve content for.
        [Parameter(Mandatory = $true)]
        [string]$PageId,

        # The OneNote.Application COM object. Required for performance.
        [Parameter(Mandatory = $true)]
        [Alias('App')]
        $OneNoteApplication,

        # If specified, adds the OneNote.Page type to the returned element.
        [switch]$Annotate
    )

    [xml]$xml = ''
    $OneNoteApplication.GetPageContent($PageId, [ref]$xml)
    $page = $xml.Page

    if ($Annotate) {
        $page.PSObject.TypeNames.Insert(0, 'OneNote.Page')
    }

    $page
}
