#
# .SYNOPSIS
# Applies spacing fixes to Level 1 headings on a OneNote page.
#
# .DESCRIPTION
# Scans a OneNote page for Level 1 headings (h1 style) and applies consistent
# spacing by setting spaceBefore to 0.23 on the style definition and removing
# individual spaceBefore attributes from heading elements. This ensures uniform
# heading spacing across the page.
#
# If a lightweight page element (metadata only) is provided, the full page
# content is retrieved automatically. The modified page element is passed
# through for pipeline chaining to Update-OneNotePage.
#
# .EXAMPLE
# # Apply spacing fixes to the current page.
# Get-OneNotePage -Current -Content | Set-OneNotePageSpacing | Update-OneNotePage
#
# .EXAMPLE
# # Works with lightweight page elements too (fetches content internally).
# Get-OneNotePage -Current | Set-OneNotePageSpacing | Update-OneNotePage
#
# .OUTPUTS
# System.Xml.XmlElement. The modified page element for pipeline chaining to
# Update-OneNotePage.
#
# .NOTES
# This cmdlet modifies the spaceBefore attribute on h1 QuickStyleDef elements
# and removes spaceBefore from individual OE elements with h1 style.
#
function Set-OneNotePageSpacing {
    [CmdletBinding()]
    param(
        # The page XML element. Can be a lightweight element (metadata only) or
        # a full page element from Get-OneNotePage -Content.
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [System.Xml.XmlElement]$Page,

        # An existing OneNote.Application COM object. If not provided, a new COM
        # object will be created and automatically disposed after the operation.
        [Alias('App')]
        [Parameter()]
        $OneNoteApplication = $null
    )

    begin {
        $comObjectCreated = $false
        if (-not $OneNoteApplication) {
            $comObjectCreated = $true
            $OneNoteApplication = New-Object -ComObject OneNote.Application
        }
    }

    process {
        $app = $OneNoteApplication
        $ns = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
        $nsMap = @{ "one" = $ns }

        $pageId = $Page.ID
        if (-not $pageId) {
            throw "Page ID not found. Ensure the element is a valid OneNote page."
        }

        Write-Verbose -Message "Processing page ID: $pageId."

        # If lightweight element, fetch full content.
        $pageElement = $Page
        if (-not ($Page | Test-OneNotePageHasContent)) {
            Write-Verbose -Message "Lightweight page element detected, fetching full content."
            $pageElement = Get-OneNotePageContent -PageId $pageId -App $app -Annotate
        }

        $doc = $pageElement.OwnerDocument

        # Finding h1 style definitions.
        $styleSelector = '//one:QuickStyleDef[@name="h1"]'
        $defs = @($doc | Select-Xml -Namespace $nsMap -XPath $styleSelector).Node
        $indexes = @($defs.index)
        Write-Verbose -Message (
            'Found {0} quick style definitions for Level 1 headings.' -f $indexes.Count)

        # Finding h1 headings.
        $headingCandidatesPattern = '/one:Page/one:Outline/one:OEChildren/one:OE[@quickStyleIndex]'
        $l1Headings = @($doc |
            Select-Xml -Namespace $nsMap -XPath $headingCandidatesPattern |
            ForEach-Object -Process { $_.Node } |
            Where-Object -Property quickStyleIndex -In -Value $indexes)

        Write-Verbose ("Found {0} Level 1 headings." -f $l1Headings.Count)

        # Apply spacing fixes.
        Write-Verbose "Applying spacing fixes to Level 1 headings."
        $defs.SetAttribute("spaceBefore", "0.23")
        $l1Headings | ForEach-Object { $_.RemoveAttribute("spaceBefore") }

        Write-Verbose -Message "Spacing fixes complete."

        # Pass through page element for pipeline chaining to Update-OneNotePage.
        $pageElement
    }

    end {
        if ($comObjectCreated) {
            Remove-ComObject -ComObject $OneNoteApplication
        }
    }
}
