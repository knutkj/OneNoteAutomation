#
# .SYNOPSIS
# Creates or updates a table of contents for a OneNote page based on h1 headings.
#
# .DESCRIPTION
# Scans a OneNote page for Level 1 headings (h1 style) and generates a clickable
# table of contents at the top of the page. Existing TOC entries are replaced
# with updated links. TOC items are marked with metadata to enable future
# updates.
#
# If a lightweight page element (metadata only) is provided, the full page
# content is retrieved automatically. The modified page element is passed
# through for pipeline chaining to Update-OneNotePage.
#
# .EXAMPLE
# # Update TOC for the current page.
# Get-OneNotePage -Current -Content | Set-OneNotePageToc | Update-OneNotePage
#
# .EXAMPLE
# # Update TOC for the current page using -Current switch.
# Set-OneNotePageToc -Current | Update-OneNotePage
#
# .EXAMPLE
# # Update and save TOC for the current page in one command.
# Set-OneNotePageToc -Current -Save
#
# .EXAMPLE
# # Works with lightweight page elements too (fetches content internally).
# Get-OneNotePage -Current | Set-OneNotePageToc | Update-OneNotePage
#
# .OUTPUTS
# System.Xml.XmlElement. The modified page element for pipeline chaining to
# Update-OneNotePage.
#
# .NOTES
# TOC items are tagged with metadata (pscomlib.kind = "toc-item") to identify
# them for future updates. Only h1 headings are included in the TOC.
#
function Set-OneNotePageToc {
    [CmdletBinding()]
    param(
        # The page XML element. Can be a lightweight element (metadata only) or
        # a full page element from Get-OneNotePage -Content.
        [Parameter(ParameterSetName = 'Page', ValueFromPipeline = $true, Mandatory = $true)]
        [System.Xml.XmlElement]$Page,

        # If specified, operates on the currently viewed page in OneNote.
        [Parameter(ParameterSetName = 'Current', Mandatory = $true)]
        [switch]$Current,

        # If specified, automatically saves the changes to OneNote.
        [switch]$Save,

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

        # Block helper: Each.
        $each = {
            param($items, $t)
            $res = $items |
            ForEach-Object { (& $t $_) -split [Environment]::NewLine } |
            ForEach-Object { $_.trim() }
            $res -join ''
        }

        # TOC Section Template using the each block helper.
        $ns = 'http://schemas.microsoft.com/office/onenote/2013/onenote'
        $tocSectionTemplate = { param($m) [xml]@"
<one:Outline xmlns:one="$ns">
  <one:OEChildren>
$(& $each $m { param($i) @"
    <one:OE>
      <one:Meta name="pscomlib.kind" content="toc-item" />
      <one:Meta name="pscomlib.toc-item.target" content="$($i.Text)" />
      <one:T>
        <![CDATA[<p><a href='$($i.HyperLinkToObject)'>
          ↓ $($i.Text)
        </a></p>]]>
      </one:T>
    </one:OE>
"@})
  </one:OEChildren>
</one:Outline>
"@}
    }

    process {
        $app = $OneNoteApplication
        $nsMap = @{ "one" = $ns }

        # Get the page to work on
        if ($PSCmdlet.ParameterSetName -eq 'Current') {
            $Page = Get-OneNotePage -Current -App $app
        }

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
        $tocItemSelector = '//one:QuickStyleDef[@name="h1"]'
        $defs = @($doc | Select-Xml -Namespace $nsMap -XPath $tocItemSelector).Node
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

        # Build TOC items with hyperlinks.
        Write-Verbose "Building TOC items."
        $tocItems = $l1Headings | ForEach-Object {
            $node = $_
            $link = ''
            $app.GetHyperlinkToObject($pageId, $node.objectID, [ref]$link)
            $text = $node.T.InnerText
            Write-Verbose ("TOC item: Text='{0}', Link='{1}'" -f $text, $link)
            [PSCustomObject]@{
                Text              = $text
                HyperLinkToObject = $link
            }
        }

        # Generate new TOC elements.
        Write-Verbose -Message "Generating new TOC elements ..."
        $newTocElements = @((& $tocSectionTemplate $tocItems).Outline.ChildNodes.OE |
            ForEach-Object -Process { $doc.ImportNode($_, $true) })
        Write-Verbose ("Generated and imported {0} new TOC elements." -f $newTocElements.Count)

        # Find and remove existing TOC elements.
        $tocItemSelector = '//one:OE[one:Meta[@name="pscomlib.kind" and @content="toc-item"]]'
        $existingTocElements = @($doc |
            Select-Xml -Namespace $nsMap -XPath $tocItemSelector |
            ForEach-Object -Process { $_.Node })
        Write-Verbose ("Found {0} existing TOC elements." -f $existingTocElements.Count)

        Write-Verbose -Message "Removing existing TOC elements."
        $existingTocElements |
        ForEach-Object -Process { $_.ParentNode.RemoveChild($_) | Out-Null }

        # Prepend new TOC elements to page.
        [Array]::Reverse($newTocElements)
        Write-Verbose -Message "Prepending TOC elements to page."
        $newTocElements |
        ForEach-Object -Process { $pageElement.Outline.ChildNodes.PrependChild($_) | Out-Null }

        Write-Verbose -Message "TOC generation complete."

        # Save changes if requested.
        if ($Save) {
            Write-Verbose -Message "Saving changes to OneNote."
            $pageElement | Update-OneNotePage -App $app
        }

        # Pass through page element for pipeline chaining to Update-OneNotePage.
        $pageElement
    }

    end {
        if ($comObjectCreated) {
            Remove-ComObject -ComObject $OneNoteApplication
        }
    }
}
