#
# .SYNOPSIS
# Updates the content of a OneNote page.
#
# .DESCRIPTION
# Accepts a page XML element from the pipeline or a page XML document via
# parameter and updates the page in OneNote. Validates that the provided XML
# has the expected OneNote page structure before updating.
#
# .EXAMPLE
# # Update a page after modifying its XML element.
# $page = Get-OneNotePage -Current -Content
# $page.Title.OE.T.'#cdata-section' = "Updated Title"
# $updatedPage = $page | Update-OneNotePage
#
# .EXAMPLE
# # Update page with shared COM object for better performance.
# Use-ComObject -ProgId OneNote.Application -Script {
#     param($OneNote)
#     $page = Get-OneNotePage -Current -Content -App $OneNote
#     # Modify content...
#     $page | Update-OneNotePage -App $OneNote
# }
#
# .OUTPUTS
# System.Xml.XmlElement. The updated page element retrieved from OneNote
# after the update is complete, allowing for pipeline chaining with the
# latest version.
#
# .NOTES
# The page must be retrieved with -Content to get a modifiable XML element.
# This cmdlet validates that the XML document element is a OneNote Page before
# calling the update.
#
function Update-OneNotePage {
  [CmdletBinding()]
  param(
    # A page XML element retrieved from Get-OneNotePage -Content. The element's
    # OwnerDocument will be validated and used for the update.
    [Parameter(ParameterSetName = 'Element', ValueFromPipeline = $true, Mandatory = $true)]
    [System.Xml.XmlElement]$Page,

    # A complete page XML document. The document element must be a OneNote Page.
    [Parameter(ParameterSetName = 'Document', Mandatory = $true)]
    [System.Xml.XmlDocument]$PageDocument,

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
    if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"]) {
      Write-Verbose "Starting process block."
      Write-Verbose ("Parameters: OneNoteApplication={0}" -f ($null -ne $OneNoteApplication))
    }

    $app = $OneNoteApplication

    # Determine the document to use based on parameter set.
    if ($PSCmdlet.ParameterSetName -eq 'Element') {
      if (-not ($Page | Test-OneNotePageHasContent)) {
        throw "Invalid page element: must be a full page element from Get-OneNotePage -Content, not a lightweight metadata element."
      }
      $doc = $Page.OwnerDocument
    }
    else {
      if (-not ($PageDocument | Test-OneNotePageHasContent)) {
        throw "Invalid document: expected 'Page' root element but found '$($PageDocument.DocumentElement.LocalName)'."
      }
      $doc = $PageDocument
    }

    Write-Verbose -Message "Calling UpdatePageContent for page '$($doc.DocumentElement.ID)'."
    $app.UpdatePageContent($doc.OuterXml)
    Write-Verbose -Message "Update complete. Retrieving updated page."
    
    # Retrieve and return the updated page.
    $pageId = $doc.DocumentElement.ID
    Get-OneNotePage -Id $pageId -Content -App $app
    
    Write-Verbose -Message "Process block complete."
  }

  end {
    if ($comObjectCreated) {
      Remove-ComObject -ComObject $OneNoteApplication
    }
  }
}
