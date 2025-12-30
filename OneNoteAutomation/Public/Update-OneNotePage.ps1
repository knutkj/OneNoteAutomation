#
# .SYNOPSIS
# Updates the content of a OneNote page using the COM API.
#
# .DESCRIPTION
# Accepts page content XML from the pipeline or via parameters and updates the
# page content in OneNote. The Content parameter should be a
# System.Xml.XmlDocument containing the page's XML structure. If no OneNote
# application object is provided, it is created and disposed automatically.
#
# .EXAMPLE
# # Update a page's content after modifying its XML.
# $page = Get-OneNotePage -Section $section -Name "MyPage" -Content
# $page.Content.Page.Title.OE.T.'#cdata-section' = "Updated Title"
# $page | Update-OneNotePage
#
# .EXAMPLE
# # Update page content with shared COM object for better performance.
# Use-ComObject -ProgId OneNote.Application -Script {
#     param($OneNote)
#     $section = Get-OneNoteSection -NotebookName "Work" -Name "Projects" -App $OneNote
#     $page = $section | Get-OneNotePage -Name "Bathroom" -Content -App $OneNote
#     # Modify content...
#     Update-OneNotePage -Content $page.Content -App $OneNote
# }
#
# .OUTPUTS
# None. This cmdlet does not return any objects.
#
# .NOTES
# This cmdlet calls the UpdatePageContent method of the OneNote COM API. The
# Content parameter must contain valid OneNote page XML structure. Always
# retrieve page content with the -Content switch when using Get-OneNotePage
# before attempting to update it.
#
function Update-OneNotePage {
  [CmdletBinding()]
  param(
    # The XML content to update the page with. Must be a System.Xml.XmlDocument
    # object containing valid OneNote page XML. Can be provided via pipeline
    # (ValueFromPipelineByPropertyName) or as a parameter.
    [Parameter(ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
    [System.Xml.XmlDocument]$Content,

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

    Write-Verbose -Message "Calling UpdatePageContent."
    $app.UpdatePageContent($Content.OuterXml)
    Write-Verbose -Message "Process block complete."
  }

  end {
    if ($comObjectCreated) {
      Remove-ComObject -ComObject $OneNoteApplication
    }
  }
}
