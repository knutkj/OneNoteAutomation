#
# .SYNOPSIS
# Retrieves OneNote pages from specified sections or from all sections using
# the OneNote COM API.
#
# .DESCRIPTION
# Retrieves OneNote pages from sections provided via the pipeline, or from all
# notebooks when no section is specified. Supports wildcard filtering by name
# and can retrieve the currently viewed page with -Current.
#
# Returns page XML elements. By default, these contain metadata only (Name, ID,
# timestamps). Use -Content to retrieve full page XML elements with content
# structures that can be modified and passed to Update-OneNotePage.
#
# .EXAMPLE
# # Get all pages from a specific section.
# Get-OneNoteSection -NotebookName "Work" -Name "Daily" | Get-OneNotePage
#
# .EXAMPLE
# # Get pages matching a name pattern with content.
# Get-OneNoteSection "Work" "Daily" | Get-OneNotePage -Name "2025-*" -Content
#
# .EXAMPLE
# # Get the currently viewed page.
# Get-OneNotePage -Current
#
# .EXAMPLE
# # Get a specific page by ID.
# Get-OneNotePage -Id "{12345678-1234-5678-9012-...}{1}{...}"
#
# .EXAMPLE
# # Pipeline flow: notebook -> sections -> pages.
# Get-OneNoteNotebook "Personal" | Get-OneNoteSection | Get-OneNotePage
#
# .OUTPUTS
# OneNote.Page XML elements. Without -Content, returns lightweight metadata
# (Name, ID, dateTime attributes). With -Content, returns the full page XML
# element including Outline, OEChildren, and other content structures that can
# be modified and passed to Update-OneNotePage.
#
# .NOTES
# This cmdlet follows the OneNote hierarchy pipeline pattern: Notebook ->
# Section -> Page. To modify page content, use -Content to get the full XML
# element, make changes, then pass to Update-OneNotePage.
#
function Get-OneNotePage {
    [CmdletBinding()]
    param(
        # If specified, retrieves only the currently viewed page in OneNote.
        [Parameter(ParameterSetName = 'Current', Mandatory = $true)]
        [switch]$Current,

        # The ID of a specific page to retrieve.
        [Parameter(ParameterSetName = 'Id', Mandatory = $true)]
        [string]$Id,

        # Name of the page to retrieve, supporting wildcards and prefix
        # matching. Default is "*" (all pages).
        [Parameter(ParameterSetName = 'Pipeline', Position = 0)]
        [SupportsWildcards()]
        [string]$Name = "*",

        # Section XML element to retrieve pages from. Can be provided via
        # pipeline or directly. If not provided, retrieves pages from all
        # sections.
        [Parameter(ParameterSetName = 'Pipeline', ValueFromPipeline = $true)]
        [System.Xml.XmlElement]$Section,

        # If specified, returns the full page XML element instead of lightweight
        # metadata. Required for page content inspection or modification.
        [Parameter()]
        [switch]$Content,

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
        $hsPages = 4 # HierarchyScope.hsPages
        $app = $OneNoteApplication
        
        # Helper to get full page content by ID
        $getPageContent = {
            param($pageId)
            [xml]$xml = ''
            $app.GetPageContent($pageId, [ref]$xml)
            $xml.Page
        }
        
        # Helper to annotate and return page
        $annotatePage = {
            param($page)
            $page.PSObject.TypeNames.Insert(0, 'OneNote.Page')
            $page
        }
        
        $pages = @()
        
        if ($Current) {
            # Get current page from current section.
            $currentSection = Get-OneNoteSection -Current -App $app
            $hierarchy = Get-OneNoteHierarchy -Scope $hsPages -StartNodeId $currentSection.ID -App $app
            $pages = @($hierarchy.Section.Page | Where-Object -Property isCurrentlyViewed -EQ true)
            
            if ($pages.Count -gt 1) {
                throw "There are currently $($pages.Count) pages that are viewed."
            }
        }
        elseif ($Id) {
            # Get page by ID.
            if ($Content) {
                $pages = @(& $getPageContent $Id)
            }
            else {
                # Get lightweight metadata from hierarchy.
                $hierarchy = Get-OneNoteHierarchy -Scope $hsPages -StartNodeId $null -App $app
                $pages = @($hierarchy.Notebooks.Notebook.Section.Page | Where-Object -Property ID -EQ $Id)
            }
        }
        else {
            # Get pages from sections or all sections.
            if ($Section) {
                $hierarchy = Get-OneNoteHierarchy -Scope $hsPages -StartNodeId $Section.ID -App $app
                $pages = $hierarchy.Section.Page
            }
            else {
                $hierarchy = Get-OneNoteHierarchy -Scope $hsPages -StartNodeId $null -App $app
                $pages = $hierarchy.Notebooks.Notebook.Section.Page
            }
            
            # Filter by name pattern.
            $pages = $pages | 
            Where-Object -FilterScript { $_ } |
            Where-Object -Property Name -Like -Value $Name
        }
        
        # Process and return pages.
        $pages | ForEach-Object -Process {
            $page = $_
            
            # Get full content if requested (but not already fetched for Id parameter set).
            if ($Content -and -not $Id) {
                $page = & $getPageContent $_.ID
            }
            
            & $annotatePage $page
        }
    }

    end {
        if ($comObjectCreated) {
            Remove-ComObject -ComObject $OneNoteApplication
        }
    }
}
