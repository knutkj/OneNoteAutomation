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
        # Section object(s) from the pipeline. Each section should have an ID
        # property. If not provided, retrieves pages from all sections.
        [Parameter(ParameterSetName = 'Pipeline', ValueFromPipeline = $true)]
        $Section,

        # Name of the page to retrieve, supporting wildcards and prefix
        # matching. Default is "*" (all pages).
        [Parameter(Position = 0)]
        [SupportsWildcards()]
        [string]$Name = "*",

        # If specified, retrieves only the currently viewed page in OneNote.
        [Parameter(ParameterSetName = 'Current', Mandatory = $true)]
        [switch]$Current,

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
        $pages = @()
        if ($Current) {
            $rootHierarchy = Get-OneNoteHierarchy `
                -Scope $hsPages `
                -StartNodeId $null `
                -App $app

            $pages = @($rootHierarchy.Notebooks.Notebook.Section.Page |
                Where-Object -Property isCurrentlyViewed -EQ true)

            if ($viewedPages.Count -gt 1) {
                throw "There are currently $($viewedPages.Count) pages that are viewed.)"
            }
        }
        else {
            if ($Section) {
                # Handle section object from pipeline
                $hierarchy = Get-OneNoteHierarchy `
                    -Scope $hsPages `
                    -StartNodeId $Section.ID `
                    -App $app

                $pages = $hierarchy.Section.Page
            }
            else {
                # If no section specified, get all pages from root hierarchy.
                $rootHierarchy = Get-OneNoteHierarchy `
                    -Scope $hsPages `
                    -StartNodeId $null `
                    -App $app

                $pages = $rootHierarchy.Notebooks.Notebook.Section.Page
            }

            $pages = $pages |
            Where-Object -FilterScript { $_ } |
            Where-Object -Property Name -Like -Value $Name
        }

        $pages |
        ForEach-Object -Process {
            $page = $_
            if ($Content) {
                $page = Get-OneNotePageContent -PageId $_.ID -App $app
            }

            $page.PSObject.TypeNames.Insert(0, 'OneNote.Page')
            $page
        }
    }

    end {
        if ($comObjectCreated) {
            Remove-ComObject -ComObject $OneNoteApplication
        }
    }
}
