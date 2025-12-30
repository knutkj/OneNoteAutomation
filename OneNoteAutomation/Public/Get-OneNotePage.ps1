#
# .SYNOPSIS
# Retrieves OneNote pages from specified sections or from all sections using
# the OneNote COM API.
#
# .DESCRIPTION
# Accepts section objects from the pipeline and retrieves their pages.
# Optionally filters pages by name using wildcard patterns and case-insensitive
# matching. Can retrieve the currently viewed page with the -Current switch or
# include page content XML with the -Content switch. If no section is provided,
# retrieves all pages across all sections in all notebooks. Returns page
# objects with the custom type name 'OneNote.Page'.
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
# OneNote.Page objects with Name, ID, and other page properties. If -Content is
# specified, includes a Content property with the page's XML structure.
#
# .NOTES
# This cmdlet follows the OneNote hierarchy pipeline pattern: Notebook ->
# Section -> Page. When using -Content, the page XML can be modified and passed
# to Update-OneNotePage for content updates.
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

        # If specified, includes the page content XML in the returned page
        # objects. The XML content is added as a Content property.
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
                Where-Object -Property isCurrentlyViewed -EQ true |
                ForEach-Object -Process { $_.PSTypeNames.Insert(0, 'OneNote.Page'); $_ })

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
            $_.PSTypeNames.Insert(0, 'OneNote.Page')
            if ($Content) {
                [xml]$xml = ''
                $app.GetPageContent($_.ID, [ref]$xml)
                $_ | Add-Member -NotePropertyName Content -NotePropertyValue $xml -Force
            }
            $_
        }
    }

    end {
        if ($comObjectCreated) {
            Remove-ComObject -ComObject $OneNoteApplication
        }
    }
}
