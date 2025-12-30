using namespace Microsoft.Office.Interop.OneNote

#
# .SYNOPSIS
# Creates a new page in a specified OneNote section using the COM API and
# returns the new page object.
#
# .DESCRIPTION
# Creates a new page in the given section (by object or ID), optionally
# specifying the page style. Returns the new page's object (not just the ID) as
# retrieved from the OneNote hierarchy. If no OneNote application object is
# provided, it is created and disposed automatically.
#
# .EXAMPLE
# # Creates a new page in the specified section and returns the page object.
# New-OneNotePage -Id $sectionId
#
# .EXAMPLE
# # Creates a new page with a title and returns the page object.
# New-OneNotePage -Id $sectionId -Title "Meeting Notes"
#
# .EXAMPLE
# # Creates a new page in the first section of the "Diary" notebook.
# Get-OneNoteSection -NotebookName "Diary" |
#     Select-Object -First 1 | 
#     New-OneNotePage -Title "Today's Thoughts"
#
function New-OneNotePage {
    [CmdletBinding(DefaultParameterSetName = 'BySectionId')]
    param(
        # The ID of the section in which to create the new page.
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [string]$Id,

        # Optional title for the new page.
        [Parameter(Position = 0)]
        [string]$Title,

        # Optional position for the new page in the section (0-based index).
        [Parameter(Position = 1)]
        [int]$Position,
        
        # Optional page level (1=top, 2=subpage, etc).
        [Alias('Level')]
        [Parameter(Position = 2)]
        [int]$PageLevel,

        # The style of the new page.
        # 0 = npsDefault (default page style)
        # 1 = npsBlankPageWithTitle (blank page with title)
        # 2 = npsBlankPageNoTitle (blank page with no title)
        [Parameter()]
        [int]$PageStyle = 0, # NewPageStyle.npsDefault.

        # The OneNote application object. If not provided, it will be created.
        [Alias('App')]
        [Parameter()]
        $OneNoteApplication = $null
    )

    begin {
        $hsSelf = 0 # HierarchyScope.hsSelf
        $hsPages = 4 # HierarchyScope.hsPages
        $disposeApp = $false
        if (-not $OneNoteApplication) {
            $disposeApp = $true
            $OneNoteApplication = New-Object -ComObject OneNote.Application
        }
    }

    process {
        $app = $OneNoteApplication
        [string]$newPageId = ''
        $app.CreateNewPage($Id, [ref]$newPageId, $PageStyle)

        # Set title if requested.
        if ($Title) {
            [xml]$pageDoc = ''
            $app.GetPageContent($newPageId, [ref]$pageDoc)
            $pageDoc.Page.Title.OE.T.'#cdata-section' = $Title
            $app.UpdatePageContent($pageDoc.OuterXml)
        }

        # Set page level if requested.
        if ($PageLevel) {
            $hierarchy = Get-OneNoteHierarchy -Scope $hsPages -Node $Id -App $app
            $pageNode = $hierarchy.Section.Page | Where-Object -Property ID -EQ $newPageId
            $pageNode.pageLevel = "$PageLevel"
            $OneNoteApplication.UpdateHierarchy($hierarchy.OuterXml)
        }

        # Move page to a specific position if requested.
        if ($PSBoundParameters.ContainsKey('Position')) {
            $hierarchy = Get-OneNoteHierarchy -Scope $hsPages -Node $Id -App $app
            $pagesList = @($hierarchy.Section.Page)
            $pageNode = $pagesList | Where-Object -Property ID -EQ $newPageId
            $targetIndex = $Position
            if ($targetIndex -lt 0 -or $targetIndex -ge $pagesList.Count) {
                Write-Warning -Message `
                    "Position $targetIndex is out of bounds (0..$($pagesList.Count-1)). Page was not moved."
            }
            else {
                $currentIndex = $pagesList.IndexOf($pageNode)
                if ($currentIndex -eq $targetIndex) {
                    # Already at correct position.
                }
                else {
                    $hierarchy.Section.RemoveChild($pageNode) | Out-Null
                    if ($targetIndex -eq 0) {
                        $hierarchy.Section.PrependChild($pageNode) | Out-Null
                    }
                    else {
                        $beforeNode = $pagesList[$targetIndex]
                        $hierarchy.Section.InsertBefore($pageNode, $beforeNode) | Out-Null
                    }
                    $OneNoteApplication.UpdateHierarchy($hierarchy.OuterXml)
                }
            }
        }

        # Fetch the new page.
        (Get-OneNoteHierarchy `
            -Scope $hsSelf `
            -StartNodeId $newPageId `
            -OneNoteApplication $OneNoteApplication).Page
    }

    end {
        if ($disposeApp -and $OneNoteApplication) {
            Remove-ComObject -ComObject $OneNoteApplication
        }
    }
}
