#
# .SYNOPSIS
# Retrieves the OneNote hierarchy as an XML object using the OneNote COM API.
#
# .DESCRIPTION
# Connects to the OneNote COM API and retrieves the hierarchy XML for the
# specified node and scope. Returns the hierarchy as a PowerShell XML object.
# The scope parameter determines how deep into the hierarchy to retrieve,
# from just the specified node to all pages in all notebooks.
#
# .EXAMPLE
# # Get all notebooks from root.
# $hierarchy = Get-OneNoteHierarchy -Scope 2
# $notebooks = $hierarchy.Notebooks.Notebook
#
# .EXAMPLE
# # Get all sections within a specific notebook.
# $hierarchy = Get-OneNoteHierarchy -Scope 3 -StartNodeId $notebookId
# $sections = $hierarchy.Notebook.Section
#
# .EXAMPLE
# # Get all pages within a specific section.
# $hierarchy = Get-OneNoteHierarchy -Scope 4 -StartNodeId $sectionId
# $pages = $hierarchy.Section.Page
#
# .EXAMPLE
# # Get just the specified node without descendants
# $node = Get-OneNoteHierarchy -Scope 0 -StartNodeId $pageId
#
# .OUTPUTS
# System.Xml.XmlDocument. Containing the OneNote hierarchy XML structure. The
# structure varies based on scope:
# - Scope 2: Contains Notebooks with nested Section and Page elements.
# - Scope 3: Contains Notebook/Section elements.
# - Scope 4: Contains Section/Page elements.
#
# .NOTES
# This is a low-level function used by other OneNote cmdlets. Most users should
# use Get-OneNoteNotebook, Get-OneNoteSection, or Get-OneNotePage instead.
# The COM object is automatically managed when not provided.
#
function Get-OneNoteHierarchy {
    [CmdletBinding()]
    param(
        # The hierarchy scope that specifies the lowest level to get in the
        # notebook node hierarchy. Valid values:
        #   0 (hsSelf) - Gets just the start node specified and no descendants.
        #   1 (hsChildren) - Gets immediate child nodes of the start node only.
        #   2 (hsNotebooks) - Gets all notebooks below the start node or root.
        #   3 (hsSections) - Gets all sections below the start node, including
        #     sections in section groups.
        #   4 (hsPages) - Gets all pages below the start node, including all
        #     pages in section groups.
        [Parameter(Mandatory = $true)]
        [int]$Scope, # HierarchyScope

        # The ID of the node to start from. Use $null or omit to start from the
        # root. When specified, retrieves hierarchy starting from that specific
        # notebook, section, or page node.
        [Alias('Node')]
        [Parameter()]
        [string]$StartNodeId = $null,

        # An existing OneNote.Application COM object. If not provided, a new COM
        # object will be created using Use-ComObject for automatic cleanup.
        [Alias('App')]
        [Parameter()]
        $OneNoteApplication = $null
    )

    $script = {
        param($App)
        $xml = ''
        $App.GetHierarchy($StartNodeId, $Scope, [ref]$xml)
        [xml]$doc = $xml
        return $doc
    }

    if (-not $OneNoteApplication) {
        . $PSScriptRoot/Use-ComObject.ps1
        return Use-ComObject -ProgId OneNote.Application -Script $script
    }
    else {
        return & $script -App $OneNoteApplication
    }
}
