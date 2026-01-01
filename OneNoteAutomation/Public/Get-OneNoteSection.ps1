#
# .SYNOPSIS
# Retrieves OneNote sections from a specified notebook using the OneNote COM
# API.
#
# .DESCRIPTION
# Accepts notebook objects from the pipeline or a notebook name via the
# notebook name parameter. Optionally filters for a section by name using
# wildcard patterns and case-insensitive matching. If neither notebook objects
# nor notebook name is provided, retrieves all sections across all notebooks.
#
function Get-OneNoteSection {
    [CmdletBinding(DefaultParameterSetName = 'ByNotebookName')]
    param(
        # Notebook object(s) from the pipeline only.
        [Parameter(
            ValueFromPipeline = $true,
            ParameterSetName = 'FromPipeline',
            Mandatory = $true
        )]
        [System.Xml.XmlElement]$Notebook,

        # Name of the notebook to search.
        [Parameter(ParameterSetName = 'ByNotebookName', Position = 0)]
        [string]$NotebookName,

        # Name of the section to retrieve, supporting wildcards and prefix
        # matching.
        [Parameter(ParameterSetName = 'ByNotebookName', Position = 1)]
        [Parameter(ParameterSetName = 'FromPipeline')]
        [SupportsWildcards()]
        [string]$Name = "*",

        # If specified, retrieves only the currently viewed section in OneNote.
        [Parameter(ParameterSetName = 'Current', Mandatory = $true)]
        [switch]$Current,

        # The OneNote application object. If not provided, it will be created.
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
        $hsSections = 3 # HierarchyScope.hsSections
        $app = $OneNoteApplication
        $sections = @()

        if ($PSCmdlet.ParameterSetName -eq 'Current') {
            # First get the current notebook,
            # then search for current section within it.
            $currentNotebook = Get-OneNoteNotebook -Current -App $app
            $hierarchy = Get-OneNoteHierarchy `
                -Scope $hsSections `
                -StartNodeId $currentNotebook.ID `
                -OneNoteApplication $app

            $sections = @($hierarchy.Notebook.Section |
                Where-Object -Property isCurrentlyViewed -EQ true)

            if ($sections.Count -gt 1) {
                throw "There are currently $($sections.Count) sections that are viewed."
            }
        }
        else {
            # Determine which notebooks to retrieve sections from.
            if ($NotebookName) {
                $notebooks = Get-OneNoteNotebook -Name $NotebookName -App $app
            }
            elseif ($Notebook) {
                $notebooks = @($Notebook)
            }
            else {
                # Get all notebooks
                $notebooks = Get-OneNoteNotebook -App $app
            }

            # Fetch sections for each notebook.
            foreach ($nb in $notebooks) {
                $hierarchy = Get-OneNoteHierarchy `
                    -Scope $hsSections `
                    -StartNodeId $nb.ID `
                    -OneNoteApplication $app

                $sections += @($hierarchy.Notebook.Section)
            }

            # Apply name filter.
            $sections = @($sections | Where-Object -Property Name -Like -Value $Name)
        }

        # Tag each section with custom type name.
        $sections | ForEach-Object -Process { 
            $_.PSTypeNames.Insert(0, 'OneNote.Section'); $_ 
        }
    }

    end {
        if ($comObjectCreated) {
            Remove-ComObject -ComObject $OneNoteApplication
        }
    }
}

Get-Command Get-OneNoteSection | Register-ArgumentCompleterMap -Map @{
    NotebookName = { Get-OneNoteNotebookNameCompletion @args }
    Name         = { Get-OneNoteSectionNameCompletion @args }
}