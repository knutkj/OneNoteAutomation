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
        $parameterSet = $PSCmdlet.ParameterSetName
        $app = $OneNoteApplication
        $notebooks = @()

        if ($NotebookName) {
            $notebooks = Get-OneNoteNotebook -Name $NotebookName -App $App
        }
        elseif ($Notebook) {
            $notebooks = @($Notebook)
        }
        else {
            # If no notebook specified, get all sections from root hierarchy.
            $rootHierarchy = Get-OneNoteHierarchy `
                -Scope $hsSections `
                -StartNodeId $null `
                -OneNoteApplication $App

            $sections = $rootHierarchy.Notebooks.Notebook.Section |
            Where-Object -Property Name -Like -Value $Name

            # Tag each section with custom type name
            return $sections | ForEach-Object `
                -Process { $_.PSTypeNames.Insert(0, 'OneNote.Section'); $_ }
        }

        foreach ($nb in $notebooks) {
            $hierarchy = Get-OneNoteHierarchy `
                -Scope $hsSections `
                -StartNodeId $nb.ID `
                -OneNoteApplication $App

            $sections = $hierarchy.Notebook.Section |
            Where-Object -Property Name -Like -Value $Name

            # Tag each section with custom type name.
            return $sections | ForEach-Object `
                -Process { $_.PSTypeNames.Insert(0, 'OneNote.Section'); $_ }
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