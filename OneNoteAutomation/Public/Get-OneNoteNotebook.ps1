#
# .SYNOPSIS
# Uses the OneNote COM API to enumerate notebooks.
#
# .EXAMPLE
# Get-OneNoteNotebook
#   
# .EXAMPLE
# Get-OneNoteNotebook -Name Personal
#
function Get-OneNoteNotebook {
    [CmdletBinding()]
    param(
        # The name of the notebook to retrieve.
        [Parameter(Position = 0)]
        [SupportsWildcards()]
        [ArgumentCompleter({ __NotebookNameArgumentCompleter @args })]
        [string]$Name = '*',

        # The OneNote application object. If not provided, it will be created.
        [Alias('App')]
        [Parameter()]
        $OneNoteApplication = $null
    )

    begin {
        $hsNotebooks = 2 # HierarchyScope.hsNotebooks
        $disposeApp = $false

        # Instantiate OneNote if needed.
        if (-not $OneNoteApplication) {
            $disposeApp = $true
            $OneNoteApplication = New-Object -ComObject OneNote.Application
        }
    }

    process {
        (Get-OneNoteHierarchy -Scope $hsNotebooks -App $OneNoteApplication).Notebooks.Notebook |
        Where-Object -Property Name -Like -Value $Name |
        ForEach-Object -Process { $_.PSTypeNames.Insert(0, 'OneNote.Notebook'); $_ }
    }

    end { if ($disposeApp) { Remove-ComObject -ComObject $OneNoteApplication } }
}
