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
# .EXAMPLE
# # Get a specific notebook by ID.
# Get-OneNoteNotebook -Id '{12345678-1234-5678-9012-...}{1}{...}'
#
function Get-OneNoteNotebook {
    [CmdletBinding()]
    param(
        # If specified, retrieves only the currently viewed notebook in OneNote.
        [Parameter(ParameterSetName = 'Current', Mandatory = $true)]
        [switch]$Current,

        # The ID of a specific notebook to retrieve.
        [Parameter(ParameterSetName = 'Id', Mandatory = $true)]
        [string]$Id,

        # The name of the notebook to retrieve.
        [Parameter(ParameterSetName = 'Name', Position = 0)]
        [SupportsWildcards()]
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
        $notebooks = @()

        if ($Current) {
            $hierarchy = Get-OneNoteHierarchy -Scope $hsNotebooks -App $OneNoteApplication
            $notebooks = @($hierarchy.Notebooks.Notebook |
                Where-Object -Property isCurrentlyViewed -EQ true)

            if ($notebooks.Count -gt 1) {
                throw "There are currently $($notebooks.Count) notebooks that are viewed."
            }
        }
        elseif ($Id) {
            # Get notebook by ID using hsSelf scope.
            $hsSelf = 0 # HierarchyScope.hsSelf
            $hierarchy = Get-OneNoteHierarchy -Scope $hsSelf -StartNodeId $Id -App $OneNoteApplication
            $notebooks = @($hierarchy.Notebook)
        }
        else {
            $hierarchy = Get-OneNoteHierarchy -Scope $hsNotebooks -App $OneNoteApplication
            $notebooks = @($hierarchy.Notebooks.Notebook |
                Where-Object -Property Name -Like -Value $Name)
        }

        $notebooks |
        ForEach-Object -Process { $_.PSTypeNames.Insert(0, 'OneNote.Notebook'); $_ }
    }

    end { if ($disposeApp) { Remove-ComObject -ComObject $OneNoteApplication } }
}

Get-Command Get-OneNoteNotebook | Register-ArgumentCompleterMap -Map @{
    Name = { Get-OneNoteNotebookNameCompletion @args }
}
