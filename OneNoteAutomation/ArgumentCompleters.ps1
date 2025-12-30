function __SectionNameArgumentCompleter {
    param(
        $CommandName,
        $ParameterName,
        $WordToComplete,
        $CommandAst,
        $FakeBoundParameters
    )

    $hsSections = 3 # HierarchyScope.hsSections
    $notebook = $FakeBoundParameters['NotebookName']
    # if (-not $notebook) { return }

    Use-ComObject -ProgId OneNote.Application -Script {
        param($app)

        if ($notebook) {
            $notebook = Get-OneNoteNotebook -Name $notebook -App $app
            $sections = (Get-OneNoteHierarchy -Scope $hsSections `
                    -StartNodeId $notebook.ID -App $app).Notebook.Section
        }
        else {
            $sections = (Get-OneNoteHierarchy -Scope $hsSections -App $app
            ).Notebooks.Notebook.Section
        }

        $sections |
        Where-Object -FilterScript {
            $toMatch = $WordToComplete -replace "'", ''
            $_.Name -like $toMatch -or
            $_.Name.StartsWith($toMatch, [StringComparison]::OrdinalIgnoreCase)
        } |
        New-OneNoteCompletionResult
    }
}

function __NotebookNameArgumentCompleter {
    param(
        $CommandName,
        $ParameterName,
        $WordToComplete,
        $CommandAst,
        $FakeBoundParameters
    )

    $hsNotebooks = 2 # HierarchyScope.hsNotebooks

    Use-ComObject -ProgId OneNote.Application -Script {
        param($app)

        (Get-OneNoteHierarchy -Scope $hsNotebooks -App $app).Notebooks.Notebook |
        Where-Object -FilterScript {
            $toMatch = $WordToComplete -replace "'", ''
            $_.Name -like $toMatch -or
            $_.Name.StartsWith($toMatch, [StringComparison]::OrdinalIgnoreCase)
        } |
        New-OneNoteCompletionResult
    }
}

filter New-OneNoteCompletionResult {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Id,

        [Parameter(ValueFromPipelineByPropertyName = $true)]
        [string]$Name
    )

    [CompletionResult]::new(
        "'$Name'", $Name, [CompletionResultType]::ParameterValue, "ID: '$Id'.")
}
