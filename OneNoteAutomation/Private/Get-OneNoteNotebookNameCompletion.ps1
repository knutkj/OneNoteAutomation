#
# .SYNOPSIS
# Provides argument completion for OneNote notebook names.
#
# .DESCRIPTION
# Internal function used by Register-ArgumentCompleter to provide tab completion
# for notebook name parameters.
#
function Get-OneNoteNotebookNameCompletion {
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
        New-OneNoteCompletion
    }
}