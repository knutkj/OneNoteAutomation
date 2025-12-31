#
# .SYNOPSIS
# Provides argument completion for OneNote section names.
#
# .DESCRIPTION
# Internal function used by Register-ArgumentCompleter to provide tab completion
# for section name parameters. Filters sections by notebook context if specified.
#
function Get-OneNoteSectionNameCompletion {
    param(
        $CommandName,
        $ParameterName,
        $WordToComplete,
        $CommandAst,
        $FakeBoundParameters
    )

    $hsSections = 3 # HierarchyScope.hsSections
    $notebook = $FakeBoundParameters['NotebookName']

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
        New-OneNoteCompletion
    }
}