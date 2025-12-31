#
# .SYNOPSIS
# Registers argument completers for a cmdlet's parameters using a hashtable map.
# 
# .DESCRIPTION
# Convenience function to register multiple argument completers for a cmdlet.
# Takes the cmdlet from the pipeline and a hashtable mapping parameter names to 
# completer scriptblocks.
# 
# .EXAMPLE
# Get-Command Get-OneNoteSection | Register-ArgumentCompleterMap -Map @{
#     NotebookName = { __NotebookNameArgumentCompleter @args }
#     Name         = { __SectionNameArgumentCompleter @args }
# }
#
function Register-ArgumentCompleterMap {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [System.Management.Automation.CommandInfo]$Command,

        [Parameter(Mandatory)]
        [hashtable]$Map
    )

    process {
        foreach ($entry in $Map.GetEnumerator()) {
            Register-ArgumentCompleter `
                -CommandName $Command.Name `
                -ParameterName $entry.Key `
                -ScriptBlock $entry.Value
        }
    }
}
