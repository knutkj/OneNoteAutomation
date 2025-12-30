using namespace System.Management.Automation

#
# .SYNOPSIS
# Shows/navigates to a OneNote entity (notebook, section, page, ...) in the
# OneNote application.
#
# .DESCRIPTION
# Takes OneNote objects from the pipeline or by ID parameter and navigates to
# them in the OneNote application window. This command works with notebooks,
# sections, and pages.
#
# When multiple entities are piped to the command, an Out-GridView picker is
# displayed allowing the user to choose which entity to navigate to. Use the
# -First switch to skip the picker and automatically navigate to the first
# entity.
#
# .EXAMPLE
# # Shows an Out-GridView picker if multiple notebooks are found.
# Get-OneNoteNotebook | Show-OneNote
#
# .EXAMPLE
# # Navigates to the first section found without showing a picker.
# Get-OneNoteSection -NotebookName "Work" | Show-OneNote -First
#
# .EXAMPLE
# # Navigates directly to a specific OneNote entity using its ID.
# Show-OneNote -ID "..."
#
function Show-OneNote {
    [CmdletBinding(DefaultParameterSetName = 'ByObject')]
    param(
        # OneNote object from the pipeline (Notebook or Section).
        [Parameter(
            ValueFromPipeline = $true,
            ParameterSetName = 'ByObject'
        )]
        [PSObject]$InputObject,

        # The ID of the OneNote entity to navigate to.
        [Parameter(
            ParameterSetName = 'ById',
            Mandatory = $true,
            Position = 0
        )]
        [string]$ID,

        # Skip the picker and always navigate to the first item when multiple
        # items are piped.
        [Parameter()]
        [switch]$First,

        # The OneNote application object. If not provided, it will be created.
        [Alias('App')]
        [Parameter()]
        $OneNoteApplication = $null
    )

    begin {
        $disposeApp = $false
        $allEntities = @()

        # Instantiate OneNote if needed.
        if (-not $OneNoteApplication) {
            $disposeApp = $true
            $OneNoteApplication = New-Object -ComObject OneNote.Application
        }
    }

    process {
        # For the ID parameter set, we navigate immediately
        if ($PSCmdlet.ParameterSetName -eq 'ById') {
            $app = $OneNoteApplication
            try {
                $app.NavigateTo($ID)
                Write-Verbose "Navigated to OneNote entity with ID: $ID."
                return
            }
            catch {
                Write-Error "Failed to navigate to OneNote entity: $_."
                return
            }
        }

        # For pipeline input, collect entities as-is for potential picker selection
        if ($InputObject -and (Get-Member -InputObject $InputObject -Name "ID" -MemberType Properties)) {
            # Just add the original object to our collection
            $allEntities += $InputObject
        }
        elseif ($InputObject) {
            Write-Warning "Input object does not have an ID property. Cannot navigate to it."
        }
    }

    end {
        $app = $OneNoteApplication

        # Handle the collected entities based on count and parameters
        if ($allEntities.Count -gt 0) {
            # When multiple entities are found and -First isn't specified, use Out-GridView picker
            if ($allEntities.Count -gt 1 -and -not $First) {
                # Use Out-GridView with the original objects - simple and idiomatic
                $selectedEntity = $allEntities | Out-GridView -Title "Select a OneNote entity to navigate to" -OutputMode Single

                if ($selectedEntity) {
                    $targetId = $selectedEntity.ID
                }
                else {
                    Write-Verbose "No entity was selected. Navigation canceled."
                    if ($disposeApp) { Remove-ComObject -ComObject $OneNoteApplication }
                    return
                }
            }
            # If only one entity was collected or -First is specified, take the first one
            else {
                $targetId = $allEntities[0].ID
                Write-Verbose "Navigating to entity with ID: $targetId"
            }

            # Navigate to the selected entity
            if ($targetId) {
                try {
                    $app.NavigateTo($targetId)
                    Write-Verbose "Navigated to OneNote entity with ID: $targetId"
                }
                catch {
                    Write-Error "Failed to navigate to OneNote entity: $_"
                }
            }
        }
        else {
            # No entities to navigate to
            if ($PSCmdlet.ParameterSetName -eq 'ByObject') {
                Write-Warning "No valid OneNote entities were provided to navigate to."
            }
        }

        if ($disposeApp) {
            Remove-ComObject -ComObject $OneNoteApplication
        }
    }
}
