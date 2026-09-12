BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../OneNoteAutomation/OneNoteAutomation.psd1') -Force
}

Describe 'Get-OneNoteNotebookNameCompletion cleanup' {
    It 'releases its application when retrieving the hierarchy throws' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{ Name = 'Test application' }
            Mock New-OneNoteApplication { $script:createdApplication }
            Mock New-Object { throw 'Unexpected direct COM activation' }
            Mock Get-OneNoteHierarchy { throw 'Simulated hierarchy failure' }
            Mock Remove-ComObject {}

            { Get-OneNoteNotebookNameCompletion -WordToComplete 'Test' } |
            Should -Throw -ExpectedMessage '*Simulated hierarchy failure*'

            Should -Invoke New-OneNoteApplication -Times 1 -Exactly -Scope It
            Should -Invoke New-Object -Times 0 -Exactly -Scope It
            Should -Invoke Get-OneNoteHierarchy -Times 1 -Exactly -Scope It -ParameterFilter {
                $Scope -eq 2 -and
                [object]::ReferenceEquals($OneNoteApplication, $script:createdApplication)
            }
            Should -Invoke Remove-ComObject -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($ComObject, $script:createdApplication)
            }
        }
    }
}