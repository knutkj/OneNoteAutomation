BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../OneNoteAutomation/OneNoteAutomation.psd1') -Force
}

Describe 'Get-OneNoteSectionNameCompletion cleanup' {
    It 'releases its application when looking up the notebook throws' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{ Name = 'Test application' }
            Mock New-OneNoteApplication { $script:createdApplication }
            Mock New-Object { throw 'Unexpected direct COM activation' }
            Mock Get-OneNoteNotebook { throw 'Simulated notebook lookup failure' }
            Mock Get-OneNoteHierarchy { throw 'Unexpected hierarchy lookup' }
            Mock Remove-ComObject {}

            { Get-OneNoteSectionNameCompletion -WordToComplete 'Test' -FakeBoundParameters @{
                    NotebookName = 'Test notebook'
                } } | Should -Throw -ExpectedMessage '*Simulated notebook lookup failure*'

            Should -Invoke New-OneNoteApplication -Times 1 -Exactly -Scope It
            Should -Invoke New-Object -Times 0 -Exactly -Scope It
            Should -Invoke Get-OneNoteNotebook -Times 1 -Exactly -Scope It -ParameterFilter {
                $Name -eq 'Test notebook' -and
                [object]::ReferenceEquals($OneNoteApplication, $script:createdApplication)
            }
            Should -Invoke Get-OneNoteHierarchy -Times 0 -Exactly -Scope It
            Should -Invoke Remove-ComObject -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($ComObject, $script:createdApplication)
            }
        }
    }
}