BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../OneNoteAutomation/OneNoteAutomation.psd1') -Force
}

Describe 'Get-OneNoteSection cleanup' {
    It 'releases its own application when reading the section hierarchy throws' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{ Name = 'Test application' }
            Mock New-OneNoteApplication { $script:createdApplication }
            Mock Get-OneNoteHierarchy { throw 'Simulated hierarchy failure' }
            Mock Remove-ComObject {}

            { Get-OneNoteSection -Id 'test-section-id' } |
            Should -Throw -ExpectedMessage 'Simulated hierarchy failure'

            Should -Invoke New-OneNoteApplication -Times 1 -Exactly -Scope It
            Should -Invoke Get-OneNoteHierarchy -Times 1 -Exactly -Scope It -ParameterFilter {
                $Scope -eq 0 -and $StartNodeId -eq 'test-section-id' -and
                [object]::ReferenceEquals($OneNoteApplication, $script:createdApplication)
            }
            Should -Invoke Remove-ComObject -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($ComObject, $script:createdApplication)
            }
        }
    }
}