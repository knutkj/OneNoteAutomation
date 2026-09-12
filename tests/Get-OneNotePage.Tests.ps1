BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../OneNoteAutomation/OneNoteAutomation.psd1') -Force
}

Describe 'Get-OneNotePage cleanup' {
    It 'releases its own application when reading the page hierarchy throws' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{ Name = 'Test application' }
            Mock New-OneNoteApplication { $script:createdApplication }
            Mock Get-OneNoteHierarchy { throw 'Simulated hierarchy failure' }
            Mock Remove-ComObject {}

            { Get-OneNotePage -Id 'test-page-id' } |
            Should -Throw -ExpectedMessage 'Simulated hierarchy failure'

            Should -Invoke New-OneNoteApplication -Times 1 -Exactly -Scope It
            Should -Invoke Get-OneNoteHierarchy -Times 1 -Exactly -Scope It -ParameterFilter {
                $Scope -eq 4 -and
                [object]::ReferenceEquals($OneNoteApplication, $script:createdApplication)
            }
            Should -Invoke Remove-ComObject -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($ComObject, $script:createdApplication)
            }
        }
    }
}