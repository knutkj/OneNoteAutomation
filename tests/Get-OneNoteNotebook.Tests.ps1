BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../OneNoteAutomation/OneNoteAutomation.psd1') -Force
}

Describe 'Get-OneNoteNotebook cleanup' {
    It 'releases its own application after successfully returning notebooks' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{ Name = 'Test application' }
            $script:notebook = [pscustomobject]@{ Name = 'Test notebook'; ID = 'test-notebook-id' }
            Mock New-OneNoteApplication { $script:createdApplication }
            Mock Get-OneNoteHierarchy {
                [pscustomobject]@{
                    Notebooks = [pscustomobject]@{ Notebook = @($script:notebook) }
                }
            }
            Mock Remove-ComObject {}

            $notebooks = @(Get-OneNoteNotebook)

            $notebooks | Should -HaveCount 1
            [object]::ReferenceEquals($notebooks[0], $script:notebook) | Should -BeTrue
            Should -Invoke New-OneNoteApplication -Times 1 -Exactly -Scope It
            Should -Invoke Get-OneNoteHierarchy -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($OneNoteApplication, $script:createdApplication)
            }
            Should -Invoke Remove-ComObject -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($ComObject, $script:createdApplication)
            }
        }
    }

    It 'does not release a caller-owned application when reading the hierarchy throws' {
        InModuleScope OneNoteAutomation {
            $script:callerApplication = [pscustomobject]@{ Name = 'Caller application' }
            Mock New-OneNoteApplication { throw 'Unexpected application creation' }
            Mock Get-OneNoteHierarchy { throw 'Simulated hierarchy failure' }
            Mock Remove-ComObject {}

            { Get-OneNoteNotebook -App $script:callerApplication } |
            Should -Throw -ExpectedMessage 'Simulated hierarchy failure'

            Should -Invoke New-OneNoteApplication -Times 0 -Exactly -Scope It
            Should -Invoke Get-OneNoteHierarchy -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($OneNoteApplication, $script:callerApplication)
            }
            Should -Invoke Remove-ComObject -Times 0 -Exactly -Scope It
        }
    }

    It 'releases its own application when reading the hierarchy throws' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{ Name = 'Test application' }
            Mock New-OneNoteApplication { $script:createdApplication }
            Mock Get-OneNoteHierarchy { throw 'Simulated hierarchy failure' }
            Mock Remove-ComObject {}

            { Get-OneNoteNotebook } | Should -Throw -ExpectedMessage 'Simulated hierarchy failure'

            Should -Invoke New-OneNoteApplication -Times 1 -Exactly -Scope It
            Should -Invoke Get-OneNoteHierarchy -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($OneNoteApplication, $script:createdApplication)
            }
            Should -Invoke Remove-ComObject -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($ComObject, $script:createdApplication)
            }
        }
    }
}