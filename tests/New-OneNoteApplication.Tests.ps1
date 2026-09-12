BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../OneNoteAutomation/OneNoteAutomation.psd1') -Force
}

Describe 'New-OneNoteApplication' {
    It 'returns the OneNote COM object created by New-Object' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{ Name = 'Test application' }
            Mock New-Object { $script:createdApplication }

            $application = New-OneNoteApplication

            [object]::ReferenceEquals($application, $script:createdApplication) | Should -BeTrue
            Should -Invoke New-Object -Times 1 -Exactly -Scope It -ParameterFilter {
                $ComObject -eq 'OneNote.Application'
            }
        }
    }
}

Describe 'Use-ComObject creation' {
    It 'preserves a creation failure without attempting cleanup' {
        InModuleScope OneNoteAutomation {
            Mock New-OneNoteApplication { throw 'Simulated application creation failure' }
            Mock New-Object { throw 'Unexpected direct COM activation' }
            Mock Remove-ComObject {}

            { Use-ComObject -ProgId 'OneNote.Application' -Script {
                    throw 'Script must not run when creation fails'
                } } | Should -Throw -ExpectedMessage '*Simulated application creation failure*'

            Should -Invoke New-OneNoteApplication -Times 1 -Exactly -Scope It
            Should -Invoke New-Object -Times 0 -Exactly -Scope It
            Should -Invoke Remove-ComObject -Times 0 -Exactly -Scope It
        }
    }

    It 'uses the OneNote factory and releases its result' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{ Name = 'Test application' }
            Mock New-OneNoteApplication { $script:createdApplication }
            Mock New-Object { throw 'Unexpected direct COM activation' }
            Mock Remove-ComObject {}

            $application = Use-ComObject -ProgId 'onenote.application' -Script {
                param($instance)
                $instance
            }

            [object]::ReferenceEquals($application, $script:createdApplication) | Should -BeTrue
            Should -Invoke New-OneNoteApplication -Times 1 -Exactly -Scope It
            Should -Invoke New-Object -Times 0 -Exactly -Scope It
            Should -Invoke Remove-ComObject -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($ComObject, $script:createdApplication)
            }
        }
    }

    It 'keeps direct COM activation for other ProgIDs' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{ Name = 'Other application' }
            Mock New-OneNoteApplication { throw 'Unexpected OneNote activation' }
            Mock New-Object { $script:createdApplication }
            Mock Remove-ComObject {}

            $application = Use-ComObject -ProgId 'Other.Application' -Script {
                param($instance)
                $instance
            }

            [object]::ReferenceEquals($application, $script:createdApplication) | Should -BeTrue
            Should -Invoke New-OneNoteApplication -Times 0 -Exactly -Scope It
            Should -Invoke New-Object -Times 1 -Exactly -Scope It -ParameterFilter {
                $ComObject -eq 'Other.Application'
            }
            Should -Invoke Remove-ComObject -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($ComObject, $script:createdApplication)
            }
        }
    }
}