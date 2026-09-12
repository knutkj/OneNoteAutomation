BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../OneNoteAutomation/OneNoteAutomation.psd1') -Force
}

Describe 'Show-OneNote cleanup' {
    It 'releases its own application when pipeline navigation fails with ErrorAction Stop' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{
                Name = 'Test application'
                NavigateToCalls = 0
            }
            $script:createdApplication | Add-Member -MemberType ScriptMethod -Name NavigateTo -Value {
                param($EntityId)
                $this.NavigateToCalls++
                throw 'Simulated pipeline navigation failure'
            }
            Mock New-OneNoteApplication { $script:createdApplication }
            Mock New-Object { throw 'Unexpected direct COM activation' }
            Mock Remove-ComObject {}

            $entity = [pscustomobject]@{ ID = 'test-page-id' }
            { $entity | Show-OneNote -ErrorAction Stop } |
                Should -Throw -ExpectedMessage '*Simulated pipeline navigation failure*'

            Should -Invoke New-OneNoteApplication -Times 1 -Exactly -Scope It
            Should -Invoke New-Object -Times 0 -Exactly -Scope It
            $script:createdApplication.NavigateToCalls | Should -Be 1
            Should -Invoke Remove-ComObject -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($ComObject, $script:createdApplication)
            }
        }
    }

    It 'releases its own application when navigation fails with ErrorAction Stop' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{
                Name = 'Test application'
                NavigateToCalls = 0
            }
            $script:createdApplication | Add-Member -MemberType ScriptMethod -Name NavigateTo -Value {
                param($EntityId)
                $this.NavigateToCalls++
                throw 'Simulated navigation failure'
            }
            Mock New-OneNoteApplication { $script:createdApplication }
            Mock Remove-ComObject {}

            { Show-OneNote -ID 'test-page-id' -ErrorAction Stop } |
                Should -Throw -ExpectedMessage '*Simulated navigation failure*'

            Should -Invoke New-OneNoteApplication -Times 1 -Exactly -Scope It
            $script:createdApplication.NavigateToCalls | Should -Be 1
            Should -Invoke Remove-ComObject -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($ComObject, $script:createdApplication)
            }
        }
    }
}