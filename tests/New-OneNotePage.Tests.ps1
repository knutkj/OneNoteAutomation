BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../OneNoteAutomation/OneNoteAutomation.psd1') -Force
}

Describe 'New-OneNotePage cleanup' {
    It 'releases its own application when creating the page throws' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{
                Name               = 'Test application'
                CreateNewPageCalls = 0
            }
            $script:createdApplication | Add-Member -MemberType ScriptMethod -Name CreateNewPage -Value {
                param($SectionId, $PageId, $PageStyle)
                $this.CreateNewPageCalls++
                throw 'Simulated page creation failure'
            }
            Mock New-OneNoteApplication { $script:createdApplication }
            Mock Remove-ComObject {}

            { New-OneNotePage -Id 'test-section-id' } |
            Should -Throw -ExpectedMessage '*Simulated page creation failure*'

            Should -Invoke New-OneNoteApplication -Times 1 -Exactly -Scope It
            $script:createdApplication.CreateNewPageCalls | Should -Be 1
            Should -Invoke Remove-ComObject -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($ComObject, $script:createdApplication)
            }
        }
    }
}