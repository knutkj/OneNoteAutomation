BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../OneNoteAutomation/OneNoteAutomation.psd1') -Force
}

Describe 'Get-OneNoteHierarchy cleanup' {
    It 'releases its own application when retrieving the hierarchy throws' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{
                Name              = 'Test application'
                GetHierarchyCalls = 0
            }
            $script:createdApplication | Add-Member -MemberType ScriptMethod -Name GetHierarchy -Value {
                param($StartNodeId, $Scope, $HierarchyXml)
                $this.GetHierarchyCalls++
                throw 'Simulated hierarchy failure'
            }
            Mock New-OneNoteApplication { $script:createdApplication }
            Mock New-Object { throw 'Unexpected direct COM activation' }
            Mock Remove-ComObject {}

            { Get-OneNoteHierarchy -Scope 0 -StartNodeId 'test-section-id' } |
            Should -Throw -ExpectedMessage '*Simulated hierarchy failure*'

            Should -Invoke New-OneNoteApplication -Times 1 -Exactly -Scope It
            Should -Invoke New-Object -Times 0 -Exactly -Scope It
            $script:createdApplication.GetHierarchyCalls | Should -Be 1
            Should -Invoke Remove-ComObject -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($ComObject, $script:createdApplication)
            }
        }
    }
}