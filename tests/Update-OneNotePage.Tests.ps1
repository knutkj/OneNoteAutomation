BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../OneNoteAutomation/OneNoteAutomation.psd1') -Force
}

Describe 'Update-OneNotePage cleanup' {
    It 'releases its own application when updating page content throws' {
        InModuleScope OneNoteAutomation {
            $script:createdApplication = [pscustomobject]@{
                Name                   = 'Test application'
                UpdatePageContentCalls = 0
            }
            $script:createdApplication | Add-Member -MemberType ScriptMethod -Name UpdatePageContent -Value {
                param($PageXml)
                $this.UpdatePageContentCalls++
                throw 'Simulated page update failure'
            }
            Mock New-OneNoteApplication { $script:createdApplication }
            Mock Test-OneNotePageHasContent { $true }
            Mock Remove-ComObject {}

            $document = [System.Xml.XmlDocument]::new()
            $page = $document.CreateElement('one', 'Page', 'http://schemas.microsoft.com/office/onenote/2013/onenote')
            $page.SetAttribute('ID', 'test-page-id')
            [void]$document.AppendChild($page)

            { Update-OneNotePage -PageDocument $document } |
            Should -Throw -ExpectedMessage '*Simulated page update failure*'

            Should -Invoke New-OneNoteApplication -Times 1 -Exactly -Scope It
            Should -Invoke Test-OneNotePageHasContent -Times 1 -Exactly -Scope It
            $script:createdApplication.UpdatePageContentCalls | Should -Be 1
            Should -Invoke Remove-ComObject -Times 1 -Exactly -Scope It -ParameterFilter {
                [object]::ReferenceEquals($ComObject, $script:createdApplication)
            }
        }
    }
}