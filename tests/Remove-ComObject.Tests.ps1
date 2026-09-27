BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../OneNoteAutomation/OneNoteAutomation.psd1') -Force
}

Describe 'Remove-ComObject' {
    It 'disposes a managed application without attempting COM release' {
        $application = [System.IO.MemoryStream]::new()
        try {
            Remove-ComObject -ComObject $application

            $application.CanRead | Should -BeFalse
        }
        finally {
            $application.Dispose()
        }
    }
}