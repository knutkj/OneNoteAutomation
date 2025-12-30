@{
    RootModule           = 'OneNoteAutomation.psm1'
    
    # Module version is dynamically updated from GitHub release tags during CI/CD.
    # Example: release tag "v1.2.3" becomes ModuleVersion "1.2.3"
    # The placeholder value below is overwritten by scripts/Update-ModuleManifest.ps1
    ModuleVersion        = '0.1.0'
    
    GUID                 = '2a6e65f3-7d7b-43bd-9ea2-8ccf1e1d7c2b'
    Author               = 'Knut Kristian Johansen'
    Copyright            = '© 2025 Knut Kristian Johansen'
    Description          = 'PowerShell module for automating Microsoft OneNote using the COM API.
    
See https://github.com/knutkj/OneNoteAutomation for more information.'

    # Compatibility.
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = 'Desktop'

    # Functions are dynamically discovered from Public/*.ps1 files during CI/CD.
    # Convention: one function per file, where filename matches function name.
    # Example: Public/Write-HelloWorld.ps1 exports function "Write-HelloWorld"
    # The empty array below is populated by scripts/Update-ModuleManifest.ps1
    # FunctionsToExport = @()

    # Defines table views for OneNote.Notebook, OneNote.Section, and
    # OneNote.Page custom objects produced by the cmdlets.
    FormatsToProcess     = @('OneNote.format.ps1xml')

    PrivateData          = @{
        PSData = @{
            Tags       = 'OneNote', 'Automation', 'PSEdition_Desktop', 'Windows'
            ProjectUri = 'https://github.com/knutkj/OneNoteAutomation'
            LicenseUri = 'https://github.com/knutkj/OneNoteAutomation/blob/main/LICENSE'

            # Prerelease versions are dynamically added when release tags
            # contain suffixes. Example: tag "v1.2.3-preview1" adds
            # Prerelease = 'preview1'. Standard releases (e.g., "v1.2.3") do not
            # include a Prerelease field.
            # Prerelease = 'preview1'

            # Release notes are dynamically extracted from GitHub release
            # descriptions. The release body text becomes the ReleaseNotes field.
            # ReleaseNotes = 'Release description'
        }
    }
}
