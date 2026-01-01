[![Publish PowerShell Module](https://github.com/knutkj/OneNoteAutomation/actions/workflows/publish.yml/badge.svg)](https://github.com/knutkj/OneNoteAutomation/actions/workflows/publish.yml)
[![PowerShell Gallery Compatibility](https://img.shields.io/powershellgallery/p/OneNoteAutomation)](https://www.powershellgallery.com/packages/OneNoteAutomation)

# Getting started with OneNoteAutomation

OneNoteAutomation is an open-source PowerShell module (not affiliated with or
supported by Microsoft) that provides a set of cmdlets for automating OneNote
operations from PowerShell scripts. The module is published to the
[PowerShell Gallery](https://www.powershellgallery.com/packages/OneNoteAutomation)
(Microsoft's official repository for PowerShell modules) and can be installed
with `Install-Module`. Instead of working with COM objects directly, you can use
familiar PowerShell patterns like piping, filtering, and parameter binding for a
more natural scripting experience.

## Prerequisites

- **PowerShell 5.1** on Windows.
- **Microsoft OneNote** desktop client application.

## Installation

**Assumptions:** [Execution policy](#execution-policy) is `RemoteSigned` or less
restrictive, [NuGet provider](#nuget-provider) is installed, and
[PSGallery](#psgallery-trust) is trusted or you accept the prompt.

### Quick Start (Admin)

1. Open PowerShell as Administrator.

2. Install and verify:
   ```powershell
   Install-Module -Name OneNoteAutomation
   Get-OneNoteNotebook
   ```

The module auto-imports when you run any of its cmdlets. For non-admin
installation, use `Install-Module -Name OneNoteAutomation -Scope CurrentUser`.

### Setup Details

If installation fails, check the assumptions below.

#### Execution Policy

PowerShell's execution policy must allow running scripts. Check with
`Get-ExecutionPolicy`. If it returns `Restricted`, run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### NuGet Provider

The first time you install any module from the PowerShell Gallery, PowerShell
needs the NuGet provider to download packages. If prompted, type `Y` to install
it.

#### PSGallery Trust

PSGallery (PowerShell Gallery) is not trusted by default as a security
measure—you must confirm you want to install code from the internet. Type `Y`
when prompted, or trust it permanently:

```powershell
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
```

## Core Concepts

### COM Object Lifecycle

The module handles COM object cleanup automatically—each cmdlet can create its
own OneNote.Application COM object if you don't provide one. However, for better
control and to reuse a single connection across multiple cmdlets, use the
`Use-ComObject` cmdlet:

```powershell
Use-ComObject -ProgId OneNote.Application -Script {
    param($OneNote)
    $sections = Get-OneNoteNotebook -App $OneNote |
        Get-OneNoteSection -App $OneNote
    # ... more operations
}
```

`Use-ComObject` implements a C#-like `using` pattern: it creates the COM object,
passes it to your script block, and guarantees cleanup in a `finally` block even
if an error occurs. This is the recommended approach for batch operations.

### PowerShell Pipeline

Most cmdlets support pipeline input, allowing you to chain operations:

```powershell
Get-OneNoteNotebook "Work" |
  Get-OneNoteSection |
  Show-OneNote
```

This pipeline retrieves all sections from the "Work" notebook and displays an
Out-GridView picker to select which section to navigate to in the OneNote
application. It demonstrates how cmdlets compose together: each cmdlet passes
its output to the next, creating a fluent workflow.

### Hierarchy Levels

OneNote has a hierarchy: **Notebooks** → **Sections** → **Pages**

- Notebooks are top-level containers.
- Sections live inside notebooks (or section groups).
- Pages live inside sections.

## Common Workflows

### List All Notebooks

```powershell
Get-OneNoteNotebook
```

### Find a Specific Notebook

```powershell
Get-OneNoteNotebook -Name "Diary"
```

Wildcards are supported:

```powershell
Get-OneNoteNotebook -Name "Work*"
```

### Get All Sections in a Notebook

```powershell
Get-OneNoteNotebook "Diary" | Get-OneNoteSection
```

### Get the Current Page

Get the currently active page in OneNote:

```powershell
Get-OneNotePage -Current
```

### Get Pages from a Section

```powershell
Get-OneNoteNotebook "Diary" |
  Get-OneNoteSection -Name "Daily" |
  Get-OneNotePage
```

Or use `Get-OneNoteSection` directly for a more convenient approach:

```powershell
Get-OneNoteSection -NotebookName "Diary" -Name "Daily" | Get-OneNotePage
```

This skips the need to call `Get-OneNoteNotebook` first, allowing you to specify
both the notebook and section name in a single command.

### Get a Specific Page with Content

```powershell
$page = Get-OneNotePage -Name "2025-01-15" -Content
```

The `-Content` switch retrieves the full page XML element instead of lightweight
metadata. This is required for inspecting or modifying page content. This
example assumes there's only one page with that name across all your notebooks
and sections.

### Navigate to a Page in the UI

```powershell
Get-OneNoteNotebook "Diary" |
  Get-OneNoteSection -Name "Daily" |
  Get-OneNotePage -Name "2025-01-15" |
  Show-OneNote
```

This retrieves a specific page and navigates to it in the OneNote application
window.

### Create a New Page

```powershell
$section = Get-OneNoteSection -NotebookName "Diary" -Name "Daily"
$newPage = New-OneNotePage -Id $section.ID -Title "2025-01-15"
```

Create a page with a specific level (subpage):

```powershell
Get-OneNoteSection -NotebookName "Diary" -Name "Daily" |
  Select-Object -First 1 |  # Select the first matching section (in case multiple exist).
  New-OneNotePage -Title "Subtopic" -PageLevel 2 |
  Show-OneNote
```

### Update Page Content

```powershell
$page = Get-OneNotePage -Current -Content

# Modify the XML as needed.
$page.Title.OE.T.'#cdata-section' = "New Title"

# Update the page.
$page | Update-OneNotePage
```

## Samples

The [`samples/`](samples/) directory contains reusable scripts demonstrating
common page manipulation patterns. These scripts accept page elements from the
pipeline and pass them through, enabling powerful composition:

```powershell
Use-ComObject -ProgId OneNote.Application -Script {
    param($app)
    Get-OneNotePage -Current -Content -App $app |
        Set-OneNotePageToc -App $app |
        Set-OneNotePageSpacing -App $app |
        Update-OneNotePage -App $app
}
```

This pipeline retrieves the current page, generates a table of contents from h1
headings, applies consistent heading spacing, and saves the changes—all in a
single composable workflow with a shared COM object.

- [**Set-OneNotePageToc**](samples/Set-OneNotePageToc.ps1) — Creates or updates
  a clickable table of contents from h1 headings.
- [**Set-OneNotePageSpacing**](samples/Set-OneNotePageSpacing.ps1) — Applies
  consistent spacing to h1 headings.

## Tips and Best Practices

### Store Objects for Later Use

Store objects themselves (not just IDs) since most cmdlets support pipeline by
value:

```powershell
$notebook = Get-OneNoteNotebook "Diary"

# Later, pipe the object to retrieve sections within that notebook.
$sections = $notebook | Get-OneNoteSection

# Or pipe directly to Show-OneNote.
$notebook | Show-OneNote
```

This approach is more idiomatic PowerShell and leverages the pipeline-first
design of the module.

### Use Argument Completion

The module leverages PowerShell's argument completion feature to provide
intelligent suggestions. Press Tab or Ctrl+Space after parameter names to see
all available options queried directly from your OneNote data. When you type
partial values, the suggestions automatically narrow to matching entries:

```powershell
Get-OneNoteSection -NotebookName <Tab>  # Suggests notebook names.
Get-OneNoteSection -Name <Tab>          # Suggests section names.
```

The completers work together contextually. When you specify a notebook name
first, the section name completion filters suggestions to only that notebook:

```powershell
# Shows only sections in "Work" notebook.
Get-OneNoteSection -NotebookName "Work" -Name <Tab>
```

This integration makes the available parameter values more discoverable and
reduces the need to remember exact notebook or section names.

For maximum convenience, you can also use positional parameters with argument
completion:

```powershell
# First parameter is notebook name, second is section name.
Get-OneNoteSection <Tab> <Tab>  # Suggests notebook names, then section names.
```

This allows for very concise commands once you're familiar with the parameter
order, combining the brevity of positional parameters with the discoverability
of argument completion.

### Handle XML Content Carefully

Page content is returned as an XML element. When modifying:

1. Always retrieve with `-Content` flag to get the full page XML element.
2. Use XPath queries or XML navigation to find elements.
3. Pipe the modified element to `Update-OneNotePage`.

Example:

```powershell
$page = Get-OneNotePage -Name "MyPage" -Content  # Assumes unique page title.
$page.Title.OE.T.'#cdata-section' = "Updated Title"
$page | Update-OneNotePage
```

### Logging and Debugging

Use `-Verbose` to see detailed operation logs:

```powershell
Update-OneNotePage -Content $content -Verbose
```

## Next Steps

- Read full cmdlet help: `Get-Help Get-OneNoteNotebook -Full`.
- Explore the hierarchy:
  `Get-OneNoteNotebook | Get-OneNoteSection | Get-OneNotePage`.
- Build automation scripts for your OneNote workflows.
- Contribute improvements or new functionality.

## Resources

- [PowerShell Gallery OneNoteAutomation Package][1]
- [GitHub Repository][2]
- [OneNote developer reference | Microsoft Learn][3]

[1]: https://www.powershellgallery.com/packages/OneNoteAutomation
[2]: https://github.com/knutkj/OneNoteAutomation
[3]:
  https://learn.microsoft.com/en-us/office/client-developer/onenote/onenote-developer-reference
