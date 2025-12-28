# Getting Started with OneNoteAutomation PowerShell Module

## Overview

OneNoteAutomation is an open-source PowerShell module (not affiliated with or
supported by Microsoft) that provides a set of cmdlets for automating OneNote
operations from PowerShell scripts. Instead of working with COM objects
directly, you can use familiar PowerShell patterns like piping, filtering, and
parameter binding for a more natural scripting experience.

## Prerequisites

- **PowerShell** on Windows.
- **Microsoft OneNote** desktop client application.

## Installation

1. Install the module from the PowerShell Gallery:

   ```powershell
   Install-Module -Name OneNoteAutomation
   ```

2. Import the module:

   ```powershell
   Import-Module OneNoteAutomation
   ```

3. Verify the installation:
   ```powershell
   Get-OneNoteNotebook
   ```
   This should display your OneNote notebooks without errors.

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

### Get Pages from a Section

```powershell
Get-OneNoteNotebook "Diary" |
  Get-OneNoteSection -Name "Daily" |
  Get-OneNotePage
```

Or using argument completion for a more concise approach:

```powershell
Get-OneNoteSection Diary Daily | Get-OneNotePage
```

### Get a Specific Page with Content

```powershell
Get-OneNotePage -ID $sectionId -Name "2025-01-15" -Content
```

The `-Content` switch retrieves the page's XML content, useful for modification.

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
$page = Get-OneNotePage -ID $section.ID -Name "2025-01-15" -Content

# Modify the XML as needed
$page.Content.Page.Title.OE.T.'#cdata-section' = "New Title"

# Update the page
Update-OneNotePage -ID $page.ID -Content $page.Content
```

## Tips and Best Practices

### Store Objects for Later Use

Store objects themselves (not just IDs) since most cmdlets support pipeline by
value:

```powershell
$notebook = Get-OneNoteNotebook "Diary"

# Later, pipe the object to retrieve sections within that notebook.
$sections = $notebook | Get-OneNoteSection

# Or pipe directly to Show-OneNote
$notebook | Show-OneNote
```

This approach is more idiomatic PowerShell and leverages the pipeline-first
design of the module.

### Use Argument Completion

Argument completers suggest available notebook and section names:

```powershell
Get-OneNoteSection -NotebookName <Tab>  # Suggests notebook names
```

### Handle XML Content Carefully

Page content is returned as XML. When modifying:

1. Always retrieve content with `-Content` flag.
2. Use XPath queries or XML navigation to find elements.
3. Update and save back with `Update-OneNotePage`.

Example:

```powershell
$page = Get-OneNotePage -ID $sectionId -Name "MyPage" -Content
$xml = $page.Content
$xml.Page.Title.OE.T.'#cdata-section' = "Updated Title"
Update-OneNotePage -ID $page.ID -Content $xml
```

### Logging and Debugging

Use `-Verbose` to see detailed operation logs:

```powershell
Update-OneNotePage -ID $pageId -Content $content -Verbose
```

## Next Steps

- Read full cmdlet help: `Get-Help Get-OneNoteNotebook -Full`.
- Explore the hierarchy:
  `Get-OneNoteNotebook | Get-OneNoteSection | Get-OneNotePage`.
- Build automation scripts for your OneNote workflows.
- Contribute improvements or new functionality.

## Resources

- [OneNote developer reference | Microsoft Learn](https://learn.microsoft.com/en-us/office/client-developer/onenote/onenote-developer-reference)
