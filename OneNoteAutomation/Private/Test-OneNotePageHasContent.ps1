#
# .SYNOPSIS
# Tests whether a OneNote page element or document has full content.
#
# .DESCRIPTION
# Checks if the input is a full page with content structures (from
# Get-OneNotePage -Content) or a lightweight metadata-only element (from
# hierarchy). Accepts both XmlElement and XmlDocument inputs.
#
# For XmlElement: checks that OwnerDocument.DocumentElement is the Page element.
# For XmlDocument: checks that DocumentElement is a Page element.
#
# .OUTPUTS
# System.Boolean. Returns $true if the page has full content, $false if
# lightweight or invalid.
#
filter Test-OneNotePageHasContent {
    $input = $_
    
    if ($input -is [System.Xml.XmlDocument]) {
        # Document: check if root element is a Page
        return $input.DocumentElement.LocalName -eq 'Page'
    }
    elseif ($input -is [System.Xml.XmlElement]) {
        # Element: check if it's the document root and is a Page
        $doc = $input.OwnerDocument
        return $doc.DocumentElement.LocalName -eq 'Page' -and
               $doc.DocumentElement -eq $input
    }
    
    return $false
}
