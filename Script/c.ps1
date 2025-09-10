# Create a new Word application instance
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Open the source and template documents
$sourceDoc = $word.Documents.Open("C:\Users\SiddharthSrivastava\Downloads\TestS.docx")
$templateDoc = $word.Documents.Open("C:\Users\SiddharthSrivastava\Downloads\TestT.docx")

# Initialize variables
$startText = "You Have the Right to Appeal Our Decision"
$endText = "Plan share may include a reduction for sequestration"
$copying = $false

# Iterate through paragraphs in the source document
foreach ($paragraph in $sourceDoc.Paragraphs) {
    if ($paragraph.Range.Text -like "*$startText*") {
        $copying = $true  # Start copying text
    }
    if ($copying) {
        # Copy the paragraph to the template document
        $newParagraph = $templateDoc.Content.Paragraphs.Add()
        $newParagraph.Range.FormattedText = $paragraph.Range.FormattedText
    }
    if ($paragraph.Range.Text -like "*$endText*") {
        break  # Stop copying after reaching the end text
    }
}

# Save the modified template document
$document.Save()
# Close the documents and quit Word
$sourceDoc.Close()
$templateDoc.Close()
$word.Quit()