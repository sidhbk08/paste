$folderPath = Read-Host -Prompt "Please enter the folder path"
$searchWords = @(
    "[insert_CustomerServiceNum]", "[insert_FastAppealPhoneVerbiage]", "[insert_PlanName]",
    "[insert_PlanNameCapitalized]", "[insert_DenialCodeDescription]", "[insert_QR Logo]",
    "[insert_FastAppealHours]", "[insert_FastAppealDays]", "[insert_DecisionDays]",
    "[insert_TollFreeTTY]", "[insert_TollFreeNumber]", "[insert_PlanName2]",
    "[insert_CustomerServiceTTY]", "[insert_CustomerServiceNumFull]", "[insert_PlanAppealDays]",
    "[insert_StandardAppealDays]", "[insert_WrittenDecisionDays]", "[insert_StandardAppealMailingAddr]",
    "[insert_StandardAppealDeliveryAddr]", "[insert_StandardAppealFax]", "[insert_TollFreeVerbiage]",
    "[insert_KeyTerms]"
)

# Excel Setup
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $false
$workbook = $excelApp.Workbooks.Add()
$worksheet = $workbook.Sheets.Item(1)
$worksheet.Cells.Item(1, 1).Value = "Filename"
$worksheet.Cells.Item(1, 2).Value = "Found Words"

# Start Word once
$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false
$wordApp.DisplayAlerts = 0

$files = Get-ChildItem -Path $folderPath -Filter *.docx
$row = 2
$fileCounter = 0

foreach ($file in $files) {
    $fileCounter++
    $document = $wordApp.Documents.Open($file.FullName, $false, $true)

    # Collect all searchable text (main content + headers + footers)
    $fullText = $document.Content.Text

    foreach ($section in $document.Sections) {
        foreach ($header in $section.Headers) {
            $fullText += "`n" + $header.Range.Text
        }
        foreach ($footer in $section.Footers) {
            $fullText += "`n" + $footer.Range.Text
        }
    }

    # Check for search terms
    $foundWords = @()
    foreach ($word in $searchWords) {
        if ($fullText -match [regex]::Escape($word)) {
            $foundWords += $word
        }
    }

    if ($foundWords.Count -gt 0) {
        $worksheet.Cells.Item($row, 1).Value = $file.Name
        $worksheet.Cells.Item($row, 2).Value = [string]::Join(", ", $foundWords)
        $row++
    }

    $document.Close($false)
}

$wordApp.Quit()

# Save Excel
$excelFilePath = "C:\Users\SiddharthSrivastava\OneDrive - BIG Language Solutions\Book1.xlsx"
$workbook.SaveAs($excelFilePath)
$workbook.Close($false)
$excelApp.Quit()

Write-Host "Total files processed: $fileCounter" -ForegroundColor Green
Write-Host "Results saved to: $excelFilePath" -ForegroundColor Green
