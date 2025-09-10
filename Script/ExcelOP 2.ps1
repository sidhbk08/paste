$folderPath = "C:\Users\SiddharthSrivastava\Downloads\New folder"  # Set your folder path here
$searchWords = @("[insert_CustomerServiceNum]", "[insert_FastAppealPhoneVerbiage]", "[insert_PlanName]", "[insert_PlanNameCapitalized]", "[insert_DenialCodeDescription]", "[insert_QR-Logo]", "[insert_FastAppealHours]", "[insert_FastAppealDays]", "[insert_DecisionDays]", "[insert_TollFreeTTY]", "[insert_TollFreeNumber]", "[insert_PlanName2]", "[insert_CustomerServiceTTY]" , "[insert_CustomerServiceNumFull]", "[insert_PlanAppealDays]", "[insert_StandardAppealDays]", "[insert_WrittenDecisionDays]", "[insert_StandardAppealMailingAddr]", "[insert_StandardAppealDeliveryAddr]", "[insert_StandardAppealFax]", "[insert_TollFreeVerbiage]")  # Array of words to search for

# Create Excel application and a new workbook
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $false  # Set this to $false if you don't want Excel to be visible
$workbook = $excelApp.Workbooks.Add()
$worksheet = $workbook.Sheets.Item(1)

# Set header row in Excel sheet
$worksheet.Cells.Item(1, 1).Value = "Filename"
$worksheet.Cells.Item(1, 2).Value = "Found Words"

# Get all Word documents in the folder
$files = Get-ChildItem -Path $folderPath -Filter *.docx

$fileCounter = 0  # Initialize the counter for processed files
$row = 2  # Start writing from row 2 in Excel

foreach ($file in $files) {
    $fileCounter++  # Increment the counter for each file processed

    # Load the Word application
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $document = $wordApp.Documents.Open($file.FullName)

    # Initialize a list to keep track of found words for the current document
    $foundWords = @()

    foreach ($word in $searchWords) {
        $found = $document.Content.Find.Execute($word)
        if ($found) {
            $foundWords += $word  # Add the found word to the list
        }
    }

    # If any words were found, write them to the Excel sheet
    if ($foundWords.Count -gt 0) {
        $worksheet.Cells.Item($row, 1).Value = $file.Name  # Write the filename
        $worksheet.Cells.Item($row, 2).Value = [string]::Join(", ", $foundWords)  # Join found words with commas and write to Excel
        $row++  # Move to the next row in Excel
        Write-Host "Found words in file: $($file.Name)" -ForegroundColor Green
    }

    # Close the Word document
    $document.Close()
    $wordApp.Quit()
}

# Save the Excel workbook to a file
$excelFilePath = "C:\Users\SiddharthSrivastava\Downloads\New folder"
$workbook.SaveAs($excelFilePath)
$workbook.Close()
$excelApp.Quit()

Write-Host "Total files processed: $fileCounter" -ForegroundColor Green
Write-Host "Results saved to: $excelFilePath" -ForegroundColor Green
