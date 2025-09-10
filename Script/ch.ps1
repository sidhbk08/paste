$folderPath = Read-Host -Prompt "Please enter the folder path"  # Set your folder path here
$findText = "[insert_QR-Logo]"  # Text to find
$replaceText = ""  # Text to replace with

# Get all Word documents in the folder
$files = Get-ChildItem -Path $folderPath -Filter *.docx

foreach ($file in $files) {
    # Load the Word application
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $document = $wordApp.Documents.Open($file.FullName)

    # Find and replace specific words
    $find = $document.Content.Find
    $find.Text = $findText
    $find.Replacement.Text = $replaceText
    $find.Execute()  # Execute the find and replace

    # Locate the QR code
    $shapes = $document.Shapes
    foreach ($shape in $shapes) {
        if ($shape.Type -eq [Microsoft.Office.Interop.Word.WdShapeType]::wdInlineShapePicture) {
            # Replace the QR code with a new one (example: replace with a new image)
            $newQRImagePath = "C:\Users\SiddharthSrivastava\Downloads\New folder (2)\SS.png"  # Path to the new QR code image
            $shape.Delete()  # Remove the old QR code
            $document.Shapes.AddPicture($newQRImagePath, $false, $true, $shape.Left, $shape.Top, $shape.Width, $shape.Height)  # Add new QR code
        }
    }

    # Save the document with the changes
    $document.Save()
    # Close the Word document
    $document.Close()
    $wordApp.Quit()
}

Write-Host "Words replaced and QR codes updated in documents." -ForegroundColor Green