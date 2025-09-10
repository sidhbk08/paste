$folderPath = Read-Host -Prompt "Please enter the folder path"  # Set your folder path here

# Get all Word documents in the folder
$files = Get-ChildItem -Path $folderPath -Filter *.docx

foreach ($file in $files) {
    # Load the Word application
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $document = $wordApp.Documents.Open($file.FullName)

    # Locate the QR code
    $shapes = $document.Shapes
    foreach ($shape in $shapes) {
        if ($shape.Type -eq [Microsoft.Office.Interop.Word.WdShapeType]::wdInlineShapePicture) {
            # Check if the shape is a QR code (you may need to adjust this condition)
            if ($shape.AlternativeText -ne "QR Code") {  # If alternative text is not set
                $shape.AlternativeText = "QR Code"  # Set the alternative text
            }
        }
    }

    # Save the document with the changes
    $document.Save()
    # Close the Word document
    $document.Close()
    $wordApp.Quit()
}

Write-Host "Alternative text set for QR codes in documents." -ForegroundColor Green