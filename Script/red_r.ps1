$folderPath = Read-Host -Prompt "Please enter the folder path"  # Set your folder path here
$searchWords = @("[insert_QR-Logo]")  # Array of words to search for
# $newQRCodePath = "C:\Users\SiddharthSrivastava\Downloads\New folder (2)\SS.png"  # Path to the new QR code image

# Get all Word documents in the folder
$files = Get-ChildItem -Path $folderPath -Filter *.docx

$fileCounter = 0  # Initialize the counter for processed files

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
            # Remove the found word from the document
            $range = $document.Content
            $range.Find.Execute($word, $false, $false, $false, $false, $false, $false, $false, $false, "", 2)  # wdReplaceAll
        }
    }

    # Replace the QR code
  #  $shapes = $document.Shapes
  #  foreach ($shape in $shapes) {
   #     if ($shape.Type -eq [Microsoft.Office.Interop.Word.WdShapeType]::wdInlineShapePicture) {
    #        # Check if the shape is a QR code (you may need to adjust this condition)
     #       if ($shape.AlternativeText -eq "QR Code") {  # Assuming you have set an alternative text for the QR code
      #          $shape.Delete()  # Remove the old QR code
       #         $newQRCode = $document.InlineShapes.AddPicture($newQRCodePath, $false, $true, $shape.Range)  # Add new QR code
        #        $newQRCode.Width = $shape.Width  # Set the new QR code width
         #       $newQRCode.Height = $shape.Height  # Set the new QR code height
          #      break  # Exit the loop after replacing the QR code
           # }
        #}
    #}

    # Save the document with the same name
    $document.Save()
    # Close the Word document
    $document.Close()
    $wordApp.Quit()
}

Write-Host "Total files processed: $fileCounter" -ForegroundColor Green