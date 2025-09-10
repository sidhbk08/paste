Attribute VB_Name = "UHG_CMOS_ExpPDF_CC"
' Working for UHG Dental3430
' Updated to allow the macro to work from any of the subfolder within the PM_ folder, and output will always be the PM_XX's Folder-4.
' Updated to automatically export as PDF into folder 5
' Updated to use PDFMaker to export as PDFs
' Updated to remove "_HIDDEN" and "_ERROR" from the filename
' Updated to detect if the 'StartTagline' exist, and sort it out if not.
' Updated to remove all active hyperlinks.
' Updated to unhide & fix font formatting
' Yin Khong (yin.khong@biglanguage.com)
' Last updated: Jan14, 2025
' Requires PDFMaker reference enabled.
' Requires PDF Maker conversion's preferences set
 
 
Sub UHG_COSMOS_ExpPDF_CC(SourceDir As String, TargetDir As String, Parameters As String)
    subName = "UHG_COSMOS_ExportPDF_CC"
    Dim rng As Range
    Dim i As Integer
    Dim curr_doc As Document
    Dim curr_path As String, curr_file As String
    curr_path = SourceDir & "\"
    'curr_path = "Z:\UHC-COSMOS3784\Received - To Be Translated\20250109\B5752-STD-SPA-SPA-LP-168\PM_01\4-QC-ed Word files\"
    error_Path = curr_path + "Requires_Review\"
    curr_parentFolderPath = Left(curr_path, InStr(curr_path, "PM_") + 4) + "\"
    PDFOutput_Path = curr_parentFolderPath + "6-Ready for delivery\"
    If Dir(PDFOutput_Path, vbDirectory) = "" Then
        MsgBox "6-Ready for delivery folder could not be located."
        End
    End If
    curr_file = Dir(curr_path & "*.doc?")
    unhiddenFileCount = 0

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .MultiLine = False
        .Pattern = "(.*)\d{8}\\(.*)\\PM"
    End With
    Set matchCollection = regex.Execute(curr_path)
    If matchCollection.Count > 0 Then
        projInfo = Split(matchCollection(0).Submatches(1), "-")
        If Not (UBound(projInfo) = 5) Then
          MsgBox "Incorrect folder path! Terminating..."
          Exit Sub
        Else
          receivedFolder = matchCollection(0).Submatches(0)
          FolderName = matchCollection(0).Submatches(1)
          LOBID = projInfo(1)
          sourceLang = projInfo(2)
          targetLang = projInfo(3)
          fontType = projInfo(4)
        End If
    End If

    Set matchCollection = Nothing
    Dim mainFSO As New FileSystemObject
    Set mainFSO = CreateObject("Scripting.FileSystemObject")
    exportPath = curr_path & "ExportedPDF\"
    If Not (mainFSO.FolderExists(exportPath)) Then
        MkDir exportPath
    End If
    macroLogPath = exportPath & "Log_Macro_ExpPDF" & Format(Now(), "yyMMdd_hhmmss") & ".log"
    Set txtstream = mainFSO.CreateTextFile(macroLogPath, True, True)
    txtstream.Write Format(Now(), "yyyyMMdd hh:mm:ssAM/PM") & vbNewLine & Environ("USERNAME") & vbNewLine & subName & vbNewLine
    txtstream.Write "---------" & vbNewLine

    While curr_file <> ""
        Set curr_doc = Documents.Open(curr_path + curr_file)
        filenameOri = Left(curr_doc.Name, InStrRev(curr_doc.Name, ".") - 1)
        FileName = Replace(filenameOri, "_TEMPLATED", "")
        FileName = Replace(FileName, "_REVIEW", "")
        FileName = Replace(FileName, "_HIDDEN", "")
        FileName = Replace(FileName, "_ERROR", "")
        FileName = Replace(FileName, "_FixBkmrk", "")
        FileName = Replace(FileName, "_ReviewPHI", "")
        FileName = Replace(FileName, "_ReviewADDR", "")
        FileName = Replace(FileName, "_Review_NoSourceFile", "")
        FileName = Replace(FileName, "_ReviewPHI-and-ADDR", "")
        FileName = Replace(FileName, "_FixCostTable", "")
        FileName = Replace(FileName, "_MissingCostTable", "")
        FileName = Replace(FileName, "_VARIABLETEXT", "")

        alphanumFont = "Arial"
        targetFont = "Arial"
        If targetLang = "YUE" Then
            targetFont = "Microsoft YaHei"
        ElseIf targetLang = "CMN" Then
            targetFont = "SimSun"
        ElseIf targetLang = "KOR" Then
            targetFont = "Batang"
        End If

        For i = curr_doc.Sections.Count To 1 Step -1
            Set rng = curr_doc.Sections(i).Footers(wdHeaderFooterPrimary).Range
            curr_doc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Font.Name = targetFont
            curr_doc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Font.Name = alphanumFont
            curr_doc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Font.Name = alphanumFont
            Unhide_AllText rng
            Set rng = Nothing
            Set rng = curr_doc.Sections(i).Footers(wdHeaderFooterFirstPage).Range
            curr_doc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Font.Name = targetFont
            curr_doc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Font.Name = alphanumFont
            curr_doc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Font.Name = alphanumFont
            Unhide_AllText rng

            Set rng = Nothing
            Set rng = curr_doc.Sections(i).Footers(wdHeaderFooterEvenPages).Range
            curr_doc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Font.Name = targetFont
            curr_doc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Font.Name = alphanumFont
            curr_doc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Font.Name = alphanumFont
            Unhide_AllText rng

            Set rng = Nothing
            Set rng = curr_doc.Sections(i).Headers(wdHeaderFooterPrimary).Range
            curr_doc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Font.Name = targetFont
            curr_doc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Font.Name = alphanumFont
            curr_doc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Font.Name = alphanumFont
            Unhide_AllText rng

            Set rng = Nothing
            Set rng = curr_doc.Sections(i).Headers(wdHeaderFooterFirstPage).Range
            curr_doc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Font.Name = targetFont
            curr_doc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Font.Name = alphanumFont
            curr_doc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Font.Name = alphanumFont
            Unhide_AllText rng

            Set rng = Nothing
            Set rng = curr_doc.Sections(i).Headers(wdHeaderFooterEvenPages).Range
            curr_doc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Font.Name = targetFont
            curr_doc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Font.Name = alphanumFont
            curr_doc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Font.Name = alphanumFont
            Unhide_AllText rng

            Set rng = Nothing
        Next

        Set rng = curr_doc.Content
        curr_doc.Range.Font.Name = targetFont
        curr_doc.Range.Font.Name = alphanumFont
        Unhide_AllText rng

        If curr_doc.Comments.Count > 0 Then
            curr_doc.DeleteAllComments
        End If

        curr_doc.Save
        '----- Save the target file
        placeholderError = 0
        With curr_doc.Content.find
            .text = "[insert_"
            .ClearFormatting
            .Wrap = wdFindStop
            .MatchCase = CaseSensitive
            Do While .Execute
                placeholderError = placeholderError + 1
            Loop
        End With

        With curr_doc.Content.find
            .text = "{{"
            .MatchWildcards = False
            .ClearFormatting
            .Wrap = wdFindStop
            .MatchCase = CaseSensitive
            Do While .Execute
                placeholderError = placeholderError + 1
            Loop
        End With

        With curr_doc.Content.find
            .text = "}}"
            .MatchWildcards = False
            .ClearFormatting
            .Wrap = wdFindStop
            .MatchCase = CaseSensitive
            Do While .Execute
                placeholderError = placeholderError + 1
            Loop
        End With

        If placeholderError > 0 Then
            If Not (mainFSO.FolderExists(error_Path)) Then
                MkDir error_Path
            End If
            ActiveDocument.Close SaveChanges:=False
            mainFSO.MoveFile curr_path & curr_file, error_Path
            GoTo partC
        End If

 
        Dim pdfname, i2, a
        Dim pmkr As AdobePDFMakerForOffice.PDFMaker
        Dim stng As AdobePDFMakerForOffice.ISettings
        'If Not ActiveWorkbook.Saved Then
        'MsgBox "You must save the document before converting it to PDF", vbOKOnly, ""
        'Exit Sub
        'End If
        Set pmkr = Nothing ' locate PDFMaker object
        For Each a In Application.COMAddIns
            If InStr(UCase(a.Description), "PDFMAKER") > 0 Then
                Set pmkr = a.Object
                Exit For
            End If
        Next

        If pmkr Is Nothing Then
            MsgBox "Error! Unable to trigger pdf maker! Terminating..."
            Exit Sub
        End If

        Set fso = New Scripting.FileSystemObject
        'pdfname = ActiveDocument.Name
        i2 = InStrRev(curr_file, ".")
        pdfname = PDFOutput_Path & IIf(i2 = 0, FileName, Left(FileName, i2 - 1)) & ".pdf"
        ' delete PDF file if it exists
        'If Dir(pdfname) <> "" Then Kill pdfname
        pmkr.GetCurrentConversionSettings stng
        'stng.AddBookmarks = True
        'stng.AddLinks = True
        'stng.AddTags = False
        stng.ConvertAllPages = True
        stng.FitToOnePage = False
        'stng.CreateFootnoteLinks = True
        'stng.CreateXrefLinks = True
        stng.OutputPDFFileName = pdfname
        stng.PromptForPDFFilename = False
        stng.ShouldShowProgressDialog = True
        stng.ViewPDFFile = False
        pmkr.CreatePDF 0 ' perform conversion

        If Not (fso.FileExists(pdfname)) Then       ' see if conversion failed
            txtStreamLog.Write "*** Warning: " & FileName & " failed PDF creation..."
            Application.Quit
        End If
        unhiddenFileCount = unhiddenFileCount + 1
PartB:
        ActiveDocument.Close SaveChanges:=True
        mainFSO.MoveFile curr_path & curr_file, exportPath
        Set curr_doc = Nothing

partC:
        curr_file = Dir
    Wend

    If failedFile > 0 Then
        'MsgBox "WARNING: " & failedFile & " file(s) failed PDF Conversion!" & vbNewLine & vbNewLine & unhiddenFileCount & " files have been exported as PDF."
        txtstream.Write "WARNING: " & failedFile & " file(s) failed PDF Conversion (due to missing bookmark). " & vbNewLine & vbNewLine & unhiddenFileCount & " files have been exported as PDF."
    Else
        'MsgBox unhiddenFileCount & " files have been unhidden and exported as PDF." & vbNewLine & vbNewLine & "NOTE:" & vbNewLine & "Output files are saved in the \6-Ready for delivery"
        txtstream.Write "All " & unhiddenFileCount & " file(s) converted PDF into Folder-6."
    End If
    txtstream.Close
End Sub
 
Private Function Unhide_AllText(ByRef r As Range)
    Dim s As Shape
    Dim s_rng As ShapeRange
    r.Font.hidden = False
    Set s_rng = r.ShapeRange
    If s_rng.Count > 0 Then
        For Each s In s_rng
            If s.TextFrame.HasText Then
                s.TextFrame.TextRange.Font.hidden = False
            End If
            Set s = Nothing
        Next
    End If
    Set s_rng = Nothing
End Function
