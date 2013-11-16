Attribute VB_Name = "BatchPDFConverter"

Sub BatchPDFConverter()

'''''DATE OF CHANGE: 11/16/2013 @ 17:42 EDT
'''''AUTHOR: UDIT NARAYANAN'''''

'''''INITIALIZATION'''''
Dim sourceFolder As String
Dim saveFolder As String
Dim sourceFile As String
Dim saveFile As String
Dim myTitle As String
Dim myMsg As String
Dim chrChk As String
Dim i As Integer
Dim j As Integer
''''''''''

On Error GoTo ErrHandler

'''''FIRST DIALOG BOX TO START THE APPLICATION'''''
myTitle = "Batch convert MS Word documents to PDF"
myMsg = "This operation will convert Word documents stored in a folder to PDF." & _
" Have you moved all the required MS Word documents into one single folder?"
Response = MsgBox(myMsg, vbExclamation + vbYesNoCancel, myTitle)
Select Case Response
    Case Is = vbYes
        GoTo Application
    Case Is = vbNo
        ActiveDocument.Close SaveChanges:=False
        GoTo ErrHandler
    Case Is = vbCancel
        Exit Sub
End Select
''''''''''

'''''<<<<<MAIN CODE OF THE APPLICATION'''''
Application:
'''''LOCATION OF THE FOLDERS USED IN THIS MACRO'''''
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Navigate to the the folder where the MS Word documents are located"
    .Show
    sourceFolder = .SelectedItems.Item(1)
End With
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Select the folder where the converted PDF files should be saved"
    .Show
    saveFolder = .SelectedItems.Item(1)
End With
''''''''''

'''''CONVERSION TO PDF'''''
sourceFile = Dir(sourceFolder & "\" & "*.doc")
Do While sourceFile <> ""
    Documents.Open FileName:=sourceFolder & "\" & sourceFile
    chrChk = Right(sourceFile, 4)
    If chrChk = ".doc" Then
        saveFile = Left(sourceFile, Len(sourceFile) - 4)
        ActiveDocument.ExportAsFixedFormat OutputFileName:=saveFolder & "\" & saveFile & ".pdf", _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
        ActiveDocument.Close wdDoNotSaveChanges
        sourceFile = Dir
    Else
        saveFile = Left(sourceFile, Len(sourceFile) - 5)
        ActiveDocument.ExportAsFixedFormat OutputFileName:=saveFolder & "\" & saveFile & ".pdf", _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
        ActiveDocument.Close wdDoNotSaveChanges
        sourceFile = Dir
    End If
Loop
''''''''''

'''''COUNT AND DISPLAY THE NUMBER OF FILES CONVERTED TO PDF'''''
sourceFile = Dir(sourceFolder & "\" & "*.doc")
Do While Len(sourceFile) > 0
    i = i + 1
    sourceFile = Dir
Loop
saveFile = Dir(saveFolder & "\" & "*.pdf")
Do While Len(saveFile) > 0
    j = j + 1
    saveFile = Dir
Loop
MsgBox (j & " PDF files were created from " & i & " Word documents")
''''''''''

Exit Sub
'''''MAIN CODE OF THE APPLICATION>>>>>'''''

'''''ERROR HANDLER'''''
ErrHandler:
MsgBox "An error occurred trying to convert the document(s) to PDF " & Chr(13) & _
       "Error Number: " & Err.Number & Chr(13) & _
       "Description: " & Err.description, vbOKOnly + vbCritical, "Error"
Exit Sub
''''''''''

End Sub
