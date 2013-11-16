Attribute VB_Name = "BatchPDFConverter"

Sub BatchPDFConverter()

Dim sourceFolder As String
Dim saveFolder As String
Dim vFile As String
Dim saveFile As String
Dim myTitle As String
Dim myMsg As String
Dim ChrChk As String
Dim i As Integer
Dim j As Integer

On Error GoTo ErrHandler

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
            
Application:
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

vFile = Dir(sourceFolder & "\" & "*.doc")

Do While vFile <> ""
    Documents.Open FileName:=sourceFolder & "\" & vFile
    ChrChk = Right(vFile, 4)
    If ChrChk = ".doc" Then
        saveFile = Left(vFile, Len(vFile) - 4)
        ActiveDocument.ExportAsFixedFormat OutputFileName:=saveFolder & "\" & saveFile & ".pdf", _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
        ActiveDocument.Close wdDoNotSaveChanges
        vFile = Dir
    Else
        saveFile = Left(vFile, Len(vFile) - 5)
        ActiveDocument.ExportAsFixedFormat OutputFileName:=saveFolder & "\" & saveFile & ".pdf", _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument, Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
        ActiveDocument.Close wdDoNotSaveChanges
        vFile = Dir
    End If
Loop

vFile = Dir(sourceFolder & "\" & "*.doc")
Do While Len(vFile) > 0
    i = i + 1
    vFile = Dir
Loop

strFile = Dir(saveFolder & "\" & "*.pdf")
Do While Len(strFile) > 0
    j = j + 1
    strFile = Dir
Loop
MsgBox (j & " PDF files were created from " & i & " Word documents")

Exit Sub

ErrHandler:
MsgBox "An error occurred trying to convert the document(s) to PDF " & Chr(13) & _
       "Error Number: " & Err.Number & Chr(13) & _
       "Description: " & Err.description, vbOKOnly + vbCritical, "Error"
Exit Sub

End Sub
