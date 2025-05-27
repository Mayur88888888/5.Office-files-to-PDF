Attribute VB_Name = "Module1"
Option Explicit

Sub ConvertFilesToPDF()
    Dim inputFolder As String, outputFolder As String
    Dim fso As Object, folder As Object, file As Object
    Dim excelApp As Object, wordApp As Object, pptApp As Object
    Dim wb As Workbook
    Dim doc As Object, ppt As Object
    Dim pdfPath As String
    MsgBox "Now provide the input folder where all Excel, Word, Powerpoint files are saved", vbInformation, "Input folder"
    
    ' Prompt user to select input folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Files to Convert"
        If .Show <> -1 Then Exit Sub
        inputFolder = .SelectedItems(1)
    End With
  MsgBox "Now provide the Output folder where PDF files to be saved", vbInformation, "Output folder"
    ' Prompt user to select output folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder to Save PDFs"
        If .Show <> -1 Then Exit Sub
        outputFolder = .SelectedItems(1)
    End With

    If Right(inputFolder, 1) <> "\" Then inputFolder = inputFolder & "\"
    If Right(outputFolder, 1) <> "\" Then outputFolder = outputFolder & "\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(inputFolder)

    ' Process Excel Files
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) Like "xls*" Then
            Set wb = Workbooks.Open(file.Path)
            pdfPath = outputFolder & fso.GetBaseName(file.Name) & ".pdf"
            wb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath
            wb.Close False
        End If
    Next file

    ' Start Word
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set wordApp = CreateObject("Word.Application")
    End If
    wordApp.Visible = False
    On Error GoTo 0

    ' Process Word Files
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) Like "doc*" Then
            Set doc = wordApp.Documents.Open(file.Path, ReadOnly:=True)
            pdfPath = outputFolder & fso.GetBaseName(file.Name) & ".pdf"
            doc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=17
            doc.Close False
        End If
    Next file

    ' Start PowerPoint
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If Err.Number <> 0 Then
        Set pptApp = CreateObject("PowerPoint.Application")
    End If
    pptApp.Visible = True
    On Error GoTo 0

    ' Process PowerPoint Files
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) Like "ppt*" Then
            Set ppt = pptApp.Presentations.Open(file.Path, WithWindow:=msoFalse)
            pdfPath = outputFolder & fso.GetBaseName(file.Name) & ".pdf"
            ppt.SaveAs pdfPath, 32 ' 32 = ppSaveAsPDF
            ppt.Close
        End If
    Next file

    ' Cleanup
    If Not wordApp Is Nothing Then wordApp.Quit
    If Not pptApp Is Nothing Then pptApp.Quit

    MsgBox "Conversion complete!", vbInformation
End Sub

