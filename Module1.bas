Attribute VB_Name = "Module1"

Option Explicit

Dim fso As Object

Sub ConvertFilesToPDF()
    Dim inputFolder As String, outputFolder As String
    Dim selectedTypes As Collection
    Dim answer As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Select input folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Files to Convert"
        If .Show <> -1 Then Exit Sub
        inputFolder = .SelectedItems(1)
    End With
    
    ' Select output folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder to Save PDFs"
        If .Show <> -1 Then Exit Sub
        outputFolder = .SelectedItems(1)
    End With
    
    If Right(inputFolder, 1) <> "\" Then inputFolder = inputFolder & "\"
    If Right(outputFolder, 1) <> "\" Then outputFolder = outputFolder & "\"
    
    ' Ask which file types to convert
    Set selectedTypes = New Collection
    answer = MsgBox("Convert Excel files? (Yes = OK, No = Cancel)", vbOKCancel + vbQuestion, "Select File Types")
    If answer = vbOK Then selectedTypes.Add "excel"
    
    answer = MsgBox("Convert Word files? (Yes = OK, No = Cancel)", vbOKCancel + vbQuestion, "Select File Types")
    If answer = vbOK Then selectedTypes.Add "word"
    
    answer = MsgBox("Convert PowerPoint files? (Yes = OK, No = Cancel)", vbOKCancel + vbQuestion, "Select File Types")
    If answer = vbOK Then selectedTypes.Add "ppt"
    
    If selectedTypes.count = 0 Then
        MsgBox "No file types selected. Exiting.", vbExclamation
        Exit Sub
    End If
    
    ' Start processing with progress indicator
    Application.StatusBar = "Starting conversion..."
    Application.ScreenUpdating = False
    
    ProcessFolder fso.GetFolder(inputFolder), outputFolder, selectedTypes
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "Conversion complete!", vbInformation
End Sub

Sub ProcessFolder(folder As Object, outputFolder As String, selectedTypes As Collection)
    Dim file As Object, subFolder As Object
    Dim totalFiles As Long, processedFiles As Long
    
    totalFiles = CountFiles(folder, selectedTypes)
    processedFiles = 0
    
    Dim excelApp As Object, wordApp As Object, pptApp As Object
    Dim wb As Workbook, doc As Object, ppt As Object
    Dim pdfPath As String
    
    ' Initialize apps only if needed
    On Error Resume Next
    If CollectionContains(selectedTypes, "word") Then
        Set wordApp = GetObject(, "Word.Application")
        If Err.Number <> 0 Then Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = False
    End If
    Err.Clear
    
    If CollectionContains(selectedTypes, "ppt") Then
        Set pptApp = GetObject(, "PowerPoint.Application")
        If Err.Number <> 0 Then Set pptApp = CreateObject("PowerPoint.Application")
        pptApp.Visible = False
    End If
    Err.Clear
    On Error GoTo 0
    
    ' Process files in current folder
    For Each file In folder.Files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(file.Name))
        
        If CollectionContains(selectedTypes, "excel") And ext Like "xls*" Then
            Application.StatusBar = "Converting Excel: " & file.Name & " (" & processedFiles & "/" & totalFiles & ")"
            Set wb = Workbooks.Open(file.Path, ReadOnly:=True)
            pdfPath = outputFolder & fso.GetBaseName(file.Name) & ".pdf"
            wb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath
            wb.Close False
            processedFiles = processedFiles + 1
        
        ElseIf CollectionContains(selectedTypes, "word") And ext Like "doc*" Then
            Application.StatusBar = "Converting Word: " & file.Name & " (" & processedFiles & "/" & totalFiles & ")"
            Set doc = wordApp.Documents.Open(file.Path, ReadOnly:=True)
            pdfPath = outputFolder & fso.GetBaseName(file.Name) & ".pdf"
            doc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=17
            doc.Close False
            processedFiles = processedFiles + 1
        
        ElseIf CollectionContains(selectedTypes, "ppt") And ext Like "ppt*" Then
            Application.StatusBar = "Converting PowerPoint: " & file.Name & " (" & processedFiles & "/" & totalFiles & ")"
            Set ppt = pptApp.Presentations.Open(file.Path, WithWindow:=msoFalse)
            pdfPath = outputFolder & fso.GetBaseName(file.Name) & ".pdf"
            ppt.SaveAs pdfPath, 32 ' 32 = ppSaveAsPDF
            ppt.Close
            processedFiles = processedFiles + 1
        End If
    Next file
    
    ' Process subfolders recursively
    For Each subFolder In folder.SubFolders
        ProcessFolder subFolder, outputFolder, selectedTypes
    Next subFolder
    
    ' Cleanup
    If Not wordApp Is Nothing Then wordApp.Quit
    If Not pptApp Is Nothing Then pptApp.Quit
    
    Application.StatusBar = False
End Sub

' Helper to count files to process for progress
Function CountFiles(folder As Object, selectedTypes As Collection) As Long
    Dim count As Long
    Dim file As Object, subFolder As Object
    Dim ext As String
    
    count = 0
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If (CollectionContains(selectedTypes, "excel") And ext Like "xls*") _
            Or (CollectionContains(selectedTypes, "word") And ext Like "doc*") _
            Or (CollectionContains(selectedTypes, "ppt") And ext Like "ppt*") Then
            count = count + 1
        End If
    Next file
    
    For Each subFolder In folder.SubFolders
        count = count + CountFiles(subFolder, selectedTypes)
    Next subFolder
    
    CountFiles = count
End Function

' Helper function to check if collection contains a value
Function CollectionContains(col As Collection, val As String) As Boolean
    Dim itm As Variant
    On Error Resume Next
    For Each itm In col
        If LCase(itm) = LCase(val) Then
            CollectionContains = True
            Exit Function
        End If
    Next
    CollectionContains = False
End Function


