Attribute VB_Name = "Module1"
Option Explicit

' Global FileSystemObject for reuse
Dim fso As Object

' Main sub called by UserForm
Public Sub ConvertFilesToPDF(inputFolder As String, outputFolder As String, _
                             convertExcel As Boolean, convertWord As Boolean, convertPPT As Boolean, _
                             frm As Object)
    Dim selectedTypes As Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set selectedTypes = New Collection
    
    If Right(inputFolder, 1) <> "\" Then inputFolder = inputFolder & "\"
    If Right(outputFolder, 1) <> "\" Then outputFolder = outputFolder & "\"
    
    If convertExcel Then selectedTypes.Add "excel"
    If convertWord Then selectedTypes.Add "word"
    If convertPPT Then selectedTypes.Add "ppt"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Initializing..."
    
    Dim wordApp As Object, pptApp As Object
    
    On Error Resume Next
    ' Open Word if needed
    If CollectionContains(selectedTypes, "word") Then
        Set wordApp = GetObject(, "Word.Application")
        If Err.Number <> 0 Then Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = False
    End If
    Err.Clear
    
    ' Open PowerPoint if needed
    If CollectionContains(selectedTypes, "ppt") Then
        Set pptApp = GetObject(, "PowerPoint.Application")
        If Err.Number <> 0 Then Set pptApp = CreateObject("PowerPoint.Application")
        pptApp.Visible = False
    End If
    Err.Clear
    On Error GoTo 0
    
    ' Start recursive processing
    ProcessFolder fso.GetFolder(inputFolder), outputFolder, selectedTypes, wordApp, pptApp, frm
    
    ' Cleanup
    If Not wordApp Is Nothing Then wordApp.Quit
    If Not pptApp Is Nothing Then pptApp.Quit
    
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Sub ProcessFolder(folder As Object, outputFolder As String, selectedTypes As Collection, _
                          wordApp As Object, pptApp As Object, frm As Object)
    Dim file As Object, subFolder As Object
    Dim wb As Workbook, doc As Object, ppt As Object
    Dim pdfPath As String
    Dim totalFiles As Long, processedFiles As Long
    
    totalFiles = CountFiles(folder, selectedTypes)
    processedFiles = 0
    
    For Each file In folder.Files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(file.Name))
        
        If CollectionContains(selectedTypes, "excel") And ext Like "xls*" Then
            frm.lblStatus.Caption = "Converting Excel: " & file.Name
            frm.Repaint
            Set wb = Workbooks.Open(file.Path, ReadOnly:=True)
            pdfPath = outputFolder & fso.GetBaseName(file.Name) & ".pdf"
            wb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath
            wb.Close False
            processedFiles = processedFiles + 1
        
        ElseIf CollectionContains(selectedTypes, "word") And ext Like "doc*" Then
            frm.lblStatus.Caption = "Converting Word: " & file.Name
            frm.Repaint
            Set doc = wordApp.Documents.Open(file.Path, ReadOnly:=True)
            pdfPath = outputFolder & fso.GetBaseName(file.Name) & ".pdf"
            doc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=17
            doc.Close False
            processedFiles = processedFiles + 1
        
        ElseIf CollectionContains(selectedTypes, "ppt") And ext Like "ppt*" Then
            frm.lblStatus.Caption = "Converting PowerPoint: " & file.Name
            frm.Repaint
            Set ppt = pptApp.Presentations.Open(file.Path, WithWindow:=msoFalse)
            pdfPath = outputFolder & fso.GetBaseName(file.Name) & ".pdf"
            ppt.SaveAs pdfPath, 32 ' ppSaveAsPDF
            ppt.Close
            processedFiles = processedFiles + 1
        End If
    Next file
    
    For Each subFolder In folder.SubFolders
        ProcessFolder subFolder, outputFolder, selectedTypes, wordApp, pptApp, frm
    Next subFolder
    
End Sub

Private Function CountFiles(folder As Object, selectedTypes As Collection) As Long
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

Private Function CollectionContains(col As Collection, val As String) As Boolean
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

Sub openform()
frmConvertPDF.Show
End Sub
Sub OpenFolder()
Dim folderPath As String
folderPath = txtOutputFolder.Text
    If folderPath = "" Then
        MsgBox "Folder path is empty!", vbExclamation
        Exit Sub
    End If
    
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Folder does not exist: " & folderPath, vbExclamation
        Exit Sub
    End If
    
    ' Use Shell to open folder in Explorer
    
End Sub

