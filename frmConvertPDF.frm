VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConvertPDF 
   Caption         =   "Convert Office Files to PDF V1"
   ClientHeight    =   4908
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9180.001
   OleObjectBlob   =   "frmConvertPDF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConvertPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnBrowseInput_Click()
    Dim fldr As FileDialog
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select Input Folder"
        If .Show = -1 Then
            txtInputFolder.Text = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub btnBrowseOutput_Click()
    Dim fldr As FileDialog
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select Output Folder"
        If .Show = -1 Then
            txtOutputFolder.Text = .SelectedItems(1)
        End If
    End With
End Sub

Private Sub btnStart_Click()
    If txtInputFolder.Text = "" Then
        MsgBox "Please select an input folder.", vbExclamation
        Exit Sub
    End If
    If txtOutputFolder.Text = "" Then
        MsgBox "Please select an output folder.", vbExclamation
        Exit Sub
    End If
    If Not (chkExcel.Value Or chkWord.Value Or chkPPT.Value) Then
        MsgBox "Please select at least one file type to convert.", vbExclamation
        Exit Sub
    End If
    
    Me.Enabled = False
    lblStatus.Caption = "Starting conversion..."
    DoEvents
    
    ' Call main processing sub with parameters from form
    ConvertFilesToPDF txtInputFolder.Text, txtOutputFolder.Text, chkExcel.Value, chkWord.Value, chkPPT.Value, Me
    
    lblStatus.Caption = "Conversion complete!"
   ' Me.Enabled = True
    Unload Me
    
    Shell "explorer.exe """ & txtOutputFolder.Text & """", vbNormalFocus
    
    
End Sub

Private Sub lblStatus_Click()

End Sub

Private Sub txtInputFolder_Change()

End Sub

Private Sub txtOutputFolder_Change()

End Sub

Private Sub UserForm_Click()

End Sub


