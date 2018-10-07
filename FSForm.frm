VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FSForm 
   Caption         =   "File Scoring"
   ClientHeight    =   1845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6075
   OleObjectBlob   =   "FSForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FSForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub FTBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub OKButton_Click()
    If FolderBox = "" Then
        MsgBox ("Please select a folder")
    Else
        If FTBox = "" Then
            MsgBox ("Please select a file type")
        Else
            PDF = PDFCheck.Value
            FileType = FTBox.Text
            Application.ScreenUpdating = False
            Call FindFiles
            Application.ScreenUpdating = True
            Unload Me
        End If
    End If
End Sub

Private Sub SFButton_Click()
    Dim Folder As FileDialog
    
    Set Folder = Application.FileDialog(msoFileDialogFolderPicker)
    With Folder
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show = -1 Then FolderPath = .SelectedItems(1) & "\"
    End With
    
    FolderBox.Text = FolderPath
End Sub

Public Sub Userform_Initialize()
    Dim i As Integer
    Dim FTypes() As String
    
    FTypes = Split(".txt .xml .tex ", " ")
    
    For i = 0 To UBound(FTypes) - 1
        FTBox.AddItem FTypes(i)
    Next i
    
    FTBox = FTypes(0)
    FolderBox = Application.DefaultFilePath
End Sub

