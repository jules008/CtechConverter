VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtMain
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Jun 20
'===============================================================
Option Explicit

' ===============================================================
' SelectFile
' Selects PDF for extraction
' ---------------------------------------------------------------
Public Sub SelectFile()
    Dim Fldr As FileDialog
    Dim FilePath As String
    
    Set Fldr = Application.FileDialog(msoFileDialogFilePicker)
    With Fldr
        .Title = "Select a File"
        .Filters.Clear
        .Filters.Add "PDF Files", "*.PDF, 1"
        .AllowMultiSelect = False
        .ButtonName = "Select"
        .InitialFileName = Application.DefaultFilePath
        
        If .Show <> -1 Then Exit Sub
        FilePath = .SelectedItems(1)
    End With
    
    If FilePath <> "" Then FrmProgress.ShowForm FilePath
End Sub
