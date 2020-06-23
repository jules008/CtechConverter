VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmProgress 
   Caption         =   "Progress"
   ClientHeight    =   870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4605
   OleObjectBlob   =   "FrmProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FilePath As String

Option Explicit

Public Sub ShowForm(LocFilePath As String)
    FilePath = LocFilePath
    
    Me.Show
End Sub

Private Sub UserForm_Activate()
    ModMain.MainConvert FilePath

End Sub
