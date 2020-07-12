VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtPhoneList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtPhoneList
' List Phone numbers extracted from Bill
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 03 Jul 20
'===============================================================
Option Explicit

Const LIST_AREA As String = "A2:I"
Const LIST_IND As String = "A2"
Const EXP_FILENAME As String = "IndSumRep.pdf"

' ===============================================================
' LogResult
' Logs the results of the comparison
' ---------------------------------------------------------------
Public Sub LogResult(AryExtract() As Variant)
    Dim NoRows As Integer
    Dim RngResults As Range
    Dim i As Integer
    
    NoRows = Application.WorksheetFunction.CountA(Me.Range(LIST_IND).EntireColumn)
    Set RngResults = Range(LIST_IND).Offset(NoRows, 0)
    
    Me.Unprotect
    
    For i = LBound(AryExtract, 1) To UBound(AryExtract, 1)
        RngResults.Offset(i, enMobNum) = AryExtract(i, enMobNum)
        RngResults.Offset(i, enName) = AryExtract(i, enName)
        RngResults.Offset(i, enPlan) = AryExtract(i, enPlan)
        RngResults.Offset(i, enServCh) = CSng(AryExtract(i, enServCh))
        RngResults.Offset(i, enUsage) = CSng(AryExtract(i, enUsage))
        RngResults.Offset(i, enTotal) = CSng(AryExtract(i, enTotal))
    Next
    
    RngResults.Columns.AutoFit
    
    Me.Protect
    
    Set RngResults = Nothing
End Sub

' ===============================================================
' ClearData
'Clears data from sheet
' ---------------------------------------------------------------
Public Sub ClearData()
    On Error Resume Next
    
    Me.Unprotect
    Me.UsedRange.SpecialCells(xlCellTypeConstants).ClearContents
    AddHeadings
    Me.Protect
End Sub

' ===============================================================
' AddHeadings
' Adds column headings
' ---------------------------------------------------------------
Public Sub AddHeadings()
    Me.Unprotect
    Range(LIST_IND).Offset(0, 0) = "Phone Number"
    Range(LIST_IND).Offset(0, 1) = "User Name"
    Range(LIST_IND).Offset(0, 2) = "Plan"
    Range(LIST_IND).Offset(0, 3) = "Service Charge"
    Range(LIST_IND).Offset(0, 4) = "Usage"
    Range(LIST_IND).Offset(0, 5) = "Mins"
    Range(LIST_IND).Offset(0, 6) = "No SMS"
    Range(LIST_IND).Offset(0, 7) = "Data(MB)"
    Range(LIST_IND).Offset(0, 8) = "Total Cost"
    Me.Protect
End Sub

' ===============================================================
' SendToPDF
' Sends list to PDF file
' ---------------------------------------------------------------
Public Sub SendToPDF()
    Dim NoRows As Integer
    Dim RngList As Range
    Dim i As Integer
    
    For i = 1 To 500
        If Me.Range(LIST_IND).Offset(i, 0).Value = "" Then
            NoRows = i
            Exit For
        End If
    Next
    
    Set RngList = Me.Range(LIST_AREA & NoRows)
    
    RngList.ExportAsFixedFormat xlTypePDF, ThisWorkbook.Path & "\" & EXP_FILENAME, xlQualityStandard, True, True, OpenAfterPublish:=False
    
    Set RngList = Nothing
End Sub

' ===============================================================
' DeletePDF
' Deletes the PDF file when it is no longer needed
' ---------------------------------------------------------------
Public Sub DeletePDF()
    Dim FSO As FileSystemObject
    
    Set FSO = New FileSystemObject
    
    With FSO
        If .FileExists(ThisWorkbook.Path & "\" & EXP_FILENAME) Then
            .DeleteFile ThisWorkbook.Path & "\" & EXP_FILENAME
        End If
    End With
    
    Set FSO = Nothing
End Sub

' ===============================================================
' BtnHomeClick
' Returns to Main screen
' ---------------------------------------------------------------
Public Sub BtnHomeClick()
    ModBtnActions.GoToHome Me.Shapes("BtnHome")
End Sub

