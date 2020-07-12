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

Const LIST_AREA As String = "A1:I"
Const LIST_IND As String = "A1"
Const EXP_FILENAME As String = "IndSumRep.pdf"

' ===============================================================
' LogResult
' Logs the results of the comparison
' ---------------------------------------------------------------
Public Sub LogResult(AryExtract() As Variant)
    Dim NoRows As Integer
    Dim RngResults As Range
    Dim i As Integer
    
    NoRows = Application.WorksheetFunction.CountA(Me.Range("A:A"))
    Set RngResults = Range("A1").Offset(NoRows, 0)
    
    For i = LBound(AryExtract, 1) To UBound(AryExtract, 1)
        RngResults.Offset(i, enMobNum) = AryExtract(i, enMobNum)
        RngResults.Offset(i, enName) = AryExtract(i, enName)
        RngResults.Offset(i, enPlan) = AryExtract(i, enPlan)
        RngResults.Offset(i, enServCh) = CSng(AryExtract(i, enServCh))
        RngResults.Offset(i, enUsage) = CSng(AryExtract(i, enUsage))
        RngResults.Offset(i, enTotal) = CSng(AryExtract(i, enTotal))
    Next
    
    NoRows = Application.WorksheetFunction.CountA(Me.Range("A:A"))
    
    RngResults.Offset(NoRows + 1, enplan) = "Total"
    RngResults.Offset(NoRows + 1, enServCh) = Application.WorksheetFunction.Sum(Me.Columns(enServCh))
    RngResults.Offset(NoRows + 1, enUsage) = Application.WorksheetFunction.Sum(Me.Columns(enUsage))
    RngResults.Offset(NoRows + 1, enTotal) = Application.WorksheetFunction.Sum(Me.Columns(enTotal))
    
    RngResults.Columns.AutoFit
    
    Set RngResults = Nothing
End Sub

' ===============================================================
' ClearData
'Clears data from sheet
' ---------------------------------------------------------------
Public Sub ClearData()
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
    Range("A1") = "Phone Number"
    Range("B1") = "User Name"
    Range("C1") = "Plan"
    Range("D1") = "Service Charge"
    Range("E1") = "Usage"
    Range("F1") = "Mins"
    Range("G1") = "No SMS"
    Range("H1") = "Data(MB)"
    Range("I1") = "Total Cost"
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

