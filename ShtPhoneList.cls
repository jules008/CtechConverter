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
' Date - 02 Jul 20
'===============================================================
Option Explicit

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
    
    RngResults.Columns.AutoFit
    
    Set RngResults = Nothing
End Sub

' ===============================================================
' ClearData
'Clears data from sheet
' ---------------------------------------------------------------
Public Sub ClearData()
    Me.UsedRange.SpecialCells(xlCellTypeConstants).ClearContents

    AddHeadings
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




