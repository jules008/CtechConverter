VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtItemList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtItemList
' Logs extracted data from bill
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
        
    
    NoRows = Application.WorksheetFunction.CountA(Me.Range("A:A"))
    Set RngResults = Me.Range(Me.Cells(1, 1), Me.Cells(UBound(AryExtract, 1) + 1, UBound(AryExtract, 2) + 1)).Offset(NoRows, 0)
    RngResults = AryExtract
    
    RngResults.EntireColumn.AutoFit
    
    Set RngResults = Nothing
End Sub

' ===============================================================
' ClearData
'Clears data from sheet
' ---------------------------------------------------------------
Public Sub ClearData()
    Me.UsedRange.ClearContents
    AddHeadings
End Sub

' ===============================================================
' AddHeadings
' Adds column headings
' ---------------------------------------------------------------
Public Sub AddHeadings()
    Range("A1") = "Index"
    Range("B1") = "Time"
    Range("C1") = "Billed Number"
    Range("D1") = "Call Category"
    Range("E1") = "Date"
    Range("F1") = "Description"
    Range("G1") = "Duration"
    Range("H1") = "MB"
    Range("I1") = "Text"
    Range("J1") = "Cost"
End Sub

