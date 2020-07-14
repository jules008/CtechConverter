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

Private Const NO_COLS As Integer = 12
Private Const LIST_IND As String = "A1"

' ===============================================================
' LogResult
' Logs the results of the comparison
' ---------------------------------------------------------------
Public Sub LogResult(AryExtract() As Variant)
    Dim NoRows As Integer
    Dim RngResults As Range
        
    Me.Unprotect
    NoRows = Application.WorksheetFunction.CountA(Me.Range(LIST_IND).EntireColumn)
    
    Set RngResults = Me.Range(Me.Cells(1, 1), Me.Cells(UBound(AryExtract, 1) + 1, UBound(AryExtract, 2) + 1)).Offset(NoRows, 0)
    RngResults = AryExtract
    
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
    Range(LIST_IND).Offset(0, 0) = "Index"
    Range(LIST_IND).Offset(0, 1) = "Time"
    Range(LIST_IND).Offset(0, 2) = "Billed Number"
    Range(LIST_IND).Offset(0, 3) = "Call Category"
    Range(LIST_IND).Offset(0, 4) = "Date"
    Range(LIST_IND).Offset(0, 5) = "Description"
    Range(LIST_IND).Offset(0, 6) = "Duration"
    Range(LIST_IND).Offset(0, 7) = "MB"
    Range(LIST_IND).Offset(0, 8) = "Text"
    Range(LIST_IND).Offset(0, 9) = "Cost"
    Range(LIST_IND).Offset(0, 10) = "From"
    Range(LIST_IND).Offset(0, 11) = "To"
    Me.Protect
End Sub

' ===============================================================
' SendItemList
' Sends item list to overseas tab for filter
' ---------------------------------------------------------------
Public Sub SendItemList()
    Dim NoRows As Integer
    Dim AryList() As Variant
    
    NoRows = Application.WorksheetFunction.CountA(Me.Range(LIST_IND).EntireColumn)
    
    ReDim AryList(0 To NoRows - 2, 0 To NO_COLS - 1)
    
    AryList = Range(Cells(1, 1), Cells(NoRows, NO_COLS))
    
    ShtOverseas.FilterList AryList
End Sub
