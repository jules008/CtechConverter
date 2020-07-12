VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtOverseas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtOverseas
' Filtered list of just overseas calls
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 07 Jul 20
'===============================================================
Option Explicit

Private Const RNG_INDEX As String = "A1"
Private Const NO_COLS As Integer = 12

' ===============================================================
' FilterList
' Filters Extract list of overseas calls
' ---------------------------------------------------------------
Public Sub FilterList(AryItemList() As Variant)
    Dim i, x As Integer
    Dim AryOutput() As Variant
    Dim RowNo As Integer
    
    ReDim AryOutput(1 To UBound(AryItemList), 1 To NO_COLS)
    
    RowNo = 1
    For i = LBound(AryItemList, 1) To UBound(AryItemList, 1)
        
        If AryItemList(i, enfrom) = "Overseas" Then
            For x = LBound(AryItemList, 2) To UBound(AryItemList, 2)
                AryOutput(RowNo, x) = AryItemList(i, x)
            Next
            RowNo = RowNo + 1
        End If
    Next
    
'    ReDim Preserve AryOutput(1 To RowNo - 1, 1 To NO_COLS)
    ClearData
    AddHeadings
    DisplayList AryOutput
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
    Range("K1") = "From"
    Range("L1") = "To"
End Sub

' ===============================================================
' DisplayList
' Displays list of overseas calls
' ---------------------------------------------------------------
Private Sub DisplayList(AryList() As Variant)
    Dim r As Integer
    
    On Error Resume Next
    
    For r = LBound(AryList, 1) To UBound(AryList, 1)
        With Range(RNG_INDEX)
            .Offset(r, 0) = AryList(r, enIndex)
            .Offset(r, 1) = CSng(AryList(r, enTime))
            .Offset(r, 2) = AryList(r, enPhoneNo)
            .Offset(r, 3) = AryList(r, enCategory)
            .Offset(r, 4) = AryList(r, enItemDate)
            .Offset(r, 5) = AryList(r, enDescription)
            .Offset(r, 6) = CSng(AryList(r, enDuration))
            .Offset(r, 7) = CSng(AryList(r, enMB))
            .Offset(r, 8) = CSng(AryList(r, enText))
            .Offset(r, 9) = CSng(AryList(r, enCost))
            .Offset(r, 10) = AryList(r, enfrom)
            .Offset(r, 11) = AryList(r, enTo)
        End With
    Next
    
End Sub
