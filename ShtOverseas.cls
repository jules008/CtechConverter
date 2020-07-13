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
Private Const NO_COLS As Integer = 13

' ===============================================================
' FilterList
' Filters Extract list of overseas calls
' ---------------------------------------------------------------
Public Sub FilterList(AryItemList() As Variant)
    Dim i, x As Integer
    Dim AryOutput() As Variant
    Dim RowNo As Integer
    
    ReDim AryOutput(0 To UBound(AryItemList) - 1, 0 To NO_COLS - 1)
    
    RowNo = 0
    For i = LBound(AryItemList, 1) To UBound(AryItemList, 1)
        
        If AryItemList(i, enfrom + 1) = "Overseas" Then
            For x = LBound(AryItemList, 2) To UBound(AryItemList, 2)
                AryOutput(RowNo, x - 1) = AryItemList(i, x)
            Next
            RowNo = RowNo + 1
        End If
    Next
    
'    ReDim Preserve AryOutput(0 To RowNo - 1, 0 To NO_COLS - 1)
    ClearData
    AddHeadings
    DisplayList AryOutput
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
    Range(RNG_INDEX).Offset(0, 0) = "Index"
    Range(RNG_INDEX).Offset(0, 1) = "Time"
    Range(RNG_INDEX).Offset(0, 2) = "Billed Number"
    Range(RNG_INDEX).Offset(0, 3) = "Call Category"
    Range(RNG_INDEX).Offset(0, 4) = "Date"
    Range(RNG_INDEX).Offset(0, 5) = "Description"
    Range(RNG_INDEX).Offset(0, 6) = "Duration"
    Range(RNG_INDEX).Offset(0, 7) = "MB"
    Range(RNG_INDEX).Offset(0, 8) = "Text"
    Range(RNG_INDEX).Offset(0, 9) = "Cost"
    Range(RNG_INDEX).Offset(0, 10) = "From"
    Range(RNG_INDEX).Offset(0, 11) = "To"
    Me.Protect
End Sub

' ===============================================================
' DisplayList
' Displays list of overseas calls
' ---------------------------------------------------------------
Private Sub DisplayList(AryList() As Variant)
    Dim r As Integer
    
    On Error Resume Next
    
    Me.Unprotect
    For r = LBound(AryList, 1) To UBound(AryList, 1)
        With Range(RNG_INDEX)
            .Offset(r + 1, 0) = AryList(r, enIndex)
            .Offset(r + 1, 1) = CSng(AryList(r, enTime))
            .Offset(r + 1, 2) = AryList(r, enPhoneNo)
            .Offset(r + 1, 3) = AryList(r, enCategory)
            .Offset(r + 1, 4) = AryList(r, enItemDate)
            .Offset(r + 1, 5) = AryList(r, enDescription)
            .Offset(r + 1, 6) = CSng(AryList(r, enDuration))
            .Offset(r + 1, 7) = CSng(AryList(r, enMB))
            .Offset(r + 1, 8) = CSng(AryList(r, enText))
            .Offset(r + 1, 9) = CSng(AryList(r, enCost))
            .Offset(r + 1, 10) = AryList(r, enfrom)
            .Offset(r + 1, 11) = AryList(r, enTo)
        End With
    Next
    Me.Protect
    
End Sub
