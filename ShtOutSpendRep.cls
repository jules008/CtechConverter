VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtOutSpendRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtOutSpendRep
' reports all international and roaming usage
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 10 Jul 20
'===============================================================
Option Explicit

Const LIST_AREA As String = "A1:L"
Const LIST_IND As String = "A2"
Const EXP_FILENAME As String = "OutSpendRep.pdf"
Const TITLE As String = "Outside Spend Report"
Const RNG_TITLE As String = "A1"
Const COL_WIDTHS As String = "11,14,7,5,5,7,7,5,7,7,7,7"

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
    Range(RNG_TITLE) = TITLE
    Range(LIST_IND).Offset(0, 0) = "Phone Number"
    Range(LIST_IND).Offset(0, 1) = "User Name"
    Range(LIST_IND).Offset(0, 2) = "Int Mins"
    Range(LIST_IND).Offset(0, 3) = "Int SMS"
    Range(LIST_IND).Offset(0, 4) = "Int MB"
    Range(LIST_IND).Offset(0, 5) = "Int Cost"
    Range(LIST_IND).Offset(0, 6) = "Roaming Mins"
    Range(LIST_IND).Offset(0, 7) = "Roaming SMS"
    Range(LIST_IND).Offset(0, 8) = "Roaming MB"
    Range(LIST_IND).Offset(0, 9) = "Roaming Cost"
    Range(LIST_IND).Offset(0, 10) = "Other Mins"
    Range(LIST_IND).Offset(0, 11) = "Purchases"
    SetCols
    Me.Protect
End Sub

' ===============================================================
' SetCols
' Sets the column widths
' ---------------------------------------------------------------
Public Sub SetCols()
    Dim Widths() As String
    Dim i As Integer
    
    Widths = Split(COL_WIDTHS, ",")
    
    Me.Unprotect
    
    For i = LBound(Widths) To UBound(Widths)
        Range(LIST_IND).Offset(0, i).ColumnWidth = CInt(Widths(i))
    Next
        
        
    Me.Protect
End Sub

