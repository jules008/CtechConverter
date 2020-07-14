VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtCountryRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtCountryRep
' Country report
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 09 Jul 20
'===============================================================
Option Explicit

Const COUNTRY_REP As String = "A1:B"
Const EXP_FILENAME As String = "CountryRep.pdf"
Private Const COUNTRY_LIST As String = "A3:A"
Private Const LIST_IND As String = "A2"
Const TITLE As String = "Country Report"
Const RNG_TITLE As String = "A1"
Const COL_WIDTHS As String = "50,10"

' ===============================================================
' GetCountryList
' gets country list and removes duplicates
' ---------------------------------------------------------------
Sub GetCountryList()
    Dim DuplicateValues As Range
    Dim AryCountries() As Variant
    
    ClearData
    Me.Unprotect
    
    AryCountries = ShtLookup.GetCountryList
    Me.Range(COUNTRY_LIST & UBound(AryCountries) + 2) = AryCountries
    
    Set DuplicateValues = Range(COUNTRY_LIST & UBound(AryCountries) + 2)
    
    DuplicateValues.RemoveDuplicates Columns:=Array(1), Header:=xlYes
    
    Me.Protect
    
    Set DuplicateValues = Nothing
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
    SetCols
    Me.Protect
End Sub

' ===============================================================
' AddHeadings
' Adds column headings
' ---------------------------------------------------------------
Public Sub AddHeadings()
    Me.Unprotect
    Range(RNG_TITLE) = TITLE
    Range(LIST_IND).Offset(0, 0) = "Country"
    Range(LIST_IND).Offset(0, 1) = "Total Cost"
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
    
    Set RngList = Me.Range(COUNTRY_REP & NoRows)
    
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

