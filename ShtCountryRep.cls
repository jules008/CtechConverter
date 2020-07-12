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

Const COUNTRY_REP As String = "A2:B"
Const EXP_FILENAME As String = "CountryRep.pdf"
Private Const COUNTRY_LIST As String = "A3:A"
Private Const COUNTRY_IND As String = "A2"

' ===============================================================
' RemoveDuplicates
' Removes duplicates from range
' ---------------------------------------------------------------
Sub RemoveDuplicates()
    Dim DuplicateValues As Range
    Dim AryCountries() As Variant
    
    ClearData
    
    AryCountries = ShtLookup.GetCountryList
    
    Me.Range(COUNTRY_LIST & UBound(AryCountries)) = AryCountries
    
    Set DuplicateValues = Range(COUNTRY_LIST & UBound(AryCountries))
    
    DuplicateValues.RemoveDuplicates Columns:=Array(1), Header:=xlYes
    
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
    Me.Protect
End Sub

' ===============================================================
' AddHeadings
' Adds column headings
' ---------------------------------------------------------------
Public Sub AddHeadings()
    Me.Protect
    Range(COUNTRY_IND).Offset(0, 0) = "Country"
    Range(COUNTRY_IND).Offset(0, 1) = "Total Cost"
    Me.Unprotect
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
        If Me.Range(COUNTRY_IND).Offset(i, 0).Value = "" Then
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

