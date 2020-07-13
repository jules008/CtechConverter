VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ShtLookUp
' Look up lists for reports
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

Const COUNTRY_LIST As String = "G2:G"
' ===============================================================
' GetCountryList
' Sends the list of countrys
' ---------------------------------------------------------------
Public Function GetCountryList() As Variant()
    Dim AryOutput() As Variant
    Dim NoRows As Integer
    
    NoRows = Application.WorksheetFunction.CountA(Me.Range(COUNTRY_LIST & "1").EntireColumn)
    
    ReDim AryOutput(1 To NoRows)
    
    AryOutput = Me.Range(COUNTRY_LIST & NoRows)
    GetCountryList = AryOutput
    
End Function
