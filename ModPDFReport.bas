Attribute VB_Name = "ModPDFReport"
'===============================================================
' Module ModPDFReport
' Generates PDF Report
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
Private AcroApp As Acrobat.AcroApp
Const FILE_NAME As String = "FullReport.pdf"

' ===============================================================
' CreateReport
' Creates the PDF report
' ---------------------------------------------------------------
Public Sub CreateReport()
    Dim AcroApp As Acrobat.CAcroApp
    Dim PDFFrontPage As Acrobat.CAcroPDDoc
    Dim Part1Document As Acrobat.CAcroPDDoc
    Dim Part2Document As Acrobat.CAcroPDDoc
    Dim Part3Document As Acrobat.CAcroPDDoc
    Dim Part4Document As Acrobat.CAcroPDDoc
    Dim numPages As Integer
    
    Set AcroApp = CreateObject("AcroExch.App")
    Set PDFFrontPage = CreateObject("AcroExch.PDDoc")
    Set Part1Document = CreateObject("AcroExch.PDDoc")
    Set Part2Document = CreateObject("AcroExch.PDDoc")
    Set Part3Document = CreateObject("AcroExch.PDDoc")
    Set Part4Document = CreateObject("AcroExch.PDDoc")
    
    ShtFrontPage.SendToPDF
    ShtCountryRep.SendToPDF
    ShtOutSpendRep.SendToPDF
    ShtRoamRep.SendToPDF
    ShtPhoneList.SendToPDF
    
    numPages = Part1Document.GetNumPages()
    
    PDFFrontPage.Open (ThisWorkbook.Path & "\" & "FrontPage.pdf")
    Part1Document.Open (ThisWorkbook.Path & "\" & "IndSumRep.pdf")
    Part2Document.Open (ThisWorkbook.Path & "\" & "OutSpendRep.pdf")
    Part3Document.Open (ThisWorkbook.Path & "\" & "RoamRep.pdf")
    Part4Document.Open (ThisWorkbook.Path & "\" & "CountryRep.pdf")

    If PDFFrontPage.InsertPages(numPages - 1, Part1Document, _
        0, Part2Document.GetNumPages(), True) = False Then
        Err.Raise 2500, Description:="Merge - Cannot insert pages"
    End If

    If PDFFrontPage.InsertPages(numPages - 1, Part2Document, _
        0, Part3Document.GetNumPages(), True) = False Then
        Err.Raise 2500, Description:="Merge - Cannot insert pages"
    End If
    
    If PDFFrontPage.InsertPages(numPages - 1, Part3Document, _
        0, Part3Document.GetNumPages(), True) = False Then
        Err.Raise 2500, Description:="Merge - Cannot insert pages"
    End If
 
    If PDFFrontPage.InsertPages(numPages - 1, Part4Document, _
        0, Part4Document.GetNumPages(), True) = False Then
        Err.Raise 2500, Description:="Merge - Cannot insert pages"
    End If
    
    If PDFFrontPage.Save(PDSaveFull, ThisWorkbook.Path & "\" & FILE_NAME) = False Then
        Err.Raise 2500, Description:="Merge - Cannot save the modified document"
    End If

    PDFFrontPage.OpenAVDoc ThisWorkbook.Path & "\" & FILE_NAME
    Part1Document.Close
    Part2Document.Close
    Part3Document.Close
    Part4Document.Close
    
    ShtFrontPage.DeletePDF
    ShtCountryRep.DeletePDF
    ShtOutSpendRep.DeletePDF
    ShtRoamRep.DeletePDF
    ShtPhoneList.DeletePDF
    
    Set Part1Document = Nothing
    Set Part2Document = Nothing
    Set Part3Document = Nothing
    Set Part4Document = Nothing
End Sub

' ===============================================================
' ClearReports
' Clears all the report sheets
' ---------------------------------------------------------------
Public Sub ClearReports()
    ShtCountryRep.ClearData
    ShtOutSpendRep.ClearData
    ShtRoamRep.ClearData
    ShtPhoneList.ClearData
    ShtItemList.ClearData
    ShtOverseas.ClearData
End Sub
