VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' Module ThisWorkbook
' Controls Workbook level activity
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 12 Jul 20
'===============================================================
Option Explicit

' ===============================================================
' Workbook_BeforeClose
' Clean up before exit
' ---------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    UnLockWorkbook
    ModLibrary.PerfSettingsOff
    Application.DisplayAlerts = False
    ShtClose.Visible = xlSheetVisible
    ShtClose.Activate
End Sub

' ===============================================================
' LogResult
' Logs the results of the comparison
' ---------------------------------------------------------------
Private Sub Workbook_Open()
    ShtClose.Visible = xlSheetVeryHidden
    ShtMain.Activate
    ModPDFReport.ClearReports
    LockWorkbook
End Sub

' ===============================================================
' LockWorkbook
' Locks Workbook
' ---------------------------------------------------------------
Private Sub LockWorkbook()
    Dim View As WorksheetView
    Dim WSheet As Worksheet
    
    With ShtMain
        .Unprotect
        .Columns("L:M").Hidden = True
        [PageNo] = 0
        [PhoneNo] = 0
        [Index] = 0
    End With
        
    ShtItemList.Visible = xlSheetVeryHidden
    ShtLookup.Visible = xlSheetHidden
    ShtOverseas.Visible = xlSheetVeryHidden
    ShtFrontPage.Visible = xlSheetVeryHidden
    ShtColours.Visible = xlSheetVeryHidden
    
    For Each View In Me.Parent.Windows(1).SheetViews
        View.DisplayGridlines = False
        View.DisplayHeadings = False
    Next
    
    For Each WSheet In ThisWorkbook.Worksheets
         With WSheet
            .DisplayPageBreaks = False
            .Protect
        End With
    Next
    
    With Application
        .DisplayFormulaBar = False
    End With
End Sub

' ===============================================================
' UnLockWorkbook
' UnLocks Workbook
' ---------------------------------------------------------------
Private Sub UnLockWorkbook()
    Dim View As WorksheetView
    Dim WSheet As Worksheet
    
    With ShtMain
        .Unprotect
        .Columns("L:M").Hidden = False
    End With
        
    ShtItemList.Visible = xlSheetVisible
    ShtLookup.Visible = xlSheetVisible
    ShtOverseas.Visible = xlSheetVisible
    ShtFrontPage.Visible = xlSheetVisible
    ShtColours.Visible = xlSheetVisible
    
    For Each View In Me.Parent.Windows(1).SheetViews
        View.DisplayGridlines = True
        View.DisplayHeadings = True
    Next
    
    For Each WSheet In ThisWorkbook.Worksheets
         With WSheet
            .DisplayPageBreaks = True
            .Unprotect
        End With
    Next
    
    With Application
        .DisplayFormulaBar = True
    End With
    
End Sub

