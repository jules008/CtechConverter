Attribute VB_Name = "ModGlobals"
'===============================================================
' Module ModGlobals
' Global declerations
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Jun 20
'===============================================================
Option Explicit

'===============================================================
Global Constants
'---------------------------------------------------------------
Public Const ERR_NO_INV_TYPE As Integer = 513
Public Const ERR_PROC_FAILED As Integer = 514
Public Const ERR_NO_INV_FILE As Integer = 515
Public Const ERR_INV_NOT_FOUND As Integer = 516
Public Const ERR_FAILED_CONV As Integer = 517
Public Const DEBUG_MODE As Boolean = False
Public Const EXPORT_FILE_PATH As String = "G:\Development Areas\CTech Converter\Library\"
Public Const PROJECT_FILE_NAME As String = "Bill Extract"
Public Const APP_NAME As String = "Ctech Bill Extract"
Public Const IMPORT_FILE_PATH = "G:\CtechConverter\"
'===============================================================
Global variables
'---------------------------------------------------------------
Public NO_ERRORS As Integer
Public NO_DIFFS As Integer
Public COLLECTIVE_SHEET_PATH As String
Public ERROR_LOG As String
Public AcroApp As Acrobat.AcroApp
Public MULTI_TAB As Boolean
'===============================================================
Global Classes
'---------------------------------------------------------------


'===============================================================
' Enumerators
'---------------------------------------------------------------
Enum enNumCols
    enMobNum = 0
    enName
    enServCh
    enPlan
    enUsage
    enTotal
End Enum

Enum enItemList
    enIndex
    enTime
    enPhoneNo
    enCategory
    enItemDate
    enDescription
    enDuration
    enCost
End Enum
