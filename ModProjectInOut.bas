Attribute VB_Name = "ModProjectInOut"
'===============================================================
' Module ModProjectInOut
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 20 Apr 18
'===============================================================

Option Explicit

' ===============================================================
' ExportModules
' Exports all VBA Modules to dev library
' ---------------------------------------------------------------
Public Sub ExportModules()
    Dim ExportYN As Boolean
    Dim SourceBook As Excel.Workbook
    Dim SourceBookName As String
    Dim ExportFileName As String
    Dim VBModule As VBIDE.VBComponent
   
    On Error Resume Next
   
    SourceBookName = ActiveWorkbook.Name
    Set SourceBook = Application.Workbooks(SourceBookName)
    
    If Not Dir(EXPORT_FILE_PATH & "*.*") = "" Then
        Kill EXPORT_FILE_PATH & "*.*"
    End If
    
    For Each VBModule In SourceBook.VBProject.VBComponents
        
        ExportYN = True
        ExportFileName = VBModule.Name

        Select Case VBModule.Type
            Case vbext_ct_ClassModule
                ExportFileName = ExportFileName & ".cls"
            Case vbext_ct_MSForm
                ExportFileName = ExportFileName & ".frm"
            Case vbext_ct_StdModule
                ExportFileName = ExportFileName & ".bas"
            Case vbext_ct_Document
                ExportFileName = ExportFileName & ".cls"
        End Select
        
        If ExportYN Then
            VBModule.Export EXPORT_FILE_PATH & ExportFileName
            
        End If
   
    Next VBModule
    
    
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs EXPORT_FILE_PATH & PROJECT_FILE_NAME, 51
    Application.DisplayAlerts = True
    
    MsgBox "Export is ready", vbInformation, APP_NAME
    Set SourceBook = Nothing
End Sub

' ===============================================================
' ImportModules
' Imports all VBA Modules from dev library
' ---------------------------------------------------------------
Public Sub ImportModules()
    Dim TargetBook As Excel.Workbook
    Dim FSO As Scripting.FileSystemObject
    Dim FileObj As Scripting.File
    Dim TargetBookName As String
    Dim ImportFileName As String
    Dim VBModules As VBIDE.VBComponents
    
    On Error Resume Next
    
    Set FSO = New Scripting.FileSystemObject
    If FSO.GetFolder(ThisWorkbook.Path).Files.Count = 0 Then
       MsgBox "There are no files to import", vbInformation
       Exit Sub
    End If

    Set VBModules = ThisWorkbook.VBProject.VBComponents
    
    For Each FileObj In FSO.GetFolder(ThisWorkbook.Path).Files
'        Debug.Print FileObj.Name
        
        If (FSO.GetExtensionName(FileObj.Name) = "cls") Or _
            (FSO.GetExtensionName(FileObj.Name) = "frm") Or _
            (FSO.GetExtensionName(FileObj.Name) = "bas") And _
            FileObj.Name <> "ModProjectInOut.bas" Then
            VBModules.Import FileObj.Path
        End If
        
    Next FileObj
'    Debug.Print "End of import"
    Set FSO = Nothing
    Set VBModules = Nothing
End Sub
 
' ===============================================================
' RemoveAllModules
' Removes all VBA Modules from project
' ---------------------------------------------------------------
Public Sub RemoveAllModules()
    Dim ExportYN As Boolean
    Dim DlgOpen As FileDialog
    Dim SourceBook As Excel.Workbook
    Dim SourceBookName As String
    Dim EXPORTFilePath As String
    Dim ImportFileName As String
    Dim VBModule As VBIDE.VBComponent
    
    On Error Resume Next
    
    SourceBookName = ActiveWorkbook.Name
    Set SourceBook = Application.Workbooks(SourceBookName)
        
    For Each VBModule In SourceBook.VBProject.VBComponents
        
        If VBModule.Type <> vbext_ct_Document Then SourceBook.VBProject.VBComponents.Remove VBModule
           
    Next VBModule
    
    Set DlgOpen = Nothing
    Set SourceBook = Nothing
End Sub

' ===============================================================
' SetReferenceLibs
' Sets all project reference libraries
' ---------------------------------------------------------------
Public Sub SetReferenceLibs()
    Dim Reference As Object
    
    On Error Resume Next
    
    For Each Reference In ThisWorkbook.VBProject.References
        With Reference
'            Debug.Print .Name
'            Debug.Print .Description
'            Debug.Print .Minor
'            Debug.Print .Major
'            Debug.Print .GUID
'            Debug.Print
        End With
    Next

    ' Visual Basic For Applications
    If Not ReferenceExists("{000204EF-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{000204EF-0000-0000-C000-000000000046}", Major:=4, Minor:=1
    End If
    
    ' Microsoft Excel 14.0 Object Library
    If Not ReferenceExists("{00020813-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{00020813-0000-0000-C000-000000000046}", Major:=1, Minor:=7
    End If
    
    ' Microsoft Forms 2.0 Object Library
    If Not ReferenceExists("{00020813-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{0D452EE1-E08F-101A-852E-02608C4D0BB4}", Major:=2, Minor:=0
    End If
    
    ' OLE Automation
    If Not ReferenceExists("{00020430-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{00020430-0000-0000-C000-000000000046}", Major:=2, Minor:=0
    End If
    
    ' Microsoft Office 14.0 Object Library
    If Not ReferenceExists("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", Major:=2, Minor:=5
    End If
    
    ' Microsoft Office 14.0 Access database engine Object Library
    If Not ReferenceExists("{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}", Major:=12, Minor:=0
    End If
    
    ' Microsoft Scripting Runtime
    If Not ReferenceExists("{420B2830-E718-11CF-893D-00A0C9054228}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{420B2830-E718-11CF-893D-00A0C9054228}", Major:=1, Minor:=0
    End If
    
    ' Microsoft Visual Basic for Applications Extensibility 5.3
    If Not ReferenceExists("{0002E157-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{0002E157-0000-0000-C000-000000000046}", Major:=5, Minor:=3
    End If
    
    ' Microsoft Outlook 14.0 Object Library
    If Not ReferenceExists("{00062FFF-0000-0000-C000-000000000046}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{00062FFF-0000-0000-C000-000000000046}", Major:=9, Minor:=4
    End If
    
    ' Adobe Acrobat 10.0 Type Library
    If Not ReferenceExists("{E64169B3-3592-47D2-816E-602C5C13F328}") Then
        ThisWorkbook.VBProject.References.AddFromGuid _
        GUID:="{E64169B3-3592-47D2-816E-602C5C13F328}", Major:=1, Minor:=1
    End If
End Sub

' ===============================================================
' ReferenceExists
' Checks to see if reference already exists
' ---------------------------------------------------------------
Public Function ReferenceExists(Ref As String) As Boolean
    Dim i As Integer
    
    With ThisWorkbook.VBProject.References
        For i = 1 To .Count
            If .Item(i).GUID = Ref Then
                ReferenceExists = True
                Exit Function
            End If
        Next
        ReferenceExists = False
    End With
End Function

' ===============================================================
' BuildProject
' main program for building project
' ---------------------------------------------------------------
Public Sub BuildProject()
    SetReferenceLibs
    ImportModules
    CopyShtCodeModule
End Sub

' ===============================================================
' CopyShtCodeModule
' Copies sheet modules and this workbook classes
' ---------------------------------------------------------------
Public Sub CopyShtCodeModule()
    Dim SourceMod As VBIDE.VBComponent
    Dim DestMod As VBIDE.VBComponent
    Dim VBModule As VBIDE.VBComponent
    Dim VBCodeMod As VBIDE.CodeModule
    Dim i As Integer

    If ModuleExists("ThisWorkbook1") Then
        Set SourceMod = ThisWorkbook.VBProject.VBComponents("Thisworkbook1")
        Set DestMod = ThisWorkbook.VBProject.VBComponents("Thisworkbook")
    
        If DestMod.CodeModule.CountOfLines > 0 Then
            DestMod.CodeModule.DeleteLines 1, DestMod.CodeModule.CountOfLines
        End If
        
        If SourceMod.CodeModule.CountOfLines > 0 Then
            DestMod.CodeModule.AddFromString SourceMod.CodeModule.Lines(1, SourceMod.CodeModule.CountOfLines)
        End If
    End If
    
    For Each VBModule In ThisWorkbook.VBProject.VBComponents

        With VBModule

'            Debug.Print VBModule.Name
            If Left(.Name, 3) = "Sht" And .Type <> vbext_ct_Document Then
                Set SourceMod = VBModule
                Debug.Print "Source: " & SourceMod.Name

                For Each DestMod In ThisWorkbook.VBProject.VBComponents
                    Debug.Print DestMod.Name
                    If Left(SourceMod.Name, Len(SourceMod.Name) - 1) = DestMod.Name Then
                        Debug.Print "Source: " & SourceMod.Name
                        Debug.Print " Dest: " & DestMod.Name

                        If SourceMod.CodeModule.CountOfLines > 0 Then
                            DestMod.CodeModule.DeleteLines 1, DestMod.CodeModule.CountOfLines
    
                            DestMod.CodeModule.AddFromString SourceMod.CodeModule.Lines(1, SourceMod.CodeModule.CountOfLines)
                        End If
                    End If
                Next
            End If
        End With
    Next

    For Each VBModule In ThisWorkbook.VBProject.VBComponents
        If Right(VBModule.Name, 1) = "1" And VBModule.Name <> "Sheet1" Then
            ThisWorkbook.VBProject.VBComponents.Remove VBModule
        End If
    Next VBModule



    Set SourceMod = Nothing
    Set DestMod = Nothing
    Set VBModule = Nothing
    Set VBCodeMod = Nothing
End Sub

' ===============================================================
' ModuleExists
' checks to see if module exists in project
' ---------------------------------------------------------------
Public Function ModuleExists(ModuleName As String) As Boolean
    Dim CodeModule As VBIDE.VBComponent
 
    For Each CodeModule In ThisWorkbook.VBProject.VBComponents
        If CodeModule.Name = ModuleName Then ModuleExists = True
    Next
End Function

' ===============================================================
' ImportModule
' Imports a sinlge VBA Modules from dev library
' ---------------------------------------------------------------
Public Sub ImportModule(ModuleName As String)
    ThisWorkbook.VBProject.VBComponents.Import IMPORT_FILE_PATH & ModuleName
End Sub


