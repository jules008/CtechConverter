Attribute VB_Name = "ModMain"
'===============================================================
' Module ModMain
' Main controlling functions
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
Private AcroApp As Acrobat.AcroApp

' ===============================================================
' MainConvert
' Main calling subroutine to extract data from bill
' ---------------------------------------------------------------
Public Sub MainConvert(PDFPath As String)
    Dim AcroAVDoc As Acrobat.AcroAVDoc
    Dim AcroPDDoc As Acrobat.AcroPDDoc
    Dim PDFPage As AcroPDPage
    Dim PDFSelection As AcroPDTextSelect
    Dim ProgPC As Single
    Dim PDFHighlight As AcroHiliteList
    Dim PageNum, TCount As Integer
    Dim ErrorFlag As Boolean
    Dim FoundStr As Boolean
    Dim StrText As String
    Dim AryPhoneList() As Variant
    Dim AryItemList() As Variant
    
    On Error Resume Next
    
    Set AcroApp = CreateObject("AcroExch.App")
    
    Set AcroAVDoc = CreateObject("AcroExch.AVDoc")
    
    ShtItemList.ClearData
    ShtPhoneList.ClearData
    
    If AcroAVDoc.Open(PDFPath, "Accessing PDF's") Then
    
        Set AcroPDDoc = AcroAVDoc.GetPDDoc()
        
        For PageNum = 0 To AcroPDDoc.GetNumPages() - 1
            Set PDFPage = AcroPDDoc.AcquirePage(PageNum)
            Set PDFHighlight = New AcroHiliteList
        
            DoEvents
            ProgPC = PageNum / (AcroPDDoc.GetNumPages() - 1) * 100
            progress ProgPC
            
            Debug.Print PageNum
            
            PDFHighlight.Add 0, 2500 ' Adjust this up if it's not getting all the text on the page
            Set PDFSelection = PDFPage.CreatePageHilite(PDFHighlight)
            
            If Not PDFSelection Is Nothing Then
            
                For TCount = 0 To PDFSelection.GetNumText - 1
                    
                    StrText = StrText & PDFSelection.GetText(TCount)
                    
                    If InStr(1, StrText, "SNM MOBILE BKDWN", vbTextCompare) Then
                        Set PDFPage = AcroPDDoc.AcquirePage(PageNum)

                        AryPhoneList = GetNumbers(PageNum, AcroPDDoc)
                        ShtPhoneList.LogResult AryPhoneList
                        StrText = ""
                    End If
                    
                    If InStr(1, StrText, "SNM ITEMISATION", vbTextCompare) Then
                        Set PDFPage = AcroPDDoc.AcquirePage(PageNum)

                        AryItemList = ItemisationExt(PageNum, AcroPDDoc)
                        ShtItemList.LogResult AryItemList
                        StrText = ""
                    End If

               Next TCount
            End If
            
        Next PageNum
             
    End If
    
    Application.Wait (Now + TimeValue("0:00:2"))
    Unload FrmProgress
    
    MsgBox "Extraction Complete"
    
    ErrorFlag = AcroApp.CloseAllDocs()
    ErrorFlag = AcroApp.Exit()
    
    Set AcroApp = Nothing
    Set AcroAVDoc = Nothing
    Set AcroPDDoc = Nothing
    Set PDFPage = Nothing
    Set PDFSelection = Nothing
    Set PDFHighlight = Nothing
    Set PDFPage = Nothing
End Sub

' ===============================================================
' GetNumbers
' Extracts individual phone numbers and their users
' ---------------------------------------------------------------
Public Function GetNumbers(ByVal PageNum As Integer, AcroPDDoc As Acrobat.AcroPDDoc) As Variant()
    Dim PDTextSelect As AcroPDTextSelect
    Dim PDFPage As AcroPDPage
    Dim AcroRect As New Acrobat.AcroRect
    Dim AryOutput() As Variant
    Dim AryPhoneNos() As Variant
    Dim AryString() As String
    Dim TmpAry() As String
    Dim i, x, y, TxtSt, TxtEnd, TxtLen, RowNo As Integer
    Dim JSO As Object
    Dim StrText As String
   
    On Error Resume Next
    
    Set PDFPage = AcroPDDoc.AcquirePage(PageNum)
    Set JSO = AcroPDDoc.GetJSObject
        
    AcroRect.bottom = 95: AcroRect.Top = 683
    AcroRect.Left = 160: AcroRect.Right = 500
    
    StrText = ""
    For x = 0 To 5
        If x > 0 Then AcroRect.Top = 740
        
        Set PDTextSelect = AcroPDDoc.CreateTextSelect(PageNum + x, AcroRect)
        
        If Not PDTextSelect Is Nothing Then
            For i = 0 To PDTextSelect.GetNumText() - 1
                StrText = StrText & PDTextSelect.GetText(i)
            Next
            If StrText = "Total" Then Exit For
            StrText = StrText & vbCrLf
        End If
    Next
    Debug.Print StrText
    
    AryString = Split(StrText, vbCrLf)
    
    ReDim AryOutput(0 To UBound(AryString) / 2, 0 To 5)
    
    RowNo = 0
    For i = LBound(AryString) To UBound(AryString)
        
        If AryString(i) <> "" Then
            TmpAry = Split(AryString(i), " ")
            
            If Left(TmpAry(0), 5) = "Total" Then Exit For
            
            If Left(TmpAry(0), 2) = "07" Then
                
                AryOutput(RowNo, enMobNum) = TmpAry(0) & " " & TmpAry(1) & " " & TmpAry(2)
                AryOutput(RowNo, enServCh) = TmpAry(3)
                AryOutput(RowNo, enUsage) = TmpAry(4)
                AryOutput(RowNo, enTotal) = TmpAry(5)
            End If
            
            If Left(TmpAry(0), 2) <> "07" Then
                
                TxtSt = InStr(1, AryString(i), " on ", vbTextCompare)
                TxtEnd = InStr(1, AryString(i), " with ", vbTextCompare)
                TxtLen = Len(AryString(i))
                AryOutput(RowNo, enName) = Replace(Left(AryString(i), TxtSt - 1), "REF: ", "")
                AryOutput(RowNo, enPlan) = Mid(AryString(i), TxtSt + 4, TxtEnd - TxtSt - 4)
                
                RowNo = RowNo + 1
            End If
        End If
    Next
    
    ReDim AryPhoneNos(0 To RowNo - 1, 0 To 5)
    
    For x = LBound(AryPhoneNos, 1) To UBound(AryPhoneNos, 1)
         For y = LBound(AryPhoneNos, 2) To UBound(AryPhoneNos, 2)
            AryPhoneNos(x, y) = AryOutput(x, y)
         Next y
     Next x
     
    GetNumbers = AryPhoneNos
    
    Debug.Print StrText
    
    Set PDTextSelect = Nothing
    Set PDFPage = Nothing
    Set AcroRect = Nothing
    Set PDTextSelect = Nothing
    Set JSO = Nothing
End Function

' ===============================================================
' ItemisationExt
' Extracts itemistion for each number
' ---------------------------------------------------------------
Public Function ItemisationExt(ByVal PageNum As Integer, AcroPDDoc As Acrobat.AcroPDDoc) As Variant()
    Dim PDTextSelect As AcroPDTextSelect
    Dim PDFPage As AcroPDPage
    Dim OddCol1 As New Acrobat.AcroRect
    Dim OddCol2 As New Acrobat.AcroRect
    Dim EvenCol1 As New Acrobat.AcroRect
    Dim EvenCol2 As New Acrobat.AcroRect
    Dim AcroRectTmp As New Acrobat.AcroRect
    Dim AryItemList() As Variant
    Dim AryString() As String
    Dim Category As String
    Static ItemDate As String
    Dim PhoneNum As String
    Dim TmpAry() As String
    Dim RetCat As String
    Dim RowNo, i, x, y, TxtSt, TxtEnd, TxtLen
    Dim JSO As Object
    Dim StrText As String
    Dim NormCostCl As Boolean
    Dim Col As Integer
    Dim StrCont As String
    
    On Error Resume Next
    
    Set PDFPage = AcroPDDoc.AcquirePage(PageNum)
    Set JSO = AcroPDDoc.GetJSObject
    
    With OddCol1
        .bottom = 70: .Top = 700: .Left = 150: .Right = 330
    End With
    
    With OddCol2
        .bottom = 70: .Top = 700: .Left = 350: .Right = 540
    End With
    
    With EvenCol1
        .bottom = 70: .Top = 785: .Left = 150: .Right = 330
    End With
    
    With EvenCol2
        .bottom = 70: .Top = 785: .Left = 350: .Right = 540
    End With
    
    With AcroRectTmp
        .bottom = 735: .Top = 745: .Left = 565: .Right = 595
    End With
    
    PhoneNum = JSO.getPageNthWord(PageNum, 0) & " " & JSO.getPageNthWord(PageNum, 1) & " " & JSO.getPageNthWord(PageNum, 2)
    
    StrText = ""
    For x = 0 To 6
        Set PDTextSelect = AcroPDDoc.CreateTextSelect(PageNum + x, AcroRectTmp)
        
        StrCont = ""
        If Not PDTextSelect Is Nothing Then
            For y = 0 To PDTextSelect.GetNumText() - 1
                StrCont = StrCont & PDTextSelect.GetText(y)
            Next
            If x <> 0 And StrCont = "SNM ITEMISATION" Then
                Exit For
            End If
        End If
                
        For Col = 1 To 2
            If Col = 1 Then
                If (PageNum + x) Mod 2 = 0 Then
                    Set PDTextSelect = AcroPDDoc.CreateTextSelect(PageNum + x, OddCol1)
                Else
                    Set PDTextSelect = AcroPDDoc.CreateTextSelect(PageNum + x, EvenCol1)
                End If
            Else
                If (PageNum + x) Mod 2 = 0 Then
                    Set PDTextSelect = AcroPDDoc.CreateTextSelect(PageNum + x, OddCol2)
                Else
                    Set PDTextSelect = AcroPDDoc.CreateTextSelect(PageNum + x, EvenCol2)
                End If
            End If
        
            If Not PDTextSelect Is Nothing Then
                For y = 0 To PDTextSelect.GetNumText() - 1
                    StrText = StrText & PDTextSelect.GetText(y)
                Next
                StrText = StrText & vbCrLf
            End If
        Next
    Next
    Debug.Print StrText
    
    AryString = Split(StrText, vbCrLf)
    
    ReDim Preserve AryString(UBound(AryString) - 1)
    ReDim AryItemList(0 To UBound(AryString), 0 To 9)
    
    Category = "UK Calls"
    
    RowNo = 0
    For i = LBound(AryString) To UBound(AryString)
        
        If InStr(1, AryString(i), "preferred network", vbTextCompare) Then AryString(i) = "skip"
        If InStr(1, AryString(i), "Continued", vbTextCompare) Then AryString(i) = "skip"
        If InStr(1, AryString(i), "Total", vbTextCompare) Then AryString(i) = "skip"
        
        TmpAry = Split(AryString(i))
        
        If UBound(TmpAry) > 6 Then
            If TmpAry(0) = "time" Then
                If TmpAry(6) = "normal" Then
                    NormCostCl = True
                Else
                    NormCostCl = False
                End If
            End If
        End If
                
        If Not HasNumber(AryString(i)) Then
            RetCat = GetCategory(AryString(i))
        
            If RetCat <> "" Then Category = RetCat
            AryString(i) = "Skip"
            TmpAry = Split(AryString(i))
        End If
        
        If NormCostCl Then
            Select Case UBound(TmpAry)
                Case Is = 2
                    If IsNumeric(TmpAry(1)) Then
                        ItemDate = TmpAry(0) & " " & TmpAry(1) & " " & TmpAry(2)
                    End If
                
                 Case Is = 3
                    If TmpAry(0) <> "time" And TmpAry(0) <> "Total" Then
                        AryItemList(RowNo, enIndex) = i
                        AryItemList(RowNo, enTime) = TmpAry(0)
                        AryItemList(RowNo, enCategory) = Category
                        AryItemList(RowNo, enDescription) = TmpAry(1)
                        AryItemList(RowNo, enDuration) = TmpAry(2)
                        AryItemList(RowNo, enCost) = TmpAry(3)
                        AryItemList(RowNo, enPhoneNo) = PhoneNum
                        AryItemList(RowNo, enItemDate) = ItemDate
                        RowNo = RowNo + 1
                    End If

                Case Is = 4
                    AryItemList(RowNo, enIndex) = i
                    AryItemList(RowNo, enTime) = TmpAry(0)
                    AryItemList(RowNo, enCategory) = Category
                    AryItemList(RowNo, enPhoneNo) = PhoneNum
                    AryItemList(RowNo, enItemDate) = ItemDate
                    AryItemList(RowNo, enDescription) = TmpAry(1)
                    AryItemList(RowNo, enDuration) = TmpAry(2)
                    AryItemList(RowNo, enCost) = TmpAry(4)
                    RowNo = RowNo + 1
                   
               Case Is = 5
                    If Category = "UK Calls" Or Category = "UK Messaging, mobile internet" Then
                        'UK Calls and UK Internet
                        AryItemList(RowNo, enIndex) = i
                        AryItemList(RowNo, enTime) = TmpAry(0)
                        AryItemList(RowNo, enCategory) = Category
                        AryItemList(RowNo, enPhoneNo) = PhoneNum
                        AryItemList(RowNo, enItemDate) = ItemDate
                        AryItemList(RowNo, enDescription) = ""
                        AryItemList(RowNo, enDuration) = TmpAry(2) & " " & TmpAry(3)
                        AryItemList(RowNo, enCost) = TmpAry(5)
                        RowNo = RowNo + 1
                    Else
                        'overseas internet
                        AryItemList(RowNo, enIndex) = i
                        AryItemList(RowNo, enTime) = TmpAry(0)
                        AryItemList(RowNo, enCategory) = Category
                        AryItemList(RowNo, enPhoneNo) = PhoneNum
                        AryItemList(RowNo, enItemDate) = ItemDate
                        AryItemList(RowNo, enDescription) = TmpAry(2)
                        AryItemList(RowNo, enDuration) = TmpAry(3)
                        AryItemList(RowNo, enCost) = TmpAry(5)
                        RowNo = RowNo + 1
                    End If
                    
                Case Is = 6
                    If TmpAry(0) <> "time" And TmpAry(0) <> "Total" Then
                        If Category = "UK Calls" Then
                            'uk calls
                            AryItemList(RowNo, enIndex) = i
                            AryItemList(RowNo, enTime) = TmpAry(0)
                            AryItemList(RowNo, enCategory) = Category
                            AryItemList(RowNo, enPhoneNo) = PhoneNum
                            AryItemList(RowNo, enItemDate) = ItemDate
                            AryItemList(RowNo, enDescription) = TmpAry(2)
                            AryItemList(RowNo, enDuration) = TmpAry(3) & " " & TmpAry(4)
                            AryItemList(RowNo, enCost) = TmpAry(6)
                            RowNo = RowNo + 1
                        Else
                            'overseas internet
                            AryItemList(RowNo, enIndex) = i
                            AryItemList(RowNo, enTime) = TmpAry(0)
                            AryItemList(RowNo, enCategory) = Category
                            AryItemList(RowNo, enPhoneNo) = PhoneNum
                            AryItemList(RowNo, enItemDate) = ItemDate
                            AryItemList(RowNo, enDescription) = TmpAry(2) & " " & TmpAry(3)
                            AryItemList(RowNo, enDuration) = TmpAry(4)
                            AryItemList(RowNo, enCost) = TmpAry(6)
                            RowNo = RowNo + 1
                        End If
                    End If
                Case Is = 7
                    If TmpAry(0) <> "time" And TmpAry(0) <> "Total" Then
                        If Category = "Overseas Mobile Internet" Then
                            AryItemList(RowNo, enIndex) = i
                            AryItemList(RowNo, enTime) = TmpAry(0)
                            AryItemList(RowNo, enCategory) = Category
                            AryItemList(RowNo, enDescription) = TmpAry(2) & " " & TmpAry(3) & " " & TmpAry(4)
                            AryItemList(RowNo, enDuration) = TmpAry(5)
                            AryItemList(RowNo, enCost) = TmpAry(7)
                            AryItemList(RowNo, enPhoneNo) = PhoneNum
                            AryItemList(RowNo, enItemDate) = ItemDate
                            RowNo = RowNo + 1
                        Else
                            AryItemList(RowNo, enIndex) = i
                            AryItemList(RowNo, enTime) = TmpAry(0)
                            AryItemList(RowNo, enCategory) = Category
                            AryItemList(RowNo, enDescription) = TmpAry(2) & " " & TmpAry(3)
                            AryItemList(RowNo, enDuration) = TmpAry(4) & " " & TmpAry(5)
                            AryItemList(RowNo, enCost) = TmpAry(7)
                            AryItemList(RowNo, enPhoneNo) = PhoneNum
                            AryItemList(RowNo, enItemDate) = ItemDate
                            RowNo = RowNo + 1
                        End If
                    End If
                Case Is = 8
                    If TmpAry(0) <> "time" And TmpAry(0) <> "Total" Then
                        If Category = "UK Calls" Then
                            AryItemList(RowNo, enIndex) = i
                            AryItemList(RowNo, enTime) = TmpAry(0)
                            AryItemList(RowNo, enCategory) = Category
                            AryItemList(RowNo, enPhoneNo) = PhoneNum
                            AryItemList(RowNo, enDescription) = TmpAry(2) & " " & TmpAry(3) & " " & TmpAry(4)
                            AryItemList(RowNo, enItemDate) = ItemDate
                            AryItemList(RowNo, enDuration) = TmpAry(5) & " " & TmpAry(6)
                            AryItemList(RowNo, enCost) = TmpAry(8)
                            RowNo = RowNo + 1
                        Else
                            AryItemList(RowNo, enIndex) = i
                            AryItemList(RowNo, enTime) = TmpAry(0)
                            AryItemList(RowNo, enCategory) = Category
                            AryItemList(RowNo, enPhoneNo) = PhoneNum
                            AryItemList(RowNo, enDescription) = TmpAry(3) & " " & TmpAry(4)
                            AryItemList(RowNo, enItemDate) = ItemDate
                            AryItemList(RowNo, enDuration) = TmpAry(5) & " " & TmpAry(6)
                            AryItemList(RowNo, enCost) = TmpAry(8)
                            RowNo = RowNo + 1
                        End If
                    End If
                Case Is = 9
                    If TmpAry(0) <> "time" And TmpAry(0) <> "Total" Then
                        AryItemList(RowNo, enIndex) = i
                        AryItemList(RowNo, enTime) = TmpAry(0)
                        AryItemList(RowNo, enCategory) = Category
                        AryItemList(RowNo, enPhoneNo) = PhoneNum
                        AryItemList(RowNo, enDescription) = TmpAry(2) & " " & TmpAry(3)
                        AryItemList(RowNo, enItemDate) = ItemDate
                        AryItemList(RowNo, enDuration) = TmpAry(4) & " " & TmpAry(5)
                        AryItemList(RowNo, enCost) = TmpAry(6)
                        RowNo = RowNo + 1
                    End If
                                     
            End Select
        Else
            Select Case UBound(TmpAry)
                Case Is = 2
                    If IsNumeric(TmpAry(1)) Then
                        ItemDate = TmpAry(0) & " " & TmpAry(1) & " " & TmpAry(2)
                    End If
                
                Case Is = 3
                    If TmpAry(0) <> "time" And TmpAry(0) <> "Total" Then
                        AryItemList(RowNo, enIndex) = i
                        AryItemList(RowNo, enTime) = TmpAry(0)
                        AryItemList(RowNo, enCategory) = Category
                        AryItemList(RowNo, enPhoneNo) = PhoneNum
                        AryItemList(RowNo, enItemDate) = ItemDate
                        AryItemList(RowNo, enDescription) = TmpAry(1)
                        AryItemList(RowNo, enDuration) = TmpAry(2)
                        AryItemList(RowNo, enCost) = TmpAry(3)
                        RowNo = RowNo + 1
                    End If
                Case Is = 4
                    If Category = "Overseas Mobile Internet" Or Category = "UK Messaging, mobile internet" Then
                        AryItemList(RowNo, enIndex) = i
                        AryItemList(RowNo, enTime) = TmpAry(0)
                        AryItemList(RowNo, enCategory) = Category
                        AryItemList(RowNo, enPhoneNo) = PhoneNum
                        AryItemList(RowNo, enItemDate) = ItemDate
                        AryItemList(RowNo, enDescription) = TmpAry(1) & " " & TmpAry(2)
                        AryItemList(RowNo, enDuration) = TmpAry(3)
                        AryItemList(RowNo, enCost) = TmpAry(4)
                        RowNo = RowNo + 1
                    Else
                        AryItemList(RowNo, enIndex) = i
                        AryItemList(RowNo, enTime) = TmpAry(0)
                        AryItemList(RowNo, enCategory) = Category
                        AryItemList(RowNo, enPhoneNo) = PhoneNum
                        AryItemList(RowNo, enItemDate) = ItemDate
                        AryItemList(RowNo, enDescription) = ""
                        AryItemList(RowNo, enDuration) = TmpAry(2) & " " & TmpAry(3)
                        AryItemList(RowNo, enCost) = TmpAry(4)
                        RowNo = RowNo + 1
                    End If
                Case Is = 5
                    If Category = "Overseas Mobile Internet" Then
                        AryItemList(RowNo, enIndex) = i
                        AryItemList(RowNo, enTime) = TmpAry(0)
                        AryItemList(RowNo, enCategory) = Category
                        AryItemList(RowNo, enPhoneNo) = PhoneNum
                        AryItemList(RowNo, enItemDate) = ItemDate
                        AryItemList(RowNo, enDescription) = TmpAry(2) & " " & TmpAry(3)
                        AryItemList(RowNo, enDuration) = TmpAry(4)
                        AryItemList(RowNo, enCost) = TmpAry(5)
                        RowNo = RowNo + 1
                    Else
                        AryItemList(RowNo, enIndex) = i
                        AryItemList(RowNo, enTime) = TmpAry(0)
                        AryItemList(RowNo, enCategory) = Category
                        AryItemList(RowNo, enPhoneNo) = PhoneNum
                        AryItemList(RowNo, enItemDate) = ItemDate
                        AryItemList(RowNo, enDescription) = TmpAry(2)
                        AryItemList(RowNo, enDuration) = TmpAry(3) & " " & TmpAry(4)
                        AryItemList(RowNo, enCost) = TmpAry(5)
                        RowNo = RowNo + 1
                    End If
                Case Is = 6
                    If TmpAry(0) <> "time" And TmpAry(0) <> "Total" Then
                        AryItemList(RowNo, enIndex) = i
                        AryItemList(RowNo, enTime) = TmpAry(0)
                        AryItemList(RowNo, enCategory) = Category
                        AryItemList(RowNo, enPhoneNo) = PhoneNum
                        AryItemList(RowNo, enItemDate) = ItemDate
                        AryItemList(RowNo, enDescription) = TmpAry(2) & " " & TmpAry(3)
                        AryItemList(RowNo, enDuration) = TmpAry(4) & " " & TmpAry(5)
                        AryItemList(RowNo, enCost) = TmpAry(6)
                        RowNo = RowNo + 1
                    End If
                    
                Case Is = 7
                    If TmpAry(0) <> "time" And TmpAry(0) <> "Total" Then
                        If Category = "Received Calls Overseas" Or Category = "Overseas Texts Received" Then
                            AryItemList(RowNo, enIndex) = i
                            AryItemList(RowNo, enTime) = TmpAry(0)
                            AryItemList(RowNo, enCategory) = Category
                            AryItemList(RowNo, enDescription) = TmpAry(3) & " " & TmpAry(4)
                            AryItemList(RowNo, enDuration) = TmpAry(5) & " " & TmpAry(6)
                            AryItemList(RowNo, enCost) = TmpAry(7)
                            AryItemList(RowNo, enPhoneNo) = PhoneNum
                            AryItemList(RowNo, enItemDate) = ItemDate
                            RowNo = RowNo + 1
                        Else
                            AryItemList(RowNo, enIndex) = i
                            AryItemList(RowNo, enTime) = TmpAry(0)
                            AryItemList(RowNo, enCategory) = Category
                            AryItemList(RowNo, enDescription) = TmpAry(2) & " " & TmpAry(3)
                            AryItemList(RowNo, enDuration) = TmpAry(4) & " " & TmpAry(5)
                            AryItemList(RowNo, enCost) = TmpAry(7)
                            AryItemList(RowNo, enPhoneNo) = PhoneNum
                            AryItemList(RowNo, enItemDate) = ItemDate
                            RowNo = RowNo + 1
                        End If
                    End If
                Case Is = 8
                    If TmpAry(0) <> "time" And TmpAry(0) <> "Total" Then
                        AryItemList(RowNo, enIndex) = i
                        AryItemList(RowNo, enTime) = TmpAry(0)
                        AryItemList(RowNo, enCategory) = Category
                        AryItemList(RowNo, enDescription) = TmpAry(2) & " " & TmpAry(3)
                        AryItemList(RowNo, enDuration) = TmpAry(4) & " " & TmpAry(5)
                        AryItemList(RowNo, enCost) = TmpAry(6)
                        RowNo = RowNo + 1
                    End If
                Case Is = 9
                    If TmpAry(0) <> "time" And TmpAry(0) <> "Total" Then
                        AryItemList(RowNo, enIndex) = i
                        AryItemList(RowNo, enTime) = TmpAry(0)
                        AryItemList(RowNo, enCategory) = Category
                        AryItemList(RowNo, enDescription) = TmpAry(2) & " " & TmpAry(3)
                        AryItemList(RowNo, enDuration) = TmpAry(4) & " " & TmpAry(5)
                        AryItemList(RowNo, enCost) = TmpAry(6)
                        RowNo = RowNo + 1
                    End If
            End Select
        End If
    Next
    
    ItemisationExt = AryItemList
    
    Set PDTextSelect = Nothing
    Set PDFPage = Nothing
    Set OddCol1 = Nothing
    Set OddCol2 = Nothing
    Set EvenCol1 = Nothing
    Set EvenCol2 = Nothing
    Set AcroRectTmp = Nothing
    Set PDTextSelect = Nothing
    Set JSO = Nothing
End Function

' ===============================================================
' GetQuads
' Gets co-ordinates of words on page
' ---------------------------------------------------------------
Public Sub GetQuads()
    Dim AcroAVDoc As Acrobat.AcroAVDoc
    Dim AcroPDDoc As Acrobat.AcroPDDoc
    Dim PDTextSelect As AcroPDTextSelect
    Dim PDFPage As AcroPDPage
    Dim AcroRect As New Acrobat.AcroRect
    Dim AcroPointxy As Acrobat.AcroPoint
    Dim i As Integer
    Dim JSO As Object
    Dim StrText As String
    Dim PDFPath As String
    Dim PageNo As Integer
    Dim Quad

    PageNo = 26
    
    Set AcroApp = CreateObject("AcroExch.App")
    
    PDFPath = "C:\Users\Julian\OneDrive\Documents\OneSheet\Customers\CTech Group\Vodafone PDFs\Vodafone_Example_bill.pdf"
    
    Set AcroAVDoc = CreateObject("AcroExch.AVDoc")
   
    If AcroAVDoc.Open(PDFPath, "Accessing PDF's") Then
    
        Set AcroPDDoc = AcroAVDoc.GetPDDoc()
        Set PDFPage = AcroPDDoc.AcquirePage(PageNo)
        Set JSO = AcroPDDoc.GetJSObject
            
        For i = 0 To JSO.getPageNumWords(PageNo) - 1
            DoEvents
            Quad = JSO.getPageNthWordQuads(PageNo, i)
            Debug.Print "Element " & i, JSO.getPageNthWord(PageNo, i), _
                "Pt1 " & Format(Quad(0)(0), 0) & "," & Format(Quad(0)(1), 0), _
                ; "Pt2 " & Format(Quad(0)(2), 0) & "," & Format(Quad(0)(3), 0), _
                ; "Pt3 " & Format(Quad(0)(4), 0) & "," & Format(Quad(0)(5), 0), _
                ; "Pt4 " & Format(Quad(0)(6), 0) & "," & Format(Quad(0)(7), 0)
        Next
    End If
    
    Set AcroApp = Nothing
    Set AcroAVDoc = Nothing
    Set AcroPDDoc = Nothing
    Set PDFPage = Nothing
    Set JSO = Nothing
End Sub

' ===============================================================
' progress
' Updates progress bar
' ---------------------------------------------------------------
Sub progress(pctCompl As Single)

    FrmProgress.LblText.Caption = Format(pctCompl, "0") & "% Completed"
    FrmProgress.LblBar.Width = pctCompl * 2
    
    DoEvents

End Sub

' ===============================================================
' GetCategory
' Returns call category
' ---------------------------------------------------------------
Private Function GetCategory(InString As String) As String
        
        Select Case InString
            Case "Calls while in the UK"
                GetCategory = "UK Calls"
                
            Case "Calls"
                GetCategory = "UK Calls"
        
            Case "Calls continued"
                GetCategory = "UK Calls"
       
            Case "Messaging, mobile internet"
                GetCategory = "UK Messaging, mobile internet"
        
            Case "Messaging, mobile internet sent while in the"
                GetCategory = "UK Messaging, mobile internet"
        
            Case "Messaging, mobile internet sent while abroad"
                GetCategory = "Overseas Messaging, mobile internet"
            
            Case "Mobile internet sent while in the UK"
                GetCategory = "UK Mobile Internet"
                
            Case "Mobile internet sent while abroad"
                GetCategory = "Overseas Mobile Internet"
                
            Case "Text messaging received while abroad"
                GetCategory = "Overseas Texts Received"
                
             Case "Text messaging sent while abroad"
                GetCategory = "Overseas Texts Sent"
                
            Case "Calls made while abroad"
                GetCategory = "Overseas Calls"
            
            Case "Calls while abroad"
                GetCategory = "Overseas Calls"
            
            Case "Calls received while abroad"
                GetCategory = "Received Calls Overseas"
            
            Case "Mobile internet sent while abroad"
                GetCategory = "Mobile internet sent while abroad"
                
            Case "Purchases"
                GetCategory = "Purchases"
                
        End Select
        
End Function

' ===============================================================
' HasNumber
' checks to see if string has a number
' ---------------------------------------------------------------
Function HasNumber(strData As String) As Boolean
 Dim i As Integer
 
 For i = 1 To Len(strData)
    If IsNumeric(Mid(strData, i, 1)) Then
        HasNumber = True
        Exit Function
    End If
 Next i
 
End Function