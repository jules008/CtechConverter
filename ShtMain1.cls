VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtMain1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' ShtMain
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 13 May 20
'===============================================================
Option Explicit

' ===============================================================
' BtnImportClick
' ---------------------------------------------------------------
Public Sub BtnImportClick()
    Me.Unprotect
    ModReports.DataImport
    Me.Protect
End Sub

' ===============================================================
' BtnLARGenClick
' ---------------------------------------------------------------
Public Sub BtnLARGenClick()
    ActionButtonClick Me.Shapes("BtnLARGen")
    LookAheadRep [mpp_filepath]
End Sub

' ===============================================================
' BtnExcRepGenClick
' ---------------------------------------------------------------
Public Sub BtnExcRepGenClick()
    ActionButtonClick Me.Shapes("BtnExcRepGen")
    ExceptionReport
End Sub

' ===============================================================
' BtnDepRepGenClick
' ---------------------------------------------------------------
Public Sub BtnDepRepGenClick()
    ActionButtonClick Me.Shapes("BtnDepRepGen")
    ModReports.DependancyRep
End Sub

' ===============================================================
' ActionButtonClick
' ---------------------------------------------------------------
Public Sub ActionButtonClick(Button As Shape)
    Dim DTime As Double

    On Error GoTo ErrorHandler

    If [mpp_filepath] = "" Then
        MsgBox "Please import a Project File", vbOKOnly + vbExclamation
        Exit Sub
    End If

    FormatButton Button, True

    DTime = Time
    Do While Time < DTime + 1 / 24 / 60 / 60 / 2
        DoEvents
    Loop

    FormatButton Button, False


Exit Sub
ErrorHandler:
    MsgBox "An error has occured, unable to open the Project File", vbOKCancel + vbCritical

End Sub

' ===============================================================
' BtnLvl1Click
' ---------------------------------------------------------------
Public Sub BtnLvl1Click()
    Me.Unprotect
    [level] = 1
    FormatButton Me.Shapes("BtnLvl1"), True
    FormatButton Me.Shapes("BtnLvl2"), False
    FormatButton Me.Shapes("BtnLvl3"), False
    FormatButton Me.Shapes("BtnLvl4"), False
    FormatButton Me.Shapes("BtnLvlAll"), False
    Me.Protect
End Sub

' ===============================================================
' BtnLvl2Click
' ---------------------------------------------------------------
Public Sub BtnLvl2Click()
    Me.Unprotect
    [level] = 2
    FormatButton Me.Shapes("BtnLvl1"), False
    FormatButton Me.Shapes("BtnLvl2"), True
    FormatButton Me.Shapes("BtnLvl3"), False
    FormatButton Me.Shapes("BtnLvl4"), False
    FormatButton Me.Shapes("BtnLvlAll"), False
    Me.Protect
End Sub

' ===============================================================
' BtnLvl3Click
' ---------------------------------------------------------------
Public Sub BtnLvl3Click()
    Me.Unprotect
    [level] = 3
    FormatButton Me.Shapes("BtnLvl1"), False
    FormatButton Me.Shapes("BtnLvl2"), False
    FormatButton Me.Shapes("BtnLvl3"), True
    FormatButton Me.Shapes("BtnLvl4"), False
    FormatButton Me.Shapes("BtnLvlAll"), False
    Me.Protect
End Sub

' ===============================================================
' BtnLvl4Click
' ---------------------------------------------------------------
Public Sub BtnLvl4Click()
    Me.Unprotect

    [level] = 4
    FormatButton Me.Shapes("BtnLvl1"), False
    FormatButton Me.Shapes("BtnLvl2"), False
    FormatButton Me.Shapes("BtnLvl3"), False
    FormatButton Me.Shapes("BtnLvl4"), True
    FormatButton Me.Shapes("BtnLvlAll"), False

    Me.Protect
End Sub

' ===============================================================
' BtnLvlAllClick
' ---------------------------------------------------------------
Public Sub BtnLvlAllClick()
    Me.Unprotect

    [level] = 9999
    FormatButton Me.Shapes("BtnLvl1"), False
    FormatButton Me.Shapes("BtnLvl2"), False
    FormatButton Me.Shapes("BtnLvl3"), False
    FormatButton Me.Shapes("BtnLvl4"), False
    FormatButton Me.Shapes("BtnLvlAll"), True

    Me.Protect
End Sub

' ===============================================================
' FormatButton
' Formats the control to be selected or not
' ---------------------------------------------------------------
Private Sub FormatButton(Btn As Shape, OnState As Boolean)
    Me.Unprotect
    With Btn
        If OnState Then
            With .Fill
                .TwoColorGradient msoGradientHorizontal, 1
                .ForeColor.RGB = COLOUR_2
                .BackColor.RGB = COLOUR_2
            End With

            With .Line
                .ForeColor.RGB = COLOUR_4
                .Weight = 0
                If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
            End With

            With .TextFrame
                .Characters.Font.Bold = False
                .Characters.Font.Name = "Calibri"
                .Characters.Font.Color = COLOUR_7
                .Characters.Font.Size = 10
                .HorizontalAlignment = xlHAlignCenter
            End With

            With .Shadow
                .Visible = msoTrue
                .Type = msoShadow30
            End With

        Else
            With .Fill
                .TwoColorGradient msoGradientHorizontal, 1
                .ForeColor.RGB = COLOUR_8
                .BackColor.RGB = COLOUR_8
            End With

            With .Line
                .ForeColor.RGB = COLOUR_2
                .Weight = 0
                If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
            End With

            With .TextFrame
                .Characters.Font.Bold = False
                .Characters.Font.Name = "Calibri"
                .Characters.Font.Color = COLOUR_7
                .Characters.Font.Size = 10
                .HorizontalAlignment = xlHAlignCenter
            End With

            With .Shadow
                .Visible = msoTrue
                .Type = msoShadow30
               .Visible = msoFalse
            End With


        End If
    End With
    Me.Protect
End Sub

' ===============================================================
' TxtLAValue_Change
' Text box for number of weeks look ahead
' ---------------------------------------------------------------
Private Sub TxtLAValue_Change()
    Dim LAValue As Variant

    Sleep 200
    LAValue = TxtLAValue
    If LAValue = "" Then Exit Sub

    If Not IsNumeric(TxtLAValue) Then
        MsgBox "Please only enter numbers", vbOKOnly + vbExclamation, "Look Ahead Value"
        Exit Sub
    End If

    If TxtLAValue < 1 Or TxtLAValue > 500 Then
        MsgBox "Please only enter number between 1 and 500", vbOKOnly + vbExclamation, "Look Ahead Value"
        Exit Sub
    End If
    Me.Unprotect
    [LA_PERIOD] = TxtLAValue
    Me.Protect
End Sub
