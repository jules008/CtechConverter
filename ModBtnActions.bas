Attribute VB_Name = "ModBtnActions"
'===============================================================
' Module ModBtnActions
'---------------------------------------------------------------
' Created by Julian Turner
' OneSheet Consulting
' julian.turner@onesheet.co.uk
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 11 Jul 20
'===============================================================
Option Explicit

' ===============================================================
' FormatButton
' Formats the control to be selected or not
' ---------------------------------------------------------------
Private Sub FormatButton(Btn As Shape, OnState As Boolean)
    
    Btn.Parent.Unprotect
    With Btn
        If OnState Then
            With .Fill
                If Btn.Name = "BtnLoadBill" Then
                    .TwoColorGradient msoGradientHorizontal, 1
                    .ForeColor.RGB = COLOUR_6
                    .BackColor.RGB = COLOUR_6
                Else
                    .TwoColorGradient msoGradientHorizontal, 1
                    .ForeColor.RGB = COLOUR_2
                    .BackColor.RGB = COLOUR_2
                End If
            End With

            With .Line
                If Btn.Name = "BtnLoadBill" Then
                    .ForeColor.RGB = COLOUR_4
                    .Weight = 0
                Else
                    .ForeColor.RGB = COLOUR_4
                    .Weight = 0
                End If
                If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
            End With

            With .TextFrame
                If Btn.Name = "BtnLoadBill" Then
                    .Characters.Font.Bold = False
                    .Characters.Font.Name = "Calibri"
                    .Characters.Font.Color = COLOUR_7
                    .Characters.Font.Size = 11
                    .HorizontalAlignment = xlHAlignCenter
                    .MarginBottom = 0
                    .MarginTop = 0
                Else
                    .Characters.Font.Bold = False
                    .Characters.Font.Name = "Calibri"
                    .Characters.Font.Color = COLOUR_7
                    .Characters.Font.Size = 11
                    .HorizontalAlignment = xlHAlignCenter
                    .MarginBottom = 0
                    .MarginTop = 0
                End If
            End With

            With .Shadow
                .Visible = msoTrue
                .Type = msoShadow30
            End With

        Else
            With .Fill
                If Btn.Name = "BtnLoadBill" Then
                    .TwoColorGradient msoGradientHorizontal, 1
                    .ForeColor.RGB = COLOUR_2
                    .BackColor.RGB = COLOUR_2
                Else
                    .TwoColorGradient msoGradientHorizontal, 1
                    .ForeColor.RGB = COLOUR_1
                    .BackColor.RGB = COLOUR_1
                End If
            End With

            With .Line
                If Btn.Name = "BtnLoadBill" Then
                    .ForeColor.RGB = COLOUR_6
                    .Weight = 1.5
                Else
                    .ForeColor.RGB = COLOUR_2
                    .Weight = 0.75
                End If
                If .Weight = 0 Then .Visible = msoFalse Else .Visible = msoCTrue
            End With

            With .TextFrame
                 If Btn.Name = "BtnLoadBill" Then
                    .Characters.Font.Bold = False
                    .Characters.Font.Name = "Calibri"
                    .Characters.Font.Color = COLOUR_7
                    .Characters.Font.Size = 11
                    .HorizontalAlignment = xlHAlignCenter
                    .MarginBottom = 0
                    .MarginTop = 0
                Else
                    .Characters.Font.Bold = False
                    .Characters.Font.Name = "Calibri"
                    .Characters.Font.Color = COLOUR_3
                    .Characters.Font.Size = 11
                    .HorizontalAlignment = xlHAlignCenter
                    .MarginBottom = 0
                    .MarginTop = 0
                End If
            End With

            With .Shadow
                .Visible = msoTrue
                .Type = msoShadow30
               .Visible = msoFalse
            End With


        End If
    End With
    Btn.Parent.Protect
End Sub

' ===============================================================
' ActionButtonClick
' Animates control of button
' ---------------------------------------------------------------
Public Sub ActionButtonClick(Button As Shape)
    Dim DTime As Double

    FormatButton Button, True

    DTime = Time
    Do While Time < DTime + 1 / 24 / 60 / 60 / 2
        DoEvents
    Loop

    FormatButton Button, False

End Sub

' ===============================================================
' GoToHome
' Activates main screen
' ---------------------------------------------------------------
Public Sub GoToHome(BtnShape As Shape)
    ActionButtonClick BtnShape
    ShtMain.Activate
End Sub




