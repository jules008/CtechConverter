Attribute VB_Name = "Module1"
Option Explicit

Dim formApp As AFORMAUTLib.AFormApp
Dim acroForm As AFORMAUTLib.Fields
Dim field As AFORMAUTLib.field
Dim bOK As Boolean
Dim avDoc As CAcroAVDoc
Dim pdDoc As CAcroPDDoc

Private Sub CreateFieldsButton_Click()


' Open the sample PDF file
    Set avDoc = CreateObject("AcroExch.AVDoc")
    bOK = avDoc.Open("G:\Development Areas\CTech Converter\test.pdf", "Forms Automation Demo")

'If everything was OK opening the PDF, we now instantiate the Forms
    'Automation object.
    If (bOK) Then
        Set formApp = CreateObject("AFormAut.App")
        Set acroForm = formApp.Fields
    Else
        Set avDoc = Nothing
        MsgBox "Failed to open PDF Document. Aborting..."
        End
    End If


'Create a button that represents an image
    Set field = acroForm.Add("Logo Button", "button", 0, 175, 50, 225, 100)
    field.BorderStyle = "beveled"
    'field.Highlight = "push"
    field.IsReadOnly = True
    field.SetButtonIcon "N", "G:\Development Areas\CTech Converter\test.pdf", 0


'****************
    'This section converts a AVDoc to a PDDoc. Then we use PDDoc.Save to
    'save the PDF since there isn't a viewer level method to save an AVDoc.

'These lines are commented out to provide user feedback
    'If present, no doc is visible and original form is changed
    'so it isn't obvious what the sample did, and it doesn't work
    'twice in a row.
    'It's best to have Acrobat open and visible to see form creation
    '''''''''
    Set pdDoc = avDoc.GetPDDoc
    bOK = pdDoc.Save(1, "G:\Development Areas\CTech Converter\test.pdf")
    If bOK = False Then
        MsgBox "Unable to Save the PDF file"
    End If

' Close the AVDoc
    avDoc.Close (False)
    '''''''''
    'End provide User Feedback
    '****************

CleanUp:

End Sub

Private Sub ExitButton_Click()
    ' End the program
    End

End Sub
