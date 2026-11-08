VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Review_Prodnotes_Form 
   Caption         =   "Review Prodnotes (426)"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640.001
   OleObjectBlob   =   "Lp_Review_Prodnotes_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Review_Prodnotes_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Lp_Review_Prodnotes_Form
'
' Version: 1.0  Date: 11/7/2026
'
' Review Prodnotes (426), the large print box opened by the Review/Edit/Delete Prodnotes button
' on Prodnote Review (425) - Jerry, 11/7/2026. The large print twin of the braille Review
' Prodnotes (418): the same Next, Previous, Delete This Prodnote and Done, and in place of
' Convert This to TN Style, Convert This to a Box Red Style and Convert This to a Box Blue Style.
' It stays open - shown with .Show vbModeless by Lp_Review_Prodnotes - so the transcriber can read
' and edit each prodnote in the book between presses. Every button calls a macro in
' LPandBrlMacros; nothing here touches the book.
'
' KeyDown on every button, because while this box has the keyboard Word never sees F6 at all.
' QueryClose sends the title bar's X through the same Done as the button.

Private Sub NextButton_Click()
    Dx_Rvp_Next
End Sub

Private Sub PrevButton_Click()
    Dx_Rvp_Previous
End Sub

Private Sub DeleteButton_Click()
    Dx_Rvp_Delete
End Sub

Private Sub BoxRedButton_Click()
    Lp_Rvp_Box_Red
End Sub

Private Sub BoxBlueButton_Click()
    Lp_Rvp_Box_Blue
End Sub

Private Sub DoneButton_Click()
    Dx_Rvp_Done
End Sub

Private Sub NextButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Dx_Rvp_KeyToDocument
    End If
End Sub

Private Sub PrevButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Dx_Rvp_KeyToDocument
    End If
End Sub

Private Sub DeleteButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Dx_Rvp_KeyToDocument
    End If
End Sub

Private Sub BoxRedButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Dx_Rvp_KeyToDocument
    End If
End Sub

Private Sub BoxBlueButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Dx_Rvp_KeyToDocument
    End If
End Sub

Private Sub DoneButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Dx_Rvp_KeyToDocument
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Dx_Rvp_Done
    End If
End Sub

Private Sub UserForm_Initialize()
    KeyHelp.Caption = "Use F6 or Shift+F6 to move between the document and this box. " _
                    & "Edit the selected prodnote in the document; Delete and the two Convert " _
                    & "buttons act on it."
    TallyText.Caption = ""

    ' CENTERED ON THE WORD WINDOW, guarded as 404 and 418 are: reading Application.Left has hung
    ' a Word with no desktop.
    Dim wantLeft As Single, wantTop As Single
    Dim placed As Boolean

    On Error Resume Next
    wantLeft = Application.Left + (Application.Width - Me.Width) / 2
    wantTop = Application.Top + (Application.Height - Me.Height) / 2
    placed = (Err.Number = 0)
    Err.Clear
    If placed Then
        Me.StartUpPosition = 0
        Me.Left = wantLeft
        Me.Top = wantTop
        Err.Clear
    End If
    On Error GoTo 0
End Sub
