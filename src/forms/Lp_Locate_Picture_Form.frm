VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Locate_Picture_Form 
   Caption         =   "Locate This Picture (404)"
   ClientHeight    =   2220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "Lp_Locate_Picture_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Locate_Picture_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Lp_Locate_Picture_Form
'
' Version: 1.0  Date: 11/1/2026
'
' The box for Locate the Picture on the Clipboard (Jerry, 11/1/2026). It stays open - shown with
' .Show vbModeless by Lp_Locate_Picture - so the transcriber can work on each copy in the book
' between presses. Every button calls a macro in LPandBrlMacros; nothing here touches the book.
'
' KeyDown on every button, because while this box has the keyboard Word never sees F6 at all.
' QueryClose sends the title bar's X through the same Done as the button.

Private Sub NextButton_Click()
    Lp_Locp_Next
End Sub

Private Sub PrevButton_Click()
    Lp_Locp_Previous
End Sub

Private Sub DoneButton_Click()
    Lp_Locp_Done
End Sub

Private Sub NextButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Locp_KeyToDocument
    End If
End Sub

Private Sub PrevButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Locp_KeyToDocument
    End If
End Sub

Private Sub DoneButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Locp_KeyToDocument
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Lp_Locp_Done
    End If
End Sub

Private Sub UserForm_Initialize()
    KeyHelp.Caption = "Use F6 or Shift+F6 to move between the document and this box."

    ' CENTERED ON THE WORD WINDOW, guarded as 375 is: reading Application.Left has hung a Word
    ' with no desktop.
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
