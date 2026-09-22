VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Same_Pic_Range_Form 
   Caption         =   "Resize Pictures in a Selected Range (375)"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "Lp_Same_Pic_Range_Form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Same_Pic_Range_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Lp_Same_Pic_Range_Form
'
' Version: 1.0  Date: 9/22/2026
'
' The box for Resize Pictures Used Throughout a Selected Range (Jerry, 9/22/2026). It is MODELESS
' - ShowModal is False in the designer header - so the transcriber can click pictures in her book
' while it is up. That is the whole point: Word allows one selection at a time, so the range is
' held by the macro and the selection is left free for the picture she is working on.
'
' Every button does one thing: call the macro that does the work. Nothing here touches the book,
' and nothing here holds state - the option buttons and the tick box hold their own between
' presses, which is what makes the choices carry over.
'
' KeyDown on every button, because while this box has the keyboard Word never sees the key at all.
' QueryClose sends the title bar's X through the same Done as the button, or the box would be
' unloaded behind the macro's back and F6 left pointing at a box that has gone.

Private Sub ApplyButton_Click()
    Lp_Rst_Apply
End Sub

Private Sub DoneButton_Click()
    Lp_Rst_Done
End Sub

Private Sub LastSizeButton_Click()
    Lp_Rst_Apply_Last_Size
End Sub

Private Sub LastSizeButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Rst_KeyToDocument
    End If
End Sub

Private Sub ApplyButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Rst_KeyToDocument
    End If
End Sub

Private Sub DoneButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Rst_KeyToDocument
    End If
End Sub

' EVERY control that can hold the keyboard needs this, not only the two buttons. A UserForm has no
' KeyPreview, so while the tick box or one of the position buttons has the focus, F6 reaches
' neither this code nor Word, and a transcriber working without a mouse is stranded in the box.
Private Sub JoinNextParaButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Rst_KeyToDocument
    End If
End Sub

Private Sub LeftAlignButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Rst_KeyToDocument
    End If
End Sub

Private Sub CenterButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Rst_KeyToDocument
    End If
End Sub

Private Sub LeaveAsIsButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Rst_KeyToDocument
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Lp_Rst_Done
    End If
End Sub

Private Sub UserForm_Initialize()
    ' ALT+V FOR "LEAVE AS IS". Alt+A belongs to Apply, and two controls on one letter makes Alt+A
    ' cycle between them instead of pressing either. Set here rather than in the .frx so the
    ' choice can be read in the source.
    LeaveAsIsButton.Accelerator = "v"

    ' Jerry's wording, 9/22/2026. Set here so it reads in the source rather than only in the .frx.
    JoinNextParaButton.Caption = "Move the picture to the paragraph which follows it."
    JoinNextParaButton.Accelerator = "M"

    ' Nothing to copy from until a size has been used, so the button starts dead and
    ' Lp_Rst_Offer_Last_Size brings it to life after the first Apply.
    LastSizeButton.Enabled = False

    ' The F6 line is a promise, so it is only made when the key was really taken. A $pg validation
    ' already running keeps F6, and this box stands aside rather than taking it away.
    If Lp_Rst_Keys_Are_Bound() Then
        KeyHelp.Caption = "Use F6 or Shift+F6 to move between the document and this box."
    Else
        KeyHelp.Caption = "Close the $pg validation box to use F6 with this one."
    End If

    ' CENTERED ON THE WORD WINDOW - Jerry, 9/22/2026. Guarded, because reading Application.Left
    ' has hung a Word with no desktop, and the numbers are read BEFORE CenterOwner is given up, so
    ' a box that cannot be placed still opens where the designer put it rather than in the corner.
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
