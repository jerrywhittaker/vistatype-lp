VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_Valid_Ref_Pg_No_4_Form 
   Caption         =   "Delete/change/add $pg"
   ClientHeight    =   1155
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9690.001
   OleObjectBlob   =   "Sh_Valid_Ref_Pg_No_4_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sh_Valid_Ref_Pg_No_4_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Sh_Valid_Ref_Pg_No_4_Form
' This for is showing when the user in the main doc for validating  and correcting $pg tags
'
' Version 1.3:  Date: 8/23/2026 - the F6 line is in every button's hover text as well as on
'                                 the form, because a screen reader does not read a Label (Jerry)
' Version 1.2:  Date: 8/23/2026 - the title-bar X now means Done - Exit validation. It used to
'                                 unload the form behind the whole feature's back and leave F6
'                                 pointing at a menu that was gone (review)
' Version 1.1:  Date: 8/23/2026 - titled "Delete/change/add $pg" instead of "UserForm1", new
'                                 KeyHelp line, F6 out of the menu, and the box now sits ABOVE
'                                 the page instead of on top of it (Jerry)
' Version 1.0:  Date: 7/27/2026
'
' --- at the very top of the UserForm module (declarations section) ---
#If VBA7 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
#End If
Private Const LOGPIXELSY As Long = 90

' Said in three places on this form - the line under the buttons, and every button's hover text -
' so it is written once. 8/23/2026.
Private Const KEY_HELP_TEXT As String = _
    "Use F6 or Shift+F6 to move between the document and this menu."

Private Sub DoneButton_Click()
    'Closes the tag list without saving and leaves you in your document.
    Sh_PgVal_Done
End Sub

Private Sub LocateInDocButton_Click()
    'Despite the control name, on THIS form the button returns to the validation list -
    'and lands on the NEXT tag down, so your place in the list is never lost.
    Sh_PgVal_ReturnToTempAndAdvance
End Sub

Private Sub ShowDirectionsButton_Click()
    'Me.Hide
    Sh_Valid_Ref_Pg_No_3_Form.Show ' directions for use
    'Me.Show
End Sub

' F6 and Shift+F6 OUT of this menu and back into the document.
'
' On every button, because while this form has the keyboard Word never sees the key at all - so
' the key binding that brings the transcriber here cannot take her back. A Label cannot hold the
' focus, so there is nowhere else to put these. The whole loop is written up in ShNonModalMessage
' above Sh_PgVal_BindKeys.
'
' KeyCode is set to 0 afterwards so the keystroke stops here rather than also being handed on.
'
' Version: 1.0  Date: 8/23/2026
Private Sub LocateInDocButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Sh_PgVal_KeyToDocument
    End If
End Sub

Private Sub DoneButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Sh_PgVal_KeyToDocument
    End If
End Sub

Private Sub ShowDirectionsButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Sh_PgVal_KeyToDocument
    End If
End Sub

' THE TITLE-BAR X MEANS "DONE - EXIT VALIDATION". Without this it meant something far worse: the
' X unloads the form without going near Sh_PgVal_Done or either swap, so the F6 key assignment
' stayed in force with no menu to go to, and every F6 for the rest of the session ran at a form
' that was not there. Found in review, 8/23/2026, before any build carried it.
'
' Cancel = True stops Word's own unload and hands the whole teardown to Sh_PgVal_Done, which
' closes the tag list, puts the transcriber back in her document, gives F6 back to Word and
' unloads both menus - in that order.
'
' Version: 1.0  Date: 8/23/2026
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Sh_PgVal_Done
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim pxLeft As Long, pxTop As Long, pxWidth As Long, pxHeight As Long
    Dim pxPerPt As Double, zoomF As Double
    Dim topRange As Range
    #If VBA7 Then
        Dim hDC As LongPtr
    #Else
        Dim hDC As Long
    #End If

    ' FIRST, before anything that can raise. Everything below reads the active window and the
    ' active document, and if any of it ever failed the line telling a keyboard user how to reach
    ' this menu would be whatever the layout file happens to hold.
    KeyHelp.Caption = KEY_HELP_TEXT

    ' AND IN EVERY BUTTON'S HOVER TEXT (Jerry, 8/23/2026). The line on the form is a Label, and a
    ' Label is the one control a screen reader does not announce - it reads the window title and
    ' whatever has the focus. Hover text on a button IS announced, and F6 lands the keyboard on a
    ' button, so this is the form in which the instruction actually reaches the transcriber it was
    ' written for. It stays on the label as well, for a low-vision transcriber who is reading the
    ' screen rather than listening to it.
    '
    ' Built from each caption rather than appended to what is already there, so it is the same
    ' whether this runs once or a hundred times.
    LocateInDocButton.ControlTipText = LocateInDocButton.Caption & ". " & KEY_HELP_TEXT
    DoneButton.ControlTipText = DoneButton.Caption & ". " & KEY_HELP_TEXT
    ShowDirectionsButton.ControlTipText = ShowDirectionsButton.Caption & ". " & KEY_HELP_TEXT

    On Error Resume Next
    Me.StartUpPosition = 0

    ' Screen DPI: device pixels per point (this is screen/form space)
    hDC = GetDC(0)
    pxPerPt = GetDeviceCaps(hDC, LOGPIXELSY) / 72
    ReleaseDC 0, hDC

    ' Document zoom - document space is scaled by this on screen
    zoomF = ActiveWindow.View.Zoom.Percentage / 100

    ' Collapsed range at the start of the body text
    Set topRange = ActiveDocument.Content
    topRange.Collapse wdCollapseStart

    On Error Resume Next
    ActiveWindow.GetPoint pxLeft, pxTop, pxWidth, pxHeight, topRange
    On Error GoTo 0

    If pxWidth > 0 Then
        ' Center over the text column (screen px -> form points: DPI only)
        Me.Left = (pxLeft + pxWidth / 2) / pxPerPt - Me.Width / 2

        ' Step up from body-text top to the page's top edge.
        ' TopMargin is a document measure -> scale by DPI *and* zoom.
        '
        ' AND THEN THE FORM'S OWN HEIGHT AGAIN, so the box sits entirely ABOVE the page instead
        ' of on top of it (Jerry, 8/23/2026: it should not cover any of the document). Its bottom
        ' edge lands on the page's top edge.
        Me.Top = pxTop / pxPerPt - ActiveDocument.PageSetup.TopMargin * zoomF - Me.Height

        ' UNLESS THERE IS NOT ROOM, and then it goes back where it used to be - on the page's top
        ' edge. There often is not room: at the 150% to 200% zoom large print work is done at, the
        ' page top sits high in the window and a menu's height above it lands on the ribbon, the
        ' tab row, and the Quick Access Toolbar that VistaType's own eight icons live on. Covering
        ' the top margin of the page is the smaller loss - and the directions on the other dialog
        ' tell her to press a ribbon button. Found in review, 8/23/2026.
        If Me.Top < Application.Top Then
            Me.Top = pxTop / pxPerPt - ActiveDocument.PageSetup.TopMargin * zoomF
        End If
    Else
        ' Fallback if the top of the doc is scrolled off-screen
        Me.Left = Application.Left + (Application.Width - Me.Width) / 2
        Me.Top = Application.Top + 120
    End If
End Sub

