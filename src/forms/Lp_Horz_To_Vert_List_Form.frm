VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Horz_To_Vert_List_Form 
   Caption         =   "Convert Hozizontal List to Vertical List (348)"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8205.001
   OleObjectBlob   =   "Lp_Horz_To_Vert_List_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Horz_To_Vert_List_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Lp_Horz_To_Vert_List_Form

' Version 3.0 9/23/2026 - THE BOX STAYS OPEN - Jerry. Shown vbModeless by Lp_Hvb_Start; Cancel is
'                        Done (his layout, with the short-cut key and the F6 line on the box).
'                        Okay hands the choices to Lp_Hvb_Okay, which converts the list selected
'                        at that moment and gives the keyboard back to the book. F6 and Shift+F6
'                        go back to the book from every control.
' Version 2.0 8/12/2026 - the OK button only records the three choices and closes. The work,
'                        and the temporary document that never appears, are in
'                        Lp_Horz_To_Vert_Hidden
' Version 1.8 8/11/2026 - serves BOTH ribbon tabs; Dx_Horz_To_Vert_List_Form is gone. The
'                        round trip through the temp document picks its route from the
'                        document being worked on - see Sh_Copy_To_Temp_Doc, which was
'                        removed on 8/31/2026 once nothing called it
' Version 1.7 8/10/2026 - an ordered list is recognized by its SEQUENCE first
'                        (Lp_Split_Ordered_List_Sequence), falling back to the passes below
'                        when nothing convincing is found
' Version 1.6 8/2/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
' Version 1.5 3/10/2026 - trapped crash on sort of non-sortable selection
' Version 1.4 10/3/2018 - added ordinary bullet
' Version 1.3 8/23/2018 - added check for ending para mark
' Version 1.2 7/5/2018
' Version 1.1 5/17/2018
' Version 1.0 4/3/2018
'
' Every button does one thing: call the macro that does the work. Nothing here touches the book,
' and the option buttons and the tick box hold their own choices between presses.
'
' KeyDown on EVERY control that can hold the keyboard: F6 and Shift+F6 back to the book (Jerry's
' rule, 9/23/2026). A UserForm has no KeyPreview, and while the box has the keyboard Word never
' sees the key, so a control without this strands a transcriber working without a mouse.
' Shift+F6 arrives here as F6 with Shift set.
'
' QueryClose sends the title bar's X through the same Done as the button, or the box would be
' unloaded behind the macro's back and F6 left pointing at a box that has gone.

Private Sub Cmd_Cancel_Click()
    Lp_Hvb_Done
End Sub

Private Sub Cmd_Ok_Click()
    Dim listKind As String

    If Ordered_List Then
        listKind = "ORDERED"
    ElseIf Spaced_List Then
        listKind = "SPACED"
    Else
        listKind = "TABBED"
    End If
    Lp_Hvb_Okay listKind, (AscendingOrderCheckbox = True)
End Sub

Private Sub Cmd_Ok_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Hvb_KeyToDocument
    End If
End Sub

Private Sub Cmd_Cancel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Hvb_KeyToDocument
    End If
End Sub

Private Sub AscendingOrderCheckbox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Hvb_KeyToDocument
    End If
End Sub

Private Sub Ordered_List_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Hvb_KeyToDocument
    End If
End Sub

Private Sub Spaced_List_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Hvb_KeyToDocument
    End If
End Sub

Private Sub TabbedList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Hvb_KeyToDocument
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Lp_Hvb_Done
    End If
End Sub

Sub UserForm_Initialize()
    Me.AscendingOrderCheckbox.Value = True
    
    ' 8/2/2026 - removed "MS_Set_Word_Config_For_Large_Print": opening the document already
    '            configures Word for large print, and re-running it here cost ~40 Options and
    '            AutoCorrect writes plus 19 AutoCorrect entry deletions on every form open.
    '            The other 14 LP forms dropped this call on 7/24/2026; this one was missed.
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub
