VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Horz_To_Vert_List_Form 
   Caption         =   "Convert Hozizontal List to Vertical List"
   ClientHeight    =   4728
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

' Version 2.0 8/12/2026 - the OK button only records the three choices and closes. The work,
'                        and the temporary document that never appears, are in
'                        Lp_Horz_To_Vert_Hidden
' Version 1.8 8/11/2026 - serves BOTH ribbon tabs; Dx_Horz_To_Vert_List_Form is gone. The
'                        round trip through the temp document picks its route from the
'                        document being worked on - see Sh_Copy_To_Temp_Doc
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
Private Sub Cmd_Cancel_Click()
    Unload Me
    End
End Sub

Private Sub Cmd_Ok_Click()
    ' This handler does NOTHING to the document. It records what the transcriber chose and
    ' closes; Lp_Horz_List_To_Vertical does the work once the dialog is gone. Editing a document
    ' from inside a modal form's event handler is a bad place to be and there is no reason to be
    ' there - the choices are all this dialog is for.
    If Ordered_List Then
        Lp_Hv_Kind = "ORDERED"
    ElseIf Spaced_List Then
        Lp_Hv_Kind = "SPACED"
    Else
        Lp_Hv_Kind = "TABBED"
    End If
    Lp_Hv_Sort_Wanted = (AscendingOrderCheckbox = True)
    Lp_Hv_Go = True

    Me.Hide
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

Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub
