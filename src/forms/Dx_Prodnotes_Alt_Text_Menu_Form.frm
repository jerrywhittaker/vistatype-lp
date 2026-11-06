VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_Prodnotes_Alt_Text_Menu_Form 
   Caption         =   "Prodnotes and Alt-Text (417)"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5805
   OleObjectBlob   =   "Dx_Prodnotes_Alt_Text_Menu_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_Prodnotes_Alt_Text_Menu_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Dx_Prodnotes_Alt_Text_Menu_Form - Prodnotes and Alt-Text (417), braille File Cleanup group
' Opened by Dx_Prodnotes_Alt_Text_Menu_Starter.
'
' Version: 1.1  Date: 11/6/2026 - Jerry's layout, made in the VBA editor: Delete All Prodnotes
'                                 (Alt+D), Convert All Prodnotes to TN Style (Alt+C), Resize
'                                 Throughout Pictures in a Selected Range - Edit Alt-Text (Alt+T),
'                                 Locate the Picture in the Clipboard (Alt+L), Cancel (Alt+C) full
'                                 width at the bottom; an icon beside each button; the form taller.
' Version: 1.0  Date: 11/6/2026 - new (Jerry). Delete Prodnotes and Prodnote to TN moved here off
'                                 the braille ribbon; Resize Throughout Pictures in a Selected
'                                 Range and Locate the Picture in the Clipboard are the same
'                                 macros as the buttons on the large-print Background and Picture
'                                 Tools (344). Alt+D, Alt+T, Alt+R, Alt+L, Alt+C.
'
' EVERY BUTTON GOES THROUGH THE BRIDGE. This menu is modal, and the two picture tools open boxes
' that stay open, which cannot paint or be clicked while a modal form's handler is still running.
' So the macro is named in Sh_BridgeTargetMacro and started by Application.OnTime once the menu
' has gone, as Background and Picture Tools (344) does. The prodnote macros go the same way so
' that all four start alike, with nothing of this menu still on screen.
'
' The menu checks no template: each macro checks what it needs itself. Prodnote to TN asks for
' the BANA template (394); Delete Prodnotes and the two picture tools work on any book.

Private Sub Run_After_Menu(ByVal macroName As String)
    Dx_Prodnotes_Alt_Text_Menu_Form.Hide
    Sh_BridgeTargetMacro = macroName
    Application.OnTime Now, "Sh_StartSpinnerBridge"
    Unload Me
End Sub

Private Sub DeleteProdnotesButton_Click()
    Run_After_Menu "Sh_Delete_Prodnote_Paragraphs"
End Sub

Private Sub ProdnoteToTNButton_Click()
    Run_After_Menu "Dx_Change_Prodnotes_To_Transcriber_Notes"
End Sub

Private Sub SamePictureRangeButton_Click()
    Run_After_Menu "Lp_Resize_Pictures_In_Range"
End Sub

Private Sub LocatePictureButton_Click()
    Run_After_Menu "Lp_Locate_Picture"
End Sub

Private Sub Cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub userform_terminate() 'red X was clicked - nothing runs, the same as Cancel
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Start the form centered inside the Word window (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub
