VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_File_Cleanup_Sub_Menu_Form 
   Caption         =   "Large Print File Cleanup"
   ClientHeight    =   2388
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4110
   OleObjectBlob   =   "Lp_File_Cleanup_Sub_Menu_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_File_Cleanup_Sub_Menu_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lp_File_Cleanup_Sub_Menu_Form
'
' Version: 1.2  Date: 7/24/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
' Version: 1.1  Date: 1/22/2026 - refresh screen before messages
' Version: 1.0  Date: 5/7/2025 - full rewright
'

Private Sub CancelButton_Click()
    Selection.Collapse 'clear selection
    Unload Me
End Sub

Private Sub OkayButton_Click()

    Lp_File_Cleanup_Sub_Menu_Form.Hide

    ' Hold screen updating off across the selected cleanup(s); restored below.
    '
    ' ScreenUpdating alone is NOT enough. With background repagination and live
    ' spell/grammar checking left on, Word re-lays-out and repaints the document AFTER the
    ' macro finishes, painting it in patches - which reads as coloured table rows splashing
    ' across the screen (Jerry, 7/26/2026). Sh_Convert_XML_File_To_Word_Document already
    ' silences the same three for the same reason; this matches it.
    Dim su_Prev As Boolean, pag_Prev As Boolean, spell_Prev As Boolean, gram_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    pag_Prev = Application.Options.Pagination
    spell_Prev = Application.Options.CheckSpellingAsYouType
    gram_Prev = Application.Options.CheckGrammarAsYouType

    Application.ScreenUpdating = False
    Application.Options.Pagination = False
    Application.Options.CheckSpellingAsYouType = False
    Application.Options.CheckGrammarAsYouType = False
    
    If FixCommonErrors Then
        Application.Run MacroName:="Lp_Fix_Common_File_Errors"
    End If
    
    If RemoveParaMarks Then
        Application.Run MacroName:="Lp_Replace_Multiple_Para_Marks_With_Warning"
    End If
    
    ' Restore proofing/pagination BEFORE the screen comes back, so the one repaint that
    ' follows is the finished document rather than a progressive re-layout.
    Application.Options.CheckGrammarAsYouType = gram_Prev
    Application.Options.CheckSpellingAsYouType = spell_Prev
    Application.Options.Pagination = pag_Prev

    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

    If FixCommonErrors Or RemoveParaMarks Then
        MsgBox "Selected cleanup(s) complete", , "VistaType LP (187)"
    Else
        If Not (FixCommonErrors Or RemoveParaMarks) Then
            MsgBox "Canceled... No cleanup selected", , "VistaType LP (188)"
        End If
    End If
    
    Selection.Collapse 'clear selection
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' 7/24/2026 - removed "MS_Set_Word_Config_For_Large_Print": opening the document already
    '             configures Word for large print, and re-running it here reset the user's
    '             Styles-pane options (show filter / sort order) every time this form opened.
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub


