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
' Version: 1.4  Date: 9/6/2026 - shows the PROGRESS BAR instead of Sh_Please_Wait_Form, and
'                               gained an error handler: the old three lines had none, so a
'                               failure in the cleanup left the box on screen for good
' Version: 1.3  Date: 8/3/2026 - shows Sh_Please_Wait_Form, "Fixing common file errors", while Lp_Fix_Common_File_Errors runs
' Version: 1.2  Date: 7/24/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
' Version: 1.1  Date: 1/22/2026 - refresh screen before messages
' Version: 1.0  Date: 5/7/2025 - full rewright
'

Private Sub CancelButton_Click()
    Selection.Collapse 'clear selection
    Unload Me
End Sub

Private Sub OkayButton_Click()

    Dim failNumber As Long
    Dim failText As String

    Lp_File_Cleanup_Sub_Menu_Form.Hide

    ' Hold screen updating off across the selected cleanup(s); restored below.
    ' Do NOT also disable Options.Pagination here - see the note in Lp_Attach_The_Template.
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    If FixCommonErrors Then
        ' THE BAR, from 9/6/2026 - Jerry's call: one progress indicator across the add-in.
        ' This showed Sh_Please_Wait_Form from 8/3/2026, which turned a spinner and said one
        ' fixed sentence. The bar counts the 32 passes and names each one as it starts, and
        ' its own spinner keeps turning through the slow ones - Lp_Fix_Normal_Styles above
        ' all. Modeless, so the macro carries straight on.
        '
        ' THE HANDLER IS WHY THIS IS NOT THREE LINES. Sh_Show_Please_Wait / Run / Hide had no
        ' error trap between them, so a failure inside the cleanup left the box on screen for
        ' good: Application.Run swallows the error, so RibbonAction never sees it and
        ' Sh_Report_Error - the only thing that would have taken the box down - never runs.
        ' A stranded box over a Word that will not repaint is the "it hung" report.
        On Error GoTo CleanupFailed
        Sh_Progress_Open "Fixing common file errors"
        Lp_Fix_Common_File_Errors
        Sh_Progress_Say 100, "Finished"
        Sh_Progress_Close
        On Error GoTo 0
    End If
    
    If RemoveParaMarks Then
        Application.Run MacroName:="Lp_Replace_Multiple_Para_Marks_With_Warning"
    End If
    
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
    Exit Sub

CleanupFailed:
    ' TAKE A COPY OF THE ERROR FIRST, before anything else is called. Sh_Progress_Close
    ' runs On Error Resume Next and Err.Clear inside itself - as every one of these
    ' helpers does, so that a cleanup path can never raise a second time - and that wipes
    ' Err. Reporting Err.Number after it would report 0 and describe nothing.
    ' Lp_Table_Convert_Options_Form takes the same copy for the same reason.
    failNumber = Err.Number
    failText = Err.Description

    ' Then put the screen back and take the bar down BEFORE saying anything, in that
    ' order - the same order Sh_Report_Error uses, and for the same reason: a message
    ' drawn over a frozen screen is what the transcriber reports as a hang.
    Sh_Progress_Close
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    Sh_Report_Error "Lp_Fix_Common_File_Errors", failNumber, failText

    ' An error raised INSIDE a handler is not caught by that handler, so these two would
    ' reach the transcriber as Word's own Run-time error dialog - the one offering Debug,
    ' which is what Sh_Report_Error exists to keep off the screen. Collapsing a selection
    ' the failure may have destroyed is exactly the shape that raises. Found in review.
    On Error Resume Next
    Selection.Collapse
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


