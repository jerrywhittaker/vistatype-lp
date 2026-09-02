VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Section_Brk_Caution 
   Caption         =   "Caution"
   ClientHeight    =   3036
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4815
   OleObjectBlob   =   "Lp_Section_Brk_Caution.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Section_Brk_Caution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Version: 1.5  Date: 7/26/2026 - returns the user to where the cursor was when Okay was clicked
' Version: 1.4  Date: 7/24/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
' Version 1.3  Date: 1/6/2019
' Version 1.2  Date: 11/15/2018
'
Private Sub CmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub CmdOkay_Click()
'
' Version: 3.0  Date: 9/1/2026 - it changes a section START rather than deleting the section
'                               break. The reasoning, and what the old way cost Jerry's own test
'                               book, is written out above Lp_Replace_Section_Break_With_Page_Break
'                               in LPandBrlMacros. In short: a section break is the thing that
'                               CARRIES a section's page setup, so replacing it with a manual
'                               page break threw the setup of every section away and left the
'                               book unable to repaginate. Whether a section begins on the next
'                               page or on the next ODD page is a property of that section, and
'                               setting the property changes nothing else.
'
'                               Everything the old round trip needed went with the find:
'                               Lp_Copy_To_Temp_Doc and Lp_Copy_From_Temp_Doc, the two
'                               delete-one-character calls that tidied up after the paste, and
'                               the Selection.EndKey / Delete pair at the end - which ran on
'                               EVERY route out, not only the one that had pasted something, so
'                               on the nothing-selected route it deleted the last character of
'                               the transcriber's document. The same fault Lp_Fix_Para_Space_-
'                               Errors shed on 8/12/2026.
'
'                               ActiveDocument.UndoClear is gone too. It threw away the WHOLE
'                               undo history rather than this macro's part of it, on the one
'                               macro here that most needs undoing - a section start is a page
'                               setup decision, and the transcriber must be able to take it back.
'                               One custom undo record does that in a single Ctrl+Z.
'
'                               MS_Clear_F_and_R_Params_and_Clipboard went with the find as well:
'                               there is no find to reset now, and it emptied the transcriber's
'                               clipboard on purpose every time.
' Version: 3.2  Date: 9/1/2026 - THE WHOLE BOOK, ALWAYS. Jerry: "there is no reason to have
'                               anything selected... it would be unwise not to process the whole
'                               book at one time when the purpose is to convert it to a tablet
'                               screen. Processing a portion of the book makes no sense."
'
'                               Version 3.0 honoured a selection, because that is what the
'                               Selection Cleanup menu it sits on does. Wrong here: half a book
'                               converted for a tablet and half still laid out for two-sided
'                               printing is not a state anybody wants, and it would be arrived at
'                               silently. The selection is not read at all now.
'
'                               Lp_Is_Text_Selected went with it, in Lp_Selected_Cleanup_Form -
'                               that option no longer needs anything selected, so demanding it
'                               would be asking for something that is then ignored.
' Version: 3.1  Date: 9/1/2026 - collapses the selection on the way out (Jerry). Nothing in here
'                               moves the cursor any more, so 3.0 left the whole document
'                               highlighted after the usual Ctrl+A route.
' Version: 1.5  Date: 7/26/2026 - returns the user to where the cursor was when Okay was clicked
'
    Dim doc As Document
    Dim sec As Section
    Dim k As Long
    Dim changed As Long
    Dim su_Prev As Boolean
    Dim objUndo As UndoRecord
    Dim recording As Boolean
    ' errNum, not eNum: VBA identifiers are case-insensitive, so a variable called eNum IS the
    ' reserved word Enum as far as the compiler is concerned, and the Dim will not compile.
    Dim errNum As Long
    Dim errText As String

    Set doc = ActiveDocument

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    On Error GoTo eom

    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord "Section Starts to New Page"
    recording = True

    Sh_Last_Activity = "Section starts: odd or even page to new page"
    ' EVERY section, and nothing is asked about the selection - see the version note above.
    For k = 1 To doc.Sections.count
        Set sec = doc.Sections(k)
        With sec.PageSetup
            If .SectionStart = wdSectionOddPage Or .SectionStart = wdSectionEvenPage Then
                .SectionStart = wdSectionNewPage
                changed = changed + 1
            End If
        End With
    Next k

    objUndo.EndCustomRecord
    recording = False
    Sh_Last_Activity = ""

    ' Handed back to the macro that showed this dialog, which reports it once the dialog is gone.
    Lp_Sec_Starts_Changed = changed

    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

    ' Let the selection go. Nothing here reads it, but the transcriber may well have pressed
    ' Ctrl+A out of habit on the way in - and leaving a whole book highlighted puts her one
    ' keystroke away from replacing it. Jerry, 9/1/2026.
    '
    ' Collapse to the START, which is what every other macro here does at the end of a job.
    Selection.Collapse Direction:=wdCollapseStart

    Unload Me
    Exit Sub

eom:
    ' Close the undo record and put the screen back BEFORE anything else. An undo record left
    ' open swallows everything the transcriber does afterwards into one Ctrl+Z.
    '
    ' The error is NOT reported from in here. Sh_Report_Error shows dialog 240, which is another
    ' UserForm, and this is a UserForm's own button handler - a modal dialog on top of a modal
    ' dialog, with this one already unloaded. Lp_Sec_Starts_Changed is left at -1 so the macro
    ' that showed this dialog says nothing and simply stops.
    errNum = Err.Number
    errText = Err.Description
    On Error Resume Next
    If recording Then objUndo.EndCustomRecord
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh
    Sh_Last_Activity = "Section starts: failed with error " & errNum & " - " & errText
    Unload Me
End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Me
    Application.ScreenUpdating = True ' Turn screen updating on
    End
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
