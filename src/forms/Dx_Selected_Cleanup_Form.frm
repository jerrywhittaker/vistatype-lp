VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_Selected_Cleanup_Form 
   Caption         =   "Selected Cleanup (333)"
   ClientHeight    =   9780.001
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520.001
   OleObjectBlob   =   "Dx_Selected_Cleanup_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_Selected_Cleanup_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Dx_Selected_Cleanup_Form
'
' Version: 1.3  Date: 3/4/2024 - removed "End" from CmdCancel_Click() added "MS_Set_Word_Config_For_Braille"  to "UserForm_Initialize"
' version: 1.2  Date: 4/7/2023 - show form then notify that no text is selected
' Version: 1.1  Date:8/21/2018
'

Private Sub CmdCancel_Click()
    Unload Dx_Selected_Cleanup_Form
End Sub

Private Sub Okay_Click()
'
' Version: 2.8  Date: 10/20/2026 - the selection's MathML is no longer parked and put back: braille
'                               gets Word equations (OMML) now, not MathML text (Jerry).
' Version: 2.7  Date: 10/18/2026 - LEAVES THE MATHML ALONE (Jerry, issue #37): every MathML element
'                               in the selection - what every pass here works on - is parked
'                               before the first pass and put back after the last, and in
'                               CleanupFailed too.
' Version: 2.6  Date: 10/12/2026 - Okay with nothing selected says 317 and stops with Exit Sub, not
'                               End, which closed any box that stays open (issue #33).
' Version: 2.5  Date: 10/12/2026 - the two passes that ask a question - Remove Bullets (329) and
'                               Replace Manual Line Breaks (362) - now ask it HERE, before the
'                               progress bar opens (Jerry, 10/12/2026: "The progress bar should
'                               not appear until the selection is made"). The answers are handed
'                               to the passes, which then ask nothing. Cancel on 329 drops only
'                               bullet removal, as before; Cancel on 362 stops the whole cleanup,
'                               as before - but now nothing has run yet, and without End.
' Version: 2.4  Date: 10/11/2026 - words only, set in Word's designer: the note under the title
'                               loses its doubled "the the" and "use use", and its hover text
'                               now says the same thing; the automatic-lists hover text was the
'                               same sentence twice and is one copy now; Remove all images' hover
'                               text matches its caption.
' Version: 2.3  Date: 10/11/2026 - Jerry's layout edits, second round: the note under the title
'                               rewritten (select the range to clean; select the whole document to
'                               run a cleanup on all of it), 42 -> 54 high; the automatic-lists box
'                               30.6 -> 45 high and "with a space"; the tabs box 15.6 -> 21.75 high;
'                               everything below moved down to suit; form 9150 -> 9780 high.
' Version: 2.2  Date: 10/11/2026 - layout edited by Jerry: "Remove all images" now reads
'                               "Remove all images (pictures) from select text"; the run of
'                               spaces in the automatic-lists caption closed up to two; the
'                               non-breaking spaces box 20 -> 21.75 high; form 9156 -> 9150 high.
' Version: 2.1  Date: 10/11/2026 - EVERY PASS GETS THE SELECTION (issue #31). Most passes end by
'                               collapsing the selection, so every pass after the first ticked one
'                               found nothing selected and ran on the whole book. The selection is
'                               held as a range and selected again before each pass, and at the end.
'                               If the passes empty it, the rest are skipped rather than run on the book.
' Version: 2.0  Date: 9/6/2026 - reads every choice BEFORE unloading the form, runs only the
'                               passes that were ticked, and reports them on the progress bar
' Version: 1.0  Date: earlier - thirteen If blocks, each unloading the form and running one macro
'
' THE CHOICES ARE READ FIRST, AND THAT IS THE POINT OF THIS VERSION.
'
' Every one of the thirteen blocks used to read its own check box and then Unload the form. The
' blocks BELOW the first ticked one therefore read a check box on a form that had already been
' unloaded - and touching any member of a UserForm's default instance CREATES the form again,
' with its design-time values. So a transcriber who ticked three boxes could get the first one
' only, silently.
'
' THIS HAS NOT BEEN SEEN IN WORD AND IS NOT RECORDED AS A REPORTED FAULT. A UserForm cannot be
' exercised headlessly, so the only machine that can settle it is Jerry's; the same question was
' asked about Change Picture Color on 9/4/2026 and the answer there was that it had always
' worked. Reading the choices before unloading is the right order either way, which is why it is
' done here without waiting for an answer.
'
' THE BAR COUNTS ONLY WHAT WAS TICKED, so it means something: tick two boxes and it moves in
' halves, tick nine and it moves in ninths. Counting all thirteen would leave it stopping at a
' third for a job that had finished.

    Dim wanted(1 To 13) As Boolean
    Dim macroName(1 To 13) As String
    Dim alsoRun(1 To 13) As String
    Dim saying(1 To 13) As String
    Dim total As Long, done As Long, i As Long
    Dim keep As Range
    Dim bulletStyle As String, bulletHyper As String, breakAnswer As String

    If Selection.Type <> wdSelectionNormal Then
        Sh_Say "Text must be selected first!", "Braille Macros (317)"
        Unload Dx_Selected_Cleanup_Form
        Exit Sub
    End If

    ' Read the form. Nothing below this point may touch a control.
    wanted(1) = Replace_Multi_Para_Marks_With_Single
    wanted(2) = Remove_Spaces_Before_And_After_Para_Marks
    wanted(3) = Clean_Auto_List_And_Tabs
    wanted(4) = Small_Caps_To_All_Caps
    wanted(5) = Replace_Tabs_With_Single_Space
    wanted(6) = Replace_Mult_Spaces_With_Single
    wanted(7) = Kill_Hyperlinks
    wanted(8) = Remove_Pictures
    wanted(9) = Remove_Bullets
    wanted(10) = Text_Boxes_And_Frames
    wanted(11) = ReplaceHyper
    wanted(12) = Manual_Line_Break
    wanted(13) = NonBreakingSpace

    macroName(1) = "Sh_Replace_Multiple_Para_Marks_No_Warning"
    macroName(2) = "Dx_Fix_Para_Space_Errors"
    macroName(3) = "Dx_Convert_Auto_List_To_Text"
    macroName(4) = "Sh_Replace_Small_Caps_With_All_Caps"
    macroName(5) = "Dx_Replace_Tabs_With_Single_Space"
    macroName(6) = "Sh_Remove_Multi_Spaces"
    macroName(7) = "Sh_Kill_The_Hyperlinks"
    macroName(8) = "Dx_Delete_Images"
    macroName(9) = "Dx_Remove_Bullets"
    macroName(10) = "Sh_Remove_Txt_Bxs_And_Frames"
    macroName(11) = "Dx_Convert_Hyper_To_Addresses"
    macroName(12) = "Sh_Replace_Manual_Line_Break"
    macroName(13) = "Sh_ReplaceNonBreakingSpacesWithNormalSpace"

    ' Two of the thirteen tidy up after themselves with a second pass, exactly as before.
    alsoRun(5) = "Sh_Remove_Multi_Spaces"
    alsoRun(13) = "Sh_Remove_Multi_Spaces"

    saying(1) = "Removing consecutive empty paragraph marks"
    saying(2) = "Fixing spaces around paragraph marks"
    saying(3) = "Turning automatic lists into text"
    saying(4) = "Replacing small caps with all caps"
    saying(5) = "Replacing tabs with a single space"
    saying(6) = "Removing multiple spaces"
    saying(7) = "Removing hyperlinks"
    saying(8) = "Deleting images"
    saying(9) = "Removing bullets"
    saying(10) = "Removing text boxes and frames"
    saying(11) = "Converting hyperlinks to addresses"
    saying(12) = "Replacing manual line breaks"
    saying(13) = "Replacing non-breaking spaces"

    For i = 1 To 13
        If wanted(i) Then total = total + 1
    Next i

    Unload Dx_Selected_Cleanup_Form

    ' The questions come first, so the bar never sits behind a box waiting for an answer.
    If wanted(9) Then
        Dx_Bullet_Removal_Form.Show
        If Dx_GP_String_2 = "" Then
            wanted(9) = False                     'Cancel: no bullet removal; the rest still run
            total = total - 1
        Else
            bulletStyle = Dx_GP_String_2
            bulletHyper = Dx_GP_String_1
        End If
    End If
    If wanted(12) Then
        Sh_Space_Or_Para_Form.Show
        If Sh_GP_String_1 <> "Para" And Sh_GP_String_1 <> "Space" Then
            Application.ScreenUpdating = True     'Cancel stops the whole cleanup; nothing has run
            Exit Sub
        End If
        breakAnswer = Sh_GP_String_1
    End If

    If total = 0 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Held as a range so it follows the text as the passes delete; each pass collapses the
    ' selection when it finishes, so it is selected again before the next one (issue #31).
    Set keep = Selection.Range.Duplicate

    On Error GoTo CleanupFailed
    Sh_Progress_Open "Cleaning up the selection"

    For i = 1 To 13
        If wanted(i) Then
            ' Announced BEFORE the pass runs, so the bar never reads 100% while work is still
            ' going on - the same reasoning as Dx_Ffc_Step and Lp_Ffc_Step.
            Sh_Progress_Say 100# * done / total, saying(i)
            ' If the passes so far have emptied the held text (or left only a picture), selecting
            ' it gives no text selection, and the next pass would take that to mean the whole
            ' book. There is nothing left to clean, so stop.
            keep.Select
            If Selection.Type <> wdSelectionNormal Then Exit For
            Select Case i
                Case 9: Application.Run "Dx_Remove_Bullets", bulletStyle, bulletHyper
                Case 12: Application.Run "Sh_Replace_Manual_Line_Break", breakAnswer
                Case Else: Application.Run MacroName:=macroName(i)
            End Select
            If Len(alsoRun(i)) > 0 Then
                keep.Select
                If Selection.Type <> wdSelectionNormal Then Exit For
                Application.Run MacroName:=alsoRun(i)
            End If
            done = done + 1
        End If
    Next i

    Sh_Progress_Say 100, "Finished"
    Sh_Progress_Close
    On Error GoTo 0
    keep.Select

    Application.ScreenUpdating = True ' Turn screen updating on
    Exit Sub

CleanupFailed:
    ' The copy of the error comes first: Sh_Progress_Close runs On Error Resume Next and
    ' Err.Clear inside itself, which wipes Err.
    Dim failNumber As Long
    Dim failText As String
    failNumber = Err.Number
    failText = Err.Description

    Sh_Progress_Close
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    Sh_Report_Error "Dx_Selected_Cleanup_Form", failNumber, failText
End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Dx_Selected_Cleanup_Form
End Sub

Private Sub UserForm_Initialize()
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    Me.ScrollBars = fmScrollBarsNone

    

End Sub

