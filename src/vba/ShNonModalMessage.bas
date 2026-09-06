Attribute VB_Name = "ShNonModalMessage"
Option Explicit
Public Sh_BridgeTargetMacro As String

' --- $pg validation helper state ------------------------------------------------------
' VBA requires every module-level declaration to sit ABOVE the first procedure. These
' lived at the foot of the module with the rest of the helper and would not compile:
' "Compile error in hidden module: ShNonModalMessage" (7/27/2026).
Private Sh_PgVal_TempDoc As Document        ' the tag list
Private Sh_PgVal_SourceDoc As Document      ' the document being validated
Private Sh_PgVal_LastParaStart As Long      ' character start of the tag last worked from
Private Sh_PgVal_TitleText As String        ' "VistaType LP" or "Braille Macros"

' Which menu is on screen: 0 none, 2 Sh_Valid_Ref_Pg_No_2_Form (the tag list), 4
' Sh_Valid_Ref_Pg_No_4_Form (the document being validated). Tracked rather than asked, because
' READING .Visible ON AN UNLOADED FORM LOADS IT - touching any member of a UserForm's default
' instance creates it and runs its UserForm_Initialize there and then. Asking "is the menu up?"
' would therefore put one up. 8/23/2026.
Private Sh_PgVal_MenuIsOn As Long

' The two menus' window titles, written out here rather than read off the forms. Reading
' Sh_Valid_Ref_Pg_No_4_Form.Caption would CREATE the form if it were not loaded - the same trap as
' .Visible above - and the one moment that matters is exactly the moment it might not be loaded.
' These must match the Caption in each form's designer header, and they are also what the
' transcriber sees in the title bar. 8/23/2026.
Private Const SH_PGVAL_TITLE_LIST As String = "Validate $pg Tags"
Private Const SH_PGVAL_TITLE_DOC As String = "Delete/change/add $pg"

' The macro is named in FULL - project, module, procedure - the same shape src/keymap uses, and
' for the same reason: a bare name is resolved when the key is PRESSED, against whatever projects
' are loaded then, and if it does not resolve the transcriber gets Word's "the macro cannot be
' found or has been disabled" instead of a menu. Note the module is ShNonModalMessage, not
' LPandBrlMacros - the natural mistake, since every other key assignment in this project points at
' LPandBrlMacros.
Private Const SH_PGVAL_KEY_MACRO As String = "LPandBRL.ShNonModalMessage.Sh_PgVal_ToggleFocus"

' The document the please-wait box was opened over, so focus can be handed back to it.
Private Sh_PleaseWait_Doc As Document

' The progress bar's two pieces of state - see Sh_Progress_Open near the bottom of this module.
' Sh_Progress_Up says whether a bar is on screen, so Sh_Progress_Say never has to ask the FORM
' (reading a property of an unloaded UserForm instantiates it and runs its Initialize, and this
' form's Initialize reads Application.Left, which hangs a Word with no desktop). Sh_Progress_Doc
' is the document the bar was opened over, so focus can be handed back to it - the same reason
' Sh_PleaseWait_Doc exists above.
'
' Module variables, and the End statement this project runs on ordinary paths wipes them - which
' is correct here rather than a trap, because End unloads every UserForm too. Flag and form go
' together either way.
Private Sh_Progress_Up As Boolean

' The slice of the bar the current sequence owns, 0 to 100. Added 9/6/2026, and it is what
' lets one counted sequence run INSIDE another without the two fighting over the bar.
'
' Attaching the template runs Lp_Fix_Common_File_Errors (32 counted passes) and
' Lp_Normalize_Styles (13) inside itself. Both report 0 to 100 of their own work, and
' without a span the bar would race to full twice and drop back twice - worse than no bar.
' With one, the attach says "you own 3 to 35" and the inner sequence's own 0-100 is mapped
' into it, so the bar only ever goes forwards.
'
' Defaults to the whole bar, and Sh_Progress_Open resets it - so a macro that knows nothing
' about spans behaves exactly as before.
Private Sh_Progress_Lo As Single
Private Sh_Progress_Hi As Single
Private Sh_Progress_Doc As Document

' Handing focus back to the document needs the Windows API. Activating the document window
' through the object model is NOT enough: a modeless UserForm keeps the keyboard focus, so
' Word draws no caret and the arrow keys walk the form's buttons instead of the text
' (Jerry, 7/27/2026). Window.Hwnd is available here - verified on Word 16.0 build 20131.
' Sh_FindWindowApi is the other direction, added 8/23/2026 for the F6 loop: it turns a modeless
' UserForm's title into its window handle so the keyboard can be handed TO it. A modeless VBA
' UserForm's window class is "ThunderDFrame" (a modal one is ThunderXFrame), and the title is the
' form's Caption - which is one reason Sh_Valid_Ref_Pg_No_4_Form could not stay called "UserForm1".
#If VBA7 Then
    Private Declare PtrSafe Function Sh_SetFocusApi Lib "user32" Alias "SetFocus" _
        (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function Sh_SetForegroundWindowApi Lib "user32" Alias "SetForegroundWindow" _
        (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function Sh_FindWindowApi Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Private Declare Function Sh_SetFocusApi Lib "user32" Alias "SetFocus" _
        (ByVal hwnd As Long) As Long
    Private Declare Function Sh_SetForegroundWindowApi Lib "user32" Alias "SetForegroundWindow" _
        (ByVal hwnd As Long) As Long
    Private Declare Function Sh_FindWindowApi Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

' Show/hide a command bar by name, tolerating bars that don't exist in this Word version
' (e.g. the legacy "Styles"/"Navigation" bars raise error 5 in Word 2016+). Non-critical UI.
Public Sub Sh_SetBarVisible(ByVal barName As String, ByVal vis As Boolean)
    On Error Resume Next
    Application.CommandBars(barName).Visible = vis
End Sub

Public Sub Sh_SpinTick()
    On Error Resume Next
    If Sh_NonModalMessageForm.Visible Then
        Sh_NonModalMessageForm.SpinTick
    End If
End Sub

' --- Sh_Please_Wait_Form: the same three moving parts, for the smaller "please wait" box ---
'
' Its own tick. Application.OnTime can only name a MODULE-level macro, never a method on a
' form, so each spinner form needs one of these. Sh_Please_Wait_Form was written from
' Sh_NonModalMessageForm and arrived pointing its OnTime at Sh_SpinTick above -- which ticks
' the OTHER form. Left that way its spinner advanced exactly one frame and stopped.
'
' Version: 1.0  Date: 8/3/2026

Public Sub Sh_PleaseWaitTick()
    On Error Resume Next
    If Sh_Please_Wait_Form.Visible Then
        Sh_Please_Wait_Form.SpinTick
    End If
End Sub

Public Sub Sh_Progress_Tick()
' The progress bar's own OnTime tick. One per spinner form, because OnTime can only name a
' module-level macro - see Sh_Convert_Progress_Form.SpinTick for what sharing one costs.
'
' Gated on Sh_Progress_Up, the FLAG, and never on the form's .Visible: reading a property of
' an unloaded UserForm instantiates it and runs its Initialize, and this form's Initialize
' reads Application.Left, which hangs a Word with no desktop.
'
' Version: 1.0  Date: 9/6/2026
    On Error Resume Next
    If Sh_Progress_Up Then Sh_Convert_Progress_Form.SpinTick
End Sub

Public Sub Sh_Spin_DoEvents()
' A DoEvents that also moves the spinner on one frame. Use it in place of a bare DoEvents
' inside a long macro.
'
' Why it is needed: DoEvents does NOT advance the spinner. The frame only moves when
' Application.OnTime fires, and OnTime cannot be scheduled closer than one SECOND apart, so
' during a short macro the spinner twitches two or three times and looks stuck no matter how
' many DoEvents the macro contains (Jerry, 8/3/2026). Driving it from the macro's own yields
' is what makes it turn.
'
' Calls Advance, not SpinTick: SpinTick queues another OnTime every time it runs, so using it
' here would leave one pending timer per DoEvents.
'
' Turns WHICHEVER progress box is showing. Lp_Fix_Common_File_Errors runs from two places and
' each brings its own: File Cleanup on the ribbon shows Sh_Please_Wait_Form, while the attach
' sequence already has Sh_NonModalMessageForm up and merely changes its message line. One
' helper covers both, so the macro does not need to know which way it was called.
'
' Does nothing but yield when neither is showing, so a macro carrying these calls still runs
' normally on its own.
'
' Version: 1.1  Date: 8/3/2026 - also advances Sh_NonModalMessageForm, for the attach sequence
' Version: 1.0  Date: 8/3/2026
    On Error Resume Next
    ' Sh_Please_Wait_Form is NOT asked any more, from 9/6/2026. Its last two callers moved
    ' onto the bar, so nothing shows it - and merely reading .Visible on an unloaded form
    ' LOADS it and runs its Initialize, which is the very trap noted in Sh_StartSpinnerBridge
    ' one screen above. That was happening on every one of the 43 yields in a File Cleanup
    ' run. The form and Sh_Show_Please_Wait / Sh_Hide_Please_Wait / Sh_PleaseWaitTick are
    ' left in place, unused, until the last spinner box goes - taking them out is a deletion
    ' to make on purpose, not a side effect of this change.
    ' Sh_NonModalMessageForm is not asked either, from 9/6/2026 - the two Export Selection
    ' macros were its last users and they are on the bar now. Same reason as the please-wait
    ' form above: reading .Visible on an unloaded form LOADS it and runs its Initialize.
    ' NEITHER FORM IS DELETED YET. Sh_Report_Error still takes both down, which costs
    ' nothing and is the right kind of net; and deleting a UserForm is a deliberate job -
    ' the code that names it has to go FIRST, or it reaches the transcriber as run-time
    ' error 424 when a dialog opens.
    ' The progress bar, from 9/6/2026 - and asked a different way on purpose. The two above
    ' are asked whether they are VISIBLE; this one is asked the FLAG, because reading any
    ' property of Sh_Convert_Progress_Form when it is not loaded runs its Initialize, which
    ' reads Application.Left and hangs a Word with no desktop. Same rule as Sh_Progress_Say.
    If Sh_Progress_Up Then Sh_Convert_Progress_Form.Advance
    On Error GoTo 0
    DoEvents
End Sub

Public Sub Sh_Show_Please_Wait(ByVal sMessage As String)
' Opens the please-wait box modeless and starts the spinner turning. The caller carries on
' immediately; the long macro's own DoEvents calls are what let the OnTime tick fire.
'
' Safe to call with ScreenUpdating already off, which is the normal case -- the form holds and
' restores it around its own Repaint so the document underneath does not flash.
'
' Remembers the document it was opened over and hands the focus straight back to it. Showing
' a modeless form takes the focus, and Word could be left with a DIFFERENT document active --
' Jerry saw Full File Cleanup jump to the blank startup document part way through the run and
' again at the end (8/3/2026). The macro then works on the wrong file.
'
' Version: 1.1  Date: 8/3/2026 - remembers the document and gives the focus back to it
' Version: 1.0  Date: 8/3/2026
    On Error Resume Next
    Set Sh_PleaseWait_Doc = ActiveDocument
    On Error GoTo 0

    With Sh_Please_Wait_Form
        .ActivityMsg.Caption = sMessage
        .SpinnerBox.Caption = ""
        .Show vbModeless
        .StartSpinner
    End With

    Sh_Focus_Document Sh_PleaseWait_Doc
    DoEvents
End Sub

Public Sub Sh_Hide_Please_Wait()
' Stops the spinner and closes the box. Stopping first matters: it clears the flag SpinTick
' tests, so a tick already queued by OnTime does nothing instead of reopening the form.
'
' Then hands the focus back to the document the box was opened over, the same way
' Lp_Attach_The_Template re-asserts its document after unloading its own progress form.
' Without it Word is left on whatever window it fancies -- in Jerry's case the blank startup
' document (8/3/2026).
'
' Version: 1.1  Date: 8/3/2026 - gives the focus back to the document afterwards
' Version: 1.0  Date: 8/3/2026
    On Error Resume Next
    Sh_Please_Wait_Form.StopSpinner
    Unload Sh_Please_Wait_Form
    On Error GoTo 0
    DoEvents

    Sh_Focus_Document Sh_PleaseWait_Doc
    Set Sh_PleaseWait_Doc = Nothing
    DoEvents
End Sub

Public Sub Sh_ShowNonModalMessage(sCaption As String, sMessage As String)
    With Sh_NonModalMessageForm
        .StartUpPosition = 0
        .Caption = sCaption
        .LblMessage.Caption = sMessage
        .ActivityMsg.Caption = ""
        .SpinnerBox.Caption = ""
        .Show vbModeless
    End With
End Sub

Public Sub Sh_StartSpinnerBridge()
    ' It started Sh_NonModalMessageForm's spinner until 9/6/2026. Both macros reached
    ' through this bridge - the LP attach and the DAISY converter - now show the progress
    ' bar, which starts its own spinner in Sh_Progress_Open. Naming the old form here would
    ' LOAD it, since touching any member of a UserForm's default instance creates it: an
    ' invisible form running its Initialize for nothing.
    '
    ' The bridge itself stays exactly as it was, and must. A modeless box cannot paint until
    ' the modal dialog's Okay handler has reached End Sub, so the work has to be handed to
    ' Application.OnTime and started from out here.
    'Sh_NonModalMessageForm.SetActivityMessage "Attaching the LP template…"

    If Len(Sh_BridgeTargetMacro) > 0 Then
        Application.Run Sh_BridgeTargetMacro 'Sh_BridgeTargetMacro is the name of the macro to be run
    End If
End Sub

'=========================================================================================
' $pg reference-page validation helper
'
' Sh_Copy_Ref_Pg_Tags_To_Temp_File builds a temporary document listing every tagged
' paragraph, so the transcriber can spot missing numbers, bad tags and gaps in the sequence.
' Fixing one used to mean: Alt+Tab to the real document, paste the tag into the Navigation
' pane to find it, correct it, Alt+Tab back, and then hunt for your place in the list again.
'
' Two modeless forms replace that clerical half. Sh_Valid_Ref_Pg_No_2_Form rides above the
' temp document and Sh_Valid_Ref_Pg_No_4_Form above the document being validated; each
' button swaps them, so the form on screen always matches the document you are looking at.
'
' The cycle: Locate -> fix it -> Return -> already sitting on the NEXT tag down. Not losing
' your place in the list is the whole point (Jerry, 7/27/2026).
'
' Shared Sh_ code: Validate $pg Tags on BOTH the LP and Braille tabs comes through here.
' The dialog title is captured when the list is built, because once the temp document is
' active Dx_Is_The_Attached_Template_BANA_Braille would be answering about the wrong file.
'=========================================================================================

' Called by Sh_Copy_Ref_Pg_Tags_To_Temp_File once the list document exists.
Public Sub Sh_PgVal_Start(ByVal SourceDoc As Document, ByVal TempDoc As Document, _
                          ByVal TitleText As String)
    Set Sh_PgVal_SourceDoc = SourceDoc
    Set Sh_PgVal_TempDoc = TempDoc
    Sh_PgVal_TitleText = TitleText
    Sh_PgVal_LastParaStart = -1
    Sh_PgVal_BindKeys
    Sh_PgVal_SwapToTempForm
    Sh_PgVal_CursorToTopOfList
End Sub

' --- form 2 (temp document): "Locate the selected $pg code in the Document" ---------------
Public Sub Sh_PgVal_LocateInDocument()
    Dim TagText As String
    Dim Hit As Range

    If Not Sh_PgVal_Ready() Then Exit Sub

    Sh_PgVal_TempDoc.Activate
    TagText = Sh_PgVal_TextOfCurrentTag()

    If Len(TagText) = 0 Then
        MsgBox "Click anywhere on the line holding the $pg tag you want to find, " _
             & "then press this button again.", vbInformation, Sh_PgVal_TitleText
        Exit Sub
    End If

    'Remember where we are in the list BEFORE leaving, so Return can move on from here.
    Sh_PgVal_LastParaStart = Selection.Paragraphs(1).Range.start

    Set Hit = Sh_PgVal_FindInSource(TagText)

    If Hit Is Nothing Then
        MsgBox "This tag is in the list but was not found in the document:" _
             & vbCrLf & vbCrLf & "    " & TagText & vbCrLf & vbCrLf _
             & "That is worth noting in itself - the tag may have been deleted or changed " _
             & "since this list was made.", vbExclamation, Sh_PgVal_TitleText
        Exit Sub
    End If

    Sh_PgVal_SourceDoc.Activate
    Hit.Select
    Selection.Collapse Direction:=wdCollapseStart
    Sh_Keep_Cursor_In_View
    Sh_PgVal_SwapToSourceForm
End Sub

' --- form 4 (source document): "Return to the validation document" ------------------------
Public Sub Sh_PgVal_ReturnToTempAndAdvance()
    Dim p As Paragraph
    Dim NextPara As Paragraph

    If Not Sh_PgVal_Ready() Then Exit Sub

    Sh_PgVal_TempDoc.Activate

    'The next NON-EMPTY line after the one we last worked from. The list is built by pasting
    'a Find-All selection, so blank paragraphs can turn up between entries.
    For Each p In Sh_PgVal_TempDoc.Paragraphs
        If p.Range.start > Sh_PgVal_LastParaStart Then
            If Len(Trim$(Replace(p.Range.Text, Chr(13), ""))) > 0 Then
                Set NextPara = p
                Exit For
            End If
        End If
    Next p

    If NextPara Is Nothing Then
        MsgBox "That was the last tag in the list." & vbCrLf & vbCrLf _
             & "Press 'Done - Exit validation' to close the list, or keep working through " _
             & "it by hand.", vbInformation, Sh_PgVal_TitleText
    Else
        NextPara.Range.Select
        Selection.Collapse Direction:=wdCollapseStart
        Sh_PgVal_LastParaStart = NextPara.Range.start
        Sh_Keep_Cursor_In_View
    End If

    Sh_PgVal_SwapToTempForm
End Sub

' --- either form: "Done - Exit validation" ------------------------------------------------
Public Sub Sh_PgVal_Done()
    On Error Resume Next

    'F6 belongs to Word again the moment the validation is over - see Sh_PgVal_BindKeys.
    Sh_PgVal_MenuIsOn = 0
    Sh_PgVal_UnbindKeys

    'Hide before unloading: these run from a button on one of the forms being closed.
    Sh_Valid_Ref_Pg_No_2_Form.Hide
    Sh_Valid_Ref_Pg_No_4_Form.Hide

    If Sh_PgVal_DocIsOpen(Sh_PgVal_TempDoc) Then
        Sh_PgVal_TempDoc.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Sh_PgVal_DocIsOpen(Sh_PgVal_SourceDoc) Then
        Sh_PgVal_SourceDoc.Activate

        'Clear the Find All multiple selection before handing the document back. The temp-file
        'route selects EVERY $pg paragraph at once to copy them (Find In > Main Document), and
        'that selection outlives the validation. Word refuses most Selection work across a
        'multiple selection, so the next macro the user runs died with run-time error 4605.
        Sh_Clear_Multi_Selection
    End If

    Unload Sh_Valid_Ref_Pg_No_1_Form
    Unload Sh_Valid_Ref_Pg_No_2_Form
    Unload Sh_Valid_Ref_Pg_No_3_Form
    Unload Sh_Valid_Ref_Pg_No_4_Form

    Set Sh_PgVal_TempDoc = Nothing
    Set Sh_PgVal_SourceDoc = Nothing
    Sh_PgVal_LastParaStart = -1
End Sub

' --- F6 and Shift+F6: moving between the document and the menu ----------------------------
'
' Jerry, 8/23/2026, and it is for transcribers who are blind or have low vision - several of
' the people using this are. A modeless UserForm beside a document is easy to reach with a
' mouse and impossible to reach from the keyboard: Word's own F6 walks Word's panes and knows
' nothing about a UserForm.
'
' SO THE LOOP HAS TWO HALVES, AND NEITHER CAN DO THE OTHER'S JOB:
'
'   document -> menu   a Word key binding. Word only ever sees the keystroke while WORD has
'                      the focus, which is exactly when this direction is wanted.
'   menu -> document   the forms' own KeyDown handlers, calling Sh_PgVal_KeyToDocument below.
'                      While the form has the focus Word never sees the key at all, so no key
'                      binding could ever fire. This is why the handlers are on the buttons.
'
' F6 and Shift+F6 do the SAME thing, and that is not laziness: the loop has two stops, so
' forwards and backwards are the same move. Both are bound because a screen-reader user reaches
' for either out of habit.
'
' THE BINDING IS ONLY IN FORCE WHILE A VALIDATION IS RUNNING, because F6 is Word's own "next
' pane" key and she is entitled to have it back. Three separate things see to that: Sh_PgVal_Done
' removes it; Sh_PgVal_ToggleFocus removes it itself if it ever fires with no menu on screen; and
' it is written into the ADD-IN's own template in memory only - never saved - so closing Word
' forgets it whatever happened in between.
'
' ThisDocument.Saved is put back to True after every change. Adding a key binding marks the
' template dirty, and a dirty global template makes Word ask "do you want to save changes to
' LPandBRL.dotm?" on the way out - a question no transcriber should ever be asked.
'
' Version: 1.0  Date: 8/23/2026
' Version: 1.1  Date: 8/23/2026 - the macro name is qualified, and CustomizationContext is put back
' Version: 1.0  Date: 8/23/2026
Private Sub Sh_PgVal_BindKeys()
    Dim ctxWas As Object

    On Error Resume Next
    ' CustomizationContext is APPLICATION state and outlives this macro, like ScreenUpdating. Left
    ' pointing at VistaType's add-in, the next thing that adds a key assignment or a toolbar
    ' customization without setting it - Word's own, another add-in, one of her own macros - would
    ' write it in here instead of into Normal.
    Set ctxWas = CustomizationContext

    CustomizationContext = ThisDocument
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF6), _
                    KeyCategory:=wdKeyCategoryMacro, Command:=SH_PGVAL_KEY_MACRO
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyShift, wdKeyF6), _
                    KeyCategory:=wdKeyCategoryMacro, Command:=SH_PGVAL_KEY_MACRO
    ThisDocument.Saved = True

    If Not ctxWas Is Nothing Then CustomizationContext = ctxWas
    Err.Clear
End Sub

' Version: 1.1  Date: 8/23/2026 - CustomizationContext is put back; matches on the qualified name
' Version: 1.0  Date: 8/23/2026
Private Sub Sh_PgVal_UnbindKeys()
    Dim i As Long
    Dim ctxWas As Object

    On Error Resume Next
    Set ctxWas = CustomizationContext
    CustomizationContext = ThisDocument

    ' Backwards. Clearing a binding takes it out of the collection and moves everything above it
    ' down, so a forward loop would step over the second one.
    '
    ' InStr rather than "=" so a binding written by an older build - which named the macro without
    ' its project and module - is still recognized and cleared. That build never shipped, but the
    ' rule holds generally: this has to be able to clean up after itself.
    For i = KeyBindings.count To 1 Step -1
        If InStr(1, KeyBindings(i).Command, "Sh_PgVal_ToggleFocus", vbTextCompare) > 0 Then
            KeyBindings(i).Clear
        End If
    Next i

    ThisDocument.Saved = True
    If Not ctxWas Is Nothing Then CustomizationContext = ctxWas
    Err.Clear
End Sub

' What F6 and Shift+F6 run while a validation is going on. Word has the focus - that is the only
' way this can have been reached - so the move is always "to the menu".
'
' Version: 1.0  Date: 8/23/2026
Public Sub Sh_PgVal_ToggleFocus()
    On Error Resume Next

    If Sh_PgVal_MenuIsOn = 0 Then
        ' No menu on screen, so this key is not ours. Hand F6 back to Word and do nothing else.
        ' This is the safety net: if a validation ever ends without Sh_PgVal_Done running, the
        ' very first F6 afterwards undoes the binding.
        Sh_PgVal_UnbindKeys
        Exit Sub
    End If

    Sh_PgVal_FocusMenu
End Sub

' The other half, called from every button's KeyDown on both menus. Puts the keyboard back in
' whichever document the menu on screen belongs to.
'
' Version: 1.0  Date: 8/23/2026
Public Sub Sh_PgVal_KeyToDocument()
    On Error Resume Next

    ' If the document this menu belongs to has been closed, Sh_PgVal_FocusDocument gives up
    ' quietly - and quietly is the wrong answer for the one transcriber who cannot reach for the
    ' mouse instead. Sh_PgVal_Ready says what happened and ends the validation tidily.
    If Sh_PgVal_MenuIsOn = 2 Then
        If Sh_PgVal_DocIsOpen(Sh_PgVal_TempDoc) Then
            Sh_PgVal_FocusDocument Sh_PgVal_TempDoc
        Else
            Sh_PgVal_Ready
        End If
    ElseIf Sh_PgVal_MenuIsOn = 4 Then
        If Sh_PgVal_DocIsOpen(Sh_PgVal_SourceDoc) Then
            Sh_PgVal_FocusDocument Sh_PgVal_SourceDoc
        Else
            Sh_PgVal_Ready
        End If
    End If
End Sub

' Hand the keyboard to the menu that is up.
'
' Two ways of doing it, because the first one is better and the second one always works. Windows
' can be told to raise the form's own window, which needs its handle - a modeless UserForm is a
' "ThunderDFrame" window titled with the form's Caption. Failing that, hiding and re-showing a
' modeless form takes the focus by itself, which is the very behaviour complained about in
' Sh_PgVal_FocusDocument above; here it is what we want. It flickers, which is why it is second.
'
' Then the keyboard is put on the first button, so that a screen reader has something to announce
' and Tab walks the menu from a known place.
' Version: 1.1  Date: 8/23/2026 - the window is found by a CONSTANT title, and there is no
'                                 hide-and-show fallback. Both changes are the same fix: nothing
'                                 here may touch a form, because touching one CREATES it
' Version: 1.0  Date: 8/23/2026
Private Sub Sh_PgVal_FocusMenu()
    #If VBA7 Then
        Dim h As LongPtr
    #Else
        Dim h As Long
    #End If

    On Error Resume Next

    If Sh_PgVal_MenuIsOn = 2 Then
        h = Sh_FindWindowApi("ThunderDFrame", SH_PGVAL_TITLE_LIST)
    ElseIf Sh_PgVal_MenuIsOn = 4 Then
        h = Sh_FindWindowApi("ThunderDFrame", SH_PGVAL_TITLE_DOC)
    End If

    ' No window means no menu, whatever Sh_PgVal_MenuIsOn believes. DO NOTHING - and in
    ' particular do not put the form back up. An earlier version hid and re-showed it here as a
    ' fallback, which would have re-created a $pg menu over whatever unrelated document the
    ' transcriber had moved on to.
    If h = 0 Then
        Sh_PgVal_MenuIsOn = 0
        Sh_PgVal_UnbindKeys
        Exit Sub
    End If

    Sh_SetForegroundWindowApi h
    Sh_SetFocusApi h

    ' And the keyboard onto the first button, so a screen reader has something to announce and Tab
    ' walks the menu from a known place. Safe to touch the form by now: its window exists.
    If Sh_PgVal_MenuIsOn = 2 Then
        Sh_Valid_Ref_Pg_No_2_Form.LocateInDocButton.SetFocus
    Else
        Sh_Valid_Ref_Pg_No_4_Form.LocateInDocButton.SetFocus
    End If

    Err.Clear
End Sub

' --- helpers ------------------------------------------------------------------------------

' Unload before Show so UserForm_Initialize re-runs: it positions the form over the ACTIVE
' document using ActiveWindow.GetPoint, which is only right if it runs after the swap.
' The form being left is hidden, never unloaded, because its own click handler is on the stack.
Private Sub Sh_PgVal_SwapToTempForm()
    On Error Resume Next
    Unload Sh_Valid_Ref_Pg_No_2_Form
    Sh_Valid_Ref_Pg_No_2_Form.Show vbModeless
    Sh_Valid_Ref_Pg_No_4_Form.Hide
    Sh_PgVal_MenuIsOn = 2
    Sh_PgVal_FocusDocument Sh_PgVal_TempDoc
End Sub

Private Sub Sh_PgVal_SwapToSourceForm()
    On Error Resume Next
    Unload Sh_Valid_Ref_Pg_No_4_Form
    Sh_Valid_Ref_Pg_No_4_Form.Show vbModeless
    Sh_Valid_Ref_Pg_No_2_Form.Hide
    Sh_PgVal_MenuIsOn = 4
    Sh_PgVal_FocusDocument Sh_PgVal_SourceDoc
End Sub

' Showing a modeless UserForm takes the keyboard focus, so Word stops drawing a caret
' in the document behind it - the cursor looks like it has vanished. The message box
' this feature replaced used to hand focus back when it was dismissed; nothing does now,
' so every swap has to do it deliberately (Jerry, 7/27/2026).
Private Sub Sh_PgVal_FocusDocument(ByVal d As Document)
    On Error Resume Next
    If Not Sh_PgVal_DocIsOpen(d) Then Exit Sub
    Sh_Focus_Document d
End Sub

' The same job, for any caller. A modeless form takes the keyboard focus, and Word can end up
' showing a DIFFERENT document than the one being worked on - Jerry saw Full File Cleanup jump
' to the blank startup document part way through and again at the end (8/3/2026).
'
' Both halves are needed. Activating through the object model moves Word to the right document
' but leaves the keyboard focus on the form; only Windows can take that back.
'
' Version: 1.0  Date: 8/3/2026
Public Sub Sh_Focus_Document(ByVal d As Document)
    On Error Resume Next
    If d Is Nothing Then Exit Sub

    d.Activate
    d.ActiveWindow.Activate

    Sh_SetForegroundWindowApi d.ActiveWindow.hwnd
    Sh_SetFocusApi d.ActiveWindow.hwnd
End Sub

' Open the list with the cursor on its first real line, ready for Locate. Straight to
' the top, then past any leading blank paragraph the paste may have left behind.
Private Sub Sh_PgVal_CursorToTopOfList()
    Dim p As Paragraph
    On Error Resume Next
    If Not Sh_PgVal_DocIsOpen(Sh_PgVal_TempDoc) Then Exit Sub

    Sh_PgVal_FocusDocument Sh_PgVal_TempDoc
    Selection.HomeKey Unit:=wdStory

    If Len(Trim$(Replace(Selection.Paragraphs(1).Range.Text, Chr(13), ""))) = 0 Then
        For Each p In Sh_PgVal_TempDoc.Paragraphs
            If Len(Trim$(Replace(p.Range.Text, Chr(13), ""))) > 0 Then
                p.Range.Select
                Selection.Collapse Direction:=wdCollapseStart
                Exit For
            End If
        Next p
    End If

    Sh_Keep_Cursor_In_View
End Sub

' Text of the paragraph the cursor sits in, stripped of the paragraph mark. Deliberately the
' whole paragraph rather than the selection, so clicking anywhere on the line is enough.
Private Function Sh_PgVal_TextOfCurrentTag() As String
    Dim s As String
    Dim ch As Range
    Dim para As Range

    On Error Resume Next

    ' THE LEADING VISIBLE RUN, and nothing from the first hidden character onward.
    '
    ' Jerry, 8/26/2026: on the braille side an entry is a page number followed by a Duxbury code,
    ' and THOSE CODES ARE HIDDEN TEXT. Range.Text hands them over regardless of whether they are
    ' on screen, so the string being searched for held characters the reader cannot see - and
    ' neither this Locate nor Word's own Navigation pane could find them in the book.
    '
    ' Searching only the visible run sidesteps the question of whether Word's Find matches hidden
    ' text at all, which depends on whether the source document's window happens to be showing
    ' it. Not worth depending on something that changes with a display setting.
    '
    ' STOPPING at the first hidden character, rather than collecting every visible one and
    ' skipping the hidden, is deliberate: what is left has to be CONTIGUOUS in the document or
    ' Find cannot match it. A line stitched together across a hidden gap is not.
    '
    ' THE COMMON SHAPE, which this handles well - a continuation page, built by the passes in
    ' Dx_AutoTag_Page_Numbers:
    '
    '        $pg12-14[[*lec*]][[*i*]]14
    '        \_visible_/\___hidden___/\/  visible
    '
    ' The leading visible run is "$pg12-14" - the tag and the page number, which is exactly what
    ' identifies the entry, and contiguous in the book.
    '
    ' THE WEAK ONE, and it is worth knowing about: the EBAE lower-roman case inserts
    ' "$pg[[*ii*]]" ahead of the number, so the leading visible run is the bare tag "$pg" and
    ' Locate will land on the FIRST tag in the document rather than this one. That is no worse
    ' than it was - before this, the search string carried hidden characters and found nothing at
    ' all - but it is not right either. Curing it means either searching across the hidden gap
    ' (which Find cannot do) or remembering where each entry came from when the list is built.
    Set para = Selection.Paragraphs(1).Range
    For Each ch In para.Characters
        If ch.Font.Hidden Then Exit For
        s = s & ch.Text
    Next ch

    ' A line that BEGINS hidden would leave nothing to search for. Fall back to the whole text,
    ' which is what this did before - no worse than it was, and it keeps Locate answering.
    If Len(Trim$(Replace(Replace(s, Chr(13), ""), Chr(7), ""))) = 0 Then
        s = para.Text
    End If

    s = Replace(s, Chr(13), "")
    s = Replace(s, Chr(7), "")
    s = Trim$(s)

    ' AND THE ELLIPSIS OFF THE END. From 8/23/2026 a line longer than 23 characters is shown cut
    ' short with "..." on it - see Sh_PgVal_Copy_Tags_Into. That is display only: the book does not
    ' contain the ellipsis, so searching for it would find nothing and every long line would come
    ' back "in the list but not found in the document". What is left is a PREFIX of the real
    ' paragraph, which is all Find needs.
    If Right$(s, Len(SH_PGVAL_ELLIPSIS)) = SH_PGVAL_ELLIPSIS Then
        s = RTrim$(Left$(s, Len(s) - Len(SH_PGVAL_ELLIPSIS)))
    End If

    Sh_PgVal_TextOfCurrentTag = s
End Function

Private Function Sh_PgVal_FindInSource(ByVal TagText As String) As Range
    Dim r As Range
    On Error Resume Next
    Set r = Sh_PgVal_SourceDoc.Content
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = Left$(TagText, 250)        'Word's Find gives up beyond 255 characters
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        If .Execute Then Set Sh_PgVal_FindInSource = r
    End With
End Function

' Both documents still open? Touching a closed Document object raises an error rather than
' returning Nothing, so the only reliable test is to read a property and see what happens.
Private Function Sh_PgVal_DocIsOpen(ByVal d As Document) As Boolean
    Dim s As String
    If d Is Nothing Then Exit Function
    On Error Resume Next
    Err.Clear
    s = d.Name
    Sh_PgVal_DocIsOpen = (Err.Number = 0) And (Len(s) > 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function Sh_PgVal_Ready() As Boolean
    If Not Sh_PgVal_DocIsOpen(Sh_PgVal_TempDoc) Then
        MsgBox "The validation list has been closed, so there is nothing to work from." _
             & vbCrLf & vbCrLf & "Run Validate $pg Tags again to start another pass.", _
               vbInformation, Sh_PgVal_TitleText
        Sh_PgVal_Done
        Exit Function
    End If
    If Not Sh_PgVal_DocIsOpen(Sh_PgVal_SourceDoc) Then
        MsgBox "The document being validated has been closed." & vbCrLf & vbCrLf _
             & "Run Validate $pg Tags again to start another pass.", _
               vbInformation, Sh_PgVal_TitleText
        Sh_PgVal_Done
        Exit Function
    End If
    Sh_PgVal_Ready = True
End Function
'*** end of $pg reference-page validation helper ***

' --- Sh_Convert_Progress_Form: the BAR, for a macro that knows how far through it is ----------
'
' Added 8/29/2026 at Jerry's request: braille Full File Cleanup and Remove Empty Paragraphs show
' a bar on a large file, so nobody thinks Word has frozen. Same reason the DAISY converter was
' given one on 8/3/2026 - his words then were that without a visible indicator "many will think
' the computer is frozen and well... reboot time".
'
' NOT the same thing as the please-wait box above. That one turns a spinner and says nothing
' about how far along the work is; this says how far along it is. Use the spinner where the
' macro cannot know, and the bar where it can.
'
' Three wrappers rather than reaching for the form directly, for two reasons: a caller cannot
' leave the box on screen, and a progress box can never be the thing that stops a cleanup - every
' one of them traps. Sh_Progress_Say on a box that was never opened does nothing at all, which is
' what makes it safe to leave those calls in a macro that is sometimes run without a box.
'
' The form's designer caption reads "Converting - Please Wait", which is what the converter
' wanted. Sh_Progress_Open sets it at RUN TIME instead, so the .frx is not touched - a caption is
' one of the few things about a UserForm that can be changed without opening the designer in Word.
'
' Version: 1.0  Date: 8/29/2026
'
' Author: Jerry Whittaker -  jerry@vistatypelp.org

Public Sub Sh_Progress_Open(ByVal boxTitle As String)
' Puts the bar on screen, empty, and titles it. Modeless, so the macro carries straight on.
'
' HANDS FOCUS BACK TO THE DOCUMENT afterwards, and that is not tidiness. Showing a modeless form
' takes the focus, and Word can then make a DIFFERENT document active - Jerry saw exactly that on
' 8/3/2026 with the please-wait box: "Full File Cleanup jump to the blank startup document part
' way through the run", after which the macro works on the wrong file. Every pass these bars sit
' over addresses ActiveDocument, and several use Selection.Find, so the same mistake here would
' clean the empty Document1 and leave her book untouched. Sh_Show_Please_Wait and
' Sh_Hide_Please_Wait have carried this call since that day; these do too.

    On Error Resume Next
    Set Sh_Progress_Doc = ActiveDocument
    Sh_Convert_Progress_Form.Caption = boxTitle
    Sh_Convert_Progress_Form.SetProgress 0, ""
    Sh_Convert_Progress_Form.Show vbModeless
    Sh_Progress_Up = True
    Sh_Progress_Lo = 0
    Sh_Progress_Hi = 100
    ' Started AFTER the flag is raised: SpinTick queues Sh_Progress_Tick, which is gated on
    ' that flag and would stop the spinner dead on its first tick if it were still False.
    Sh_Convert_Progress_Form.StartSpinner
    Sh_Focus_Document Sh_Progress_Doc
    Err.Clear
    On Error GoTo 0

End Sub  '*** end of Sh_Progress_Open ***

Public Sub Sh_Progress_Say(ByVal pct As Single, ByVal what As String)
' Moves the bar and says what is happening. SetProgress clamps pct to 0-100 itself and holds
' ScreenUpdating across its own repaint, so a caller with the screen switched off - which is
' every caller here - does not get a flash of half-cleaned document.
'
' Gated on Sh_Progress_Up, a FLAG, and not on the form's own .Visible: reading any property of an
' unloaded UserForm instantiates it and runs UserForm_Initialize, which for this form reads
' Application.Left - and that is the line that hangs a Word driven over SSH with no desktop. A
' flag means a Say with no bar open costs nothing and touches nothing, which is what makes it
' safe to leave these calls in a macro that is sometimes run without a box.

    ' The "Where:" line in the error log, recorded BEFORE the flag test and deliberately.
    ' Lp_Normalize_Styles runs from Lp_Import_Exported_Selection_File with no bar open at
    ' all, and its thirteen messages used to go through Sh_NonModalMessageForm, which
    ' records this unconditionally. Testing the flag first would mean an import that failed
    ' logged a "Where:" line belonging to whatever ran before it - worse than none, because
    ' it points at the wrong macro.
    If Len(what) > 0 Then Sh_Last_Activity = what

    If Not Sh_Progress_Up Then Exit Sub

    On Error Resume Next

    ' That line is the ONLY position marker a failure report carries - VBA gives a handler a
    ' number and a description and nothing else. Sh_NonModalMessageForm.SetActivityMessage
    ' has written it since the log was built; writing it here too is what stops a macro
    ' losing its "Where:" line when it moves onto the bar, and gives one to every macro
    ' that was already on the bar - the braille cleanups and the DAISY converter.

    Sh_Convert_Progress_Form.SetProgress Sh_Progress_Lo _
                                       + (Sh_Progress_Hi - Sh_Progress_Lo) * (pct / 100#), what
    Err.Clear
    On Error GoTo 0

End Sub  '*** end of Sh_Progress_Say ***

Public Sub Sh_Progress_Span(ByVal lo As Single, ByVal hi As Single)
' Hand the next stretch of the bar to a sequence that counts its own work from 0 to 100.
'
'     Sh_Progress_Say 3, "Fixing common file errors"
'     Sh_Progress_Span 3, 35
'     Lp_Fix_Common_File_Errors           ' its own 0-100 now lands between 3 and 35
'     Sh_Progress_Span 0, 100             ' give the whole bar back
'
' GIVE THE BAR BACK. A span left in place makes every later Say land inside it, so the bar
' stops a third of the way along and stays there. Reset it the moment the inner sequence
' returns - including on the way out of an error handler.
'
' Refuses a backwards or empty span rather than accepting one: hi <= lo would freeze the bar
' at lo for the whole of the inner sequence, which reads exactly like a hang.
'
' Version: 1.0  Date: 9/6/2026

    If hi <= lo Then Exit Sub
    If lo < 0 Then lo = 0
    If hi > 100 Then hi = 100

    Sh_Progress_Lo = lo
    Sh_Progress_Hi = hi

End Sub  '*** end of Sh_Progress_Span ***

Public Sub Sh_Progress_Hide()
' Take the bar off the screen for a moment WITHOUT closing it, and put it back with
' Sh_Progress_Show. Needed where a modal dialog of Word's own has to come forward - the Save As
' dialog in Lp_Attach_The_Template - because a modeless form sitting over it confuses the focus.
'
' Hide, not Close, and the difference matters: Close lowers Sh_Progress_Up, and every Say after
' that does nothing. Closing here instead of hiding is what made the attach's last two stages
' and its "Finished" vanish before anyone saw them (found in review, 9/6/2026).
'
' Gated on the flag like every other helper here, so it is safe to call when no bar is open and
' can never be the thing that instantiates the form. Reading a property of an unloaded UserForm
' runs its Initialize, which for this form reads Application.Left and hangs a headless Word.
'
' Version: 1.0  Date: 9/6/2026

    If Not Sh_Progress_Up Then Exit Sub

    On Error Resume Next
    Sh_Convert_Progress_Form.Hide
    Err.Clear
    On Error GoTo 0

End Sub  '*** end of Sh_Progress_Hide ***

Public Sub Sh_Progress_Show()
' Put a hidden bar back on screen. The counterpart to Sh_Progress_Hide; see the note there.
'
' Version: 1.0  Date: 9/6/2026

    If Not Sh_Progress_Up Then Exit Sub

    On Error Resume Next
    Sh_Convert_Progress_Form.Show vbModeless
    Sh_Focus_Document Sh_Progress_Doc
    Err.Clear
    On Error GoTo 0

End Sub  '*** end of Sh_Progress_Show ***

Public Sub Sh_Progress_Close()
' Takes the bar off screen. Unload rather than Hide, so the next Open starts from a fresh form
' and cannot show the last run's message for a moment before the first step arrives.
'
' Safe to call when no bar is open, and safe to call twice - which is the point: every path out
' of a macro that opened one can end with this, including an error handler.

    If Not Sh_Progress_Up Then Exit Sub

    On Error Resume Next
    ' Stop the spinner BEFORE lowering the flag and unloading. A tick already queued with
    ' Application.OnTime fires up to a second after this runs, and a tick reaching an
    ' unloaded form would load it again - the box coming back after the macro finished.
    ' Sh_Hide_Please_Wait has done it in this order since 8/3/2026.
    Sh_Convert_Progress_Form.StopSpinner
    Sh_Progress_Up = False
    Sh_Progress_Lo = 0
    Sh_Progress_Hi = 100
    Unload Sh_Convert_Progress_Form
    Sh_Focus_Document Sh_Progress_Doc
    Set Sh_Progress_Doc = Nothing
    Err.Clear
    On Error GoTo 0

End Sub  '*** end of Sh_Progress_Close ***
'*** end of the progress bar helpers ***
