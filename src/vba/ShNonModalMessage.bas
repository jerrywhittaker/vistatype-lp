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
    If Sh_Please_Wait_Form.Visible Then Sh_Please_Wait_Form.Advance
    If Sh_NonModalMessageForm.Visible Then Sh_NonModalMessageForm.Advance
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
    Sh_NonModalMessageForm.StartSpinner
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
    On Error Resume Next
    s = Selection.Paragraphs(1).Range.Text
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
