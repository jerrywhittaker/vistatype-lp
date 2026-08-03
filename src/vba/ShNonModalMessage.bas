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

' Handing focus back to the document needs the Windows API. Activating the document window
' through the object model is NOT enough: a modeless UserForm keeps the keyboard focus, so
' Word draws no caret and the arrow keys walk the form's buttons instead of the text
' (Jerry, 7/27/2026). Window.Hwnd is available here - verified on Word 16.0 build 20131.
#If VBA7 Then
    Private Declare PtrSafe Function Sh_SetFocusApi Lib "user32" Alias "SetFocus" _
        (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function Sh_SetForegroundWindowApi Lib "user32" Alias "SetForegroundWindow" _
        (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function Sh_SetFocusApi Lib "user32" Alias "SetFocus" _
        (ByVal hwnd As Long) As Long
    Private Declare Function Sh_SetForegroundWindowApi Lib "user32" Alias "SetForegroundWindow" _
        (ByVal hwnd As Long) As Long
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
' Version: 1.0  Date: 8/3/2026
    With Sh_Please_Wait_Form
        .ActivityMsg.Caption = sMessage
        .SpinnerBox.Caption = ""
        .Show vbModeless
        .StartSpinner
    End With
    DoEvents
End Sub

Public Sub Sh_Hide_Please_Wait()
' Stops the spinner and closes the box. Stopping first matters: it clears the flag SpinTick
' tests, so a tick already queued by OnTime does nothing instead of reopening the form.
'
' Version: 1.0  Date: 8/3/2026
    On Error Resume Next
    Sh_Please_Wait_Form.StopSpinner
    Unload Sh_Please_Wait_Form
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

    'Hide before unloading: these run from a button on one of the forms being closed.
    Sh_Valid_Ref_Pg_No_2_Form.Hide
    Sh_Valid_Ref_Pg_No_4_Form.Hide

    If Sh_PgVal_DocIsOpen(Sh_PgVal_TempDoc) Then
        Sh_PgVal_TempDoc.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Sh_PgVal_DocIsOpen(Sh_PgVal_SourceDoc) Then
        Sh_PgVal_SourceDoc.Activate
    End If

    Unload Sh_Valid_Ref_Pg_No_1_Form
    Unload Sh_Valid_Ref_Pg_No_2_Form
    Unload Sh_Valid_Ref_Pg_No_3_Form
    Unload Sh_Valid_Ref_Pg_No_4_Form

    Set Sh_PgVal_TempDoc = Nothing
    Set Sh_PgVal_SourceDoc = Nothing
    Sh_PgVal_LastParaStart = -1
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
    Sh_PgVal_FocusDocument Sh_PgVal_TempDoc
End Sub

Private Sub Sh_PgVal_SwapToSourceForm()
    On Error Resume Next
    Unload Sh_Valid_Ref_Pg_No_4_Form
    Sh_Valid_Ref_Pg_No_4_Form.Show vbModeless
    Sh_Valid_Ref_Pg_No_2_Form.Hide
    Sh_PgVal_FocusDocument Sh_PgVal_SourceDoc
End Sub

' Showing a modeless UserForm takes the keyboard focus, so Word stops drawing a caret
' in the document behind it - the cursor looks like it has vanished. The message box
' this feature replaced used to hand focus back when it was dismissed; nothing does now,
' so every swap has to do it deliberately (Jerry, 7/27/2026).
Private Sub Sh_PgVal_FocusDocument(ByVal d As Document)
    On Error Resume Next
    If Not Sh_PgVal_DocIsOpen(d) Then Exit Sub

    d.Activate
    d.ActiveWindow.Activate

    'The object-model activation above moves Word to the right document but leaves the
    'keyboard focus on the modeless form. Only Windows can take it back.
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
    Sh_PgVal_TextOfCurrentTag = Trim$(s)
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
