VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_Convert_Progress_Form 
   Caption         =   "Converting - Please Wait"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8955.001
   OleObjectBlob   =   "Sh_Convert_Progress_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sh_Convert_Progress_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

' Full width of the bar track, in points. Set from the designer value on load so the two stay
' in step if the form is ever resized.
Private TrackWidth As Single

' The spinner, added 9/6/2026. The bar says how far through the job the user is; the
' spinner says the job is still alive. It earns its place because three steps in this
' project are ONE long call into Word - Repaginate, InsertFile and the Save As dialog -
' and VBA is single-threaded, so nothing can move the fill while one of them runs. Until
' now the only answer was to say so in the message text ("the bar cannot move during this
' step"), which reads like an apology for a hang.
'
' Chr$(151) is the em dash, and it is written as a code rather than typed because this
' file is read and rewritten by tools on Linux and a stray byte in a .frm is the kind of
' fault that only shows on the transcriber's machine. The other two spinner forms carry
' the raw byte; do not copy that.
Private SpinFrames As Variant
Private Frame As Long
Private SpinnerRunning As Boolean

Private Sub UserForm_Initialize()
'
' The conversion progress box. Replaces Sh_NonModalMessageForm in
' Sh_Convert_XML_File_To_Word_Document (Jerry, 8/3/2026): the old one carried a spinner that
' had never actually turned, under a message telling the user not to touch the keyboard and to
' "Wait for the BEEP!" - written when a conversion took minutes.
'
' MSForms has no progress bar, so the bar is two labels: a sunken track, and a colored fill
' whose Width the code drives. No ActiveX control, so nothing to fail to register on the
' transcriber''s machine.
'
' Version: 1.0  Date: 8/3/2026
'
    SpinFrames = Array("|", "/", Chr$(151), "\")
    Frame = 0
    SpinnerRunning = False

    TrackWidth = Me.BarTrack.Width - 4      ' the fill sits 2 points inside the track
    Me.BarFill.Width = 0
    Me.PctLabel.Caption = ""
    Me.MsgLabel.Caption = ""
    Me.SpinnerBox.Caption = ""

    ' 20 point, twice the form's Tahoma 10 - Jerry, 9/6/2026, seeing the first build.
    ' Set HERE and not in the designer because the form-building tools apply Tahoma 10 to
    ' everything they add and a control's font can only be written from VBA anyway. The
    ' form was widened to 448 points to give the spinner a gutter of its own beside the
    ' bar: at 20 point it does not fit on the percentage row.
    Me.SpinnerBox.Font.Size = 20

    ' Start centered inside the Word window, which also puts it on the right screen when there
    ' are two. Same approach as the other progress dialogs.
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

Public Sub SetProgress(ByVal pct As Single, ByVal sText As String)
'
' Move the bar to pct (0 to 100) and say what is happening.
'
' Holds ScreenUpdating across the repaint, for the same reason the other progress forms do:
' repainting a modeless form lets Word paint the document underneath it, which shows as a flash
' of half-converted content.
'
' Version: 1.0  Date: 8/3/2026
'
    Dim su_Prev As Boolean

    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100

    su_Prev = Application.ScreenUpdating

    If Len(sText) > 0 Then Me.MsgLabel.Caption = sText
    Me.BarFill.Width = TrackWidth * (pct / 100)
    Me.PctLabel.Caption = Format(pct, "0") & "%"
    Me.Repaint

    If Application.ScreenUpdating <> su_Prev Then Application.ScreenUpdating = su_Prev
End Sub

Public Sub StartSpinner()
' Begin turning. Called by Sh_Progress_Open, so every bar gets one without its callers asking.
'
' Version: 1.0  Date: 9/6/2026
    If SpinnerRunning Then Exit Sub
    SpinnerRunning = True
    Frame = 0
    SpinTick
End Sub

Public Sub StopSpinner()
' Stop turning. Called by Sh_Progress_Close BEFORE the Unload, so a tick already queued with
' Application.OnTime cannot bring the form back after it has gone - the trap Sh_Hide_Please_Wait
' was written to avoid on 8/3/2026.
'
' Version: 1.0  Date: 9/6/2026
    SpinnerRunning = False
End Sub

Public Sub SpinTick()
' One frame, then queue the next second.
'
' Sh_Progress_Tick, NOT Sh_SpinTick or Sh_PleaseWaitTick. Application.OnTime can only name a
' module-level macro, so each spinner form needs its own - and this form arriving pointed at
' another form's tick is exactly how Sh_Please_Wait_Form shipped with a spinner that advanced
' one frame and stopped (8/3/2026).
'
' Version: 1.0  Date: 9/6/2026
    If Not SpinnerRunning Then Exit Sub

    Advance
    Application.OnTime Now + TimeValue("0:00:01"), "Sh_Progress_Tick"
End Sub

Public Sub Advance()
' Move on ONE frame and repaint, scheduling nothing. Split out of SpinTick so a long macro can
' turn the spinner from its own DoEvents calls without queueing a timer for every one of them.
' OnTime cannot fire closer than a second apart, so on its own the spinner barely moves.
'
' Holds ScreenUpdating across the repaint, for the same reason SetProgress does: repainting a
' modeless form lets Word paint the document underneath, which shows as a flash of
' half-converted content.
'
' Version: 1.0  Date: 9/6/2026
    Dim su_Prev As Boolean
    If Not SpinnerRunning Then Exit Sub

    su_Prev = Application.ScreenUpdating
    Me.SpinnerBox.Caption = SpinFrames(Frame)
    Frame = (Frame + 1) Mod (UBound(SpinFrames) + 1)
    Me.Repaint
    If Application.ScreenUpdating <> su_Prev Then Application.ScreenUpdating = su_Prev
End Sub
