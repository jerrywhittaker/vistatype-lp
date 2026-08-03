VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_Please_Wait_Form 
   Caption         =   "Working - Please Wait"
   ClientHeight    =   1560
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7305
   OleObjectBlob   =   "Sh_Please_Wait_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sh_Please_Wait_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private SpinFrames As Variant
Private Frame As Long
Private SpinnerRunning As Boolean

Private Sub UserForm_Initialize()
    SpinFrames = Array("|", "/", "—", "\")
    Frame = 0
    SpinnerRunning = False
End Sub

Private Sub UserForm_Activate()
    CenterOnActiveScreen
End Sub

Public Sub StartSpinner()
    If SpinnerRunning Then Exit Sub
    SpinnerRunning = True
    Frame = 0
    SpinTick
End Sub

Public Sub StopSpinner()
    SpinnerRunning = False
End Sub

Public Sub SpinTick()
    ' Repainting this modeless form lets Word paint the document underneath it. During
    ' a long macro that shows as a flash of half-restyled content, once per tick - the
    ' yellow splashes seen all through Lp_Normalize_Styles (Jerry, 7/26/2026). OnTime
    ' re-fires this every second for the whole run, so it is a steady drip, not one hit.
    ' Hold ScreenUpdating across the repaint: the spinner still animates, the document
    ' stays put.
    If Not SpinnerRunning Then Exit Sub

    Advance

    ' Sh_PleaseWaitTick, not Sh_SpinTick. OnTime can only name a module-level macro, and
    ' Sh_SpinTick ticks Sh_NonModalMessageForm -- this form arrived pointing at it, so the
    ' spinner advanced one frame and stopped. 8/3/2026.
    Application.OnTime Now + TimeValue("0:00:01"), "Sh_PleaseWaitTick"
End Sub

Public Sub Advance()
    ' Move the spinner on ONE frame and repaint, scheduling nothing. Split out of SpinTick so a
    ' long macro can turn the spinner from its own DoEvents calls (Sh_Spin_DoEvents) without
    ' queueing an OnTime for every one of them - 28 DoEvents would mean 28 pending timers.
    '
    ' OnTime only fires once a SECOND, so on its own the spinner barely moves during a short
    ' macro. Driving it from the DoEvents is what makes it look alive. 8/3/2026.
    '
    ' Holds ScreenUpdating across the repaint, for the reason given in SpinTick above.
    Dim su_Prev As Boolean
    If Not SpinnerRunning Then Exit Sub

    su_Prev = Application.ScreenUpdating
    Me.SpinnerBox.Caption = SpinFrames(Frame)
    Frame = (Frame + 1) Mod (UBound(SpinFrames) + 1)
    Me.Repaint
    If Application.ScreenUpdating <> su_Prev Then Application.ScreenUpdating = su_Prev
End Sub

Public Sub SetActivityMessage(ByVal sText As String)
    ' Same reason as SpinTick above - the repaint must not let the document redraw.
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Me.ActivityMsg.Caption = sText
    Me.Repaint
    If Application.ScreenUpdating <> su_Prev Then Application.ScreenUpdating = su_Prev
End Sub

Private Sub CenterOnActiveScreen()
    Dim w As Long, h As Long
    w = Application.UsableWidth
    h = Application.UsableHeight

    Me.StartUpPosition = 0
    Me.Left = (w - Me.Width) / 2
    Me.Top = (h - Me.Height) / 2
End Sub

