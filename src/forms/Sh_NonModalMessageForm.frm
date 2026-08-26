VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_NonModalMessageForm 
   ClientHeight    =   5520
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7530
   OleObjectBlob   =   "Sh_NonModalMessageForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sh_NonModalMessageForm"
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
    SizeToContent
    CenterOnActiveScreen
End Sub

' THE BOX TAKES THE HEIGHT IT NEEDS, RATHER THAN ALWAYS THE HEIGHT IT MIGHT NEED.
'
' Jerry, 8/26/2026: too big. It was 276 points tall whatever it had to say, because LblMessage
' is laid out 180 points high - two thirds of the box - and the commonest caller never puts
' anything in it at all. Lp_Attach_The_Template shows this form directly with Show vbModeless
' and uses only the activity line underneath, so what was on screen through the whole attach
' was 180 points of blank label. The designer caption is empty too, so there was nothing there
' even in principle.
'
' Done here rather than by resizing the form in the designer, because the two callers want
' genuinely different boxes: Sh_ShowNonModalMessage (the export/import progress) puts two lines
' in LblMessage, and the attach puts none. One designer height has to be wrong for one of them.
'
' It works in BOTH directions and is safe to run twice - a second pass computes the same height,
' gets a shift of zero and leaves. It can never grow past LBL_LAID_OUT, the designer's own 180,
' so this can only ever make the box smaller than it was drawn.
'
' Version: 1.0  Date: 8/26/2026
'
Private Sub SizeToContent()
    Const LBL_LAID_OUT As Single = 180   ' LblMessage's designer height; the ceiling
    Dim wanted As Single
    Dim shift As Single

    ' A box that will not resize must still open. Nothing here is worth failing a macro over.
    On Error Resume Next

    If Len(Trim$(LblMessage.Caption)) = 0 Then
        wanted = 0
    Else
        wanted = WrappedHeight(LblMessage.Caption, LblMessage.Width) + 6
        If wanted > LBL_LAID_OUT Then wanted = LBL_LAID_OUT
    End If

    shift = LblMessage.Height - wanted
    If shift = 0 Then Exit Sub

    LblMessage.Height = wanted
    ActivityIndicator.Top = ActivityIndicator.Top - shift
    Label1.Top = Label1.Top - shift
    SpinnerBox.Top = SpinnerBox.Top - shift
    ActivityMsg.Top = ActivityMsg.Top - shift
    Me.Height = Me.Height - shift
End Sub

' How tall a caption needs to be, counting the rows the label will wrap it onto.
'
' ESTIMATED, and deliberately generously. MSForms will not measure text for you - there is no
' TextWidth on a UserForm - so the two constants below are Tahoma 10 rounded UP: a real average
' character is narrower than 5.5 points and a real line shorter than 13. Erring that way spends
' a few points of height and never clips a word, which is the right direction for a box read by
' people who work in large print.
'
' Version: 1.0  Date: 8/26/2026
'
Private Function WrappedHeight(ByVal s As String, ByVal boxWidth As Single) As Single
    Const CHAR_WIDTH As Single = 5.5     ' Tahoma 10, rounded up
    Const LINE_HEIGHT As Single = 13#    ' Tahoma 10, rounded up
    Dim parts As Variant
    Dim i As Long
    Dim rows As Long
    Dim perRow As Long

    perRow = Int(boxWidth / CHAR_WIDTH)
    If perRow < 1 Then perRow = 1

    ' Every flavor of line break the callers use. Sh_ShowNonModalMessage sends vbCrLf.
    s = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
    parts = Split(s, vbLf)

    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) = 0 Then
            rows = rows + 1              ' a blank line is still a line
        Else
            rows = rows + ((Len(parts(i)) - 1) \ perRow) + 1
        End If
    Next i

    WrappedHeight = rows * LINE_HEIGHT
End Function

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

    Application.OnTime Now + TimeValue("0:00:01"), "Sh_SpinTick"
End Sub

Public Sub Advance()
    ' Move the spinner on ONE frame and repaint, scheduling nothing. Split out of SpinTick so a
    ' long macro can turn the spinner from its own DoEvents calls (Sh_Spin_DoEvents) without
    ' queueing an OnTime for every one of them.
    '
    ' OnTime only fires once a SECOND, so on its own the spinner barely moves during a short
    ' macro and looks stuck. Driving it from the macro's yields is what makes it turn. 8/3/2026.
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
    ' Recorded as well as shown. VBA gives an error handler no line number and no call
    ' stack, so this is the only thing that can tell Sh_Report_Error WHERE in a long
    ' macro it stopped - and these messages are written for the transcriber to read
    ' anyway, so the log gets a plain-English position for one line of code.
    Sh_Last_Activity = sText
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

