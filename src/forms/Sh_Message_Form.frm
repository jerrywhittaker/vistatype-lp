VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_Message_Form 
   Caption         =   "VistaType LP"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360.001
   OleObjectBlob   =   "Sh_Message_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sh_Message_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' The one message dialog for the whole add-in. Everything a MsgBox used to say goes through
' here from 8/23/2026, because a MsgBox cannot do ANY of the three things Jerry fixed as the
' rule that day: its text is whatever font Windows draws a message box in, its button says
' "OK" and cannot be made to say "Okay", and it has no accelerators.
'
' Used through Sh_Say and Sh_Ask in LPandBrlMacros - callers do not touch this form directly.
'
'   Sh_Say  "text", "VistaType LP (nnn)"          one Okay button
'   If Sh_Ask("text", "VistaType LP (nnn)") ...   Okay and Cancel, True when Okay was chosen
'
' The nnn is the dialog's own number and every message has a different one - written as nnn here
' on purpose, so that hunting for the next unused number does not turn up this comment.
'
' WHAT IS LOST WITH THE MsgBox, and it is a real loss: the information / warning / question
' ICON, and the sound that went with the warning one. A UserForm has neither. The title bar
' still carries the dialog number, and the text still says what happened.
'
' THE CALLER UNLOADS IT, NOT THE FORM. Okay and Cancel HIDE the form rather than unloading it,
' because a hidden form can still be asked what was pressed. The red X does the same through
' UserForm_QueryClose, which is why that handler cancels the close instead of allowing it.
'
' Author: Jerry Whittaker -  jerry@vistatypelp.org
'
' Version: 1.0  Date: 8/23/2026

' WHAT IT SAYS COMES IN THROUGH FOUR VARIABLES IN LPandBrlMacros - Sh_Msg_Body, Sh_Msg_Title,
' Sh_Msg_Offer_Cancel and Sh_Msg_Answer - and NOT through properties of this form. That is not a
' style choice. Touching any member of a UserForm's default instance is what creates it, and
' creating it runs UserForm_Initialize THERE AND THEN: a caller writing Sh_Message_Form.Msg_Body
' would have the form size itself and caption itself from an empty message on that very line,
' before the second line ran. Variables outside the form are set before the form exists at all.

' How the form re-sizes itself to the message. One message is a line and the next is a screen,
' so the height in the layout file is only a starting point.
'
' CHARS_PER_LINE and LINE_HEIGHT are ESTIMATES, not measurements - a UserForm cannot be
' exercised without a screen, so neither number was measured in Word. They do not have to be
' exact: the text box scrolls, so guessing too few lines costs a scroll bar, and guessing too
' many costs some white space at the bottom.
'
' Msg_Text is 444 points wide, less its scroll bar and margins - call it 424 points, and Tahoma
' 10 averages a little over 5 points a character, so about 80 characters fit. Word wrapping
' throws away up to three more at each break. 72 was break-even at the width this form was
' first built to; 68 against 444 points is a real margin.
Private Const CHARS_PER_LINE As Long = 68
Private Const LINE_HEIGHT As Single = 13
Private Const MIN_LINES As Long = 2
Private Const MAX_LINES As Long = 26      ' past this it scrolls, so the form fits a small screen
Private Const PAD As Single = 12

Private Sub Okay_Button_Click()
    Sh_Msg_Answer = True
    Me.Hide
End Sub

Private Sub Cancel_Button_Click()
    Sh_Msg_Answer = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' The red X counts as Cancel. Hide rather than unload - see the note at the top.
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Sh_Msg_Answer = False
        Me.Hide
    End If
End Sub

' Lays the dialog out for the message it is about to show.
'
' IT TRAPS EVERYTHING, and that is not routine caution. Every message in the product now comes
' through here, so a line that raises would put a raw VBA error in front of the transcriber
' INSTEAD of the message - on the one code path whose whole job is to say something plainly. The
' handler falls back to the layout in the .frx, which is a perfectly usable dialog: the message
' is in it, the buttons are on it, and only the height and the centering are missed.
Private Sub UserForm_Initialize()
    Dim lines As Long
    Dim textH As Single
    Dim chrome As Single

    On Error GoTo LayoutFailed

    Me.Caption = Sh_Msg_Title

    ' PARAGRAPH BREAKS BECOME CrLf, whatever the caller wrote. A MsgBox is happy with a bare Cr
    ' and every message in this project is written with vbCr; an MSForms text box is a different
    ' control, and the one place in this project that already fills one - Sh_Prodnote_Info_Form -
    ' uses vbCrLf throughout. Fixing it HERE rather than in the callers means no message written
    ' from here on can get it wrong, and no message already written has to be edited to be safe.
    ' Everything is taken down to Cr first so that a caller who did write CrLf does not end up
    ' with two blank lines.
    Msg_Text.Text = Replace(Replace(Replace(Sh_Msg_Body, vbCrLf, vbCr), vbLf, vbCr), vbCr, vbCrLf)

    ' Open at the first line, not the last, on a message long enough to scroll.
    Msg_Text.selStart = 0

    ' Only Okay when there is nothing to cancel - and then Esc has to press Okay, or the dialog
    ' would have no keyboard way out at all. Setting Cancel on one button is what takes it off
    ' the other; MSForms allows only one, so the hidden Cancel button cannot answer Esc from
    ' where it cannot be seen.
    Cancel_Button.Visible = Sh_Msg_Offer_Cancel
    Okay_Button.Cancel = Not Sh_Msg_Offer_Cancel

    ' And it moves over into the empty place, so the one button sits at the corner every other
    ' dialog puts a button in rather than a button's width to the left of it.
    If Sh_Msg_Offer_Cancel Then
        Okay_Button.Left = Cancel_Button.Left - Okay_Button.Width - 8
    Else
        Okay_Button.Left = Cancel_Button.Left
    End If

    lines = Sh_Msg_Line_Count(Sh_Msg_Body)
    If lines < MIN_LINES Then lines = MIN_LINES
    If lines > MAX_LINES Then lines = MAX_LINES
    textH = lines * LINE_HEIGHT

    Msg_Text.Top = PAD
    Msg_Text.Height = textH
    Okay_Button.Top = PAD + textH + PAD
    Cancel_Button.Top = Okay_Button.Top

    ' The title bar and borders, which are not the form's to lay out. Measured rather than
    ' assumed, because it changes with the Windows theme and the display scale - it is 29.25
    ' points on the build box at 100%.
    '
    ' RANGE-CHECKED, because this runs in UserForm_Initialize, before the form has ever been on
    ' screen, and InsideHeight answering 0 there would take the whole of Me.Height off and leave
    ' a dialog a title bar tall with the message and both buttons outside it. Nothing here can
    ' be tested without a screen, so the one number that could ruin the form is not trusted.
    chrome = Me.Height - Me.InsideHeight
    If chrome < 6 Or chrome > 120 Then chrome = 30

    Me.Height = Okay_Button.Top + Okay_Button.Height + PAD + chrome

    ' The same again sideways. It should change nothing - the layout file already leaves 12
    ' points at each side - and it is here because it did NOT until 8/23/2026: New-UserForm.ps1
    ' allowed 36 points for a border that takes 12, so every control stopped 24 points short of
    ' the right edge and the dialog looked shoved into its left side. Both were fixed, and this
    ' line means a form built by the older script cannot come out looking like that again.
    chrome = Me.Width - Me.InsideWidth
    If chrome < 2 Or chrome > 120 Then chrome = 12
    Me.Width = Cancel_Button.Left + Cancel_Button.Width + PAD + chrome

    ' Centered on the Word window, which also puts it on the right screen when there are two.
    ' Same approach as Sh_Prodnote_Info_Form. Done LAST, so it centers the final height.
    '
    ' NOT WHEN WORD IS MINIMIZED. A minimized Word answers about -32000 for Left and Top, and
    ' the message would open that far off the side of the screen - which looks exactly like Word
    ' hanging, because a modal dialog nobody can see is still waiting to be answered. Leaving
    ' StartUpPosition alone puts it in the middle of the Word window instead, which is where the
    ' layout file already says. This form is shown with no document open, which is one of the
    ' few times Word can be minimized while a macro runs.
    If Application.WindowState <> wdWindowStateMinimize Then
        Me.StartUpPosition = 0
        Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
        Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    End If

    Exit Sub

LayoutFailed:
    ' Say it anyway. The .frx already holds a working dialog; all that is lost is the fitted
    ' height, the centering and the moved button. Nothing here can raise a second time - these
    ' three are the only properties the message actually needs - and each is set on its own line
    ' so that the two after a failure still run.
    On Error Resume Next
    Me.Caption = Sh_Msg_Title
    Msg_Text.Text = Sh_Msg_Body
    Cancel_Button.Visible = Sh_Msg_Offer_Cancel
End Sub

' How many lines the message will take once the text box has wrapped it. Paragraph marks are
' counted first, then each paragraph is divided by what fits on a line - a paragraph of no
' characters at all is still a line, which is what makes a blank line between paragraphs count.
Private Function Sh_Msg_Line_Count(ByVal body As String) As Long
    Dim paras As Variant
    Dim i As Long
    Dim n As Long
    Dim total As Long

    ' vbCrLf first: splitting on vbCr alone would leave a stray vbLf at the head of the next
    ' paragraph, which counts as a character and is invisible in the source.
    paras = Split(Replace(Replace(body, vbCrLf, vbCr), vbLf, vbCr), vbCr)
    For i = LBound(paras) To UBound(paras)
        n = (Len(paras(i)) + CHARS_PER_LINE - 1) \ CHARS_PER_LINE
        If n < 1 Then n = 1
        total = total + n
    Next i

    Sh_Msg_Line_Count = total
End Function
