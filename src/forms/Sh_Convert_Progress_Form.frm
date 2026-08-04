VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_Convert_Progress_Form 
   Caption         =   "Converting - Please Wait"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8400.001
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

Private Sub UserForm_Initialize()
'
' The conversion progress box. Replaces Sh_NonModalMessageForm in
' Sh_Convert_XML_File_To_Word_Document (Jerry, 8/3/2026): the old one carried a spinner that
' had never actually turned, under a message telling the user not to touch the keyboard and to
' "Wait for the BEEP!" - written when a conversion took minutes.
'
' MSForms has no progress bar, so the bar is two labels: a sunken track, and a coloured fill
' whose Width the code drives. No ActiveX control, so nothing to fail to register on the
' transcriber''s machine.
'
' Version: 1.0  Date: 8/3/2026
'
    TrackWidth = Me.BarTrack.Width - 4      ' the fill sits 2 points inside the track
    Me.BarFill.Width = 0
    Me.PctLabel.Caption = ""
    Me.MsgLabel.Caption = ""

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
