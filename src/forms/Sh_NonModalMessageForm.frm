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
    If Not SpinnerRunning Then Exit Sub

    Me.SpinnerBox.Caption = SpinFrames(Frame)
    Frame = (Frame + 1) Mod (UBound(SpinFrames) + 1)
    Me.Repaint

    Application.OnTime Now + TimeValue("0:00:01"), "Sh_SpinTick"
End Sub

Public Sub SetActivityMessage(ByVal sText As String)
    Me.ActivityMsg.Caption = sText
    Me.Repaint
End Sub

Private Sub CenterOnActiveScreen()
    Dim w As Long, h As Long
    w = Application.UsableWidth
    h = Application.UsableHeight

    Me.StartUpPosition = 0
    Me.Left = (w - Me.Width) / 2
    Me.Top = (h - Me.Height) / 2
End Sub

