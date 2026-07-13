Attribute VB_Name = "ShNonModalMessage"
Option Explicit
Public Sh_BridgeTargetMacro As String

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


