Attribute VB_Name = "ShNonModalMessage"
Option Explicit
Public Sh_BridgeTargetMacro As String

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
    'Sh_NonModalMessageForm.SetActivityMessage "Attaching the LP template"

    If Len(Sh_BridgeTargetMacro) > 0 Then
        Application.Run Sh_BridgeTargetMacro 'Sh_BridgeTargetMacro is the name of the macro to be run
    End If
End Sub


