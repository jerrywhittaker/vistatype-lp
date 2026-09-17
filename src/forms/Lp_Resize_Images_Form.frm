VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Resize_Images_Form 
   Caption         =   "Resize Selected Pictures or All Pictures (350)"
   ClientHeight    =   5892
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5835
   OleObjectBlob   =   "Lp_Resize_Images_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Resize_Images_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOkay_Click()

    ' check the percent value range
    If Val(Percent_Requested) > 575 Or Val(Percent_Requested) < 1 Then
        MsgBox "Percent is out of range", , "VistaType LP (300)"
    Else
        If Resize_Selected Then
            Lp_Pic_All_Selectd = "S" 'save the choice
            Lp_Pic_Percent = Str(Int(Percent_Requested))  'save the choice
            'Unload Me 'close this form
            Lp_Resize_Images_Form.Hide
        End If
        
        If Resize_All Then
            Lp_Pic_All_Selectd = "A"  'save the choice
            Lp_Pic_Percent = Percent_Requested   'save the choice
            'Unload Me 'close this form
            Lp_Resize_Images_Form.Hide
        End If
    End If
    
    Application.Run MacroName:="Lp_Resize_Images"
    Unload Me 'close this form
    
End Sub

Private Sub UserForm_Initialize()

    If Lp_Pic_All_Selectd = "" And Lp_Pic_All_Selectd = "" Then
            Resize_Selected.Value = True
    End If
    
    If Lp_Pic_All_Selectd = "A" Then
        Resize_All.Value = True
    End If
    
    If Lp_Pic_All_Selectd = "S" Then
        Resize_Selected.Value = True
    End If
    
    If Val(Lp_Pic_Percent) < 1 Then
        Lp_Pic_Percent = "50"
    Else
        Percent_Requested = Lp_Pic_Percent
    End If
    
    Percent_Requested = Val(Lp_Pic_Percent)
    
    With Me.Percent_Requested
        .selStart = 0
        .SelLength = 3
    End With
    
    Resize_Selected = True
    
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)


End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub
