VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Type_Fill_In_Line_Form 
   Caption         =   "Type Fill-In Lines (357)"
   ClientHeight    =   9045.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4545
   OleObjectBlob   =   "Lp_Type_Fill_In_Line_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Type_Fill_In_Line_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Lp_Type_Fill_In_Line_Form
'
' Version: 2.1  Date: 9/23/2026 - Jerry's layout: taller, Done, the short-cut key and the F6 line
'                               on the box itself.
' Version: 2.0  Date: 9/23/2026 - THE BOX STAYS OPEN - Jerry. Shown vbModeless by Lp_Fil_Start;
'                               Cancel is now Done. Every button calls Lp_Fil_Type, which types
'                               the line at the cursor and hands the keyboard back to the book.
'                               F6 and Shift+F6 go back to the book from every control.
' Version: 1.3  Date: 7/24/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
' Version: 1.2  Date: 5/30/2025 - set max lines to 99
' Version: 1.1  Date: 9/26/2023 - added shortcut key to form
' Version: 1.0  Date: 1/3/2019

'
' Every button does one thing: call the macro that does the work. Nothing here touches the book.
'
' KeyDown on EVERY control that can hold the keyboard: F6 and Shift+F6 back to the book (Jerry's
' rule, 9/23/2026). A UserForm has no KeyPreview, and while the box has the keyboard Word never
' sees the key, so a control without this strands a transcriber working without a mouse.
' Shift+F6 arrives here as F6 with Shift set.
'
' QueryClose sends the title bar's X through the same Done as the button, or the box would be
' unloaded behind the macro's back and F6 left pointing at a box that has gone.

Private Sub CancelButton_Click()
    Lp_Fil_Done
End Sub

Private Sub CommandButton1_Click()
    Lp_Fil_Type 1, 1
End Sub

Private Sub CommandButton2_Click()
    Lp_Fil_Type 1, 2
End Sub

Private Sub CommandButton3_Click()
    Lp_Fil_Type 1, 3
End Sub

Private Sub CommandButton4_Click()
    Lp_Fil_Type 1, 4
End Sub

Private Sub CommandButton5_Click()
    Lp_Fil_Type 1, 5
End Sub

Private Sub CommandButton6_Click()
    Lp_Fil_Type 1, 6
End Sub

Private Sub CommandButton7_Click()
    Lp_Fil_Type 1, 7
End Sub

Private Sub CommandButton8_Click()
    Lp_Fil_Type 1, 8
End Sub

Private Sub CommandButton9_Click()
    Lp_Fil_Type 1, 9
End Sub

Private Sub CommandButton10_Click()
    Lp_Fil_Type 1, 10
End Sub

Private Sub CommandButton11_Click()
    Lp_Fil_Type 1, 11
End Sub

Private Sub CommandButton12_Click()
    Lp_Fil_Type 1, 12
End Sub

Private Sub CommandButton13_Click()
    Lp_Fil_Type 1, 13
End Sub

Private Sub CommandButton14_Click()
    Lp_Fil_Type 1, 14
End Sub

Private Sub CommandButton15_Click()
    Lp_Fil_Type 1, 15
End Sub

Private Sub CommandButton16_Click()
    Lp_Fil_Type 1, 16
End Sub

Private Sub CommandButton17_Click()
    Lp_Fil_Type 1, 17
End Sub

Private Sub CommandButton18_Click()
    Lp_Fil_Type 1, 18
End Sub

Private Sub CommandButton19_Click()
    Lp_Fil_Type 1, 19
End Sub

Private Sub CommandButton20_Click()
    Lp_Fil_Type 1, 20
End Sub

Private Sub CommandButton21_Click()
    Lp_Fil_Type 1, 21
End Sub

Private Sub FillToRightMargin_Click()
    Lp_Fil_Type 2, SpinButton1.Value
End Sub

Private Sub TextBox1_Change()
    NewVal = Val(TextBox1.Text)
    If NewVal >= SpinButton1.min And _
        NewVal <= SpinButton1.Max Then _
        SpinButton1.Value = NewVal
End Sub

Private Sub SpinButton1_Change()
    TextBox1.Text = SpinButton1.Value
End Sub

Private Sub CommandButton1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton7_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton8_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton9_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton10_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton11_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton12_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton13_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton14_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton15_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton16_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton17_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton18_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton19_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton20_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CommandButton21_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub FillToRightMargin_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub CancelButton_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub SpinButton1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyF6 Then
        KeyCode = 0
        Lp_Fil_KeyToDocument
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Lp_Fil_Done
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim NewVal, Cntr As Integer

    ' CANCEL IS DONE - Jerry, 9/23/2026. The box stays open until this is pressed, so there is
    ' nothing to cancel. Set here so it reads in the source rather than only in the .frx; the
    ' control keeps its name so the .frx does not have to change.
    CancelButton.Caption = "Done"
    CancelButton.Accelerator = "D"
    With SpinButton1
        .min = 0
        .Max = 99
        .Value = 0
        TextBox1.Text = .Value
    End With
    
    ' 7/24/2026 - removed "MS_Set_Word_Config_For_Large_Print": opening the document already
    '             configures Word for large print, and re-running it here reset the user's
    '             Styles-pane options (show filter / sort order) every time this form opened.
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub
