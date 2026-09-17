VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Type_Fill_In_Line_Form 
   Caption         =   "Type Fill-In Lines (357)"
   ClientHeight    =   7932
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4485
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
' Version: 1.3  Date: 7/24/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
' Version: 1.2  Date: 5/30/2025 - set max lines to 99
' Version: 1.1  Date: 9/26/2023 - added shortcut key to form
' Version: 1.0  Date: 1/3/2019

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub CommandButton1_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 1
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton2_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 2
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton3_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 3
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton4_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 4
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton5_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 5
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton6_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 6
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton7_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 7
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton8_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 8
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton9_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 9
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton10_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 10
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton11_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 11
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton12_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 12
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton13_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 13
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton14_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 14
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton15_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 15
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton16_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 16
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton17_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 17
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton18_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 18
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton19_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 19
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton20_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 20
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub CommandButton21_Click()
    Lp_Type_Fill_In_Line_Form.Hide
    Lp_GP_Counter_1 = 21
    Application.Run MacroName:="Lp_Type_Counted_Fill_In_Lines"
End Sub

Private Sub FillToRightMargin_Click()
    Lp_GP_Counter_1 = SpinButton1.Value
    Lp_Type_Fill_In_Line_Form.Hide
    Application.Run MacroName:="Lp_Type_Fill_In_Line_To_Margin"
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

Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim NewVal, Cntr As Integer
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
