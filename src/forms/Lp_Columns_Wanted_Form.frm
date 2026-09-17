VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Columns_Wanted_Form 
   Caption         =   "Column Settings (346)"
   ClientHeight    =   3984
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4980
   OleObjectBlob   =   "Lp_Columns_Wanted_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Columns_Wanted_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Version: 1.2  Date: 7/24/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
' Version: 1.1  Date: 12/13/2021 - fixed fatal error in spin box code
' Version: 1.0  Date: 1/14/2020
'

Private Sub BorderWanted_Click()
If BorderWanted.Value = True Then
    Black.Enabled = True
    Red.Enabled = True
    Orange.Enabled = True
    Blue.Enabled = True
    Violet.Enabled = True
    Green.Enabled = True
    Black.Value = True
Else
    Black.Enabled = False
    Red.Enabled = False
    Orange.Enabled = False
    Blue.Enabled = False
    Violet.Enabled = False
    Green.Enabled = False
    Black.Value = False
End If
    
End Sub

Private Sub CancelButton_Click()
    End
End Sub

Private Sub OkayButton_Click()
    Lp_GP_String_1 = TextBox1
    
     Lp_GP_String_3 = ""
     
    If BorderWanted.Value = True Then
        Lp_GP_String_3 = "Y"
    End If
    
    If Black Then
        Lp_GP_String_3 = Lp_GP_String_3 + "1 "
    End If
    
    If Red Then
        Lp_GP_String_3 = Lp_GP_String_3 + "2 "
    End If
    
    If Orange Then
        Lp_GP_String_3 = Lp_GP_String_3 + "3 "
    End If
    
    If Blue Then
        Lp_GP_String_3 = Lp_GP_String_3 + "4 "
    End If
    
    If Violet Then
        Lp_GP_String_3 = Lp_GP_String_3 + "5 "
    End If
    
    If Green Then
        Lp_GP_String_3 = Lp_GP_String_3 + "6 "
    End If
    
    Unload Me
End Sub

Private Sub TextBox1_Change()
    NewVal = Val(TextBox1.Text)
    If NewVal > 8 Then
        NewVal = 8
        TextBox1 = "8"
    End If
    If NewVal >= SpinButton1.min And _
        NewVal <= SpinButton1.Max Then _
        SpinButton1.Value = NewVal
End Sub

Private Sub SpinButton1_Change()
    TextBox1.Text = SpinButton1.Value
End Sub

Private Sub UserForm_Initialize()
    Dim NewVal As Integer
    Dim ColumnCount As Integer

    ' for form testing - uncomment the next line
    'Lp_GP_Counter_1 = 3
    
    ' for form testing - uncomment the next line
    ' Lp_GP_Counter = 3
    'MsgBox Lp_GP_Counter_1
    'End

    ColumnCount = Lp_GP_Counter_1
    CurrentColumnCountLabel.Caption = "The selected table currently has" + Str(ColumnCount) + " columns"
    CurrentColumnCountLabel.ControlTipText = "The selected table currently has" + Str(ColumnCount) + " columns"

    With SpinButton1
        .min = 1
        .Max = 8
       TextBox1.Text = .Value
     End With

    Black.Enabled = False
    Red.Enabled = False
    Orange.Enabled = False
    Blue.Enabled = False
    Violet.Enabled = False
    Green.Enabled = False

    ' 7/24/2026 - removed "MS_Set_Word_Config_For_Large_Print": opening the document already
    '             configures Word for large print, and re-running it here reset the user's
    '             Styles-pane options (show filter / sort order) every time this form opened.
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub userform_terminate() 'red X was clicked
        Lp_GP_String_1 = TextBox1
    Unload Me
End Sub






