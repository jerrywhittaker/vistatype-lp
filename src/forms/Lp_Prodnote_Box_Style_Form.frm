VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Prodnote_Box_Style_Form 
   Caption         =   "Convert All Prodnotes (427)"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   OleObjectBlob   =   "Lp_Prodnote_Box_Style_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Prodnote_Box_Style_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Lp_Prodnote_Box_Style_Form - Convert All Prodnotes (427), opened by Convert All Prodnotes to LP
' Style on Prodnote Review (425), through Lp_Convert_All_Prodnotes_To_Box
'
' Asks which box style every prodnote in the book goes to, Box Red Style or Box Blue Style, and
' hands the answer back in Lp_Cap_Style: the style's name on Okay, "" on Cancel and on the red X,
' which acts like Cancel (Jerry, 10/11/2026). Box Red Style is chosen when the box opens.
'
' Box Red Prefix and Box Blue Prefix (Alt+E, Alt+L): the bold word every converted prodnote starts
' with. Saved as it is typed, so Cancel does not undo a change to them.
'
' Version: 1.1  Date: 11/8/2026 - Box Red Prefix and Box Blue Prefix
' Version: 1.0  Date: 11/8/2026

Private filling As Boolean          ' UserForm_Initialize is putting the saved prefixes in

Private Sub OkayButton_Click()
    If BoxBlueOption.Value Then
        Lp_Cap_Style = "Box Blue"
    Else
        Lp_Cap_Style = "Box Red"
    End If
    Unload Me
End Sub

Private Sub CancelButton_Click()
    Lp_Cap_Style = ""                 'no answer - nothing is converted
    Unload Me
End Sub

Private Sub BoxRedPrefixText_Change()
    If Not filling Then Lp_Rvp_Prefix_Set "Box Red", BoxRedPrefixText.Text
End Sub

Private Sub BoxBluePrefixText_Change()
    If Not filling Then Lp_Rvp_Prefix_Set "Box Blue", BoxBluePrefixText.Text
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'The red X is Cancel: it leaves no answer, and the box then closes by itself. Do not set
    'Cancel = True and unload from here - Word ignores an Unload inside QueryClose.
    If CloseMode = vbFormControlMenu Then Lp_Cap_Style = ""
End Sub

Private Sub UserForm_Initialize()
    Lp_Cap_Style = ""
    QuestionLabel.Caption = "To which style should all prodnotes be converted?"
    BoxRedOption.Value = True

    ' The two prefixes (Jerry, 11/8/2026): the words a converted prodnote starts with, in bold.
    ' Shared with Review Prodnotes (426) through VistaType.ini (Lp_Rvp_Prefix_Get). Each box's Change writes it
    ' back; filling keeps these two lines from writing what was just read.
    BoxRedPrefixLabel.Caption = "Box Red Prefix"
    BoxRedPrefixLabel.Accelerator = "E"
    BoxBluePrefixLabel.Caption = "Box Blue Prefix"
    BoxBluePrefixLabel.Accelerator = "L"
    filling = True
    BoxRedPrefixText.Text = Lp_Rvp_Prefix_Get("Box Red")
    BoxBluePrefixText.Text = Lp_Rvp_Prefix_Get("Box Blue")
    filling = False

    ' Start the form centered inside the Word window (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub
