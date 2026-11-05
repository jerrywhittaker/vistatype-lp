VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Pic_Description_Form 
   Caption         =   "Picture Description (414)"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6960
   OleObjectBlob   =   "Lp_Pic_Description_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Pic_Description_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Lp_Pic_Description_Form - Picture Description (414)
'
' Version: 1.0  Date: 11/5/2026
'
' Jerry, 11/5/2026: from Edit Description on Resize Pictures in a Selected Range (375). The
' description typed here is given to the clicked picture and every copy of it on the next Apply.
' Nothing here touches the book: the text comes in and goes out through two Publics in
' LPandBrlMacros, Lp_Rst_Desc_Text and Lp_Rst_Desc_Ok, and Lp_Rst_Edit_Description does the rest.
'
' MODAL. Okay is the default button, so Enter is Okay (EnterKeyBehavior is False on the text box);
' Esc, Cancel and the red X all leave Lp_Rst_Desc_Ok False, and every way out unloads the box.

Private Sub OkayButton_Click()
    Lp_Rst_Desc_Text = DescText.Text
    Lp_Rst_Desc_Ok = True
    Unload Me
End Sub

Private Sub CancelButton_Click()
    Lp_Rst_Desc_Ok = False
    Unload Me
End Sub

' The red X is Cancel (Jerry, 10/11/2026). Cancel is left alone, so the box closes itself.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Lp_Rst_Desc_Ok = False
End Sub

Private Sub UserForm_Initialize()
    PromptLabel.Caption = "Type the description for this picture. When you press Apply, the " _
                        & "picture and every copy of it in your range are given it."
    OkayButton.Caption = "Okay"
    OkayButton.Accelerator = "O"
    CancelButton.Caption = "Cancel"
    CancelButton.Accelerator = "C"
    ' Enter is Okay and Esc is Cancel. Set in the .frx as well; here so the source says so.
    OkayButton.Default = True
    CancelButton.Cancel = True
    DescText.EnterKeyBehavior = False

    DescText.Text = Lp_Rst_Desc_Text
End Sub

' The text box has the keyboard with all of its text selected, so typing replaces it and an arrow
' key edits it. In Activate, not Initialize: SetFocus before the box is on screen can raise.
Private Sub UserForm_Activate()
    On Error Resume Next
    DescText.SetFocus
    DescText.SelStart = 0
    DescText.SelLength = Len(DescText.Text)
    Err.Clear
    On Error GoTo 0
End Sub
