VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Selected_Cleanup_Form 
   Caption         =   "Selected Cleanup"
   ClientHeight    =   8424.001
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8310.001
   OleObjectBlob   =   "Lp_Selected_Cleanup_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Selected_Cleanup_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Lp_Selected_Cleanup_Form
'
' Version: 1.7 Date: 1/20/2024 - added Lp_Italics_To_Dashed_Underline
' Version: 1.6 Date: 10/12/2023 - Revsion of items - added msg if no selection made
' Version: 1.5 Date: 8/5/2021 - added highlight remove, lp box remove, lp word background color remove, lp para background color remove
' Version: 1.4 Date: 1/28/2020 - added "is text selected" to each selection - added remove spaces befors punction
' Version: 1.3 Date: 2/13/2019 - added Abbyy 14 selection to list
' Version: 1.2 Date: 1/8/2019 - added check for selected text so that user could see menu
' Version: 1.1 Date: 10/19/2018 - added selected text question to calling menu

Private Sub CancelButton_Click()
    Unload Lp_Selected_Cleanup_Form
    End
End Sub


Private Sub OkayButton_Click()
    
    Unload Lp_Selected_Cleanup_Form
    
    If Replace_Mult_Spaces_With_Single Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Lp_Remove_Multi_Spaces"
    ElseIf Remove_Spaces_Before_And_After_Para_Marks Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Lp_Fix_Para_Space_Errors"
    ElseIf Clean_Auto_List_And_Tabs Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Lp_Convert_Auto_List_To_Text"
    ElseIf Text_Boxes_And_Frames Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Lp_Remove_Txt_Bxs_And_Frames"
    ElseIf ReplaceHyper Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Lp_Convert_Hyper_To_Addresses"
    ElseIf Maunal_Line_Break Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Lp_Replace_Manual_Line_Break"
    ElseIf Kill_Hyperlinks Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Lp_Kill_The_Hyperlinks"
    ElseIf Small_Caps_To_All_Caps Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Lp_Replace_Small_Caps_With_All_Caps"
    ElseIf Replace_Tabs_With_Single_Space Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Lp_Replace_Tabs_With_Single_Space"
    ElseIf Box_Bullets_Bullets_and_Numbers Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Lp_Remove_Box_Bullets_Bullets_and_Numbers"
    ElseIf Replace_Multi_Para_Marks_With_Single Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Lp_Replace_Multiple_Para_Marks_With_Warning"
    ElseIf Replace_Sec_Brk Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Lp_Replace_Section_Break_With_Page_Break"
    ElseIf SpacesBeforePunctuation Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Sh_Remove_Spaces_Before_Punctuation"
    ElseIf ItalicsToDashedUnderline Then
        Application.Run MacroName:="Lp_Is_Text_Selected"
        Application.Run MacroName:="Lp_Italics_To_Dashed_Underline"
    Else
        MsgBox "No selection made from Selected Cleanup menu!", , "VistaType LP (163)"
        Lp_Selected_Cleanup_Form.Show
    End If

End Sub


Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub
  
Private Sub UserForm_Initialize()
    Application.Run MacroName:="MS_Set_Word_Config_For_Large_Print"
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

