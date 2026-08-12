VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_Selected_Cleanup_Form 
   Caption         =   "Selected Cleanup"
   ClientHeight    =   9660.001
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8670.001
   OleObjectBlob   =   "Dx_Selected_Cleanup_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_Selected_Cleanup_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Dx_Selected_Cleanup_Form
'
' Version: 1.3  Date: 3/4/2024 - removed "End" from CmdCancel_Click() added "MS_Set_Word_Config_For_Braille"  to "UserForm_Initialize"
' version: 1.2  Date: 4/7/2023 - show form then notify that no text is selected
' Version: 1.1  Date:8/21/2018
'

Private Sub CmdCancel_Click()
    Unload Dx_Selected_Cleanup_Form
End Sub

Private Sub Okay_Click()

    If Selection.Type <> wdSelectionNormal Then
        MsgBox "Text must be selected first!", , "Braille Macros"
        Unload Dx_Selected_Cleanup_Form
        End
    End If
    
    If Replace_Multi_Para_Marks_With_Single Then
        Unload Dx_Selected_Cleanup_Form
        Application.Run MacroName:="Dx_Replace_Multiple_Para_Marks_No_Warning"
    End If

    If Remove_Spaces_Before_And_After_Para_Marks Then
        Unload Dx_Selected_Cleanup_Form
        Application.Run MacroName:="Dx_Fix_Para_Space_Errors"
     End If

    If Clean_Auto_List_And_Tabs Then
        Unload Dx_Selected_Cleanup_Form
        Application.Run MacroName:="Dx_Convert_Auto_List_To_Text"
    End If
           
    If Small_Caps_To_All_Caps Then
        Unload Dx_Selected_Cleanup_Form
        Application.Run MacroName:="Dx_Replace_Small_Caps_With_All_Caps"
    End If
    
    If Replace_Tabs_With_Single_Space Then
        Unload Dx_Selected_Cleanup_Form
        Application.Run MacroName:="Dx_Replace_Tabs_With_Single_Space"
        Application.Run MacroName:="Sh_Remove_Multi_Spaces"
     End If

    If Replace_Mult_Spaces_With_Single Then
        Unload Dx_Selected_Cleanup_Form
        Application.Run MacroName:="Sh_Remove_Multi_Spaces"
    End If
    
    If Kill_Hyperlinks Then
        Unload Dx_Selected_Cleanup_Form
        Application.Run MacroName:="Dx_Kill_The_Hyperlinks"
    End If
    
    If Remove_Pictures Then
        Unload Dx_Selected_Cleanup_Form
        Application.Run MacroName:="Dx_Delete_Images"
    End If

    If Remove_Bullets Then
        Unload Dx_Selected_Cleanup_Form
        Application.Run MacroName:="Dx_Remove_Bullets"
    End If

    If Text_Boxes_And_Frames Then
        Unload Dx_Selected_Cleanup_Form
        Application.Run MacroName:="Dx_Remove_Txt_Bxs_And_Frames"
    End If
             
    If ReplaceHyper Then
        Unload Dx_Selected_Cleanup_Form
        Application.Run MacroName:="Dx_Convert_Hyper_To_Addresses"
    End If
        
    If Manual_Line_Break Then
        Unload Dx_Selected_Cleanup_Form
        Application.Run MacroName:="Sh_Replace_Manual_Line_Break"
    End If
        
    If NonBreakingSpace Then
        Unload Dx_Selected_Cleanup_Form
        Application.Run MacroName:="Sh_ReplaceNonBreakingSpacesWithNormalSpace"
        Application.Run MacroName:="Sh_Remove_Multi_Spaces"
    End If
    
    'If xx Then  ' dummy routine
    '    Unload Dx_Selected_Cleanup_Form
    '    Application.Run MacroName:=""
    'End If

    Application.ScreenUpdating = True ' Turn screen updating on

End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Dx_Selected_Cleanup_Form
End Sub

Private Sub UserForm_Initialize()
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    

End Sub

