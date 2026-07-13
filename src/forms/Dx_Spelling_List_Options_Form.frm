VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_Spelling_List_Options_Form 
   Caption         =   "Format Spelling Word List (Contracted/Uncontracted)"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5280
   OleObjectBlob   =   "Dx_Spelling_List_Options_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_Spelling_List_Options_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Date: 6/18/2018
' Version 1.4

Private Sub CmdCancel_Click()
    Unload Me
    End
End Sub


Private Sub CmdOk_Click()

    Dx_Spelling_List_Options_Form.Hide
    Application.ScreenUpdating = False ' Turn screen updating off

'**********************************************************************************
' set the BANA style to List1
'**********************************************************************************

    Selection.Style = ActiveDocument.Styles("List1")
    
'**********************************************************************************
' delete spaces before para mark
'**********************************************************************************
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

    With Selection.Find
        .Text = "^032{1,}^013"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '**********************************************************************************
    ' delete spaces after para mark
    '**********************************************************************************
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

    With Selection.Find
        .Text = "^013^032{1,}"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '**********************************************************************************
    ' replace multi spaces with single space
    '**********************************************************************************
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

    With Selection.Find
        .Text = "^032{1,}"
        .Replacement.Text = "^032"
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '**********************************************************************************
    ' replace multi periods with single period
    '**********************************************************************************
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

    With Selection.Find
        .Text = "^046{1,}"
        .Replacement.Text = "^046"
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
If Is_Ordered Then 'list is orderd

    '**********************************************************************************
    ' replace periods not followed by a space with a period plus a space
    '**********************************************************************************
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

    With Selection.Find
        .Text = "^046([{A-z){1,}])"
        .Replacement.Text = "^046^032\1"
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    '**********************************************************************************
    ' Find the target word(s)
    '**********************************************************************************

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

    With Selection.Find
        .Text = "^032(<*>)^013"
        
        If Two_Spaces Then
            .Replacement.Text = "^032\1^030^030\1^p"
         End If
      
         If One_Space Then
            .Replacement.Text = "^032\1^030\1^p"
         End If
        
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End If

If Is_Unordered Then 'list is NOT ordered

    '**********************************************************************************
    ' Find the target word(s)
    '**********************************************************************************
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
       With Selection.Find
           .Text = "(<*>)^013"
           
         If Two_Spaces Then
           .Replacement.Text = "\1^030^030\1^p"
         End If
         
         If One_Space Then
            .Replacement.Text = "\1^030\1^p"
         End If
         
           .Forward = True
           .Wrap = wdFindStop
           .Format = False
           .MatchCase = False
           .MatchWholeWord = False
           .MatchAllWordForms = False
           .MatchSoundsLike = False
           .MatchWildcards = True
       End With
       Selection.Find.Execute Replace:=wdReplaceAll
End If
    
    '**********************************************************************************
    ' Duplicate the target word(s)in uncontracted form
    '**********************************************************************************
        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Uncontracted")
        
    With Selection.Find
    
        If Two_Spaces Then
            .Text = "(^030^030<*>)^013"
        End If
      
        If One_Space Then
            .Text = "(^030<*>)^013"
        End If
      
        .Replacement.Text = "\1^p"
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '**********************************************************************************
    ' replace hard hyphens with hard spaces
    '**********************************************************************************
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        
    With Selection.Find
        .Text = "^030"
        
         If Two_Spaces Then
            .Replacement.Text = "^0160"
         Else
            .Replacement.Text = "^032" 'one_space assumed
         End If
         
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Application.ScreenUpdating = True ' Turn screen updating on
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"

End Sub

Private Sub userform_terminate() 'red X was clicked
    Application.ScreenUpdating = True ' Turn screen updating on
    Unload Me
    End
End Sub

Private Sub UserForm_Initialize()
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

