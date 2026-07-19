VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_Horz_To_Vert_List_Form 
   Caption         =   "Horizonal List to Vertical List"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8130
   OleObjectBlob   =   "Dx_Horz_To_Vert_List_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_Horz_To_Vert_List_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Version 1.0 4/3/2018
' Version 1.1 5/17/2018
' Version 1.2 7/5/2018
' Version 1.3 8/23/2018 - added check for ending para mark
' Version 1.4 10/3/2018 - added ordinary bullet

Private Sub Cmd_Cancel_Click()
    Unload Me
    End
End Sub

Private Sub Cmd_Ok_Click()

    Dx_Horz_To_Vert_List_Form.Hide
    
    Application.ScreenUpdating = False ' Turn screen updating off

    Application.Run MacroName:="Dx_Copy_To_Temp_Doc" 'move selected text to temp file
    Application.Run MacroName:="Sh_Is_End_Paragraph_Mark_Included"
    Selection.WholeStory 'select the whole document
    
    'replace manual line break with para mark
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^l"
        .Replacement.Text = "^p"
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
    
    'replace multiple periods
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^046{2,}"
        .Replacement.Text = "^046"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'replace multiple spaces
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{2,}"
        .Replacement.Text = "^032"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'replace space period with period
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032^046"
        .Replacement.Text = "^046"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    If Ordered_List Then

        ' upper or lower case followed by period followed by period and space
        'number followed by period followed by period and space
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "[A-z0-9]{1,}^046^032"
            .Replacement.Text = "^p^&"
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
 
        ' change double parens (something) to ~somthingå -Alt+0140
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "\(([A-z0-9]{1,})\)"
            .Replacement.Text = "~\1å"
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        ' fix the double paren substute followed by a space
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "(~[A-z0-9]{1,}å^032)"
            .Replacement.Text = "^p^&"
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        ' fix the double paren substute followed by a period
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "(~[A-z0-9]{1,}å^046)"
            .Replacement.Text = "^p^&"
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        ' ordinary bullet followed by space
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^0149^032*"
            .Replacement.Text = "^p^&"
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        'upper or lower case followed by period only
        'number followed by period only
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "([A-z0-9]{1,}^046)"
            .Replacement.Text = "^p\1^032"
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        'upper or lower case followed by parenthesis followed by a period
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "([A-z0-9]{1,}\)^046)"
            .Replacement.Text = "^p\1^032"
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        'upper or lower case followed by parenthesis
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "([A-z0-9]{1,}\))"
            .Replacement.Text = "^p\1^032"
        End With
        Selection.Find.Execute Replace:=wdReplaceAll

    ElseIf Spaced_List Then
        
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^032{1,}"
            .Replacement.Text = "^p"
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
        
        Selection.HomeKey Unit:=wdStory
        Selection.TypeParagraph
    
    Else  'Is a Tabbed_List
    
        'remove multi spaces
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^032{2,}"
            .Replacement.Text = "^032"
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
        
        'remove spaces following tabs
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^009^032{1,}"
            .Replacement.Text = "^009"
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        'remove spaces preceeding tabs
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^032{1,}^009"
            .Replacement.Text = "^009"
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
            
        'remove multi tabs
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^009{2,}"
            .Replacement.Text = "^009"
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        'Replace tabs with para mark
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^009{1,}"
            .Replacement.Text = "^p"
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        Selection.HomeKey Unit:=wdStory
        Selection.TypeParagraph

    End If

    ' cleanup before pasting back into document

    'remove muilt spaces
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}"
        .Replacement.Text = "^032"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' Replace ~ with (
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "~"
        .Replacement.Text = "("
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' Replace Chr(0140) with )
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "å"  'Alt+0140
        .Replacement.Text = ")"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'remove spaces before para marks
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}^013"
        .Replacement.Text = "^p"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'remove spaces after para marks
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013^032{1,}"
        .Replacement.Text = "^p"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'remove spaces preceeding periodes
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}^046"
        .Replacement.Text = "^046"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'remove muilt para marks
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013{2,}"
        .Replacement.Text = "^p"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.HomeKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1

    If AscendingOrderCheckbox = True Then
        Selection.WholeStory
        Application.Run MacroName:="Sh_Sort_Ascending"
    End If

    Application.Run MacroName:="Dx_Copy_From_Temp_Doc"

    Application.ScreenUpdating = True ' Turn screen updating on
    Application.ScreenRefresh
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
  
    Unload Me
    
End Sub

Sub UserForm_Initialize()
    Me.AscendingOrderCheckbox.Value = True
    
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub
