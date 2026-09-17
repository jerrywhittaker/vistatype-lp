VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DN_Tag_Daisy_Nimas_Form 
   Caption         =   "Add $pg tags to Reference Page Numbers"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5760
   OleObjectBlob   =   "DN_Tag_Daisy_Nimas_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DN_Tag_Daisy_Nimas_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' DN_Tag_Daisy_Nimas_Form
'
' Author: Jerry Whittaker -  jerry@vistatypelp.org
'
' Version: 1.8  Date: 9/17/2026 - THE DAISY HALF HAD NEVER REPLACED ANYTHING. Its Execute read
'                                 Replace:=wdReplaceAl - one letter short of wdReplaceAll. This form
'                                 has no Option Explicit, so that name was an empty Variant, which is
'                                 0, which is wdReplaceNone: the Find ran, selected its first match,
'                                 and replaced nothing. The NIMAS half beside it was spelled
'                                 correctly, which is why only one of the two buttons was dead. Found
'                                 in the undo survey of 9/16/2026, not by a report - nothing raises an
'                                 error here and nothing is written to the log, so a transcriber could
'                                 only see that no $pg tags appeared. THE DAISY PATTERN ITSELF HAS
'                                 THEREFORE NEVER RUN and is unproven against a real DAISY file.
' Version: 1.7  Date: 11/30/2023 - optomized - removed prodnote which deleted some NIMAS book pages
' Version: 1.6  Date: 3/24/2023 - isolated $pg+page number into its own paragraph
' Version: 1.5  Date: 11/18/2021 - added remove producton notes
' Version: 1.4  Date: 9/24/2018 - added Application.Run MacroName:="Sh_Color_Dollar_PG_Red"
' Version: 1.3  Date: 2/3/2018 -fixed end of macro message
'
Private Sub CmdCancel_Click()
    End
End Sub

Private Sub CmdOkay_Click()

    '-----------------------------------------------------------------------------
    ' for DAISY files
    '-----------------------------------------------------------------------------

     If IsDAISY = True Then
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "\>([0-9A-z]{1,})\</pagenum\>"
            .Replacement.Text = ">$pg\1<p></p></pagenum><p></p>"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
    End If

    '-----------------------------------------------------------------------------
    ' for NIMAS files
    '-----------------------------------------------------------------------------

    If IsNIMAS = True Then
        ' locate each page number and preceed it with a $pg tag
        ' and a new paragraph code after each page number code
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "\>([0-9A-z-]{1,}\</pagenum\>)"
            .Replacement.Text = ">$pg\1<p></p>^013"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    Application.Run MacroName:="Sh_Color_Dollar_PG_Red"
    
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    
    DN_Tag_Daisy_Nimas_Form.Hide
    MsgBox "Macro Completed", , "Braille Macros (302)"
    Unload DN_Tag_Daisy_Nimas_Form
        
    End Sub   '*** End of  CmdOkay_Click() ***
    
Private Sub userform_terminate() 'red X was clicked
    End
End Sub
Private Sub UserForm_Initialize()

    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

