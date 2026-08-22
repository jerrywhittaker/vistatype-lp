VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LP_Attach_An_Lp_Template_Form 
   Caption         =   "Attach LP Template & Select Output Media"
   ClientHeight    =   9132.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12360
   OleObjectBlob   =   "LP_Attach_An_Lp_Template_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LP_Attach_An_Lp_Template_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



' LP_Attach_An_Lp_Template_Form
'
' Version: 6.2  Date: 8/22/2026 - the Typeface choice is BACK, with a face that fixes what killed
'                                the last one. FontChoiceFrame / FontTahoma / FontSans, and the
'                                bundled face is VistaTypeLP Sans: Greek complete, the phonetic
'                                characters including U+025E, and a slashed zero. It is the
'                                default for a new book, grayed out and labelled when it is not
'                                installed. A book already set in the dropped VistaTypeLP Legible
'                                still keeps it - that test runs FIRST in AttachOkay_Click and
'                                outranks these buttons, and Initialize grays the box out and
'                                says so, rather than showing a choice that would be ignored
' Version: 6.1  Date: 8/20/2026 - the Typeface choice is GONE, and so is the typeface. Jerry
'                                dropped the bundled VistaTypeLP Legible from the product: its
'                                character set is Latin, and Latin proper, mathematics, the IPA
'                                and the Greek in medical work all fall back to another face at
'                                another size, silently. A large print book is Tahoma again, as
'                                it was for years before 8/8/2026, and nothing is asked.
'                                FontChoiceFrame, FontTahoma and FontLegible are off the form,
'                                with their two Click handlers and the Initialize block.
'                                Lp_Doc_Font_At_Open STAYS and now does the whole job: a book
'                                already set in the dropped face keeps it, so pressing Attach on
'                                one of the books already produced in it cannot rewrite it to
'                                Tahoma and move every page break. The missing-font warning goes
'                                with the choice - there is nothing left to warn about, since the
'                                only two outcomes are "keep what the book has" and "Tahoma",
'                                and every machine has Tahoma
' Version: 6.0  Date: 8/8/2026 - Jerry took CustomOrientationFrame off the form and moved the two
'                                orientation buttons up into CustomizeFrame, where they are still
'                                the only option buttons and so still switch each other. The three
'                                CustomOrientationFrame.Enabled lines went with it. They were
'                                redundant even before: every one of them sat beside lines that
'                                already enable or disable CustomOrientPortrait and
'                                CustomOrientLandscape one at a time. Left in, each raised run-time
'                                error 424 - this form has no Option Explicit, so a name that is no
'                                longer a control is simply an empty Variant, and .Enabled on it
'                                fails. UserForm_Initialize holds one, so it fired the moment the
'                                ribbon's Attach LP Template button was clicked
' Version: 5.9  Date: 8/8/2026 - added the Typeface choice beside the point size: Tahoma, or the
'                                bundled VistaTypeLP Legible (the default). Legible is grayed out,
'                                and says so on the button, when it is not installed - Word
'                                substitutes a missing font silently and at the wrong size.
'                                Re-attaching keeps the book's own typeface, and asks first if
'                                that typeface is missing and the book would be rewritten to
'                                Tahoma. Needs the FontChoiceFrame / FontTahoma / FontLegible
'                                controls on the form
' Version: 5.8  Date: 7/7/2026 - changes to non-modal message text - added close navigation pane
' Version: 5.7  Date: 6/22/2026 - set return key to default to "Okay" Button - fixes to labels on form
' Version: 5.6  Date: 5/21/2025 - minor fixes to directions and small visual bugs
' Version: 5.5  Date: 4/15/2025 - added routine to ask user if pictures should be set to inline (on on new Lp Docs)
' Version: 4.4  Date: 3/21/2024 - Fixed bug that showed erronous margins after customized was checked followed by another media selection
' Version: 4.3  Date: 10/23/2021 - removed superflous code calls to   "Lp_Set_TOC_and_Print_Page_Num_Tab_Stops" and "Lp_Normalize_Styles"
' Version: 4.2  Date: 10/22/2021 - added bypass for Fix Common File Error and deleteing multipal para marks when attaching to file which has obsolete LP template Attached
' Version: 4.1  Date: 10/20/2021 - added questions about Fix Common File Error and deleteing multipal para marks when template has been applied to non-Lp doc and character count is > 10
' Version: 4.0  Date: 8/25/2021 - made turning off the customize checkbox revert to the selected media params - height and width flipped on changes of orientation
' Version: 3.9  Date: 6/16/2021 - changed "gutter size" to "binding width" on user form and in the code
' Version: 3.8  Date: 4/2/2021 - added calls to routines to rebuild the "LargePrintTemplate.dotx" before attaching to document in "Private Sub AttachOkay_Click()"
' Version: 3.7  Date: 3/19/2021 - allow up to 2 inches for the left margin - change 12.9 iPad to Portrait Orientation
' Version: 3.6  Date: 2/12/2021 - changed on-screen label for gutter plus margin
' Version: 3.5  Date: 12/14/2020 - turned off customize checkbox with each predefined media selection
' Version: 3.4  Date: 11/22/2020 - complete rewrite with custom page settings

Dim IsPaper As Boolean
' The typeface this document was already using when the dialog opened, or "" if it was not a
' large print document. UserForm_Initialize records it; AttachOkay_Click uses it to decide
' whether this attach may change the book's typeface at all.
Dim Lp_Doc_Font_At_Open As String
Dim Hold_Orientation As String
Dim Hold_HeightValue As String
Dim Hold_WidthValue As String
Dim Hold_LMarginSizeValue As String
Dim Hold_RMarginSizeValue As String
Dim Hold_TMarginSizeValue As String
Dim Hold_BMarginSizeValue As String

Private Sub CancelButton_Click()
    Application.ScreenUpdating = True ' Turn screen updating on
    Unload Me
End Sub

Private Sub BindingWidthValue_Change()
    FinalGutterSize = Round(Val(LMarginSizeValue), 2) + Round(Val(BindingWidthValue), 2)
    If FinalGutterSize = 1 Then
        InchWord = " inch "
    Else
        InchWord = " inches "
    End If
    'FinalGutterSizeLabel.Caption = "There is a gutter with of " + Str(FinalGutterSize) + InchWord + "on each page."
    FinalGutterSizeLabel.Caption = "At the gutter, the left and right margins are " + Str(FinalGutterSize) + InchWord + "from the edge of the paper"
End Sub

Private Sub LMarginSizeValue_Change()
    If MirroredCheckBox = True Then
        FinalGutterSize = Round(Val(LMarginSizeValue), 2) + Round(Val(BindingWidthValue), 2)
        If FinalGutterSize = 1 Then
            InchWord = " inch "
        Else
            InchWord = " inches "
        End If
        'FinalGutterSizeLabel.Caption = "There is a gutter width of" + Str(FinalGutterSize) + InchWord + "on each page."
        FinalGutterSizeLabel.Caption = "At the gutter, the left and right margins are " + Str(FinalGutterSize) + InchWord + "from the edge of the paper"
    End If
End Sub

Private Sub MirroredCheckBox_Click()
    DocPicLandscape.Visible = False
    DocPicPortrait.Visible = False
    
    If MirroredCheckBox = True Then
        GutterSizeLabel.Enabled = True
        BindingWidthValue.Enabled = True
        BindingWidthValue.Value = Round((Round(Val(LMarginSizeValue), 2) / 2), 2)
        PMM = True
        
        If CustomOrientLandscape = True Then
            LandscapePicWithGutter.Visible = True
        Else
            PortraitPicWithGutter.Visible = True
        End If
        
        FinalGutterSize = Round(Val(LMarginSizeValue), 2) + Round(Val(BindingWidthValue), 2)
        If FinalGutterSize = 1 Then
            InchWord = " inch "
        Else
            InchWord = " inches "
        End If
        
        'FinalGutterSizeLabel.Caption = "There is a gutter width of" + Str(FinalGutterSize) + InchWord + "on each page."
        FinalGutterSizeLabel.Caption = "At the gutter, the left and right margins are " + Str(FinalGutterSize) + InchWord + "from the edge of the paper"
        FinalGutterSizeLabel.Visible = True
    Else
        PMM = False
        FinalGutterSizeLabel.Visible = False
        GutterSizeLabel.Enabled = False
        BindingWidthValue.Enabled = False
        BindingWidthValue.Value = ""
        LandscapePicWithGutter.Visible = False
        PortraitPicWithGutter.Visible = False
        If CustomOrientLandscape = True Then
            DocPicLandscape.Visible = True
        Else
            DocPicPortrait.Visible = True
        End If
    End If
End Sub

Private Sub userform_terminate() 'red X was clicked
    Application.ScreenUpdating = True ' Turn screen updating on
    Unload Me
End Sub

Private Sub FontTahoma_Click()
        Lp_Base_Font_Name = LP_FONT_TAHOMA
End Sub

Private Sub FontSans_Click()
        Lp_Base_Font_Name = LP_FONT_SANS
End Sub

Private Sub AttachOkay_Click()
    
    ' ** Validate Custom Settings **
    
    If CustomizeCheckBox = True Then
        ' ** Height, Width & Margins Value Range Checks**
        
        If (Round(Val(CustomHeightValue), 2) < 0.1 Or Round(Val(CustomHeightValue), 2) > 22) Or IsNumeric(CustomHeightValue) = False Then
            MsgBox "The size range For paper Or screen height Is .1 To 22", , "VistaType LP (189)"
            CustomHeightValue = PPH
            Exit Sub
        End If
        
        If (Round(Val(CustomWidthValue), 2) < 0.1 Or Round(Val(CustomWidthValue), 2) > 22) Or IsNumeric(CustomWidthValue) = False Then
            MsgBox "The size range For paper Or screen width Is .1 To 22", , "VistaType LP (190)"
            CustomWidthValue = PPW
            Exit Sub
        End If
        
        If Round(Val(LMarginSizeValue), 2) < 0.1 Or Round(Val(LMarginSizeValue), 2) > 2 Or IsNumeric(LMarginSizeValue) = False Then
            MsgBox "The size range For the left margin Is .1 To 2 inches.", , "VistaType LP (191)"
            LMarginSizeValue = PLM
            Exit Sub
        End If
        
        If Round(Val(RMarginSizeValue), 2) < 0.1 Or Round(Val(RMarginSizeValue), 2) > 1 Or IsNumeric(RMarginSizeValue) = False Then
            MsgBox "The size range For the right margin Is .1 To 1 inch.", , "VistaType LP (192)"
            RMarginSizeValue = PRM
            Exit Sub
        End If
        
        If Round(Val(TMarginSizeValue), 2) < 0.1 Or Round(Val(TMarginSizeValue), 2) > 1 Or IsNumeric(TMarginSizeValue) = False Then
            MsgBox "The size range For the top margin Is .1 To 1 inch.", , "VistaType LP (193)"
            TMarginSizeValue = PTM
            Exit Sub
        End If
        
        If Round(Val(BMarginSizeValue), 2) < 0.1 Or Round(Val(BMarginSizeValue), 2) > 1 Or IsNumeric(BMarginSizeValue) = False Then
            MsgBox "The size range For the bottom margin Is .1 To 1 inch.", , "VistaType LP (194)"
            BMarginSizeValue = PBM
            Exit Sub
        End If
        
        ' ** Logic Validations **
        If Round(Val(CustomWidthValue), 2) - (Val(LMarginSizeValue) + Val(RMarginSizeValue)) < 0.1 Then
            MsgBox "The sum of the left And right margins subtracted from the paper Or screen width Is less than .1 inch", , "VistaType LP (195)"
            CustomWidthValue = PPW
            Exit Sub
        End If
        
        If Round(Val(CustomHeightValue), 2) - (Val(TMarginSizeValue) + Val(BMarginSizeValue)) < 0.5 Then
            MsgBox "The sum of the top And bottom margins subracted from the paper Or screen height Is less than .5 inch", , "VistaType LP (196)"
            CustomHeightValue = PPH
            Exit Sub
        End If
        
        If PPO = "L" And Round(Val(CustomHeightValue), 2) > Round(Val(CustomWidthValue), 2) Then
            MsgBox "The custom orientation Is Set For landscape but the custom height Is greater than the custom width. " & _
                   "In landscapt mode, width must be greater than height", , "VistaType LP (197)"
            Exit Sub
        End If
        
        If PPO = "P" And Round(Val(CustomWidthValue), 2) > Round(Val(CustomHeightValue), 2) Then
            MsgBox "The custom orientation Is Set For portrait but the custom width Is greater than the custom height. " & _
                   "In portrait mode, the height must be greater than width.", , "VistaType LP (198)"
            Exit Sub
        End If
        
        If MirroredCheckBox = True And (Round(Val(BindingWidthValue), 2) < 0.1 Or Round(Val(BindingWidthValue), 2) > 1 Or IsNumeric(BindingWidthValue) = False) Then
            MsgBox "The size range For the Binding Margin Is .1 To 1 inch.", , "VistaType LP (199)"
            BindingWidthValue = LMarginSizeValue
            Exit Sub
        End If
        
    End If        ' If CustomizeCheckBox = True
    
    ' ** end of  Validate Custom Settings **
    
    ' ** move form values to public variables **
    PPW = Str(Round(Val(CustomWidthValue), 2))
    PPH = Str(Round(Val(CustomHeightValue), 2))
    PTM = Str(Round(Val(TMarginSizeValue), 2))
    PBM = Str(Round(Val(BMarginSizeValue), 2))
    PLM = Str(Round(Val(LMarginSizeValue), 2))
    PRM = Str(Round(Val(RMarginSizeValue), 2))
    TOCTabSetting = Str(Val(PPW) - (Val(PLM) + Val(PRM)))
    
    If MirroredCheckBox = True Then
        PMM = True
        PPG = Str(Round(Val(BindingWidthValue), 2))
    Else
        PMM = False
        PPG = "0"
    End If

    ' ***** Typeface *****
    ' This is the one that has to happen. Lp_Attach_The_Template runs LATER and ASYNCHRONOUSLY,
    ' through Sh_BridgeTargetMacro and Application.OnTime below, and it unloads this form before
    ' it does anything - so the controls are gone by then and the public is the only carrier.
    '
    ' A large print book is Tahoma. It was for years, it was a choice for twelve days from
    ' 8/8/2026, and from 8/20/2026 it is not a choice again - the bundled VistaTypeLP Legible was
    ' dropped because it has no Greek, no IPA and not enough mathematics, and Word fills a missing
    ' character from another face at another size without a word.
    '
    ' The ONE exception, and the reason this is not simply an assignment: a book ALREADY set in
    ' that face keeps it. Rewriting one to Tahoma would move every page break in a book that may
    ' already be printed and in a reader's hands, and nothing on screen would say it had happened
    ' - the pages would simply be different. Those books also carry the face embedded in
    ' themselves (Lp_Attach_The_Template's EmbedTrueTypeFonts line), so they still set correctly
    ' even though the installer now takes that font off the machine.
    '
    ' The test names the face outright rather than carrying forward whatever it finds. An LP
    ' document whose Normal style has drifted to Calibri is a document attaching is supposed to
    ' REPAIR, and blanket carry-forward would preserve the fault instead.
    '
    ' The two constants come from LPandBrlMacros so the name lives in ONE place. A form may use
    ' them BECAUSE they are uppercase LP_: tools/lib/check_form_calls.py looks for Lp_-prefixed
    ' tokens and its pattern is case-sensitive, so LP_FONT_LEGACY_LEGIBLE is invisible to it,
    ' while an Lp_-prefixed constant would be read as a call to a macro that does not exist and
    ' would fail the build. That is why they were declared uppercase in the first place.
    ' The legacy test stays FIRST, and deliberately outranks the buttons. A book already set in
    ' the dropped VistaTypeLP Legible keeps it whatever this dialog shows: rewriting one moves
    ' every page break in a book that may already be in a reader's hands, and nothing on screen
    ' afterwards would say why the pages changed. UserForm_Initialize grays the choice out and
    ' says so on the box, so no one is picking a face here that is then ignored.
    If UCase(Lp_Doc_Font_At_Open) = UCase(LP_FONT_LEGACY_LEGIBLE) Then
        Lp_Base_Font_Name = Lp_Doc_Font_At_Open
    ElseIf FontSans.Value = True Then
        Lp_Base_Font_Name = LP_FONT_SANS
    Else
        Lp_Base_Font_Name = LP_FONT_TAHOMA
    End If
    ' ***** end Typeface *****
    ' ** end of move form values to public variables **

   LP_Attach_An_Lp_Template_Form.Hide

'+++++++++++++++++++ Call Non-Modal Message +++++++++++++++++++++++

   Dim msgBody As String
    msgBody = "Fixing common file errors for non-LP files, attaching the large print template, setting paper or screen size, " & _
              "attaching the large print template, and adjusting font sizes and border weights." & _
              vbCrLf & vbCrLf & "Do not use mouse or keyboard in this Word window or any other Word document! " & _
              "Running applications other than Word is acceptable." & _
              vbCrLf & vbCrLf & "                         Wait for the BEEP!"

    Call Sh_ShowNonModalMessage("VistaType LP is Working", msgBody)

    DoEvents
    
    Sh_SetBarVisible "Styles", False
    Sh_SetBarVisible "Navigation", False
    
    Sh_BridgeTargetMacro = "Lp_Attach_The_Template" ' there must be a complete end of this macro in order to display the non-modal message
    Application.OnTime Now, "Sh_StartSpinnerBridge" ' after the non-modal message is displayed the "Sh_StartSpinnerBridge" will initiate the
                                                    ' second portion of the LP attachment "Lp_Attach_The_Template" will begin
    Exit Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

End Sub

Private Sub CustomOrientLandscape_Click()

    If Hold_Orientation = PPO Then ' PPO = Print Page Orientation
        CustomHeightValue = Hold_WidthValue
        CustomWidthValue = Hold_HeightValue
    Else
        CustomHeightValue = Hold_HeightValue
        CustomWidthValue = Hold_WidthValue
    End If
    
    PPO = "L" 'new setting
    
    If IsPaper Then
        DocPicPortrait.Visible = False
        DocPicLandscape.Visible = True
        iPadPicPortrait.Visible = False
        iPadPicLandscape.Visible = False
        CustomHeightValue = PPW
        CustomWidthValue = PPH
        If MirroredCheckBox = True Then
            DocPicLandscape.Visible = False
            LandscapePicWithGutter.Visible = True
            PortraitPicWithGutter.Visible = False
        End If
    Else
        DocPicPortrait.Visible = False
        DocPicLandscape.Visible = False
        iPadPicPortrait.Visible = False
        iPadPicLandscape.Visible = True
    End If
End Sub

Private Sub CustomOrientPortrait_Click()

    If Hold_Orientation = PPO Then
        CustomHeightValue = Hold_WidthValue
        CustomWidthValue = Hold_HeightValue
    Else
        CustomHeightValue = Hold_HeightValue
        CustomWidthValue = Hold_WidthValue
    End If
     
     PPO = "P" 'new setting
     
    If IsPaper Then
        DocPicPortrait.Visible = True
        DocPicLandscape.Visible = False
        iPadPicPortrait.Visible = False
        iPadPicLandscape.Visible = False
        CustomHeightValue = PPH
        CustomWidthValue = PPW
        If MirroredCheckBox = True Then
            DocPicPortrait.Visible = False
            PortraitPicWithGutter.Visible = True
            LandscapePicWithGutter.Visible = False
        End If
    Else
        DocPicPortrait.Visible = False
        DocPicLandscape.Visible = False
        iPadPicPortrait.Visible = True
        iPadPicLandscape.Visible = False
    End If
End Sub

Private Sub CustomizeCheckBox_Click()

    If CustomizeCheckBox = True Then
        ' move to temp hold in case customize is turned off
        Hold_WidthValue = CustomWidthValue
        Hold_HeightValue = CustomHeightValue
        Hold_LMarginSizeValue = LMarginSizeValue
        Hold_RMarginSizeValue = RMarginSizeValue
        Hold_TMarginSizeValue = TMarginSizeValue
        Hold_BMarginSizeValue = BMarginSizeValue
        Hold_Orientation = PPO
        OrientationLabel.Enabled = True
        ' *** Margins ***
        MarginSizeLabel.Enabled = True
        TMarginSizeValue.Enabled = True
        BMarginSizeValue.Enabled = True
        LMarginSizeValue.Enabled = True
        RMarginSizeValue.Enabled = True
        
        TMarginSizeLabel.Enabled = True
        BMarginSizeLabel.Enabled = True
        LMarginSizeLabel.Enabled = True
        RMarginSizeLabel.Enabled = True
        
        TMarginSizeValue = PTM
        BMarginSizeValue = PBM
        LMarginSizeValue = PLM
        RMarginSizeValue = PRM
        '*** Custom Orientation ***
        CustomOrientLandscape.Enabled = True
        CustomOrientPortrait.Enabled = True

         If PPO = "L" Then
             CustomOrientLandscape = True
         Else
             CustomOrientPortrait = True
         End If

        '*** Width and Height Values ***
        CustomWidthLabel.Enabled = True
        CustomHeightLabel.Enabled = True
        
        CustomHeightValue.Enabled = True
        CustomWidthValue.Enabled = True
        
        If IsPaper Then
            MirroredCheckBox.Enabled = True
        Else
            MirroredCheckBox.Value = False
            MirroredCheckBox.Enabled = False
        End If
        
    Else  ' customize is unchecked - custom screen area is turned off *************************************************

        PPO = Hold_Orientation
        PPH = Hold_HeightValue
        PPW = Hold_WidthValue
        PLM = Hold_LMarginSizeValue
        PRM = Hold_RMarginSizeValue
        PTM = Hold_TMarginSizeValue
        PBM = Hold_BMarginSizeValue

        CustomHeightValue = PPH
        CustomWidthValue = PPW
        
        LMarginSizeValue = PLM
        RMarginSizeValue = PRM
        TMarginSizeValue = PTM
        BMarginSizeValue = PBM
    
        If PPO = "P" Then
            CustomOrientPortrait = True
            CustomOrientLandscape.Value = False
            CustomOrientPortrait.Value = True
            CustomOrientPortrait.Enabled = True
        Else
            CustomOrientLandscape = True
            CustomOrientLandscape.Value = True
            CustomOrientPortrait.Value = False
        End If
       
        CustomOrientPortrait.Enabled = False
        CustomWidthLabel.Enabled = False
        CustomHeightLabel.Enabled = False
        OrientationLabel.Enabled = False
        
        MarginSizeLabel.Enabled = False
        TMarginSizeLabel.Enabled = False
        BMarginSizeLabel.Enabled = False
        LMarginSizeLabel.Enabled = False
        RMarginSizeLabel.Enabled = False
        TMarginSizeValue.Enabled = False
        BMarginSizeValue.Enabled = False
        LMarginSizeValue.Enabled = False
        RMarginSizeValue.Enabled = False
        CustomHeightValue.Enabled = False
        CustomWidthValue.Enabled = False
     
        CustomOrientLandscape.Enabled = False
        
        MirroredCheckBox.Enabled = False
        MirroredCheckBox.Value = False
        LandscapePicWithGutter.Visible = False
        PortraitPicWithGutter.Visible = False
    End If
End Sub   '*** end of CustomizeCheckBox_Click ****

Private Sub MarginHalfInch_Click()
    PTM = ".5"
    PBM = ".5"
    PLM = ".5"
    PRM = ".5"
    PPW = "8.5"
    PPH = "11"
    PPO = "P"
    Hold_WidthValue = PPW
    Hold_HeightValue = PPH
    CustomizeCheckBox.Value = False
    IsPaper = True
    DocPicPortrait.Visible = True
    DocPicLandscape.Visible = False
    iPadPicPortrait.Visible = False
    iPadPicLandscape.Visible = False
    CustomWidthValue = PPW
    CustomHeightValue = PPH
    TMarginSizeValue = PTM
    BMarginSizeValue = PBM
    LMarginSizeValue = PLM
    RMarginSizeValue = PRM
    DM = "Paper"
    If CustomizeCheckBox = True Then
        MirroredCheckBox.Enabled = True
        MirroredCheckBox.Value = False
    Else 'restores values to this preset - 3/21/2024 fix
        TMarginSizeValue = ".5"
        BMarginSizeValue = ".5"
        LMarginSizeValue = ".5"
        RMarginSizeValue = ".5"
    End If
    CustomOrientPortrait = True
End Sub

Private Sub MarginTheeFourthsInch_Click()
    PTM = ".75"
    PBM = ".75"
    PLM = ".75"
    PRM = ".75"
    PPW = "8.5"
    PPH = "11"
    PPO = "P"
    Hold_WidthValue = PPW
    Hold_HeightValue = PPH
    CustomizeCheckBox.Value = False
    IsPaper = True
    DocPicPortrait.Visible = True
    DocPicLandscape.Visible = False
    iPadPicPortrait.Visible = False
    iPadPicLandscape.Visible = False
    CustomWidthValue = PPW
    CustomHeightValue = PPH
    TMarginSizeValue = PTM
    BMarginSizeValue = PBM
    LMarginSizeValue = PLM
    RMarginSizeValue = PRM
    DM = "Paper"
    If CustomizeCheckBox = True Then
        MirroredCheckBox.Enabled = True
        MirroredCheckBox.Value = False
    Else 'restores values to this preset - 3/21/2024 fix
        TMarginSizeValue = ".75"
        BMarginSizeValue = ".75"
        LMarginSizeValue = ".75"
        RMarginSizeValue = ".75"
    End If
    CustomOrientPortrait = True
End Sub

Private Sub MarginOneInch_Click()
    PTM = "1"
    PBM = "1"
    PLM = "1"
    PRM = "1"
    PPW = "8.5"
    PPH = "11"
    PPO = "P"
    Hold_WidthValue = PPW
    Hold_HeightValue = PPH
    CustomizeCheckBox.Value = False
    IsPaper = True
    DocPicPortrait.Visible = True
    DocPicLandscape.Visible = False
    iPadPicPortrait.Visible = False
    iPadPicLandscape.Visible = False
    CustomWidthValue = PPW
    CustomHeightValue = PPH
    TMarginSizeValue = PTM
    BMarginSizeValue = PBM
    LMarginSizeValue = PLM
    RMarginSizeValue = PRM

    DM = "Paper"
    If CustomizeCheckBox = True Then
        MirroredCheckBox.Enabled = True
        MirroredCheckBox.Value = False
    Else 'restores values to this preset - 3/21/2024 fix
        TMarginSizeValue = "1"
        BMarginSizeValue = "1"
        LMarginSizeValue = "1"
        RMarginSizeValue = "1"
    End If
    CustomOrientPortrait = True
End Sub

Private Sub Nine_Point_Seven_Tablet_Click()
    PTM = ".2"
    PBM = ".2"
    PLM = ".2"
    PRM = ".2"
    PPW = "7.76"
    PPH = "5.82"
    PPO = "L"
    Hold_WidthValue = PPW
    Hold_HeightValue = PPH
    CustomizeCheckBox.Value = False
    IsPaper = False
    DocPicPortrait.Visible = False
    DocPicLandscape.Visible = False
    iPadPicPortrait.Visible = False
    iPadPicLandscape.Visible = True
    CustomWidthValue = PPW
    CustomHeightValue = PPH
    TMarginSizeValue = PTM
    BMarginSizeValue = PBM
    LMarginSizeValue = PLM
    RMarginSizeValue = PRM
    DM = "Screen"
    MirroredCheckBox.Enabled = False
    MirroredCheckBox.Value = False
    LandscapePicWithGutter.Visible = False
    PortraitPicWithGutter.Visible = False
    DocPicLandscape.Visible = False
    DocPicPortrait.Visible = False
    CustomOrientLandscape = True
End Sub

Private Sub Ten_Point_One_Tablet_Click()
    PTM = ".2"
    PBM = ".2"
    PLM = ".2"
    PRM = ".2"
    PPW = "8.52"
    PPH = "5.33"
    PPO = "L"
    Hold_WidthValue = PPW
    Hold_HeightValue = PPH
    CustomizeCheckBox.Value = False
    IsPaper = False
    DocPicPortrait.Visible = False
    DocPicLandscape.Visible = False
    iPadPicPortrait.Visible = False
    iPadPicLandscape.Visible = True
    CustomWidthValue = PPW
    CustomHeightValue = PPH
    TMarginSizeValue = PTM
    BMarginSizeValue = PBM
    LMarginSizeValue = PLM
    RMarginSizeValue = PRM
    DM = "Screen"
    MirroredCheckBox.Enabled = False
    MirroredCheckBox.Value = False
    LandscapePicWithGutter.Visible = False
    PortraitPicWithGutter.Visible = False
    DocPicLandscape.Visible = False
    DocPicPortrait.Visible = False
    CustomOrientLandscape = True
End Sub

Private Sub Ten_Point_Two_Tablet_Click()
    PTM = ".2"
    PBM = ".2"
    PLM = ".2"
    PRM = ".2"
    PPW = "8.18"
    PPH = "6.14"
    PPO = "L"
    Hold_WidthValue = PPW
    Hold_HeightValue = PPH
    CustomizeCheckBox.Value = False
    IsPaper = False
    DocPicPortrait.Visible = False
    DocPicLandscape.Visible = False
    iPadPicPortrait.Visible = False
    iPadPicLandscape.Visible = True
    CustomWidthValue = PPW
    CustomHeightValue = PPH
    TMarginSizeValue = PTM
    BMarginSizeValue = PBM
    LMarginSizeValue = PLM
    RMarginSizeValue = PRM
    DM = "Screen"
    MirroredCheckBox.Enabled = False
    MirroredCheckBox.Value = False
    LandscapePicWithGutter.Visible = False
    PortraitPicWithGutter.Visible = False
    DocPicLandscape.Visible = False
    DocPicPortrait.Visible = False
    CustomOrientLandscape = True
End Sub

Private Sub Ten_Point_Five_Tablet_Click()
    PTM = ".2"
    PBM = ".2"
    PLM = ".2"
    PRM = ".2"
    PPW = "8.42"
    PPH = "6.32"
    PPO = "L"
    Hold_WidthValue = PPW
    Hold_HeightValue = PPH
    CustomizeCheckBox.Value = False
    IsPaper = False
    DocPicPortrait.Visible = False
    DocPicLandscape.Visible = False
    iPadPicPortrait.Visible = False
    iPadPicLandscape.Visible = True
    CustomWidthValue = PPW
    CustomHeightValue = PPH
    TMarginSizeValue = PTM
    BMarginSizeValue = PBM
    LMarginSizeValue = PLM
    RMarginSizeValue = PRM
    DM = "Screen"
    MirroredCheckBox.Enabled = False
    MirroredCheckBox.Value = False
    LandscapePicWithGutter.Visible = False
    PortraitPicWithGutter.Visible = False
    DocPicLandscape.Visible = False
    DocPicPortrait.Visible = False
    CustomOrientLandscape = True
End Sub

Private Sub Eleven_Point_Zero_Tablet_Click()
    PTM = ".2"
    PBM = ".2"
    PLM = ".2"
    PRM = ".2"
    PPW = "9.01"
    PPH = "6.3"
    PPO = "L"
    Hold_WidthValue = PPW
    Hold_HeightValue = PPH
    CustomizeCheckBox.Value = False
    IsPaper = False
    DocPicPortrait.Visible = False
    DocPicLandscape.Visible = False
    iPadPicPortrait.Visible = False
    iPadPicLandscape.Visible = True
    CustomWidthValue = PPW
    CustomHeightValue = PPH
    TMarginSizeValue = PTM
    BMarginSizeValue = PBM
    LMarginSizeValue = PLM
    RMarginSizeValue = PRM
    DM = "Screen"
    MirroredCheckBox.Enabled = False
    MirroredCheckBox.Value = False
    LandscapePicWithGutter.Visible = False
    PortraitPicWithGutter.Visible = False
    DocPicLandscape.Visible = False
    DocPicPortrait.Visible = False
    CustomOrientLandscape = True
End Sub

Private Sub Twelve_Point_Nine_Tablet_Click()
    PTM = ".2"
    PBM = ".2"
    PLM = ".2"
    PRM = ".2"
    PPW = "7.76"
    PPH = "10.35"
    PPO = "P"
    
    Hold_WidthValue = PPW
    Hold_HeightValue = PPH
    
    CustomizeCheckBox.Value = False
    IsPaper = False
    DocPicPortrait.Visible = False
    DocPicLandscape.Visible = False
    iPadPicPortrait.Visible = True
    iPadPicLandscape.Visible = False
    CustomWidthValue = PPW
    CustomHeightValue = PPH
    TMarginSizeValue = PTM
    BMarginSizeValue = PBM
    LMarginSizeValue = PLM
    RMarginSizeValue = PRM
    DM = "Screen"
    MirroredCheckBox.Enabled = False
    MirroredCheckBox.Value = False
    LandscapePicWithGutter.Visible = False
    PortraitPicWithGutter.Visible = False
    DocPicLandscape.Visible = False
    DocPicPortrait.Visible = False
    CustomOrientPortrait = True
End Sub

' Lp_Base_Font_Size is a public variable

Private Sub Point14_Click()
        Lp_Base_Font_Size = "14"
End Sub

Private Sub Point16_Click()
        Lp_Base_Font_Size = "16"
End Sub

Private Sub Point18_Click()
        Lp_Base_Font_Size = "18"
End Sub

Private Sub Point20_Click()
        Lp_Base_Font_Size = "20"
End Sub

Private Sub Point22_Click()
        Lp_Base_Font_Size = "22"
End Sub

Private Sub Point24_Click()
        Lp_Base_Font_Size = "24"
End Sub

Private Sub Point26_Click()
        Lp_Base_Font_Size = "26"
End Sub

Private Sub Point28_Click()
        Lp_Base_Font_Size = "28"
End Sub

Private Sub Point30_Click()
        Lp_Base_Font_Size = "30"
End Sub

Private Sub Point32_Click()
        Lp_Base_Font_Size = "32"
End Sub

Private Sub Point34_Click()
        Lp_Base_Font_Size = "34"
End Sub

Private Sub Point36_Click()
        Lp_Base_Font_Size = "36"
End Sub

Private Sub Point38_Click()
        Lp_Base_Font_Size = "38"
End Sub

Private Sub Point40_Click()
        Lp_Base_Font_Size = "40"
End Sub

Private Sub Point42_Click()
        Lp_Base_Font_Size = "42"
End Sub

' FontTahoma_Click and FontLegible_Click stood here until 8/20/2026, when the Typeface choice
' came off the form with the typeface. Lp_Base_Font_Name is now set once, in AttachOkay_Click,
' from the LP_FONT_ constants in LPandBrlMacros - see the note there about why a form may name
' those two and may not name an Lp_-prefixed one.

Private Sub UserForm_Initialize()

    MarginHalfInch.Value = True
    MarginHalfInch.Value = True
    Point18.Enabled = True
    Point18.Value = True
    
    CustomOrientLandscape.Enabled = False
    CustomOrientLandscape.Value = False
    'CustomOrientPortrait.Value = True

    CustomOrientPortrait.Enabled = False

    CustomWidthLabel.Enabled = False
    CustomHeightLabel.Enabled = False
    MarginSizeLabel.Enabled = False
    TMarginSizeLabel.Enabled = False
    BMarginSizeLabel.Enabled = False
    LMarginSizeLabel.Enabled = False
    RMarginSizeLabel.Enabled = False
    TMarginSizeValue.Enabled = False
    BMarginSizeValue.Enabled = False
    LMarginSizeValue.Enabled = False
    RMarginSizeValue.Enabled = False
    CustomHeightValue.Enabled = False
    CustomWidthValue.Enabled = False
    
    Point18.Value = True
    MarginHalfInch.Value = True
    MirroredCheckBox.Enabled = False
    GutterSizeLabel.Enabled = False
    BindingWidthValue.Enabled = False
    
    DocPicPortrait.Visible = True
    DocPicLandscape.Visible = False
    iPadPicPortrait.Visible = False
    iPadPicLandscape.Visible = False
    PortraitPicWithGutter.Visible = False
    LandscapePicWithGutter.Visible = False
    
    FinalGutterSizeLabel.Visible = False

    ' ***** Typeface *****
    ' Nothing on the form asks about the typeface any more - a large print book is Tahoma. All
    ' this does is note what the book in front of us is ALREADY set in, so that AttachOkay_Click
    ' can leave a book made in the dropped VistaTypeLP Legible exactly as it is rather than
    ' repaginating it. See the comment there.
    '
    ' Only asked of a document that is already large print. On anything else the Normal style is
    ' whatever Word or the original author left behind, and attaching is meant to replace it.
    ' VistaTypeLP Sans is the default for a NEW book (Jerry, 8/22/2026), but only when it is
    ' actually on this machine. A font Word cannot find is substituted SILENTLY, and the
    ' substitute has different metrics - so a book that reads 18 point on screen prints at some
    ' other size, which is the exact fault the rescaled face exists to cure. Better to gray the
    ' choice out and say why on the button itself, which needs no extra room on a form that has
    ' none.
    If Sh_Is_Font_Installed(LP_FONT_SANS) Then
        FontSans.Enabled = True
        FontSans.Value = True
        Lp_Base_Font_Name = LP_FONT_SANS
    Else
        FontSans.Enabled = False
        FontTahoma.Value = True
        Lp_Base_Font_Name = LP_FONT_TAHOMA
        ' Appended, not replaced, so the button keeps whatever wording it was given in the
        ' designer. Word only rebuilds its font list at startup, hence the second half.
        FontSans.Caption = FontSans.Caption & "  --  NOT INSTALLED (or Word needs restarting)"
    End If

    Lp_Doc_Font_At_Open = ""
    On Error Resume Next
    ' The LOOSER test - 8/20/2026. A book set in the dropped VistaTypeLP Legible is exactly the
    ' kind that may also be on an older template, and it is the FACE that must not be rewritten
    ' here, whatever template it arrived on. See Lp_Was_Made_As_An_Lp_Book.
    If Lp_Was_Made_As_An_Lp_Book() = True Then
        Lp_Doc_Font_At_Open = ActiveDocument.Styles(wdStyleNormal).Font.Name
    End If
    On Error GoTo 0

    ' An EXISTING book shows the face it is already set in, so the dialog never offers to change
    ' something silently.
    If UCase(Lp_Doc_Font_At_Open) = UCase(LP_FONT_TAHOMA) Then
        FontTahoma.Value = True
        Lp_Base_Font_Name = LP_FONT_TAHOMA
    ElseIf UCase(Lp_Doc_Font_At_Open) = UCase(LP_FONT_SANS) And FontSans.Enabled = True Then
        FontSans.Value = True
        Lp_Base_Font_Name = LP_FONT_SANS
    ElseIf UCase(Lp_Doc_Font_At_Open) = UCase(LP_FONT_LEGACY_LEGIBLE) Then
        ' The book keeps the dropped face - AttachOkay_Click enforces that regardless of what is
        ' selected here. Say so instead of showing a choice that will be ignored.
        FontTahoma.Enabled = False
        FontSans.Enabled = False
        FontChoiceFrame.Caption = "Font Choice  --  this book keeps VistaTypeLP Legible"
    End If
    ' ***** end Typeface *****

    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub
