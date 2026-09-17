VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_Type_Dashes_Form 
   Caption         =   "Type Dashes/Primes/Fractions (335)"
   ClientHeight    =   8088
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11520
   OleObjectBlob   =   "Dx_Type_Dashes_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_Type_Dashes_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Dx_Type_Dashes_Form (Code)

' Version: 1.6  Date: 8/29/2026 - Encode Fractions leaves the DBT codes HIDDEN AND PLUM, as it always
'                                meant to (Jerry). It hid them in three passes and then un-hid them
'                                all again on the last two statements - see CmdEncode_Click. Also
'                                reworded the instruction above the button
' Version: 1.5  Date: 3/5/2023 - added encode selected fraction with DBT Codes - CmdEncode_Click
' Version: 1.4  Date: 2/24/2023 - fixed enter fraction button to change to automatic color after before entering space at end
' Version: 1.3  Date: 10/27/2021 - added filter for Numerator and Denominator to allow only numbers, commas, and decimal point
' Version: 1.3  Date: 10/23/2021 - added hyphen - fixed diplayed unicode values
' Version: 1.2  Date: 6/18/2018

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdEm_Click()
    Selection.TypeText Text:="2014"
    Selection.ToggleCharacterCode
    Unload Me
End Sub

Private Sub CmdEn_Click()
    Selection.TypeText Text:="2013"
    Selection.ToggleCharacterCode
    Unload Me
End Sub

Private Sub CmdEncode_Click()
'
' Author: Jerry Whittaker
'
' Version: 2.0  Date: 8/29/2026 - the work moved to Dx_Encode_Fractions_With_DBT_Codes in
'                                 LPandBrlMacros, where it now skips dates. This handler was a
'                                 hundred and fifteen lines of Find and Replace, which is not how
'                                 anything else here is arranged - forms call macros. Keeping the
'                                 layout of this form editable by hand without colliding with that
'                                 logic is the other half of the reason.
' Version: 1.0  Date: 3/5/2023

    Application.Run MacroName:="Dx_Encode_Fractions_With_DBT_Codes"
    Unload Me

End Sub

Private Sub CmdLong_Click()
    Selection.TypeText Text:="2015"
    Selection.ToggleCharacterCode
    Unload Me
End Sub

Private Sub CmdMinus_Click()
    Selection.TypeText Text:="2012"
    Selection.ToggleCharacterCode
    Unload Me
End Sub

Private Sub CmdSingle_Click()
    Selection.TypeText Text:="2032"
    Selection.ToggleCharacterCode
    Unload Me
End Sub

Private Sub CmdDouble_Click()
    Selection.TypeText Text:="2033"
    Selection.ToggleCharacterCode
    Unload Me
End Sub

Private Sub Button_Five_Eighths_Click()
    Selection.TypeText Text:="215D"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_Five_Sixths_Click()
    Selection.TypeText Text:="215A"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_Four_Fifths_Click()
    Selection.TypeText Text:="2158"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_One_Eighth_Click()
    Selection.TypeText Text:="215B"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_One_Fifth_Click()
    Selection.TypeText Text:="2155"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_One_Fourth_Click()
    Selection.TypeText Text:="00BC"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_One_Half_Click()
    Selection.TypeText Text:="00BD"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_One_Ninth_Click()
    Selection.TypeText Text:="2151"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_One_Seventh_Click()
    Selection.TypeText Text:="2150"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_One_Sixth_Click()
    Selection.TypeText Text:="2159"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_One_Thenth_Click()
    Selection.TypeText Text:="2152"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_One_Third_Click()
    Selection.TypeText Text:="2153"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_Seven_Eighths_Click()
    Selection.TypeText Text:="215E"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_Three_Eighths_Click()
    Selection.TypeText Text:="215C"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_Three_Fifths_Click()
    Selection.TypeText Text:="2157"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_Three_Fourths_Click()
    Selection.TypeText Text:="00BE"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_Two_Fifths_Click()
    Selection.TypeText Text:="2156"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Button_Two_Thirds_Click()
    Selection.TypeText Text:="2154"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Hyphen_Click()
    Selection.TypeText Text:="2010"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub NonBreakHyphen_Click()
    Selection.TypeText Text:="2011"
    Selection.ToggleCharacterCode
    Selection.TypeText Text:=" "
    Unload Me
End Sub

Private Sub Numerator_Change()
    'From: https://stackoverflow.com/questions/30451176/vba-userform-best-way-to-use-a-textbox-as-a-decimal-numerical-input-independe
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack, vbKeyClear, vbKeyDelete, _
        vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
            '~~> Check to see if there is already a decimal
            If KeyAscii = 46 Then If InStr(1, Numerator.Text, ".") Then KeyAscii = 0
        Case Else
            KeyAscii = 0
            'Beep
    End Select
End Sub

Private Sub Denominator_Change()
    'From: https://stackoverflow.com/questions/30451176/vba-userform-best-way-to-use-a-textbox-as-a-decimal-numerical-input-independe
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack, vbKeyClear, vbKeyDelete, _
        vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
            '~~> Check to see if there is already a decimal
            If KeyAscii = 46 Then If InStr(1, Denominato.Text, ".") Then KeyAscii = 0
        Case Else
            KeyAscii = 0
            'Beep
    End Select
End Sub

Private Sub EnterFractionButton_Click()
    If Val(Numerator) = 0 Or Val(Denominator) = 0 Then
        MsgBox "Invalid Fraction Values", , "Braille Macros (304)"
        Exit Sub
    Else
        With Selection.Font
            .Hidden = True
            .Color = wdColorPlum
            Selection.TypeText Text:="[[*fs*]]"
            Selection.InsertAfter Numerator
            .Hidden = False
            .Color = wdColorAutomatic
            Selection.MoveRight Unit:=wdCharacter, count:=1
            .Hidden = True
            .Color = wdColorPlum
            Selection.TypeText Text:="[[*fl*]]"
            Selection.InsertAfter Denominator
            .Hidden = False
            .Color = wdColorAutomatic
            Selection.MoveRight Unit:=wdCharacter, count:=1
            .Hidden = True
            .Color = wdColorPlum
            Selection.TypeText Text:="[[*fe*]]"
            .Hidden = False
            .Color = wdColorAutomatic
            Selection.TypeText Text:=" "
        End With
        ActiveWindow.View.ShowAll = True
    End If
    Unload Me
End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

' *** end of ' Dx_Type_Dashes_Form (Code) ***

