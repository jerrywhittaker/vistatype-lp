VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_MathSymbolsForm 
   Caption         =   "Math Symbols"
   ClientHeight    =   9870.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12735
   OleObjectBlob   =   "Dx_MathSymbolsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_MathSymbolsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Version: 1.1  Date: 8/3/2026 - removed 26 calls to "Dx_ShowMathSymbolsAgain", a forgotten
'                                development relic (Jerry, 8/3/2026). The macro existed nowhere,
'                                so every symbol button raised "the macro cannot be found" after
'                                typing its character. Found by tools/lib/check_form_calls.py.


'+++++++ Begin Comparisons ++++++++++++++++++++

Private Sub ComparisionsAlmostEqualTo_Click()
    Selection.TypeText ChrW(&H2248)
End Sub

Private Sub ComparisonsApproxEqual_Click()
    Selection.TypeText ChrW(&H2245)
End Sub

Private Sub ComparisonsEqual_Click()
    Selection.TypeText ChrW(&H3D)
End Sub

Private Sub ComparisonsGreaterThan_Click()
    Selection.TypeText ChrW(&H3E)
End Sub

Private Sub ComparisonsGreaterThanOrEqualTo_Click()
    Selection.TypeText ChrW(&H2265)
End Sub

Private Sub ComparisonsIdenticalTo_Click()
    Selection.TypeText ChrW(&H2261)
End Sub

Private Sub ComparisonsLessThan_Click()
    Selection.TypeText ChrW(&H3C)
End Sub

Private Sub ComparisonsLessThanOrEqualTo_Click()
    Selection.TypeText ChrW(&H2264)
End Sub

Private Sub ComparisonsNotEqualTo_Click()
    Selection.TypeText ChrW(&H2260)
End Sub

Private Sub ComparisonsProportion_Click()
    Selection.TypeText ChrW(&H2237)
End Sub

Private Sub ComparisonsRatio_Click()
    Selection.TypeText ChrW(&H2236)
End Sub

Private Sub ComparisonsTilde_Click()
    Selection.TypeText ChrW(&H7E)
End Sub

'+++++++ Begin Indicators ++++++++++++++++++++

Private Sub IndicatorsDecimalPoint_Click()
    Selection.Style = ActiveDocument.Styles("ExactTranslation")
    Selection.TypeText ChrW(&H2E)
    Selection.Font.Reset
End Sub

Private Sub IndicatorsMathComma_Click()
    Selection.Style = ActiveDocument.Styles("ExactTranslation")
    Selection.TypeText ChrW(&H2C)
    Selection.Font.Reset
End Sub

Private Sub IndicatorsNumeric_Click()
    Selection.Style = ActiveDocument.Styles("ExactTranslation")
    Selection.TypeText ChrW(&H23)
    Selection.Font.Reset
End Sub

Private Sub IndicatorsPunctuation_Click()
    Selection.Style = ActiveDocument.Styles("ExactTranslation")
    Selection.TypeText ChrW(&H5F)
    Selection.Font.Reset
End Sub

'+++++++ Begin Operations ++++++++++++++++++++

Private Sub OperationsBullet_Click()
    Selection.TypeText ChrW(&H2022)
End Sub

Private Sub OperationsDivisionSlash_Click()
    Selection.TypeText ChrW(&H2F)
End Sub

Private Sub OperationsDivison_Click()
    Selection.TypeText ChrW(&HF7)
End Sub

Private Sub OperationsMinus_Click()
    Selection.TypeText ChrW(&H2212)
End Sub

Private Sub OperationsMinusOrPlus_Click()
    Selection.TypeText ChrW(&H2213)
End Sub

Private Sub OperationsMinusPlus_Click()
    Selection.TypeText ChrW(&H2D)
    Selection.TypeText ChrW(&H2B)
End Sub

Private Sub OperationsMultiplication_Click()
    Selection.TypeText ChrW(&HD7)
End Sub

Private Sub OperationsPlus_Click()
    Selection.TypeText ChrW(&H2B)
End Sub


Private Sub OperationsPlusMinus_Click()
    Selection.TypeText ChrW(&H2B)
    Selection.TypeText ChrW(&H2D)
End Sub

Private Sub OperationsPlusOrMinus_Click()
    Selection.TypeText ChrW(&HB1)
End Sub

Private Sub ExitButton_Click()
    Unload Dx_MathSymbolsAgain
    Unload Me
End Sub

Sub userform_terminate() 'red X was clicked
    Unload Dx_MathSymbolsAgain
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

