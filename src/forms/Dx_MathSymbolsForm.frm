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


'+++++++ Begin Comparisons ++++++++++++++++++++

Private Sub ComparisionsAlmostEqualTo_Click()
    Selection.TypeText ChrW(&H2248)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub ComparisonsApproxEqual_Click()
    Selection.TypeText ChrW(&H2245)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub ComparisonsEqual_Click()
    Selection.TypeText ChrW(&H3D)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub ComparisonsGreaterThan_Click()
    Selection.TypeText ChrW(&H3E)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub ComparisonsGreaterThanOrEqualTo_Click()
    Selection.TypeText ChrW(&H2265)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub ComparisonsIdenticalTo_Click()
    Selection.TypeText ChrW(&H2261)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub ComparisonsLessThan_Click()
    Selection.TypeText ChrW(&H3C)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub ComparisonsLessThanOrEqualTo_Click()
    Selection.TypeText ChrW(&H2264)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub ComparisonsNotEqualTo_Click()
    Selection.TypeText ChrW(&H2260)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub ComparisonsProportion_Click()
    Selection.TypeText ChrW(&H2237)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub ComparisonsRatio_Click()
    Selection.TypeText ChrW(&H2236)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub ComparisonsTilde_Click()
    Selection.TypeText ChrW(&H7E)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

'+++++++ Begin Indicators ++++++++++++++++++++

Private Sub IndicatorsDecimalPoint_Click()
    Selection.Style = ActiveDocument.Styles("ExactTranslation")
    Selection.TypeText ChrW(&H2E)
    Selection.Font.Reset
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub IndicatorsMathComma_Click()
    Selection.Style = ActiveDocument.Styles("ExactTranslation")
    Selection.TypeText ChrW(&H2C)
    Selection.Font.Reset
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub IndicatorsNumeric_Click()
    Selection.Style = ActiveDocument.Styles("ExactTranslation")
    Selection.TypeText ChrW(&H23)
    Selection.Font.Reset
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub IndicatorsPunctuation_Click()
    Selection.Style = ActiveDocument.Styles("ExactTranslation")
    Selection.TypeText ChrW(&H5F)
    Selection.Font.Reset
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

'+++++++ Begin Operations ++++++++++++++++++++

Private Sub OperationsBullet_Click()
    Selection.TypeText ChrW(&H2022)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub OperationsDivisionSlash_Click()
    Selection.TypeText ChrW(&H2F)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub OperationsDivison_Click()
    Selection.TypeText ChrW(&HF7)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub OperationsMinus_Click()
    Selection.TypeText ChrW(&H2212)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub OperationsMinusOrPlus_Click()
    Selection.TypeText ChrW(&H2213)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub OperationsMinusPlus_Click()
    Selection.TypeText ChrW(&H2D)
    Selection.TypeText ChrW(&H2B)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub OperationsMultiplication_Click()
    Selection.TypeText ChrW(&HD7)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub OperationsPlus_Click()
    Selection.TypeText ChrW(&H2B)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub


Private Sub OperationsPlusMinus_Click()
    Selection.TypeText ChrW(&H2B)
    Selection.TypeText ChrW(&H2D)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
End Sub

Private Sub OperationsPlusOrMinus_Click()
    Selection.TypeText ChrW(&HB1)
    Application.Run MacroName:="Dx_ShowMathSymbolsAgain"
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

