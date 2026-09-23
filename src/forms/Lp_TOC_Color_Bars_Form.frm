VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_TOC_Color_Bars_Form 
   Caption         =   "TOC Color Bars (353)"
   ClientHeight    =   4308
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6405
   OleObjectBlob   =   "Lp_TOC_Color_Bars_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_TOC_Color_Bars_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Lp_TOC_Color_Bars_Form
'
' Converts lead lines in TOC styles to alternating color of choice - changes tab value of dots to none
'
' Version: 1.1 Date: 9/22/2026 - NO End. Cancel and the X only unload this box now: it is shown from
'                               the TOC box (354), which stays open, and End would take that box
'                               down and drop the TOC range it holds. Lp_TOC_Bars_Applied tells the
'                               TOC box whether the bars went on, so it says so only when they did.
' Version: 1.0 Date: 9/4/2025
'
Private Step1 As Boolean
Private Step2 As Boolean
Private BarColor1 As String
Private BarColor2 As String
Private TempFileName As String
Private ColorSwitch As String
Private NextColor As Integer

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub YellowImageButton_Click()
    If Step1 Then
        YellowButton.Value = True
        YellowButton.Enabled = True
        ColorSwitch = "Y"
    End If
    
    If Step2 And Not ColorSwitch = "Y" Then
        YellowButton.Enabled = True
        YellowButton = True
    End If
End Sub

Private Sub PinkImageButton_Click()
    If Step1 Then
        PinkButton.Value = True
        PinkButton.Enabled = True
        ColorSwitch = "P"
    End If
    
    If Step2 And Not ColorSwitch = "P" Then
        PinkButton.Enabled = True
        PinkButton = True
    End If
End Sub

Private Sub TanImageButton_Click()
    If Step1 Then
        TanButton.Value = True
        TanButton.Enabled = True
        ColorSwitch = "T"
    End If
    
    If Step2 And Not ColorSwitch = "T" Then
        TanButton.Enabled = True
        TanButton = True
    End If
End Sub

Private Sub BlueImageButton_Click()
    If Step1 Then
        BlueButton.Value = True
        BlueButton.Enabled = True
        ColorSwitch = "B"
    End If
    
    If Step2 And Not ColorSwitch = "B" Then
        BlueButton.Enabled = True
        BlueButton = True
    End If
End Sub

Private Sub AquaImageButton_Click()
    If Step1 Then
        AquaButton.Value = True
        AquaButton.Enabled = True
        ColorSwitch = "A"
    End If
    
    If Step2 And Not ColorSwitch = "A" Then
        AquaButton.Enabled = True
        AquaButton = True
    End If
End Sub

Private Sub GreenImageButton_Click()
    If Step1 Then
        GreenButton.Value = True
        GreenButton.Enabled = True
        ColorSwitch = "G"
    End If
    
    If Step2 And Not ColorSwitch = "G" Then
        GreenButton.Enabled = True
        GreenButton = True
    End If
End Sub

Private Sub OkayButton_Click()

    If Step1 Then
        Step1 = False
        Step2 = True
        Directions1.Visible = False
        Directions2.Visible = True

        If YellowButton Then
            BarColor1 = "Words Yellow"
        ElseIf PinkButton Then
            BarColor1 = "Words Pink"
        ElseIf TanButton Then
            BarColor1 = "Words Tan"
        ElseIf BlueButton Then
            BarColor1 = "Words Blue"
        ElseIf GreenButton Then
            BarColor1 = "Words Green"
        ElseIf AquaButton Then
            BarColor1 = "Words Aqua"
        End If
        
        If YellowButton Then
            YellowButton.Value = False
            YellowButton.Enabled = False
        ElseIf PinkButton Then
            PinkButton.Value = False
            PinkButton.Enabled = False
        ElseIf TanButton Then
            TanButton.Value = False
            TanButton.Enabled = False
        ElseIf BlueButton Then
            BlueButton.Value = False
            BlueButton.Enabled = False
        ElseIf GreenButton Then
            GreenButton.Value = False
            GreenButton.Enabled = False
        ElseIf AquaButton Then
            AquaButton.Value = False
            AquaButton.Enabled = False
        End If
        
        GoTo NextStep
    End If
    
    If Step2 Then
        If YellowButton Then
            BarColor2 = "Words Yellow"
        ElseIf PinkButton Then
            BarColor2 = "Words Pink"
        ElseIf TanButton Then
            BarColor2 = "Words Tan"
        ElseIf BlueButton Then
            BarColor2 = "Words Blue"
        ElseIf GreenButton Then
            BarColor2 = "Words Green"
        ElseIf AquaButton Then
            BarColor2 = "Words Aqua"
        End If
    End If

    Lp_TOC_Color_Bars_Form.Hide
      
    'convert to color bars
    Dim para As Paragraph
    Dim i As Integer
    Dim BarColor As String

    ' Loop through the paragraphs in the selected range
    For i = 1 To Selection.Range.Paragraphs.count
        Set para = Selection.Range.Paragraphs(i)
        
        ' Apply style to every other paragraph
        If Left(para.Range.Style, 3) = "TOC" And NextColor = 1 Then
            para.Range.Style = BarColor1
            NextColor = 2
        ElseIf Left(para.Range.Style, 3) = "TOC" And NextColor = 2 Then
            para.Range.Style = BarColor2
            NextColor = 1
        End If
    Next i

    'remove tab leader dots
    If Not LeaveDots Then
        ActiveDocument.Content.Select
        Dim tabStop As tabStop
    
        ' Loop through each paragraph in the selection
        For Each para In Selection.Paragraphs
            ' Check if the paragraph's style starts with "TOC"
            If Left(para.Style, 3) = "TOC" Then
                ' Loop through each tab stop in the paragraph
                For Each tabStop In para.TabStops
                    ' Remove the leader from the tab stop if it exists
                    tabStop.Leader = wdTabLeaderSpaces
                Next tabStop
            End If
        Next para
    End If
    
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    DoEvents
    ActiveWindow.View.Type = wdNormalView
    ActiveWindow.View.Type = wdPrintView
    Selection.Collapse Direction:=wdCollapseStart
    Lp_TOC_Bars_Applied = True

NextStep:

    If ColorSwitch = "Y" Then
        AquaButton = True
    ElseIf ColorSwitch = "A" Then
        YellowButton = True
    End If

End Sub

Private Sub UserForm_Initialize()
    Step1 = True
    Step2 = False
    YellowButton.Value = True
    ColorSwitch = "Y"
    Directions1.Visible = True
    Directions2.Visible = False
    NextColor = 1
    
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub




