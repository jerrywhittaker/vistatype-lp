VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LP_Set_Doc_Background_Form 
   Caption         =   "Background (342)"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5010
   OleObjectBlob   =   "LP_Set_Doc_Background_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LP_Set_Doc_Background_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Lp_Set_Doc_Background_Form
'
' Version: 1.4  Date: 7/24/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
' Version: 1.3  Date: 11/8/2025 - added  Sh_Write_Document_Variables "Media", "Paper"
' Version: 1.2  Date: 11/9/2023 - added "VistaType" to msgbox
' Version: 1.1  Date: 12/3/2020 - rewrite to check for document media type
' Version: 1.0  Date: 10/18/2020

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OkayButton_Click()

    Application.Run MacroName:="Lp_Get_Doc_Setup_Params"    'need to determine Doc is for "Screen", "Paper", or "Unknown" - places in Public Var "DM"

    If BlackBkgrndWhiteTextButton = True Then
        Unload LP_Set_Doc_Background_Form
        'DM = document media - screen or paper
        If DM = "Screen" Then
            Application.Run MacroName:="Lp_Set_Page_To_Black"
            GoTo eom
        ElseIf DM = "Paper" Then
                MsgBox "Black backgrounds are for screen documents only", , "VistaType LP (176)"
                GoTo eom
        ElseIf DM = "Unknown" Or DM = "" Then
                Dim lngChoice As Long
                lngChoice = MsgBox("Black backgounds are for documents intended for screen reading only. Is this document intended for screen reading?", vbYesNo + vbDefaultButton2, "VistaType LP (298)")
                If lngChoice = vbYes Then
                    ' writes a variable name (Media) and variable value "Screen" into the document xml file
                    Sh_Write_Document_Variables "Media", "Screen"
                    Application.Run MacroName:="Lp_Set_Page_To_Black"
                    GoTo eom
                Else ' answer was "No" writes a variable name (Media) and variable value "Paper" into the document xml file
                    Sh_Write_Document_Variables "Media", "Paper"
                    GoTo eom
                End If
        End If
        Else
        Application.Run MacroName:="Lp_Set_Page_To_White"
    End If
    
eom:
    Unload Me
                   
End Sub

Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' 7/24/2026 - removed "MS_Set_Word_Config_For_Large_Print": opening the document already
    '             configures Word for large print, and re-running it here reset the user's
    '             Styles-pane options (show filter / sort order) every time this form opened.
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub


