VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Change_Image_Color_Form 
   Caption         =   "Change Picture Color"
   ClientHeight    =   4464
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3945
   OleObjectBlob   =   "Lp_Change_Image_Color_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Change_Image_Color_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Change Image Color
'
' Version: 1.2  Date: 7/26/2026 - returns the user to where the cursor was when Okay was clicked
' Version: 1.1  Date: 7/24/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
' Version: 1.0 Date: 12/26/2023
'

Private Sub Okay_Button_Click()
    Dim pic As inlineShape
    Unload Me
    
    If All_Pictures_Radio_Button Then
        If Color_Radio_Button Then
            For Each pic In ActiveDocument.InlineShapes
                pic.PictureFormat.ColorType = msoPictureAutomatic 'color
            Next
        ElseIf Grayscale_Radio_Button Then
            For Each pic In ActiveDocument.InlineShapes
                pic.PictureFormat.ColorType = msoPictureGrayscale 'grayscale
            Next
        End If
    End If
        
    If Selected_Picture_Radio_Button Then
        On Error GoTo err1 'error when on next line when no selected picture
        Set pic = Selection.InlineShapes(1)

        If Color_Radio_Button Then
            ' modified code from: https://www.appsloveworld.com/vba/200/73/recolor-picture-to-black-and-white-75-using-word-vba
            With pic
                pic.PictureFormat.ColorType = msoPictureAutomatic 'color
            End With
            End
err1:
            MsgBox "No Picture selected", , "VistaType LP (299)"
            On Error GoTo 0
            End
        ElseIf Grayscale_Radio_Button Then
            With pic
                pic.PictureFormat.ColorType = msoPictureGrayscale 'grayscale
            End With
        ElseIf Black_and_White_Radio_Button Then
            With pic
                pic.PictureFormat.ColorType = msoPictureBlackAndWhite
            End With
        End If
    End If
        
    If Picture_Range_Radio_Button Then

        If Selection.Type <> wdSelectionNormal Then
            MsgBox "Select the text range containing pictures for color change", , "VistaType LP (164)"
            End
        End If
          
        Application.ScreenUpdating = False ' Turn screen updating off
        Sh_Save_User_Position
        Application.Run MacroName:="Lp_Copy_To_Temp_Doc"
        
        Application.Run MacroName:="Sh_Is_End_Paragraph_Mark_Included"
        
        For Each pic In ActiveDocument.InlineShapes
            If Color_Radio_Button Then
                pic.PictureFormat.ColorType = msoPictureAutomatic 'color
            ElseIf Grayscale_Radio_Button Then
                pic.PictureFormat.ColorType = msoPictureGrayscale
            End If
        Next
        
        Selection.EndKey Unit:=wdStory
        Selection.Delete Unit:=wdCharacter, count:=1
    
        Selection.HomeKey Unit:=wdStory
        Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    
        Selection.Copy 'copy the selected text to the clipboard
        ActiveDocument.Close SaveChanges:=False 'close the temp doc without saving
        Selection.Paste
        Application.ScreenUpdating = True
        Sh_Return_User_To_Start_Position
    End If
End Sub

Private Sub Cancel_Button_Click()
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






