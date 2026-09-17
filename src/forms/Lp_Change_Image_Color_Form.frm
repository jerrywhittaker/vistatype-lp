VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Change_Image_Color_Form 
   Caption         =   "Change Picture Color (345)"
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
' Version: 1.3  Date: 9/4/2026 - the Okay button no longer round-trips the transcriber's
'                               selection through a scratch document, and the work has moved out
'                               of the form into Lp_Change_Image_Color. That macro is where the
'                               whole account of what the round trip was costing lives; in short,
'                               it showed a blank document on the screen, emptied the
'                               transcriber's clipboard, reached pictures in cells nobody had
'                               selected, pasted the selection back instead of leaving it alone,
'                               and on a machine without LargePrintTemplate.dotx closed the
'                               transcriber's own book without saving.
'                               THE RADIO BUTTONS ARE READ BEFORE THE FORM IS UNLOADED. They were
'                               read after it: Unload Me stood on the first line of this handler
'                               and every choice was taken afterwards. Touching any member of a
'                               UserForm's default instance is what CREATES the form - the same
'                               mechanism written up beside Sh_Msg_Body in LPandBrlMacros - so
'                               those lines could be reading a brand new form's design-time
'                               values rather than what the transcriber had clicked.
'                               IT DID NOT BITE, and say that rather than leaving the scare in
'                               place. A UserForm cannot be exercised headlessly on the build
'                               box, so it was put to Jerry instead, 9/4/2026: choosing grayscale
'                               has always given him grayscale. So the values were surviving the
'                               unload and this dialog has been doing what it was told. Reading
'                               them first is still the right order and removes the question -
'                               it is not a repair.
'                               ITS THREE End STATEMENTS ARE GONE TOO. End was how this handler
'                               finished on the "selected picture, original color" path - a
'                               SUCCESS path - and End wipes every module-level variable in the
'                               project, Sh_Pos_Depth and the shared message dialog's own state
'                               among them. Exit Sub leaves the same way without the damage.
' Version: 1.2  Date: 7/26/2026 - returns the user to where the cursor was when Okay was clicked
' Version: 1.1  Date: 7/24/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
' Version: 1.0 Date: 12/26/2023
'

Private Sub Okay_Button_Click()
    Dim whichPics As String
    Dim newColorType As Long

    ' Read the choices FIRST - see the note above. Unload Me comes after them.
    If All_Pictures_Radio_Button Then
        whichPics = "A"                      ' every picture in the document
    ElseIf Selected_Picture_Radio_Button Then
        whichPics = "S"                      ' the picture the transcriber has selected
    ElseIf Picture_Range_Radio_Button Then
        whichPics = "R"                      ' the pictures in the selected text
    End If

    If Color_Radio_Button Then
        newColorType = msoPictureAutomatic   ' the picture's original colors
    ElseIf Grayscale_Radio_Button Then
        newColorType = msoPictureGrayscale
    End If

    Unload Me

    ' A frame with nothing chosen in it. Nothing to do, and nothing to say about it - the same
    ' as this dialog has always done.
    If whichPics = "" Or newColorType = 0 Then Exit Sub

    Lp_Change_Image_Color whichPics, newColorType
End Sub

' A Black and White choice was written into this handler and could never run: there is no
' Black_and_White_Radio_Button on this form, and there is none in this repository's history
' either. These forms carry no Option Explicit, so that name was an empty Variant and its ElseIf
' simply tested False. Removed with the rest of the rewrite on 9/4/2026. Word does have
' msoPictureBlackAndWhite, so the choice is one radio button and one line away if Jerry wants it.

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






