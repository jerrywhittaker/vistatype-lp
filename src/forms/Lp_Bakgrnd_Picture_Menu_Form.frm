VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Bakgrnd_Picture_Menu_Form 
   Caption         =   "Background and Picture Tools (344)"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5070
   OleObjectBlob   =   "Lp_Bakgrnd_Picture_Menu_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Bakgrnd_Picture_Menu_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Lp_Bakgrnd_Picture_Menu_Form
' Version: 1.6  Date: 9/15/2026 - added SamePictureSizeButton, Resize This Picture Throughout the Book (Jerry)
' Version: 1.5  Date: 7/26/2026 - Pictures-to-In-Line no longer jumps to a bookmark another form left behind; records and returns to its own position
' Version: 1.4  Date: 7/24/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
' Version 1.3 Date: 4/12/2025 - added Center_Or_Left_Align_Pictures_Form
' Version 1.2 Date: 4/9/2025
' Version 1.1 Date: 1/9/2023


Private Sub ExitButton_Click()
    Unload Me
End Sub

Private Sub BackgroundColorButton_Click()
    Lp_Bakgrnd_Picture_Menu_Form.Hide
    Application.Run MacroName:="Lp_Toggle_Page_Color"
    Unload Lp_Bakgrnd_Picture_Menu_Form
    Unload Me
End Sub

Private Sub ImageSizeButton_Click()
    Lp_Bakgrnd_Picture_Menu_Form.Hide
    Lp_Resize_Images_Form.Show
    Unload Lp_Resize_Images_Form
    Unload Me
End Sub

Private Sub PicturesToInlineButton_Click()
    Lp_Bakgrnd_Picture_Menu_Form.Hide
    ' This used to call Sh_Move_To_And_Delete_Placeholder_Bookmark, jumping to whatever
    ' bookmark another form happened to have left behind. It now records its own position.
    Sh_Save_User_Position
    Dim shp As Shape
    Dim ilShp As inlineShape

    ' Convert floating shapes to inline and lock their aspect ratio
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            ' Convert the shape to an inline shape
            Set ilShp = shp.ConvertToInlineShape
            ' Lock the aspect ratio
            ilShp.LockAspectRatio = True
        End If
    Next shp

    ' Lock aspect ratio for existing inline shapes
    For Each ilShp In ActiveDocument.InlineShapes
        If ilShp.Type = wdInlineShapePicture Or ilShp.Type = wdInlineShapeLinkedPicture Then
            ilShp.LockAspectRatio = True
        End If
    Next ilShp

    MsgBox "All pictures converted to In-Line", , "VistaType LP (177)"
    Application.ScreenUpdating = True
    Sh_Return_User_To_Start_Position
    Unload Me
End Sub

Private Sub PicturesColorGrayscaleButton_Click()
    Lp_Bakgrnd_Picture_Menu_Form.Hide
    Load Lp_Change_Image_Color_Form
    Lp_Change_Image_Color_Form.Show
    Unload Lp_Change_Image_Color_Form
    Unload Me
End Sub
Private Sub CenterOrLeftAlignPicturesButton_Click()
    Lp_Bakgrnd_Picture_Menu_Form.Hide
    Load LP_Picture_Alignment_Form
    LP_Picture_Alignment_Form.Show
    Unload LP_Picture_Alignment_Form
    Unload Me
End Sub

Private Sub SamePictureSizeButton_Click()
    ' Resize This Picture Throughout the Book (Jerry, 9/15/2026). The transcriber has already
    ' selected the picture and set its size; the macro reads that picture and gives every other
    ' copy of it in the book the same size. Hidden first so the book is what is active. Called
    ' directly rather than through Application.Run, so a missing macro stops the build instead
    ' of reaching the transcriber, and an error is reported by the macro's own handler.
    Lp_Bakgrnd_Picture_Menu_Form.Hide
    Lp_Resize_Same_Picture_Throughout
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
