VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_Video_Download_Link_Page 
   Caption         =   "Video and Practice File Links (337)"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8205.001
   OleObjectBlob   =   "Dx_Video_Download_Link_Page.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_Video_Download_Link_Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Exit_Button_Click()
    Unload Me
    End
End Sub

Private Sub View_the_Page_Click()
    Application.Run MacroName:="Sh_Is_Doc_Open"
    On Error GoTo NoInternet
    ActiveDocument.FollowHyperlink _
        Address:="https://www.dropbox.com/s/3h9ma3au93e42e0/Braille%20Macros%20Video%20and%20Practice%20File%20Links.pdf?dl=0", _
        NewWindow:=True, _
        AddHistory:=True
        GoTo EndOfSub
NoInternet:
        MsgBox "The connection has failed. There may be no internet connection or the site is blocked or unavailable. Try again later or..." & vbCr _
        & vbCr & "Write down this short URL to try in a web browser on this or a different computer." & vbCr _
        & vbCr & "https://tinyurl.com/mjawwjtd", , "Braille Macros (318)"
EndOfSub:
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

Private Sub userform_terminate() 'red X was clicked
    End
End Sub


