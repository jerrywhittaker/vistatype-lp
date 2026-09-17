VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Video_Download_Link_Page 
   Caption         =   "Video and Practice File Downloads (358)"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   OleObjectBlob   =   "Lp_Video_Download_Link_Page.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Video_Download_Link_Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Version: 1.1  Date: 7/24/2026 - no longer runs "MS_Set_Word_Config_For_Large_Print" on form open
'
Private Sub Exit_Button_Click()
    Unload Me
End Sub

Private Sub View_the_Page_Click()
    Application.Run MacroName:="Sh_Is_Doc_Open"
    On Error GoTo NoInternet
    ActiveDocument.FollowHyperlink _
        Address:="https://www.dropbox.com/s/9lu2rtuq9czgarn/Large%20Print%20Video%20and%20Practice%20File%20Links.pdf?dl=0", _
        NewWindow:=True, _
        AddHistory:=True
        GoTo EndOfSub
NoInternet:
        MsgBox "The connection has failed. There may be no internet connection or the site is blocked or unavailable. Try again later or..." & vbCr _
        & vbCr & "Write down this short URL to try in a web browser on this or a different computer." & vbCr _
        & vbCr & "https://tinyurl.com/ycn3dncw", , "VistaType LP (147)"
EndOfSub:
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

Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub
