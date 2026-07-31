VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_About_Title_And_Agreement 
   Caption         =   "UserForm1"
   ClientHeight    =   8800.001
   ClientLeft      =   105
   ClientTop       =   465
   ClientWidth     =   11340
   OleObjectBlob   =   "Lp_About_Title_And_Agreement.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_About_Title_And_Agreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False












































'Author: Jerry Whittaker -  jerry@thewhittakers.org
 
' Version: 1.4  Date: 7/24/2026 - VersionLabel caption bumped 3.0.5 -> 3.0.6 (caption lives in the .frx)
' Version: 1.3  Date: 12/10/2019 - added alternative short url
' Version: 1.2  Date: 3/17/2017

Private Sub ExitAgreement_Click()
    Unload Me
End Sub


Private Sub Label6_Click()

End Sub

Private Sub ViewVersion_Click()
    Application.Run MacroName:="Sh_Is_Doc_Open"
    Unload Me
    On Error GoTo NoInternet
    ActiveDocument.FollowHyperlink _
        Address:="https://www.dropbox.com/s/l5md4w69h942ua0/Information%20and%20Download%20Links%20for%20Large%20Print%20Templates%20and%20Macros.pdf?dl=0", _
        NewWindow:=True, _
        AddHistory:=True
        GoTo EndOfSub
NoInternet:
        MsgBox "The connection has failed. There may be no internet connection or the site is blocked or unavailable. Try again later or..." & vbCr _
        & vbCr & "Write down this short URL to try in a web browser on this or a different computer." & vbCr _
        & vbCr & "https://bit.ly/2YzXiFi", , "VistaType LP (150)"
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





