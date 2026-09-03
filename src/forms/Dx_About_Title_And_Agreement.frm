VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dx_About_Title_And_Agreement 
   Caption         =   "Software Agreement"
   ClientHeight    =   8820.001
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11460
   OleObjectBlob   =   "Dx_About_Title_And_Agreement.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dx_About_Title_And_Agreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Author: Jerry Whittaker -  jerry@vistatypelp.org

' Version: 1.1  Date: 8/2/2026 - old permissive agreement replaced by the GPLv3 summary (Sh_Software_Agreement_Text) in a scrollable box, plus a View Full License button
' Version: 1.4  Date: 7/24/2026 - Label5 version caption bumped 3.0.5 -> 3.0.6 (caption lives in the .frx)
' Version: 1.3  Date: 12/10/2019 - added alternative short url
' Version: 1.2  Date: 3/17/2017

Private Sub ExitAgreement_Click()
    Unload Me
    End
End Sub

Private Sub ViewVersion_Click()
    Application.Run macroName:="Sh_Is_Doc_Open"
    Unload Me
    On Error GoTo NoInternet
    ActiveDocument.FollowHyperlink _
        Address:="https://www.dropbox.com/s/zafnrnfhftocu1b/Current%20Information%20about%20VBA%20Macros%20for%20the%20BANA%20Template%20for%20the%20Duxbury%20Braille%20Translator.pdf?dl=0", _
        NewWindow:=True, _
        AddHistory:=True
        GoTo EndOfSub
NoInternet:
        MsgBox "The connection has failed. There may be no internet connection or the site is blocked or unavailable. Try again later or..." & vbCr _
        & vbCr & "Write down this short URL to try in a web browser on this or a different computer." & vbCr _
        & vbCr & "https://bit.ly/2YKnQUF", , "Braille Macros (314)"
EndOfSub:

End Sub

Private Sub ViewLicense_Click()
    ' The full GNU GPL, the same text the installer shows. Shared helper so both About
    ' dialogs behave identically.
    '
    ' Unload Me FIRST. This dialog is modal, so it blocks Word's own window -- the license
    ' document opened behind it and the user saw nothing happen (Jerry, 8/3/2026). The
    ' View Version button beside this one has always closed itself before following its
    ' link, for the same reason.
    Unload Me
    Sh_Show_Full_License
End Sub

Private Sub UserForm_Initialize()

    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

    ' The Software Agreement summary. One shared copy in LPandBrlMacros so the LP and
    ' Braille dialogs can never drift apart, and so the wording stays in git-tracked
    ' text instead of inside this form's binary .frx. The full license is not here --
    ' the View Full License button opens it read-only in Word, where it can be
    ' scrolled with the wheel and zoomed. See Sh_Show_Full_License.
    AgreementTextBox.Text = Sh_Software_Agreement_Text()
    ' SelStart, NOT CurLine. CurLine needs the control to already have the focus and
    ' raises run-time error 2185 from UserForm_Initialize, which runs before the form
    ' is shown. SelStart does the same job -- open at the top, not the end -- with no
    ' focus required. Same approach as Sh_Prodnote_Info_Form.
    AgreementTextBox.selStart = 0

End Sub
