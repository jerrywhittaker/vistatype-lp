VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_Valid_Ref_Pg_No_3_Form 
   Caption         =   "Delete/Change/Add $pg Tags"
   ClientHeight    =   2772
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10965
   OleObjectBlob   =   "Sh_Valid_Ref_Pg_No_3_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sh_Valid_Ref_Pg_No_3_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sh_Valid_Ref_Pg_No_3_Form
' Provides directions for operation in the temp doc to Fix the $pg tags in the main document
'
' Version 1.0: Date: 7/27/2026
'

Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    Directions.Caption = "Here, in your main document, you can delete, add, or correct the $pg tags." _
    & vbCrLf & vbCrLf & "New $pg tags may be added using the 'Manual Tag Ref Page' macro on the 'Reference Page Numbers group of the the Word ribbon.'" _
    & " New $pg tags should only be added to empty paragraphs." _
    & vbCrLf & vbCrLf & "To return the the validation document, click the 'Return to the validation document' button." _
    & " When you have completed the tag corrections click the 'Done - Exit validation' button."
   
End Sub

