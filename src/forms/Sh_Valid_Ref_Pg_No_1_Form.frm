VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_Valid_Ref_Pg_No_1_Form 
   Caption         =   "How to validate the $pg tags"
   ClientHeight    =   3036
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10800
   OleObjectBlob   =   "Sh_Valid_Ref_Pg_No_1_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sh_Valid_Ref_Pg_No_1_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sh_Valid_Ref_Pg_No_1_Form
' Provides directions for operation in the temp doc to validate the $pg tags in the main document
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
    
    Directions.Caption = "This is a temporary document to facilitate validation of the '$pg' reference page tags." _
    & " This document can be closed without saving when the validation is complete by selecting the 'Done - Exit validation' button." _
    & vbCrLf & vbCrLf & "Select the $pg item to be located in the document then click the 'Locate the selected $pg code in the Document' button." _
    & vbCrLf & vbCrLf & "To add a missing $pg tag, select the $pg tag the preceeds the missing tag, locate it in the document and scroll to the point for the insertion."
End Sub
