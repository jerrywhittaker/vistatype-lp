VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sh_Valid_Ref_Pg_No_2_Form 
   Caption         =   "Validate $pg Tags"
   ClientHeight    =   708
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9735.001
   OleObjectBlob   =   "Sh_Valid_Ref_Pg_No_2_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sh_Valid_Ref_Pg_No_2_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sh_Valid_Ref_Pg_No_2_Form
' This for is showing when the user in the temp doc for validating $pg tags
'
' Version 1.0:  Date: 7/27/2026
'
' --- at the very top of the UserForm module (declarations section) ---
#If VBA7 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
#End If
Private Const LOGPIXELSY As Long = 90

Private Sub DoneButton_Click()
    'Closes the tag list without saving and puts you back in your document.
    Sh_PgVal_Done
End Sub

Private Sub LocateInDocButton_Click()
    'Finds the tag on the line the cursor is sitting on, in the document being validated,
    'and swaps this form for Sh_Valid_Ref_Pg_No_4_Form over there.
    Sh_PgVal_LocateInDocument
End Sub

Private Sub ShowDirectionsButton_Click()
    'Me.Hide
    Sh_Valid_Ref_Pg_No_1_Form.Show ' directions for use
    'Me.Show
End Sub

Private Sub UserForm_Initialize()
    Dim pxLeft As Long, pxTop As Long, pxWidth As Long, pxHeight As Long
    Dim pxPerPt As Double, zoomF As Double
    Dim topRange As Range
    #If VBA7 Then
        Dim hDC As LongPtr
    #Else
        Dim hDC As Long
    #End If

    Me.StartUpPosition = 0

    ' Screen DPI: device pixels per point (this is screen/form space)
    hDC = GetDC(0)
    pxPerPt = GetDeviceCaps(hDC, LOGPIXELSY) / 72
    ReleaseDC 0, hDC

    ' Document zoom — document space is scaled by this on screen
    zoomF = ActiveWindow.View.Zoom.Percentage / 100

    ' Collapsed range at the start of the body text
    Set topRange = ActiveDocument.Content
    topRange.Collapse wdCollapseStart

    On Error Resume Next
    ActiveWindow.GetPoint pxLeft, pxTop, pxWidth, pxHeight, topRange
    On Error GoTo 0

    If pxWidth > 0 Then
        ' Center over the text column (screen px -> form points: DPI only)
        Me.Left = (pxLeft + pxWidth / 2) / pxPerPt - Me.Width / 2

        ' Step up from body-text top to the page's top edge.
        ' TopMargin is a document measure -> scale by DPI *and* zoom.
        Me.Top = pxTop / pxPerPt - ActiveDocument.PageSetup.TopMargin * zoomF
    Else
        ' Fallback if the top of the doc is scrolled off-screen
        Me.Left = Application.Left + (Application.Width - Me.Width) / 2
        Me.Top = Application.Top + 120
    End If
End Sub
