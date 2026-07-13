VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Table_Tools_Menu_Form 
   Caption         =   "Large Print Table Tools"
   ClientHeight    =   8300.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6855
   OleObjectBlob   =   "Lp_Table_Tools_Menu_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Lp_Table_Tools_Menu_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lp_Table_And_TOC_Tools_Menu_Form
'
' Version: 2.1  Date: 5/28/2025 - moved check for a selected table to Lp_Color_Bars_For_TOCs_Form
' Version: 2.0  Date: 4/19/2024 - error tapping added
' Version: 1.9  Date: 3/9/2023 - added additional table weight procedures
' Version: 1.8  Date: 12/9/2021 - added plain (Table Grid) button
' Version: 1.7  Date:  1/16/2020 - reduced form size - add real columns button - misc edits
' Version: 1.6  Date: 12/19/2019 - fixed font Spacing and Position
' Version: 1.5  Date: 12/16/2019 - fixed table formatting
' Version: 1.4  Date: 11/22/2019 - changed List_Button_Click to begin list user option sequence (list options and table type)
'                                - moved all checks for table selection or cursor in table to this menu
' Version: 1.3  Date: 5/31/2019 - added Lp_Get_Doc_Setup_Params
' Version: 1.2  Date: 2/26/2019
' Version: 1.1  Date: 1/14/2019
' Version: 1.0  Date: 5/22/2018

Private Sub Cmd_Cancel_Click()
    Unload Me
    End
End Sub

Private Sub Default_Color_Selected_Click()
   
    If Not Selection.Information(wdWithInTable) Then 'is the table selected
        MsgBox "Select the entire table or place the cursor within the table.", , "VistaType LP (165)"
        Unload Me
        End
    End If

    On Error GoTo eom
    Selection.Tables(1).Select
 
    Dim CurrentTable As Table
        For Each CurrentTable In Selection.Tables
            With Selection.Font
                .Name = "Tahoma"
                .Size = Lp_Base_Font_Size
            End With
            With CurrentTable
                .Range.Font.Name = "Tahoma"
                .Style = "Yellow on White Paper Table"
             End With
             Application.Run MacroName:="Lp_SetSelectedTableBorderWeight"
        Next CurrentTable
eom:
    Application.Selection.Collapse
    Unload Me
End Sub

Private Sub Yellow_Color_Selected_Click()
   
    If Not Selection.Information(wdWithInTable) Then 'is the table selected
        MsgBox "Select the entire table or place the cursor within the table first!", , "VistaType LP (166)"
        Unload Me
    End If

    On Error GoTo eom
    Selection.Tables(1).Select
 
    Dim CurrentTable As Table
        For Each CurrentTable In Selection.Tables
            With Selection.Font
                .Name = "Tahoma"
                .Size = Lp_Base_Font_Size
            End With
            With CurrentTable
                .Range.Font.Name = "Tahoma"
                .Style = "Yellow on Black Screen Table"
                Application.Run MacroName:="Lp_SetSelectedTableBorderWeight"
             End With
         Next CurrentTable
eom:
    Application.Selection.Collapse
    Unload Me
End Sub

Private Sub Gray_Color_Selected_Click()
   
    If Not Selection.Information(wdWithInTable) Then 'is the table selected
        MsgBox "Select the entire table or place the cursor within the table first!", , "VistaType LP (167)"
        Unload Me
    End If
    
    On Error GoTo eom
    Selection.Tables(1).Select
 
    Dim CurrentTable As Table
        For Each CurrentTable In Selection.Tables
            With Selection.Font
                .Name = "Tahoma"
                .Size = Lp_Base_Font_Size
            End With
            With CurrentTable
                .Range.Font.Name = "Tahoma"
                .Style = "Gray Scale Table"
                Application.Run MacroName:="Lp_SetSelectedTableBorderWeight"
             End With
        Next CurrentTable
eom:
    Application.Selection.Collapse
    Unload Me
End Sub

Private Sub Black_And_White_Selected_Click()
   
    If Not Selection.Information(wdWithInTable) Then 'is the table selected
        MsgBox "Select the entire table or place the cursor within the table first!", , "VistaType LP (168)"
        Unload Me
    End If
    
    On Error GoTo eom
    Selection.Tables(1).Select
 
    Dim CurrentTable As Table
        For Each CurrentTable In Selection.Tables
           With Selection.Font
                .Name = "Tahoma"
                .Size = Lp_Base_Font_Size
            End With
            With CurrentTable
                .Range.Font.Name = "Tahoma"
                .Style = "Table Grid"
                Application.Run MacroName:="Lp_SetSelectedTableBorderWeight"
             End With
        Next CurrentTable
eom:
    Application.Selection.Collapse
    Unload Me
End Sub

Private Sub List_Button_Click()
    If Not Selection.Information(wdWithInTable) Then 'is the table selected
        MsgBox "Select the entire table or place the cursor within the table first!", , "VistaType LP (169)"
        End
    End If
    Lp_GP_String_3 = "L " 'to tell Lp_Table_Convert_Options_Form that the request was for List
    Lp_Table_Convert_Options_Form.Show
    End
End Sub

Private Sub PseudoButton_Click()
    If Not Selection.Information(wdWithInTable) Then 'is the table selected
        MsgBox "Select the entire table or place the cursor within the table", , "VistaType LP (170)"
        End
    End If
    Application.Run MacroName:="LP_Convert_Table_To_Pseudo_Columns"
    End
End Sub

Private Sub PseudoQuestion_Click()
    LP_PseudoColumnInfoForm.Show
End Sub

Private Sub RealButton_Click()
    If Not Selection.Information(wdWithInTable) Then 'is the table selected
        MsgBox "Select the entire table or place the cursor within the table", , "VistaType LP (171)"
        End
    End If
    Application.Run MacroName:="Lp_Convert_Table_To_Real_Columns"
    End
End Sub

Private Sub RealQuestion_Click()
    LP_RealColumnInfoForm.Show
End Sub

Private Sub Rotate_Button_Click()
    Lp_GP_String_3 = "R " 'to tell Lp_Table_Convert_Options_Form that the request was for rotation
    Lp_Table_Convert_Options_Form.Show
    Unload Me
    End
End Sub
   
Private Sub Default_Color_All_Click()
    For Each t In ActiveDocument.Tables
        t.Style = "Yellow on White Paper Table"
    Next
    Application.Run MacroName:="Lp_Set_Table_Border_Weights"
    Unload Me
    End
End Sub

Private Sub Yellow_Table_All_Click()
    For Each t In ActiveDocument.Tables
        t.Style = "Yellow on Black Screen Table"
     Next
    Application.Run MacroName:="Lp_Set_Table_Border_Weights"
    Unload Me
    End
End Sub
         
Private Sub Gray_Table_All_Click()
    For Each t In ActiveDocument.Tables
        t.Style = "Gray Scale Table"
     Next
    Application.Run MacroName:="Lp_Set_Table_Border_Weights"
    Unload Me
End Sub

Private Sub Plain_Table_All_Click()
    For Each t In ActiveDocument.Tables
        t.Style = "Table Grid"
    Next
    Application.Run MacroName:="Lp_Set_Table_Border_Weights"
    Unload Me
End Sub
   
Private Sub userform_terminate() 'red X was clicked
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    Dim t As Table
    Dim rng As Range
     
    Lp_GP_String_3 = ""
    
    Application.Run MacroName:="Lp_Get_Doc_Setup_Params"
    
    Application.Run MacroName:="MS_Set_Word_Config_For_Large_Print"
    
    ' From: https://www.thespreadsheetguru.com/the-code-vault/launch-vba-userforms-in-correct-window-with-dual-monitors
    ' Start Userform Centered inside Word Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub

