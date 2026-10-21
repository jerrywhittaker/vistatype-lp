VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DN_Equations_Form 
   Caption         =   "Equations (398)"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3255
   OleObjectBlob   =   "DN_Equations_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DN_Equations_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
' DN_Equations_Form  -  "Equations (398)"
'
' Version: 1.2  Date: 10/20/2026 - a third choice, "MathType equations" (Alt+M), offered only when
'                                MathType is installed on the machine (Sh_MathType_Is_Installed);
'                                otherwise it is hidden and the box is the size it was (Jerry's
'                                MathType plan, Step 4)
' Version: 1.1  Date: 10/20/2026 - a line at the top says the file contains math equations and asks
'                                how they should go in; the first choice reads "Word equations
'                                (OMML)" (Jerry)
' Version: 1.0  Date: 10/18/2026 - new (Jerry, 10/18/2026, issue #37)
'
' Author: Jerry Whittaker - jerry@vistatypelp.org
'
' Asked by Sh_Convert_XML_File_To_Word_Document for a LARGE PRINT conversion, and only when the
' book's XML carries MathML equations. A NIMAS book names a picture of each equation in the
' altimg attribute (the "EQ" files); Word's HTML importer drops the MathML and the picture with
' it. The transcriber chooses which goes in: Word equations, made from the MathML by Word's own
' MML2OMML.XSL; MathType equations, made from the MathML by VistaType LP and drawn by MathType
' (only when MathType is installed); or the pictures.
'
' Answers through Sh_GP_String_3, "OMML", "MATHTYPE" or "PICTURE". An empty string means
' canceled, and the caller stops before anything is done to the book.

Private Sub CmdOkay_Click()
    If EquationPictures.Value = True Then
        Sh_GP_String_3 = "PICTURE"
    ElseIf MathTypeEquations.Visible And MathTypeEquations.Value = True Then
        Sh_GP_String_3 = "MATHTYPE"
    Else
        Sh_GP_String_3 = "OMML"
    End If

    Unload Me
End Sub

Private Sub CmdCancel_Click()
    Sh_GP_String_3 = vbNullString  'signal cancel
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'The red X is Cancel: it leaves no answer, and the box then closes by itself. Do not set
    'Cancel = True and unload from here - Word ignores an Unload inside QueryClose.
    If CloseMode = vbFormControlMenu Then Sh_GP_String_3 = ""
End Sub

Private Sub UserForm_Initialize()
    'Word equations are the default: text the reader can enlarge, where a picture is fixed at
    'one size (Jerry's large print rule: every character at the size set).
    WordEquations.Value = True

    'MathType equations: offered only when MathType is on this machine. Otherwise the choice is
    'hidden, the buttons move up into its place, and the box is the height it always was.
    MathTypeEquations.Caption = "MathType equations"
    MathTypeEquations.Accelerator = "M"
    MathTypeEquations.TabIndex = 2
    If Not Sh_MathType_Is_Installed() Then
        Dim lift As Single
        lift = CmdOkay.Top - MathTypeEquations.Top
        MathTypeEquations.Visible = False
        CmdOkay.Top = CmdOkay.Top - lift
        CmdCancel.Top = CmdCancel.Top - lift
        Me.Height = Me.Height - lift
    End If

    'Center UserForm within the Word application window
    Me.StartUpPosition = 0

    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub
