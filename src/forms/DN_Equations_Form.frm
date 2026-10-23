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
' Version: 1.3  Date: 10/22/2026 - "MathType equations" is always offered (Jerry). Chosen on a
'                                machine without MathType, Okay asks VistaType LP (400) whether to
'                                go on, since the equations could not be edited or corrected there;
'                                Cancel there comes back to this box. Asked for braille too, with
'                                "Pictures of the equations" grayed out (Sh_GP_String_2 = "BRL").
'                                "MathType equations" takes its font from "Word equations (OMML)";
'                                it was drawn smaller than the other two (Jerry, testing 3.0.599).
'                                The order is MathType equations (the default), Word equations
'                                (OMML), Pictures of the equations (Jerry, testing 3.0.601)
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
' Asked by Sh_Convert_XML_File_To_Word_Document for a large print or braille conversion, and only
' when the book's XML carries MathML equations. Braille cannot have the pictures, so that choice
' is grayed out when 325 was answered braille (Sh_GP_String_2 = "BRL"). A NIMAS book names a
' picture of each equation in the altimg attribute (the "EQ" files); Word's HTML importer drops
' the MathML and the picture with it. The transcriber chooses which goes in: Word equations, made from the MathML by Word's own
' MML2OMML.XSL; MathType equations, made from the MathML by VistaType LP and drawn by MathType;
' or the pictures. MathType equations can be made without MathType, but cannot then be edited,
' so the transcriber is asked (400) before going on.
'
' Answers through Sh_GP_String_3, "OMML", "MATHTYPE" or "PICTURE". An empty string means
' canceled, and the caller stops before anything is done to the book.

Private Sub CmdOkay_Click()
    If EquationPictures.Value = True And EquationPictures.Enabled Then
        Sh_GP_String_3 = "PICTURE"
    ElseIf MathTypeEquations.Value = True Then
        'Without MathType the equations still go in, but nobody can open one to correct it.
        'Cancel (or Esc, or the red X) on 400 is "No": this box stays up for another choice.
        If Not Sh_MathType_Is_Installed() Then
            If Not Sh_Ask("MathType is not installed on this computer, resulting in an inability " & _
                          "to edit or correct the equations." & vbCr & vbCr & _
                          "Do you wish to continue creating MathType equations?", _
                          "VistaType LP (400)") Then Exit Sub
        End If
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
    'MathType equations: always offered, and the default. Whether MathType is on this machine is
    'asked only when it is chosen (CmdOkay_Click).
    MathTypeEquations.Caption = "MathType equations"
    MathTypeEquations.Accelerator = "M"

    'The order, top to bottom: MathType equations, Word equations (OMML), Pictures of the
    'equations (Jerry, 10/22/2026). The three keep the places the designer gave them, handed
    'out in this order, so the box is the same size.
    Dim topFirst As Single, topThird As Single, topSecond As Single
    topFirst = WordEquations.Top
    If EquationPictures.Top < topFirst Then topFirst = EquationPictures.Top
    If MathTypeEquations.Top < topFirst Then topFirst = MathTypeEquations.Top
    topThird = WordEquations.Top
    If EquationPictures.Top > topThird Then topThird = EquationPictures.Top
    If MathTypeEquations.Top > topThird Then topThird = MathTypeEquations.Top
    topSecond = WordEquations.Top + EquationPictures.Top + MathTypeEquations.Top - topFirst - topThird
    MathTypeEquations.Top = topFirst
    WordEquations.Top = topSecond
    EquationPictures.Top = topThird
    MathTypeEquations.TabIndex = 0
    WordEquations.TabIndex = 1
    EquationPictures.TabIndex = 2
    MathTypeEquations.Value = True
    MathTypeEquations.Font.Name = WordEquations.Font.Name
    MathTypeEquations.Font.Size = WordEquations.Font.Size

    'Braille (325): the pictures are no use to a braille reader, so that choice is grayed out.
    If Sh_GP_String_2 = "BRL" Then EquationPictures.Enabled = False

    'Center UserForm within the Word application window
    Me.StartUpPosition = 0

    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub
