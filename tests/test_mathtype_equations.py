"""MathType equations in a DAISY/NIMAS conversion: Steps 3 and 4 of the MathType plan (10/20/2026).

Large print's Equations (398) gets a third choice, "MathType equations", shown only when MathType
is installed. Each equation then goes in as a MathType equation made from its MathML
(Sh_MTEF_From_MathML, Sh_MTEF_Ole_File, Sh_MathType_Package) with the book's EQ picture showing,
falling back to a Word equation and then to the picture; and MathType's Format Equations draws
them all afterwards. If MathType cannot be reached the conversion carries on, and the closing
message (313) tells the user to click Format Equations once on the MathType tab.

Measured on VistaBuild (see the plan and the memory note mathtype-from-mathml): MTInitAPI must
come before Format Equations is run from code, or it stops behind a message nobody can see.

The Word side cannot run in these tests; these are checks on the source. Sh_MathType_Package
and the picture-size reader are tested for real in tests/vba/TestMTEF.bas.
"""
import re

from test_toc_box_stays_open import form_code
from test_toc_never_deletes_a_line import module_text, procedures

NEW = ("Sh_MathType_Braille_Prefs", "Sh_MathType_Drop_Prefs", "Sh_Picture_Size_Pt", "Sh_Picture_U32BE", "Sh_Read_File_Bytes", "Sh_Points_Text",
       "Sh_MathType_Package", "Sh_MathML_Put_MathType", "Sh_MathType_Is_Installed",
       "Sh_MathType_Format_Equations")
CONVERTER = "Sh_Convert_XML_File_To_Word_Document"


def _procs(repo_root):
    return procedures(module_text(repo_root))


def _body(repo_root, name):
    return "\n".join(_procs(repo_root)[name])


def test_new_procedures_byval_and_private(repo_root):
    procs = _procs(repo_root)
    for name in NEW:
        assert name in procs, f"{name} is missing"
        head, n = procs[name][0], 1
        while head.rstrip().endswith("_"):
            head = head.rstrip()[:-1] + " " + procs[name][n]
            n += 1
        params = head[head.index("(") + 1:head.rindex(")")].split(")")[0]
        for p in filter(None, (x.strip() for x in params.split(","))):
            assert p.startswith("ByVal "), f"{name}: parameter {p!r} is not ByVal"
        # The form calls Sh_MathType_Is_Installed, so it is the one that cannot be Private.
        if name == "Sh_MathType_Is_Installed":
            assert not head.startswith("Private ")
        else:
            assert head.startswith("Private "), f"{name} is not Private"
        assert not any("MsgBox" in l for l in procs[name]), f"{name} uses a bare MsgBox"


def test_new_procedures_have_version_and_example(repo_root):
    lines = module_text(repo_root).splitlines()
    for name in NEW:
        start = next(i for i, l in enumerate(lines)
                     if re.match(r"(Private )?(Sub|Function) %s\(" % name, l))
        n = start
        while lines[n].rstrip().endswith("_"):
            n += 1
        block = []
        for l in lines[n + 1:]:
            if not l.strip().startswith("'"):
                break
            block.append(l.strip())
        assert any(re.match(r"' Version: 1\.0  Date: 10/2[02]/2026", l) for l in block), name
        assert any(l.startswith("'   ") and name in l for l in block), f"{name} has no example"


def test_the_api_is_declared_for_both_vba_versions(repo_root):
    text = module_text(repo_root)
    block = text[text.index("#If VBA7 Then\r\n    Private Declare PtrSafe Function Sh_MT_InitAPI"):]
    block = block[:block.index("#End If")]
    assert block.count('Lib "MathPage.WLL" Alias "MTInitAPI"') == 2
    assert block.count('Lib "MathPage.WLL" Alias "MTTermAPI"') == 2
    assert "#Else" in block and block.count("PtrSafe") == 2


def test_format_equations_connects_first_and_never_stops_the_conversion(repo_root):
    body = _body(repo_root, "Sh_MathType_Format_Equations")
    init = body.index("Sh_MT_InitAPI(0, 30)")
    run = body.index('Application.Run "MTFormatEquations.DoFormatEquations", False, source, "", False, False,')
    term = body.index("Sh_MT_TermAPI()")
    assert init < run < term
    assert body.index("On Error GoTo Failed") < init
    assert "targetDoc.Activate" in body


def test_braille_is_formatted_in_times_12_not_the_users_settings(repo_root):
    """Jerry, 10/22/2026: Format Equations used the transcriber's own MathType settings - set for
    large print - for braille books too. Braille now hands MathType its factory Times New Roman
    12 point in the book's MTPreferences property (source 1) and takes it out again; large print
    keeps the user's settings (source 2). Measured on VistaBuild: whichever source is chosen
    replaces each equation's own settings."""
    body = _body(repo_root, "Sh_MathType_Format_Equations")
    assert "ByVal forBraille As Boolean" in "\n".join(_procs(repo_root)["Sh_MathType_Format_Equations"][:3])
    two = body.index("source = 2")
    brl = body.index("If forBraille Then", two)
    add = body.index('CustomDocumentProperties.Add Name:="MTPreferences"', brl)
    one = body.index("source = 1", add)
    run = body.index('"MTFormatEquations.DoFormatEquations", False, source,', one)
    drop = body.index("If forBraille Then Sh_MathType_Drop_Prefs targetDoc", run)
    assert two < brl < add < one < run < drop
    fail = body[body.index("Failed:"):]
    assert "Sh_MathType_Drop_Prefs targetDoc" in fail, "the property is left behind on an error"

    prefs = _body(repo_root, "Sh_MathType_Braille_Prefs")
    for piece in ('"[Styles]"', '"Text=Times New Roman"', '"Variable=Times New Roman,I"',
                  '"[Sizes]"', '"Full=12 pt"', '"Script=58 %"', '"ScriptScript=42 %"', '"[Spacing]"'):
        assert piece in prefs, piece
    assert "vbCrLf" in prefs

    drop_body = _body(repo_root, "Sh_MathType_Drop_Prefs")
    assert "On Error Resume Next" in drop_body
    assert 'CustomDocumentProperties("MTPreferences").Delete' in drop_body

    conv = "\n".join(_procs(repo_root)[CONVERTER])
    assert "MathType is formatting the equations" in conv and "drawing the equations" not in conv
    assert "could not be reached to format the equations" in conv


def test_is_installed_reads_the_registry_and_cannot_raise(repo_root):
    body = _body(repo_root, "Sh_MathType_Is_Installed")
    assert body.index("On Error Resume Next") < body.index("RegRead")
    assert 'RegRead("HKEY_CLASSES_ROOT\\Equation.DSMT4\\CLSID\\")' in body


def test_mathtype_falls_back_to_a_word_equation_then_the_picture(repo_root):
    put_in = _body(repo_root, "Sh_MathML_Put_In")
    case = put_in[put_in.index('Case "MATHTYPE"'):put_in.index("Case Else")]
    mt = case.index("Sh_MathML_Put_MathType(r, mathml, picPath)")
    eq = case.index("Sh_MathML_Put_Equation(r, xslt, xsl, mathml)")
    pic = case.index("Sh_MathML_Put_Picture(r, picPath)")
    assert mt < eq < pic
    assert "asOMML = asOMML + 1" in case and "asPicture = asPicture + 1" in case
    # MML2OMML.XSL is loaded for the fallback too.
    assert 'If mode = "OMML" Or mode = "MATHTYPE" Then' in put_in
    assert "could not be made into a MathType equation" in put_in


def test_the_package_is_step_zeros_markup(repo_root):
    body = _body(repo_root, "Sh_MathType_Package")
    for piece in ('ProgID=""Equation.DSMT4""', 'pkg:name=""/word/_rels/document.xml.rels""',
                  'Target=""embeddings/oleObject1.bin""', '<v:imagedata r:id=""rIdImg1""',
                  'pkg:name=""/word/embeddings/oleObject1.bin""', 'pkg:compression=""store""'):
        assert piece in body, piece
    put = _body(repo_root, "Sh_MathML_Put_MathType")
    assert "r.InsertXML Sh_MathType_Package(Sh_Base64(ole)" in put


def test_the_converter_formats_after_the_walk_and_after_the_font(repo_root):
    # Format Equations came before the font pass until 10/27/2026; it now runs after the font
    # pass, the repaginate and the save (issue #42 - see test_mathtype_format_after_save.py).
    lines = _procs(repo_root)[CONVERTER]
    text = "\n".join(lines)
    assert 'If Sh_GP_String_3 <> "OMML" And Sh_GP_String_3 <> "MATHTYPE" And Sh_GP_String_3 <> "PICTURE" Then' in text
    walk = text.index("Sh_MathML_Put_In(finalDoc, mathFragments, mathPrefix, mathMode, folderPath, 45, 62)")
    fmt = text.index('If Not Sh_MathType_Format_Equations(finalDoc, (Sh_GP_String_2 = "BRL")) Then')
    font = text.index("Sh_Set_Whole_Document_Font doc")
    assert walk < font < fmt
    # A failure is told in 313, not stopped on.
    note = text[fmt:fmt + 600]
    assert "Format Equations once" in note and "MathType tab" in note
    assert "Exit Sub" not in note
    # Nothing between the walk and the save walks or unlinks the objects.
    save = text.index("dlgSaveAs.Display")
    between = text[walk:save]
    assert "Fields.Unlink" not in between and "InlineShapes" not in between


def test_the_form_always_offers_mathtype_and_asks_400_without_it(repo_root):
    """Jerry, 10/22/2026: "MathType equations" is always on 398. Chosen where MathType is not
    installed, Okay asks VistaType LP (400) whether to go on; Cancel there leaves 398 up, with no
    answer given, for another choice."""
    code = "\n".join(form_code(repo_root, "DN_Equations_Form"))
    assert "MathTypeEquations.Visible" not in code, "the MathType choice can still be hidden"
    assert 'MathTypeEquations.Accelerator = "M"' in code
    # Same font as the other choices (Jerry, testing 3.0.599: it was smaller).
    assert "MathTypeEquations.Font.Name = WordEquations.Font.Name" in code
    assert "MathTypeEquations.Font.Size = WordEquations.Font.Size" in code
    # MathType first and the default, then Word equations, then the pictures (Jerry, 10/22/2026).
    init = code[code.index("Private Sub UserForm_Initialize()"):]
    assert "MathTypeEquations.Value = True" in init and "WordEquations.Value = True" not in init
    order = [init.index(s) for s in ("MathTypeEquations.Top = topFirst", "WordEquations.Top = topSecond",
                                     "EquationPictures.Top = topThird")]
    assert order == sorted(order)
    assert "MathTypeEquations.TabIndex = 0" in init and "WordEquations.TabIndex = 1" in init
    assert "EquationPictures.TabIndex = 2" in init
    okay = code[code.index("Private Sub CmdOkay_Click()"):code.index("Private Sub CmdCancel_Click()")]
    mt = okay[okay.index("ElseIf MathTypeEquations.Value = True Then"):]
    ask = mt.index("If Not Sh_MathType_Is_Installed() Then")
    q = mt.index('Sh_Ask("MathType is not installed on this computer, resulting in an inability "')
    assert '"Do you wish to continue creating MathType equations?"' in mt
    assert re.search(r'"VistaType LP \(400\)"\) Then Exit Sub', mt), "No on 400 does not return to 398"
    answer = mt.index('Sh_GP_String_3 = "MATHTYPE"')
    assert ask < q < answer
    assert "MsgBox" not in code
    frx = (repo_root / "src" / "forms" / "DN_Equations_Form.frx").read_bytes().decode("latin-1")
    assert "MathTypeEquations" in frx and "MathType equations" in frx
    # W, P, O and C are taken by the other four controls; M is the new one's.
    for other in ("WordEquations", "EquationPictures", "CmdOkay", "CmdCancel"):
        assert f'{other}.Accelerator = "M"' not in code


def test_without_mathtype_the_converter_does_not_try_to_draw(repo_root):
    """MathType equations made on a machine without MathType are not handed to Format Equations;
    313 says they show their pictures and cannot be edited here."""
    text = "\n".join(_procs(repo_root)[CONVERTER])
    walk = text.index("Sh_MathML_Put_In(finalDoc, mathFragments, mathPrefix, mathMode, folderPath, 45, 62)")
    check = text.index("If Not Sh_MathType_Is_Installed() Then", walk)
    fmt = text.index('If Not Sh_MathType_Format_Equations(finalDoc, (Sh_GP_String_2 = "BRL")) Then')
    assert walk < check < fmt
    note = text[check:fmt]
    assert "MathType is not installed on this computer" in note and "Else" in note
