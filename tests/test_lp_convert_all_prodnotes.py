"""Convert All Prodnotes to LP Style - Jerry, 11/8/2026.

The Convert All Prodnotes to LP Style button on Prodnote Review (425) asks "To which style should
all prodnotes be converted?" on Convert All Prodnotes (427) - Box Red Style or Box Blue Style - and
gives every prodnote in the book that style (Lp_Convert_All_Prodnotes_To_Box). It needs the large
print template, checked before the question, with the same message (135) as Review Prodnotes (426).

Each part below fails silently in Word: a Replace All inside the undo record crashes Word outright
(docs/Undo-Audit.md); a restyle with no undo record costs a Ctrl+Z per paragraph; a red X that
answers converts the book when the transcriber meant to back out; a handler-less button does
nothing at all; and a MsgBox is none of 10 point Tahoma, Okay, or an Alt key.

The restyle itself, the one-press undo and message 428's words are tested in VBA
(tests/vba/TestReviewProdnotes.bas). The form cannot be shown headlessly.
"""
import re

from test_toc_box_stays_open import form_code
from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

START = "Lp_Convert_All_Prodnotes_To_Box"
FORM = "Lp_Prodnote_Box_Style_Form"
MENU = "Lp_Prodnotes_Menu_Form"


def _code(repo_root):
    procs = procedures(module_text(repo_root))
    names = [n for n in procs if n == START or n.startswith("Lp_Cap_")]
    assert START in names and len(names) == 3, names
    return {n: [strip_comment(c) for c in procs[n]] for n in names}


def _body(lines):
    return [c.strip() for c in lines if c.strip()]


def _raw_form(repo_root, name):
    return (repo_root / "src" / "forms" / f"{name}.frm").read_bytes().decode("latin-1")


def test_convert_all_on_425_runs_the_macro(repo_root):
    code = _raw_form(repo_root, MENU)
    handler = re.search(r"Private Sub CommandButton2_Click\(\)\r\n(.*?)End Sub", code, re.S)
    assert handler and handler.group(1).strip() == f'Run_After_Menu "{START}"'


def test_the_lp_template_is_checked_before_anything_else(repo_root):
    body = _body(_code(repo_root)[START])
    i = body.index('Application.Run MacroName:="Sh_Is_Doc_Open"')
    assert body[i + 1] == "If Not Lp_Is_The_Attached_Template_LP Then"
    assert body[i + 2] == 'Sh_Say "VistaType large print template is not attached", "VistaType LP (135)"'
    assert body[i + 3] == "Exit Sub"
    text = "\n".join(body)
    assert text.index("Lp_Is_The_Attached_Template_LP") < text.index(f"{FORM}.Show")
    assert "Lp_Is_Lp_Template_Attached" not in text, "that sub ends with End"
    assert not any(re.fullmatch(r"End", c) for c in body)


def test_no_prodnotes_says_419_before_the_question(repo_root):
    text = "\n".join(_body(_code(repo_root)[START]))
    say = text.index('Sh_Say Dx_Rvp_None_Found_Text(), "VistaType LP (419)"')
    assert text.index("If Not Sh_Document_Has_Prodnotes(book) Then") < say < text.index(f"{FORM}.Show")
    # the same words as Review Prodnotes' 419, from one place
    opener = "\n".join(procedures(module_text(repo_root))["Dx_Rvp_Open"])
    assert 'Sh_Say Dx_Rvp_None_Found_Text(), IIf(isLP, "VistaType LP (419)", "Braille Macros (419)")' in opener


def test_cancel_and_the_red_x_change_nothing(repo_root):
    body = _body(_code(repo_root)[START])
    show = body.index(f"{FORM}.Show")
    assert body[show - 1] == 'Lp_Cap_Style = ""'
    assert body[show + 1] == "styleName = Lp_Cap_Style"
    stop = body.index("If styleName <> LP_RVP_RED And styleName <> LP_RVP_BLUE Then Exit Sub")
    assert stop > show
    assert stop < body.index("Sh_Save_User_Position") < body.index("Set hits = Dx_Rvp_Collect(book)")
    form = procedures("\n".join(form_code(repo_root, FORM)))
    assert re.search(r'Lp_Cap_Style\s*=\s*""', "\n".join(form["CancelButton_Click"]))
    assert re.search(r'Lp_Cap_Style\s*=\s*""', "\n".join(form["UserForm_Initialize"]))
    q = "\n".join(l for l in form["UserForm_QueryClose"] if not l.strip().startswith("'"))
    assert re.search(r'CloseMode\s*=\s*vbFormControlMenu\s+Then\s+Lp_Cap_Style\s*=\s*""', q)
    assert not re.search(r"\bCancel\s*=\s*True\b|\bUnload\b", q)


def test_okay_answers_with_the_two_box_styles(repo_root):
    form = procedures("\n".join(form_code(repo_root, FORM)))
    okay = "\n".join(form["OkayButton_Click"])
    assert re.search(r'If BoxBlueOption\.Value Then\s*\n\s*Lp_Cap_Style = "Box Blue"\s*\n\s*Else'
                     r'\s*\n\s*Lp_Cap_Style = "Box Red"', okay)
    macros = module_text(repo_root)
    assert 'Private Const LP_RVP_RED As String = "Box Red"' in macros
    assert 'Private Const LP_RVP_BLUE As String = "Box Blue"' in macros
    assert re.search(r"^Public Lp_Cap_Style As String", macros, re.M)


def test_the_form_has_its_number_question_and_default(repo_root):
    raw = _raw_form(repo_root, FORM)
    assert 'Caption         =   "Convert All Prodnotes (427)"' in raw
    assert (repo_root / "src" / "forms" / f"{FORM}.frx").exists()
    init = "\n".join(procedures("\n".join(form_code(repo_root, FORM)))["UserForm_Initialize"])
    assert 'QuestionLabel.Caption = "To which style should all prodnotes be converted?"' in init
    assert "BoxRedOption.Value = True" in init
    assert "Me.StartUpPosition = 0" in init and "Application.Left" in init


def test_the_form_controls_and_alt_keys(repo_root):
    """Read from the .frx the same way test_form_handlers_match_controls does."""
    from test_form_handlers_match_controls import control_names
    names = set(control_names(repo_root / "src" / "forms" / f"{FORM}.frx"))
    assert {"QuestionLabel", "BoxRedOption", "BoxBlueOption", "OkayButton", "CancelButton"} <= names


def test_every_restyle_is_inside_one_undo_record(repo_root):
    code = _code(repo_root)
    body = "\n".join(code["Lp_Cap_Convert_All"])
    start = body.index('objUndo.StartCustomRecord "Convert All Prodnotes to " & st.NameLocal')
    loop = body.index("For i = 1 To hits.count")
    work = body.index("Dx_Rvp_Restyle_Run hits(i), st")
    end = body.index("objUndo.EndCustomRecord")
    assert start < loop < work < end
    assert body.count("StartCustomRecord") == 1
    eom = body[body.index("eom:"):]
    assert "If recording Then objUndo.EndCustomRecord" in eom and "Err.Raise errNum" in eom
    restyle = "\n".join(procedures(module_text(repo_root))["Dx_Rvp_Restyle_Run"])
    assert "For Each para In r.Paragraphs" in restyle and "para.Style = st" in restyle


def test_no_replace_all_no_undo_clear_and_no_style_deletion(repo_root):
    everything = "\n".join("\n".join(v) for v in _code(repo_root).values())
    assert "wdReplaceAll" not in everything and ".Find" not in everything
    assert "UndoClear" not in everything
    assert "Sh_Set_Prodnote_Style_Visibility" not in everything, (
        "deleting the Prodnote style after the record would take a second Ctrl+Z")


def test_the_screen_bar_and_position_are_put_back_on_every_path(repo_root):
    body = "\n".join(_body(_code(repo_root)[START]))
    main, eom = body.split("eom:")
    for part in (main, eom):
        assert "Application.ScreenUpdating = su_Prev" in part
        assert "Sh_Progress_Close" in part
        assert "Sh_Return_User_To_Start_Position" in part
    assert main.index("Sh_Progress_Open") > main.index(f"{FORM}.Show"), "the bar comes after the question"
    assert eom.index("errNum = Err.Number") < eom.index("Sh_Progress_Close")
    assert 'Sh_Report_Error "Lp_Convert_All_Prodnotes_To_Box", errNum, errText' in eom


def test_every_message_is_numbered_and_goes_through_sh_say(repo_root):
    everything = "\n".join("\n".join(v) for v in _code(repo_root).values())
    assert "MsgBox" not in everything
    says = re.findall(r"^\s*Sh_Say .*$", everything, re.M)
    assert [re.search(r'\((\d{3})\)"', s).group(1) for s in says] == ["135", "419", "428"]
    assert 'Sh_Say Lp_Cap_Summary(n, styleName), "VistaType LP (428)"' in everything


def test_every_helper_takes_its_arguments_byval(repo_root):
    for name, lines in _code(repo_root).items():
        args = re.search(r"\((.*)\)", lines[0])
        if not args or not args.group(1).strip():
            continue
        for arg in args.group(1).split(","):
            assert arg.strip().startswith("ByVal "), f"{name}: {arg.strip()} is not ByVal"
