"""Review Prodnotes (426), the large print box - Jerry, 11/7/2026.

The Review/Edit/Delete Prodnotes button on Prodnote Review (425) opens a box like the braille
Review Prodnotes (418): Next, Previous, Delete This Prodnote, Done - and in place of Convert This
to TN Style, Convert This to a Box Red Style and Convert This to a Box Blue Style. It needs the
large print template attached before it opens. 418's engine (Dx_Rvp_*) runs both boxes; what is
the large print box's own is Lp_Review_Prodnotes and Lp_Rvp_*. 418's own checks, which cover the
shared engine, are in tests/test_review_prodnotes.py.

Each part below fails silently in Word: a Replace All inside the undo record crashes Word outright
(docs/Undo-Audit.md); a press with no undo record of its own costs a Ctrl+Z per paragraph; a box
opened before the template check stands over a book it cannot convert; Lp_Is_Lp_Template_Attached
ends with End, which takes every box that stays open down with it; and a MsgBox is none of 10
point Tahoma, Okay, or an Alt key.

Lp_Rvp_Tally and Lp_Rvp_Summary, and a Box Red / Box Blue restyle in a scratch book, are tested in
VBA (tests/vba/TestReviewProdnotes.bas).
"""
import re

from test_toc_box_stays_open import form_code
from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

START = "Lp_Review_Prodnotes"
FORM = "Lp_Review_Prodnotes_Form"


def _lp_code(repo_root):
    procs = procedures(module_text(repo_root))
    names = [n for n in procs if n == START or n.startswith("Lp_Rvp_")]
    assert START in names and len(names) >= 8, names
    return {n: [strip_comment(c) for c in procs[n]] for n in names}


def _body(lines):
    return [c.strip() for c in lines if c.strip()]


def test_the_lp_template_is_checked_before_the_box_opens(repo_root):
    body = _body(_lp_code(repo_root)[START])
    assert body[1] == 'Application.Run MacroName:="Sh_Is_Doc_Open"'
    assert body[2] == "If Not Lp_Is_The_Attached_Template_LP Then"
    assert body[3] == 'Sh_Say "VistaType large print template is not attached", "VistaType LP (135)"'
    assert body[4] == "Exit Sub"
    assert body[6] == "Dx_Rvp_Open True", "the box opens only after the check"
    text = "\n".join(body)
    assert "Lp_Is_Lp_Template_Attached" not in text, (
        "Lp_Is_Lp_Template_Attached ends with End, which takes down every box that stays open")
    assert not any(re.fullmatch(r"End", c) for c in body)


def test_the_message_is_the_one_the_other_lp_buttons_give(repo_root):
    """The same words and number as Lp_Is_Lp_Template_Attached's, so the transcriber sees 135."""
    procs = procedures(module_text(repo_root))
    guard = "\n".join(procs["Lp_Is_Lp_Template_Attached"])
    assert 'MsgBox "VistaType large print template is not attached", , "VistaType LP (135)"' in guard


def test_no_replace_all_and_no_undo_clear(repo_root):
    everything = "\n".join("\n".join(v) for v in _lp_code(repo_root).values())
    assert "wdReplaceAll" not in everything and ".Find" not in everything, (
        "a Replace All inside an undo record crashes Word - restyle one paragraph at a time")
    assert "UndoClear" not in everything


def test_each_box_convert_is_one_undo_record(repo_root):
    body = "\n".join(_lp_code(repo_root)["Lp_Rvp_Box_Convert"])
    start = body.index('objUndo.StartCustomRecord "Convert Prodnote to " & styleName')
    work = body.index("Dx_Rvp_Restyle_Run Dx_Rvp_Hits(Dx_Rvp_At), st")
    assert start < work < body.index("objUndo.EndCustomRecord")
    eom = body[body.index("eom:"):]
    assert "If recording Then objUndo.EndCustomRecord" in eom
    assert 'Sh_Report_Error "Lp_Rvp_Box_Convert"' in eom
    # the shared restyle is one paragraph at a time (also held by test_review_prodnotes.py)
    restyle = "\n".join(procedures(module_text(repo_root))["Dx_Rvp_Restyle_Run"])
    assert "For Each para In r.Paragraphs" in restyle and "para.Style = st" in restyle


def test_the_two_buttons_restyle_to_box_red_and_box_blue(repo_root):
    macros = module_text(repo_root)
    assert 'Private Const LP_RVP_RED As String = "Box Red"' in macros
    assert 'Private Const LP_RVP_BLUE As String = "Box Blue"' in macros
    code = _lp_code(repo_root)
    assert _body(code["Lp_Rvp_Box_Red"])[1] == "Lp_Rvp_Box_Convert LP_RVP_RED"
    assert _body(code["Lp_Rvp_Box_Blue"])[1] == "Lp_Rvp_Box_Convert LP_RVP_BLUE"
    body = "\n".join(code["Lp_Rvp_Box_Convert"])
    assert "Set st = Dx_Rvp_Doc.Styles(styleName)" in body
    assert re.search(r"If styleName = LP_RVP_RED Then\s*\n\s*Dx_Rvp_Red = Dx_Rvp_Red \+ 1\s*\n\s*Else"
                     r"\s*\n\s*Dx_Rvp_Blue = Dx_Rvp_Blue \+ 1", body)
    # and it moves on to the next prodnote, as 418's Convert does
    assert body.index("Dx_Rvp_Busy = False\n    Dx_Rvp_Move True, False") < body.index("eom:")


def test_the_box_styles_are_in_the_lp_template(repo_root):
    """The names must be the template's own, character for character, or every press is an error report - there is no missing-style check."""
    import zipfile
    with zipfile.ZipFile(repo_root / "LargePrintTemplate.dotx") as z:
        styles = z.read("word/styles.xml").decode("utf-8")
    for name in ("Box Red", "Box Blue"):
        m = re.search(rf'<w:style w:type="(\w+)"[^>]*>\s*<w:name w:val="{name}"/>', styles)
        assert m and m.group(1) == "paragraph", f"{name} is not a paragraph style in the template"


def test_no_missing_style_check(repo_root):
    """Jerry, 11/7/2026: the LP template is required, and it brings both Box styles."""
    text = "\n".join(c.strip() for c in _lp_code(repo_root)["Lp_Rvp_Box_Convert"])
    assert "(427)" not in text and "If st Is Nothing" not in text
    assert text.index("On Error GoTo eom") < text.index("Set st = Dx_Rvp_Doc.Styles(styleName)")


def test_only_the_large_print_box_converts_to_a_box(repo_root):
    body = _body(_lp_code(repo_root)["Lp_Rvp_Box_Convert"])
    assert "If Not Dx_Rvp_ForLP Then Exit Sub" in body
    assert body.index("If Not Dx_Rvp_IsOn Then Exit Sub") < body.index("Dx_Rvp_BindKeys")


def test_no_tn_button_on_the_large_print_box(repo_root):
    code = "\n".join(form_code(repo_root, FORM))
    assert "Dx_Rvp_Convert" not in code and "ConvertButton" not in code
    for button, macro in (("NextButton", "Dx_Rvp_Next"), ("PrevButton", "Dx_Rvp_Previous"),
                          ("DeleteButton", "Dx_Rvp_Delete"), ("BoxRedButton", "Lp_Rvp_Box_Red"),
                          ("BoxBlueButton", "Lp_Rvp_Box_Blue"), ("DoneButton", "Dx_Rvp_Done")):
        m = re.search(rf"Private Sub {button}_Click\(\)(.*?)End Sub", code, re.S)
        assert m and m.group(1).strip() == macro, button
    alt = "\n".join(_lp_code(repo_root)["Lp_Rvp_AltKey"])
    assert '"T"' not in alt and "Dx_Rvp_Convert" not in alt


def test_the_form_has_its_number(repo_root):
    raw = (repo_root / "src" / "forms" / f"{FORM}.frm").read_bytes().decode("latin-1")
    assert 'Caption         =   "Review Prodnotes (426)"' in raw
    assert (repo_root / "src" / "forms" / f"{FORM}.frx").exists()
    q = re.search(r"Sub UserForm_QueryClose\(.*?\)(.*?)End Sub", raw, re.S)
    assert q and "Cancel = True" in q.group(1) and "Dx_Rvp_Done" in q.group(1)


def test_every_message_is_numbered_and_goes_through_sh_say(repo_root):
    everything = "\n".join("\n".join(v) for v in _lp_code(repo_root).values())
    assert "MsgBox" not in everything
    says = re.findall(r"Sh_Say\b", everything)
    assert len(says) == 1, says   # 135
    # Done and none-left are 418's 421 and 420, titled VistaType LP for a large print book
    procs = procedures(module_text(repo_root))
    for sub, n in (("Dx_Rvp_Done", 421), ("Dx_Rvp_None_Left", 420)):
        body = "\n".join(procs[sub])
        assert f'IIf(Dx_Rvp_IsLP, "VistaType LP ({n})", "Braille Macros ({n})")' in body
        assert "Dx_Rvp_Run_Summary()" in body, f"{sub} does not count Box Red and Box Blue"
    run = "\n".join(procs["Dx_Rvp_Run_Summary"])
    assert "Lp_Rvp_Summary(Dx_Rvp_Deleted, Dx_Rvp_Red, Dx_Rvp_Blue)" in run


def test_the_shared_engine_closes_and_shows_the_box_that_is_up(repo_root):
    procs = procedures(module_text(repo_root))
    finish = "\n".join(procs["Dx_Rvp_Finish"])
    assert re.search(r"If Dx_Rvp_ForLP Then\s*\n\s*Lp_Review_Prodnotes_Form\.Hide\s*\n"
                     r"\s*Unload Lp_Review_Prodnotes_Form", finish)
    assert finish.index("Dx_Rvp_IsOn = False") < finish.index("Unload Lp_Review_Prodnotes_Form")
    assert "Dx_Rvp_ForLP = False" not in finish, (
        "the flag stays, so the box leaves the list of boxes under the name it joined with")
    count = "\n".join(procs["Dx_Rvp_Show_Count"])
    assert "Lp_Review_Prodnotes_Form.TallyText.Caption = Lp_Rvp_Tally(Dx_Rvp_Deleted, Dx_Rvp_Red, Dx_Rvp_Blue)" in count
    opener = "\n".join(procs["Dx_Rvp_Open"])
    assert "If ActiveDocument Is Dx_Rvp_Doc And Dx_Rvp_ForLP = forLP Then" in opener, (
        "the other box's button must end this run, not bring this box back")


def test_f6_safety_nets_leave_the_list_under_their_own_name(repo_root):
    code = _lp_code(repo_root)
    toggle = "\n".join(code["Lp_Rvp_ToggleFocus"])
    assert re.search(r'If Not Dx_Rvp_IsOn Or Not Dx_Rvp_ForLP Then\s*\n\s*Sh_Box_Closed "lrvp"', toggle)
    assert "Sh_Focus_Modeless_Form LP_RVP_TITLE" in toggle
    alt = "\n".join(code["Lp_Rvp_AltKey"])
    assert re.search(r'If Not Dx_Rvp_IsOn Or Not Dx_Rvp_ForLP Then\s*\n\s*Sh_Box_Closed "lrvp"', alt)
    procs = procedures(module_text(repo_root))
    for sub in ("Dx_Rvp_ToggleFocus", "Dx_Rvp_AltKey"):
        assert re.search(r'If Not Dx_Rvp_IsOn Or Dx_Rvp_ForLP Then\s*\n\s*Sh_Box_Closed "rvp"',
                         "\n".join(procs[sub])), sub


def test_every_helper_takes_its_arguments_byval(repo_root):
    for name, lines in _lp_code(repo_root).items():
        head = lines[0]
        args = re.search(r"\((.*)\)", head)
        if not args or not args.group(1).strip():
            continue
        for arg in args.group(1).split(","):
            assert arg.strip().startswith("ByVal "), f"{name}: {arg.strip()} is not ByVal"
