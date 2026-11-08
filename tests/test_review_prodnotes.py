"""Review Prodnotes (418) - Jerry, 11/7/2026.

The Review/Edit/Delete Prodnotes button on Prodnotes and Alt-Text (417) opens a box that stays
open and goes through the prodnotes one at a time: Next, Previous, Delete This Prodnote, Convert
This to TN Style, Done. Editing is done in the book, through F6. One prodnote is a run of Prodnote
paragraphs in a row.

Each part below fails silently in Word: a Replace All inside the undo record crashes Word
outright (docs/Undo-Audit.md, docs/Reported-Errors.md 9/2/2026 and 9/3/2026); a press with no
undo record of its own costs as many Ctrl+Z presses as it has paragraphs; a Convert that skips the
BANA check restyles a large print book to a style it does not have; and a MsgBox is none of 10
point Tahoma, Okay, or an Alt key.

The step arithmetic, the labels and the run-joining rule are tested in VBA
(tests/vba/TestReviewProdnotes.bas). The rest reaches the Selection and the book, so this checks
the source.
"""
import re

from test_toc_box_stays_open import form_code
from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

# The opening, shared since 11/7/2026 with the large print box (426, tests/test_lp_review_prodnotes.py):
# Dx_Review_Prodnotes is the braille button's macro and calls Dx_Rvp_Open False.
START = "Dx_Rvp_Open"
FORM = "Dx_Review_Prodnotes_Form"


def _rvp_code(repo_root):
    procs = procedures(module_text(repo_root))
    names = [n for n in procs if n == START or n.startswith("Dx_Rvp_")]
    assert START in names and len(names) > 20, names
    return {n: [strip_comment(c) for c in procs[n]] for n in names}


def test_the_braille_button_opens_the_braille_box(repo_root):
    body = [strip_comment(c).strip() for c in procedures(module_text(repo_root))["Dx_Review_Prodnotes"]]
    body = [c for c in body if c]
    assert body[1:-1] == ['Application.Run MacroName:="Sh_Is_Doc_Open"', "Dx_Rvp_Open False"], body
    start = "\n".join(_rvp_code(repo_root)[START])
    assert re.search(r"If Dx_Rvp_ForLP Then\s*\n\s*Lp_Review_Prodnotes_Form\.Show vbModeless\s*\n"
                     r"\s*Else\s*\n\s*Dx_Review_Prodnotes_Form\.Show vbModeless", start)
    assert start.index("Dx_Rvp_ForLP = forLP") < start.index("Dx_Rvp_BindKeys")


def test_no_replace_all_anywhere_in_the_box(repo_root):
    everything = "\n".join("\n".join(v) for v in _rvp_code(repo_root).values())
    assert "wdReplaceAll" not in everything and ".Find" not in everything, (
        "a Replace All inside an undo record crashes Word - restyle and delete one Range at a time")
    assert "UndoClear" not in everything, "the transcriber's own undo list must be left alone"


def test_each_delete_and_convert_is_one_undo_record(repo_root):
    code = _rvp_code(repo_root)
    for press, work, name in (("Dx_Rvp_Delete", "Dx_Rvp_Delete_Run", "Delete Prodnote"),
                              ("Dx_Rvp_Convert", "Dx_Rvp_Restyle_Run", "Convert Prodnote to TN Style")):
        body = "\n".join(code[press])
        start = body.index(f'objUndo.StartCustomRecord "{name}"')
        assert start < body.index(work) < body.index("objUndo.EndCustomRecord"), press
        # the error path closes a record left open, or Word holds it open across the transcriber's typing
        eom = body[body.index("eom:"):]
        assert "If recording Then objUndo.EndCustomRecord" in eom, press
        assert re.search(r"Sh_Report_Error \"" + press + r"\"", eom), press


def test_restyle_is_one_paragraph_at_a_time(repo_root):
    body = "\n".join(_rvp_code(repo_root)["Dx_Rvp_Restyle_Run"])
    assert "For Each para In r.Paragraphs" in body and "para.Style = st" in body


def test_convert_asks_for_the_bana_template_first(repo_root):
    body = [c.strip() for c in _rvp_code(repo_root)["Dx_Rvp_Convert"]]
    text = "\n".join(body)
    say = 'Sh_Say "The BANA template is not attached.", "Braille Macros (394)"'
    assert say in body
    assert text.index(say) < text.index("StartCustomRecord")
    assert "Dx_Rvp_Is_BANA_Name(tmplName)" in text and "Dx_Rvp_Doc.AttachedTemplate.Name" in text


def test_a_bana_book_is_put_in_print_layout(repo_root):
    """Jerry, 11/7/2026: with a BANA template attached, the box shows the book in Print Layout -
    when it opens, and when the button brings it back. Other books keep their view."""
    code = _rvp_code(repo_root)
    start = "\n".join(code[START])
    assert start.count("Dx_Rvp_Print_Layout") == 2, "both the opening and the bring-back"
    assert start.index("Dx_Rvp_Print_Layout book") < start.index("Dx_Review_Prodnotes_Form.Show vbModeless")
    layout = "\n".join(code["Dx_Rvp_Print_Layout"])
    assert "Dx_Rvp_Is_BANA_Name(tmplName)" in layout
    assert "View.Type = wdPrintView" in layout
    assert '"BANA Braille"' in "\n".join(code["Dx_Rvp_Is_BANA_Name"])


def test_a_prodnote_is_a_run_of_paragraphs(repo_root):
    collect = "\n".join(_rvp_code(repo_root)["Dx_Rvp_Collect"])
    assert "For Each para In doc.Paragraphs" in collect
    assert re.search(r"If Dx_Rvp_Joins\(prevProd, prevEndsCell, prevInTable, inTable\) Then\s*\n"
                     r"\s*cur\.End = para\.Range\.End", collect)
    joins = "\n".join(_rvp_code(repo_root)["Dx_Rvp_Joins"])
    assert "If Not prevProd Then Exit Function" in joins and "If prevEndsCell Then Exit Function" in joins


def test_an_emptied_cell_or_last_paragraph_goes_back_to_normal(repo_root):
    body = "\n".join(_rvp_code(repo_root)["Dx_Rvp_Delete_Run"])
    assert body.index("r.Delete") < body.index(".Style = doc.Styles(wdStyleNormal)")
    assert "Len(Dx_Rvp_Bare(.Range.Text)) = 0" in body, "only an EMPTY paragraph is reset"


def test_every_message_is_numbered_and_goes_through_sh_say(repo_root):
    code = _rvp_code(repo_root)
    everything = "\n".join("\n".join(v) for v in code.values())
    assert "MsgBox" not in everything
    for n in (419, 420, 421, 422):
        assert f'IIf(isLP, "VistaType LP ({n})", "Braille Macros ({n})")' in everything \
            or f'IIf(Dx_Rvp_IsLP, "VistaType LP ({n})", "Braille Macros ({n})")' in everything, n
    assert '"Braille Macros (423)"' in everything


def test_done_and_the_x_close_the_box_the_one_way(repo_root):
    code = _rvp_code(repo_root)
    for sub in ("Dx_Rvp_Done", "Dx_Rvp_None_Left", "Dx_Rvp_Ready"):
        assert any(re.search(r"^\s*Dx_Rvp_Finish\s*$", c) for c in code[sub]), sub
    finish = "\n".join(code["Dx_Rvp_Finish"])
    assert finish.index("Dx_Rvp_IsOn = False") < finish.index("Unload Dx_Review_Prodnotes_Form")
    assert finish.index(".Hide") < finish.index("Unload")
    q = re.search(r"Sub UserForm_QueryClose\(.*?\)(.*?)End Sub", "\n".join(form_code(repo_root, FORM)), re.S)
    assert q and "Cancel = True" in q.group(1) and "Dx_Rvp_Done" in q.group(1)


def test_the_box_opens_with_the_book_holding_the_keyboard(repo_root):
    start = "\n".join(_rvp_code(repo_root)[START])
    show = start.index("Dx_Review_Prodnotes_Form.Show vbModeless")
    assert start.index("Dx_Rvp_IsOn = True") < start.index("Dx_Rvp_BindKeys") < show
    assert "Sh_Focus_Document Dx_Rvp_Doc" in start[show:]


def test_the_cursor_watch_is_hooked_and_cheap_while_down(repo_root):
    events = (repo_root / "src" / "vba" / "VtEvents.cls").read_bytes().decode("latin-1")
    handler = re.search(r"Sub App_WindowSelectionChange\(.*?\)(.*?)End Sub", events, re.S)
    assert handler and re.search(r"Dx_Rvp_Watch_Cursor\s+Sel\b", handler.group(1))
    watch = [c.strip() for c in _rvp_code(repo_root)["Dx_Rvp_Watch_Cursor"] if c.strip()]
    assert watch[1] == "Dim i As Long, n As Long, total As Long, pos As Long"
    assert watch[2] == "If Not Dx_Rvp_IsOn Then Exit Sub", "the first test must be one Boolean"
    assert any("Sh_Box_On_Top() <> Dx_Rvp_Who()" in c for c in watch)
    who = "\n".join(_rvp_code(repo_root)["Dx_Rvp_Who"])
    assert re.search(r'If Dx_Rvp_ForLP Then\s*\n\s*Dx_Rvp_Who = "lrvp"\s*\n\s*Else\s*\n\s*Dx_Rvp_Who = "rvp"', who)


def test_every_helper_takes_its_arguments_byval(repo_root):
    for name, lines in _rvp_code(repo_root).items():
        head = lines[0]
        n = 0
        while head.rstrip().endswith("_"):
            n += 1
            head = head.rstrip()[:-1] + " " + lines[n].strip()
        args = re.search(r"\((.*)\)", head)
        if not args or not args.group(1).strip():
            continue
        for arg in args.group(1).split(","):
            assert arg.strip().startswith("ByVal "), f"{name}: {arg.strip()} is not ByVal"


def test_done_and_none_left_drop_the_prodnote_style_when_none_is_left(repo_root):
    """Jerry, 11/8/2026: when the user selects Done and there are no more prodnotes left in the
    book, the Prodnote style disappears from the Styles pane. Done (the button, Alt+O and the red
    X through QueryClose) and the none-left end (420), on both boxes, 418 and 426: the box closed
    first, so every undo record it opened is closed, then the style dropped, then the message."""
    code = _rvp_code(repo_root)
    for sub in ("Dx_Rvp_Done", "Dx_Rvp_None_Left"):
        body = [c.strip() for c in code[sub] if c.strip()]
        assert "Set book = Dx_Rvp_Doc" in body, sub
        grab, finish = body.index("Set book = Dx_Rvp_Doc"), body.index("Dx_Rvp_Finish")
        drop = body.index("Dx_Rvp_Drop_Unused_Style book")
        say = next(i for i, c in enumerate(body) if c.startswith("Sh_Say "))
        assert grab < finish < drop < say, sub
    drop = [c.strip() for c in code["Dx_Rvp_Drop_Unused_Style"] if c.strip()]
    assert "If Not Dx_Rvp_Style_Can_Go(book) Then Exit Sub" in drop
    assert drop.index("If Not Dx_Rvp_Style_Can_Go(book) Then Exit Sub") \
        < drop.index("If inFront Then Sh_Set_Prodnote_Style_Visibility")
    assert "inFront = (ActiveDocument Is book)" in drop
    can = [c.strip() for c in code["Dx_Rvp_Style_Can_Go"] if c.strip()]
    assert "If Not Lp_Rst_Doc_Is_Open(book) Then Exit Function" in can   # a closed book is skipped
    assert "Dx_Rvp_Style_Can_Go = Not Sh_Document_Has_Prodnotes(book)" in can
    # A book closed during the run ends through Dx_Rvp_Ready (422) and is not touched.
    assert not any("Dx_Rvp_Drop_Unused_Style" in c for c in code["Dx_Rvp_Ready"])
    # Convert All Prodnotes (427) is left as Jerry accepted it.
    procs = procedures(module_text(repo_root))
    assert not any("Dx_Rvp_Drop_Unused_Style" in c for c in procs["Lp_Convert_All_Prodnotes_To_Box"])
