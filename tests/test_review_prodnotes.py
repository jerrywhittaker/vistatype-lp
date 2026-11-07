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

START = "Dx_Review_Prodnotes"
FORM = "Dx_Review_Prodnotes_Form"


def _rvp_code(repo_root):
    procs = procedures(module_text(repo_root))
    names = [n for n in procs if n == START or n.startswith("Dx_Rvp_")]
    assert START in names and len(names) > 20, names
    return {n: [strip_comment(c) for c in procs[n]] for n in names}


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
    assert '"BANA Braille"' in text and "Dx_Rvp_Doc.AttachedTemplate.Name" in text


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
    assert any('Sh_Box_On_Top() <> "rvp"' in c for c in watch)


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
