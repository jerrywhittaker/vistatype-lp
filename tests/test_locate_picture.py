"""Locate This Picture (404) - Jerry, 11/1/2026.

The picture on the clipboard - copied in the book, or its file copied in File Explorer from a
DAISY or NIMAS images folder - is found in the selection or the whole book, and a box that stays
open steps through the copies. Started from Background and Picture Tools (344).

The clipboard is READ, never changed: the project has removed Copy/Paste and clipboard clearing
from macros before because they destroyed the transcriber's clipboard (Lp_Same_Pic_Data's note).
So the one clipboard operation is a Paste into a hidden document of its own, and the picture is
read there by Lp_Same_Pic_Data - the one way every picture is read for a comparison. Measured on
vistabuild 11/1/2026: the clipboard held the same thing before and after, and a file copied in
File Explorer read back the same bytes as the picture in a converted book.

Not unit-testable in VBA beyond the step arithmetic (tests/vba/TestLocatePicture.bas): the rest
reaches the Selection and Documents, which hang the headless runner. So this checks the source.
"""
import re

from test_toc_box_stays_open import form_code
from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

START = "Lp_Locate_Picture"


def _locp_code(repo_root):
    procs = procedures(module_text(repo_root))
    names = [n for n in procs if n == START or n.startswith("Lp_Locp_")]
    assert START in names and len(names) > 10, names
    return {n: [strip_comment(c) for c in procs[n]] for n in names}


def test_the_clipboard_is_only_read_by_a_paste_into_a_hidden_document(repo_root):
    code = _locp_code(repo_root)
    everything = "\n".join("\n".join(v) for v in code.values())
    for banned in (r"\.Copy\b", r"\.Cut\b", r"PutInClipboard", r"SetText", r"MS_SafeClearClipboard",
                   r"\bDataObject\b", r"OpenClipboard", r"EmptyClipboard"):
        assert not re.search(banned, everything), f"Locate This Picture touches the clipboard: {banned}"
    pastes = [c for c in everything.splitlines() if re.search(r"\.Paste", c)]
    assert [p.strip() for p in pastes] == ["pasted.Content.Paste"], pastes
    start = "\n".join(code[START])
    assert re.search(r"Set pasted = Documents\.Add\(Visible:=False\)", start)
    assert start.index("Set pasted = Documents.Add") < start.index("pasted.Content.Paste")


def test_one_rule_for_the_same_picture(repo_root):
    start = "\n".join(_locp_code(repo_root)[START])
    assert re.search(r"refData = Lp_Same_Pic_Data\(pasted\.InlineShapes\(1\), reader\)", start)
    assert re.search(r"Lp_Same_Pic_Find\(r\.InlineShapes, Nothing, \"\", refData, reader, -1, -1\)", start), (
        "the copies must be found by the same procedure 375 uses, with no model picture left out")
    assert "Lp_Same_Pic_Is_Picture(pasted.InlineShapes(1))" in start, (
        "a non-picture file pastes as an icon - one in-line shape that is not a picture")


def test_every_refusal_is_a_numbered_message(repo_root):
    start = "\n".join(_locp_code(repo_root)[START])
    for n in (405, 406, 407, 408, 412):
        assert f'"VistaType LP ({n})"' in start, n
    assert "Sh_Say" in start and "MsgBox" not in start
    both = (start.index("If picCount > 1 Then"), start.index("If Not isPic Then"))
    assert both[0] < both[1], "two pictures must be told apart from no picture"


def test_whole_book_unless_something_is_selected(repo_root):
    start = "\n".join(_locp_code(repo_root)[START])
    assert re.search(r"If r\.End - r\.Start < 2 Then Set r = book\.Content", start)


def test_a_deleted_copy_is_stepped_over(repo_root):
    code = _locp_code(repo_root)
    move = "\n".join(code["Lp_Locp_Move"])
    assert "Lp_Locp_Live(Lp_Locp_Hits(n))" in move and "Do While tries < total" in move
    assert '"VistaType LP (411)"' in move
    live = "\n".join(code["Lp_Locp_Live"])
    assert "On Error Resume Next" in live and "r.End - r.Start <> 1" in live


def test_the_box_opens_from_344_through_the_bridge(repo_root):
    code = "\n".join(form_code(repo_root, "Lp_Bakgrnd_Picture_Menu_Form"))
    m = re.search(r"Sub LocatePictureButton_Click\(\)(.*?)End Sub", code, re.S)
    assert m, "344 has no Locate the Picture on the Clipboard button"
    body = m.group(1)
    assert 'Sh_BridgeTargetMacro = "Lp_Locate_Picture"' in body
    assert 'Application.OnTime Now, "Sh_StartSpinnerBridge"' in body


def test_the_box_watches_the_cursor(repo_root):
    events = (repo_root / "src" / "vba" / "VtEvents.cls").read_bytes().decode("latin-1")
    handler = re.search(r"Sub App_WindowSelectionChange\(.*?\)(.*?)End Sub", events, re.S)
    assert handler and re.search(r"Lp_Locp_Watch_Cursor\s+Sel\b", handler.group(1))
    watch = _locp_code(repo_root)["Lp_Locp_Watch_Cursor"]
    say = next(n for n, c in enumerate(watch) if "Sh_Say" in c)
    for guard in (r"If Not Lp_Locp_IsOn Then Exit Sub", r"Lp_Locp_Busy Or",
                  r"If at = Lp_Locp_Warned_At Then Exit Sub"):
        at = next((n for n, c in enumerate(watch) if re.search(guard, c)), None)
        assert at is not None and at < say, guard
    assert '"VistaType LP (409)"' in "\n".join(watch)
