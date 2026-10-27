"""The progress bar must stay in sight while MathType formats the equations.

Jerry, 10/26/2026, on 3.0.610: Convert DAISY/NIMAS to Word with "MathType equations" chosen and
a blank Document1 open at the start. Once the converted book (Document2) came forward for
MathType's Format Equations, the progress bar could not be seen; it showed only when the pointer
rested on Document1 on the taskbar.

Word 365 gives each document a window of its own, and a modeless UserForm belongs to the window
that was in front when it was shown: Sh_Progress_Open shows the bar over Document1, so when
Sh_MathType_Format_Equations brings the book forward (Format Equations works only on
ActiveDocument), the bar goes behind the book. Document1 must not be brought forward instead -
it would cover the book, which has to be in front and painted for MathType (issue #39).

The cure: Sh_Progress_Over hands the bar to the book's window (SetWindowLong(Ptr) with
GWL_HWNDPARENT) immediately before targetDoc.Activate, and Sh_Progress_Show - the converter's
next step after Format Equations, whichever way the function ends - hands it back to the window
of the document it was opened over, before bringing that document forward.

Whether the bar is seen needs Word on a screen and is not tested; these are checks on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

FUNC = "Sh_MathType_Format_Equations"
CONVERTER = "Sh_Convert_XML_File_To_Word_Document"


def _key(line):
    return " ".join(line.lower().split())


def _code(lines):
    return [l.strip() for l in lines if l.strip() and not l.strip().startswith("'")]


def _bar_module(repo_root):
    return (repo_root / "src" / "vba" / "ShNonModalMessage.bas").read_bytes().decode("utf-8")


def test_the_bar_is_handed_to_the_book_just_before_it_comes_forward(repo_root):
    code = _code(procedures(module_text(repo_root))[FUNC])
    acts = [i for i, l in enumerate(code) if _key(l) == "targetdoc.activate"]
    assert len(acts) == 1, f"expected one targetDoc.Activate in {FUNC}, found {len(acts)}"
    a = acts[0]
    assert a >= 1 and _key(code[a - 1]) == "sh_progress_over targetdoc", (
        f"{FUNC}: the statement before targetDoc.Activate must be Sh_Progress_Over targetDoc, "
        f"or the bar stays with Document1's window and is hidden behind the book for the whole "
        f"of Format Equations (Jerry, 3.0.610); found {code[max(a - 1, 0)]!r}")


def test_show_hands_the_bar_back_before_bringing_its_document_forward(repo_root):
    code = [_key(l) for l in _code(procedures(_bar_module(repo_root))["Sh_Progress_Show"])]
    assert "sh_progress_over sh_progress_doc" in code, (
        "Sh_Progress_Show must hand the bar back to its own document's window "
        "(Sh_Progress_Over Sh_Progress_Doc)")
    assert "sh_focus_document sh_progress_doc" in code
    assert code.index("sh_progress_over sh_progress_doc") < \
        code.index("sh_focus_document sh_progress_doc"), (
        "Sh_Progress_Show must hand the bar back BEFORE bringing the document forward, or the "
        "document comes forward without its bar")


def test_the_converter_shows_the_bar_again_after_format_equations(repo_root):
    code = [_key(l) for l in _code(procedures(module_text(repo_root))[CONVERTER])]
    calls = [i for i, l in enumerate(code) if "sh_mathtype_format_equations(" in l]
    assert len(calls) == 1, f"expected one Format Equations call in {CONVERTER}"
    after = code[calls[0] + 1:]
    shows = [i for i, l in enumerate(after) if l == "sh_progress_show"]
    assert shows, f"{CONVERTER} must call Sh_Progress_Show after {FUNC}"
    says = [i for i, l in enumerate(after) if l.startswith("sh_progress_say")]
    assert says and shows[0] < says[0], (
        f"{CONVERTER}: Sh_Progress_Show must be the first thing done to the bar after {FUNC}, "
        f"since it is what hands the bar back from the book")


def test_sh_progress_over_changes_the_owner_and_nothing_else(repo_root):
    procs = procedures(_bar_module(repo_root))
    assert "Sh_Progress_Over" in procs, "Sh_Progress_Over is missing from ShNonModalMessage.bas"
    head = procs["Sh_Progress_Over"][0]
    assert re.search(r"\(\s*ByVal\s+d\s+As\s+Document\s*\)", head, re.I), head
    body = " ".join(_key(l) for l in _code(procs["Sh_Progress_Over"]))
    assert "if not sh_progress_up then exit sub" in body, (
        "Sh_Progress_Over must be gated on the flag: reading the caption of an unloaded form "
        "loads it and runs its Initialize, which hangs a Word with no desktop")
    assert 'sh_findwindowapi("thunderdframe", sh_convert_progress_form.caption)' in body
    assert "sh_setownerapi bar, sh_gwl_hwndparent, d.activewindow.hwnd" in body
    for bad in ("screenupdating", "screenrefresh", ".activate", "setforegroundwindow",
                "setfocus", "unload", ".show"):
        assert bad not in body, f"Sh_Progress_Over must not touch {bad!r}"


def test_the_owner_declaration_is_right_for_32_and_64_bit_office(repo_root):
    text = _bar_module(repo_root).replace("\r\n", "\n")
    decls = re.findall(r'Declare\s+(PtrSafe\s+)?Function\s+Sh_SetOwnerApi\s+Lib\s+"user32"\s+'
                       r'Alias\s+"(\w+)"\s*_\s*\n\s*\(([^)]*)\)\s*As\s+(\w+)', text)
    assert len(decls) == 3, f"expected three declarations of Sh_SetOwnerApi, found {len(decls)}"
    (ps64, alias64, args64, ret64), (ps32, alias32, args32, ret32), (ps_old, alias_old, _, ret_old) = decls
    # 64-bit Office: the Ptr function, pointer-sized.
    assert ps64 and alias64 == "SetWindowLongPtrA" and ret64 == "LongPtr"
    assert "ByVal dwNewLong As LongPtr" in args64 and "ByVal hwnd As LongPtr" in args64
    # 32-bit Office (VBA7 too): user32 has no SetWindowLongPtrA there.
    assert ps32 and alias32 == "SetWindowLongA" and ret32 == "LongPtr"
    # Office before VBA7: no PtrSafe, no LongPtr.
    assert not ps_old and alias_old == "SetWindowLongA" and ret_old == "Long"
    # The 64-bit one must sit under #If Win64, inside #If VBA7.
    i64 = text.index('"SetWindowLongPtrA"')
    before = text[:i64]
    assert before.rfind("#If Win64 Then") > before.rfind("#If VBA7 Then") > before.rfind("#End If"), (
        "SetWindowLongPtrA must be declared under #If Win64 inside #If VBA7 - it does not exist "
        "in 32-bit user32, and 32-bit Office 2010 and later is VBA7")
    assert re.search(r"^Private Const SH_GWL_HWNDPARENT As Long = -8\s*$", text, re.M)
