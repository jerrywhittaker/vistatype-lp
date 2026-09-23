"""The TOC box (354) stays open and holds the TOC until Done.

Jerry, 9/22/2026: Table and TOC Tools' TOC box works like Resize Pictures in a Selected Range
(375). It is shown modeless, Cancel is Done, and the TOC range is held until Done collapses it.

The box itself cannot be driven headlessly (tests/vba/README.md: a UserForm blocks the runner),
and the two number decisions are tested in VBA by tests/vba/TestTocBox.bas. What this holds is
the wiring, each part of which, got wrong, fails silently in Word:

  * an `End` anywhere on the path unloads a modeless box and wipes the range it holds - in the
    button's macro after the box is shown, in either TOC form, or in Format the TOC's refusals;
  * the box must be shown vbModeless, or it is the old modal menu;
  * Done and the title bar's X must go through the same teardown;
  * Format the TOC must say it finished, or dialog 207 is said about a TOC never touched.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

END_STATEMENT = re.compile(r"^\s*End\s*$", re.IGNORECASE)


def form_code(repo_root, name):
    """The form's code, without the designer header - whose own closing line is a bare End."""
    text = (repo_root / "src" / "forms" / f"{name}.frm").read_bytes().decode("latin-1")
    lines = text.splitlines()
    first = next(n for n, line in enumerate(lines) if line.startswith("Attribute VB_Name"))
    return [strip_comment(line) for line in lines[first:]]


def has_end_statement(lines):
    return any(END_STATEMENT.match(line) for line in lines)


def test_neither_toc_form_ends_everything(repo_root):
    for name in ("Lp_TOC_Format_And_Color_Form", "Lp_TOC_Color_Bars_Form"):
        assert not has_end_statement(form_code(repo_root, name)), (
            f"{name} has an End statement again - it would unload the TOC box and drop its range")


def test_format_the_toc_does_not_end_everything(repo_root):
    procs = procedures(module_text(repo_root))
    assert not has_end_statement(procs["Lp_TOC_CleanAndFormat_TOC"])
    assert any(re.search(r"Lp_TOC_Format_Ok\s*=\s*True", c)
               for c in procs["Lp_TOC_CleanAndFormat_TOC"]), (
        "Format the TOC no longer says it finished; dialog 207 would follow a refusal")


def test_the_box_is_modeless_and_nothing_ends_after_it(repo_root):
    procs = procedures(module_text(repo_root))
    start = procs["Lp_Tocb_Start"]
    assert any(re.search(r"Lp_TOC_Format_And_Color_Form\.Show\s+vbModeless", c) for c in start)

    tools = procs["Lp_Table_Tools"]
    i = next(n for n, c in enumerate(tools) if re.search(r"\bLp_Tocb_Start\b", c))
    assert re.search(r"^\s*Exit Sub\s*$", tools[i + 1]), (
        "Lp_Table_Tools must leave with Exit Sub straight after starting the TOC box, not End")


def test_done_and_the_x_share_one_way_out(repo_root):
    code = "\n".join(form_code(repo_root, "Lp_TOC_Format_And_Color_Form"))
    done = re.search(r"Sub CancelButton_Click\(\)(.*?)End Sub", code, re.S)
    query = re.search(r"Sub UserForm_QueryClose\(.*?\)(.*?)End Sub", code, re.S)
    assert done and "Lp_Tocb_Done" in done.group(1)
    assert query and "Lp_Tocb_Done" in query.group(1) and "Cancel = True" in query.group(1)
    assert re.search(r'CancelButton\.Caption\s*=\s*"Done"', code)


def test_a_table_still_reaches_the_table_box(repo_root):
    """With the TOC box up, a cursor in a table must end the TOC run and go on to 356."""
    tools = procedures(module_text(repo_root))["Lp_Table_Tools"]
    code = "\n".join(tools)
    guard = re.search(r"If Lp_Tocb_IsOn Then(.*?)\n    End If", code, re.S)
    assert guard, "Lp_Table_Tools no longer checks for the TOC box first"
    body = guard.group(1)
    assert re.search(r"If Not Selection\.Information\(wdWithInTable\) Then", body)
    assert re.search(r"Lp_Tocb_Finish False", body), (
        "a cursor in a table must end the TOC run without moving the selection")


def test_every_press_reconciles_the_range_first(repo_root):
    procs = procedures(module_text(repo_root))
    ready = procs["Lp_Tocb_Ready"]
    rec = next(n for n, c in enumerate(ready) if "Lp_Tocb_Reconcile" in c)
    empty = next(n for n, c in enumerate(ready) if re.search(r"Lp_Tocb_Range\.End > Lp_Tocb_Range\.start", c))
    assert rec < empty, "the range must be reconciled before it is judged empty"


def test_remove_color_bars_asks_the_toc_not_the_book(repo_root):
    body = "\n".join(procedures(module_text(repo_root))["Lp_TOC_Remove_Color_Bars"])
    assert re.search(r"If Not anyChanged Then", body), (
        "Remove Color Bars must say there is nothing to remove when the TOC has no bars")


def test_the_book_gets_the_keyboard_once_the_box_is_up(repo_root):
    """Jerry, 9/23/2026: after the box first appears, the focus goes to the book."""
    start = procedures(module_text(repo_root))["Lp_Tocb_Start"]
    show = next(n for n, c in enumerate(start) if re.search(r"\.Show\s+vbModeless", c))
    assert any(re.search(r"Sh_Focus_Document\s+Lp_Tocb_Doc", c) for c in start[show + 1:]), (
        "the book must be given the focus AFTER the modeless box is shown, or the box keeps it")


def test_a_cursor_out_of_the_toc_is_said(repo_root):
    """Jerry, 9/23/2026: "Your cursor is out of the selected range" (389), from Word's
    selection event, and silent while a job is moving the selection itself."""
    events = (repo_root / "src" / "vba" / "VtEvents.cls").read_bytes().decode("latin-1")
    handler = re.search(r"Sub App_WindowSelectionChange\(.*?\)(.*?)End Sub", events, re.S)
    assert handler and re.search(r"Lp_Tocb_Watch_Cursor\s+Sel\b", handler.group(1)), (
        "nothing tells the TOC box the cursor moved")

    watch = procedures(module_text(repo_root))["Lp_Tocb_Watch_Cursor"]
    body = "\n".join(watch)
    assert "Your cursor is out of the selected range" in body
    assert "VistaType LP (389)" in body
    say = next(n for n, c in enumerate(watch) if "Sh_Say" in c)
    for guard in (r"If Not Lp_Tocb_IsOn Then Exit Sub", r"If Lp_Tocb_Busy Then Exit Sub",
                  r"If Lp_Tocb_Warned Then Exit Sub"):
        at = next((n for n, c in enumerate(watch) if re.search(guard, c)), None)
        assert at is not None and at < say, f"Lp_Tocb_Watch_Cursor lost its guard: {guard}"
    assert not any(re.search(r"Lp_Tocb_Reconcile|Set Lp_Tocb_Range\s*=", c) for c in watch), (
        "watching the cursor must not move the range the next press works on")
    table = next((n for n, c in enumerate(watch) if "wdWithInTable" in c), None)
    assert table is not None and table < say, (
        "a cursor in a table is on its way to the table box (356); 389 must not be said there")


def test_the_box_putting_the_cursor_back_clears_the_warning(repo_root):
    """Found in review, 9/23/2026: the box reselects the TOC while a job runs, when the watcher
    is silent, so the watcher never sees the cursor come back and 389 went quiet for good."""
    select = procedures(module_text(repo_root))["Lp_Tocb_Select"]
    assert any(re.search(r"Lp_Tocb_Warned\s*=\s*False", c) for c in select)


def test_the_toc_is_selected_before_the_box_appears(repo_root):
    start = procedures(module_text(repo_root))["Lp_Tocb_Start"]
    sel = next(n for n, c in enumerate(start) if re.search(r"^\s*Lp_Tocb_Select\s*$", c))
    on = next(n for n, c in enumerate(start) if re.search(r"Lp_Tocb_IsOn\s*=\s*True", c))
    assert sel < on, "a raw over-drag selection would trip 389 while the box is starting up"
