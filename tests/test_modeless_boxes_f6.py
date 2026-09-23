"""Every box that stays open moves the keyboard between book and box on F6 and Shift+F6.

Jerry's rule, 9/23/2026 (docs/UI-Conventions.md), for the TOC box (354), Resize Pictures in a
Selected Range (375), the $pg menus, and any box made modeless later. Each part fails silently in
Word - a control without a KeyDown handler strands a transcriber working without a mouse, and a
key left bound after the box closes sends F6 to a box that has gone.

Also here: 375's out-of-range message (390), the same kind as the TOC box's 389.
"""
import re

from test_toc_box_stays_open import form_code
from test_toc_never_deletes_a_line import module_text, procedures

# Controls that can hold the keyboard, as named in each form's .frx. A label cannot.
BOXES = {
    "Lp_TOC_Format_And_Color_Form": ["FormatTheTOCButton", "AddColorBarsButton",
                                     "RemoveColorBarsButton", "OkayButton", "CancelButton"],
    "Lp_Same_Pic_Range_Form": ["LastSizeButton", "ApplyButton", "DoneButton",
                               "JoinNextParaButton", "LeftAlignButton", "CenterButton",
                               "LeaveAsIsButton"],
    "Lp_Type_Fill_In_Line_Form": [f"CommandButton{n}" for n in range(1, 22)]
                                 + ["FillToRightMargin", "CancelButton", "TextBox1", "SpinButton1"],
    "Lp_Horz_To_Vert_List_Form": ["Cmd_Ok", "Cmd_Cancel", "AscendingOrderCheckbox", "Ordered_List",
                                  "Spaced_List", "TabbedList"],
}


def test_every_control_sends_f6_back_to_the_book(repo_root):
    for form, controls in BOXES.items():
        code = "\n".join(form_code(repo_root, form))
        for c in controls:
            m = re.search(rf"Sub {c}_KeyDown\(.*?\)(.*?)End Sub", code, re.S)
            assert m, f"{form}: {c} has no KeyDown - F6 cannot leave the box from it"
            assert "vbKeyF6" in m.group(1) and "KeyToDocument" in m.group(1), (
                f"{form}: {c}_KeyDown does not send F6 back to the book")


def test_one_shared_f6_for_every_box(repo_root):
    """Found in review, 9/23/2026: three boxes each binding F6 for itself left a live box without
    it. The key is bound once, to Sh_Box_ToggleFocus, and each box joins and leaves a list."""
    shared = (repo_root / "src" / "vba" / "ShNonModalMessage.bas").read_bytes().decode("latin-1")
    sprocs = procedures(shared)
    bind = "\n".join(sprocs["Sh_Box_BindKeys"])
    assert "BuildKeyCode(wdKeyF6)" in bind and "BuildKeyCode(wdKeyShift, wdKeyF6)" in bind
    assert "SH_BOX_KEY_MACRO" in bind
    toggle = "\n".join(sprocs["Sh_Box_ToggleFocus"])
    for who, sub in (("pg", "Sh_PgVal_ToggleFocus"), ("tocb", "Lp_Tocb_ToggleFocus"),
                     ("rst", "Lp_Rst_ToggleFocus"), ("fil", "Lp_Fil_ToggleFocus"),
                     ("hv", "Lp_Hvb_ToggleFocus")):
        assert f'Case "{who}"' in toggle and sub in toggle, f"F6 cannot reach the {who} box"

    # Nothing binds F6 on its own any more - every KeyBindings.Add for F6 is in Sh_Box_BindKeys.
    macros = module_text(repo_root)
    everything = macros + "\n" + shared
    adds = [line for line in everything.splitlines()
            if "KeyBindings.Add" in line and not line.lstrip().startswith("'")]
    assert len(adds) == 2, f"F6 bound somewhere other than Sh_Box_BindKeys: {adds}"

    procs = procedures(macros)
    for box, who in (("Lp_Tocb", "tocb"), ("Lp_Rst", "rst")):
        assert any(f'Sh_Box_Opened "{who}"' in c for c in procs[f"{box}_BindKeys"])
        assert any(f'Sh_Box_Closed "{who}"' in c for c in procs[f"{box}_UnbindKeys"])
    assert any('Sh_Box_Opened "pg"' in c for c in sprocs["Sh_PgVal_BindKeys"])
    assert any('Sh_Box_Closed "pg"' in c for c in sprocs["Sh_PgVal_UnbindKeys"])


def test_each_box_joins_before_it_selects_and_leaves_after(repo_root):
    """Joined late or left early, a box still up would be the newest while this one moves the
    selection, and would say its out-of-range message about it."""
    procs = procedures(module_text(repo_root))
    start = procs["Lp_Tocb_Start"]
    join = next(n for n, c in enumerate(start) if re.search(r"^\s*Lp_Tocb_BindKeys\s*$", c))
    sel = next(n for n, c in enumerate(start) if re.search(r"^\s*Lp_Tocb_Select\s*$", c))
    assert join < sel
    for fin, leave in (("Lp_Tocb_Finish", "Lp_Tocb_UnbindKeys"), ("Lp_Rst_Finish", "Lp_Rst_UnbindKeys")):
        body = procs[fin]
        at = next(n for n, c in enumerate(body) if re.search(rf"^\s*{leave}\s*$", c))
        focus = max(n for n, c in enumerate(body) if "Sh_Focus_Document" in c)
        assert at > focus, f"{fin} leaves the list of boxes before it lets the selection go"


def test_only_the_newest_box_watches_the_cursor(repo_root):
    procs = procedures(module_text(repo_root))
    for watch, who in (("Lp_Tocb_Watch_Cursor", "tocb"), ("Lp_Rst_Watch_Cursor", "rst")):
        body = procs[watch]
        say = next(n for n, c in enumerate(body) if "Sh_Say" in c)
        top = next((n for n, c in enumerate(body) if f'Sh_Box_On_Top() <> "{who}"' in c), None)
        assert top is not None and top < say, f"{watch} speaks while another box is newer"
        assert any(re.search(r"Lp_Tocb_Busy Or Lp_Rst_Busy|Lp_Rst_Busy Or Lp_Tocb_Busy", c)
                   for c in body[:say]), f"{watch} speaks while the other box is working"


def test_375_help_line_no_longer_blames_a_pg_box(repo_root):
    code = "\n".join(form_code(repo_root, "Lp_Same_Pic_Range_Form"))
    assert "Close the $pg validation box" not in code


def test_375_says_when_the_cursor_leaves_the_range(repo_root):
    events = (repo_root / "src" / "vba" / "VtEvents.cls").read_bytes().decode("latin-1")
    handler = re.search(r"Sub App_WindowSelectionChange\(.*?\)(.*?)End Sub", events, re.S)
    assert handler and re.search(r"Lp_Rst_Watch_Cursor\s+Sel\b", handler.group(1))

    watch = procedures(module_text(repo_root))["Lp_Rst_Watch_Cursor"]
    body = "\n".join(watch)
    assert "Your cursor is out of the selected range" in body and "VistaType LP (390)" in body
    assert "Lp_Rst_In_Range(" in body, "inside must be the same test Apply makes (379)"
    say = next(n for n, c in enumerate(watch) if "Sh_Say" in c)
    for guard in (r"If Not Lp_Rst_IsOn Then Exit Sub", r"If Lp_Rst_Busy Or Lp_Tocb_Busy Then Exit Sub",
                  r"If at = Lp_Rst_Warned_At Then Exit Sub"):
        at = next((n for n, c in enumerate(watch) if re.search(guard, c)), None)
        assert at is not None and at < say, f"Lp_Rst_Watch_Cursor lost its guard: {guard}"


def test_a_box_brought_back_or_used_becomes_the_newest(repo_root):
    """Found in review, 9/23/2026: an older box brought back to the front stayed below the newer
    one, so F6 and the out-of-range message still went to the other box."""
    procs = procedures(module_text(repo_root))
    tools = "\n".join(procs["Lp_Table_Tools"])
    assert re.search(r"If Not Selection\.Information\(wdWithInTable\) Then\s*(?:'.*\n\s*)*Lp_Tocb_BindKeys", tools)
    pics = "\n".join(procs["Lp_Resize_Pictures_In_Range"])
    assert re.search(r"If Lp_Rst_IsOn Then\s*(?:'.*\n\s*)*Lp_Rst_BindKeys", pics)
    assert any(re.search(r"^\s*Lp_Tocb_BindKeys\s*$", c) for c in procs["Lp_Tocb_Okay"])
    assert any(re.search(r"Lp_Rst_BindKeys", c) for c in procs["Lp_Rst_Run"])


def test_the_watchers_leave_each_others_ranges_alone(repo_root):
    procs = procedures(module_text(repo_root))
    for watch, other in (("Lp_Tocb_Watch_Cursor", "Lp_Rst_Holds"), ("Lp_Rst_Watch_Cursor", "Lp_Tocb_Holds")):
        body = procs[watch]
        say = next(n for n, c in enumerate(body) if "Sh_Say" in c)
        for need in (other, "Lp_Box_Sel_Is_A_Spot"):
            at = next((n for n, c in enumerate(body) if need in c), None)
            assert at is not None and at < say, f"{watch} does not check {need} before speaking"


def test_the_pg_validation_leaves_the_list_last(repo_root):
    shared = (repo_root / "src" / "vba" / "ShNonModalMessage.bas").read_bytes().decode("latin-1")
    done = procedures(shared)["Sh_PgVal_Done"]
    leave = next(n for n, c in enumerate(done) if re.search(r"^\s*Sh_PgVal_UnbindKeys\s*$", c))
    activate = next(n for n, c in enumerate(done) if "Sh_PgVal_SourceDoc.Activate" in c)
    assert leave > activate


def test_type_fill_in_lines_stays_open(repo_root):
    """Jerry, 9/23/2026: 357 non-modal, Cancel becomes Done, no range, F6 and Shift+F6."""
    procs = procedures(module_text(repo_root))
    start = "\n".join(procs["Lp_Fil_Start"])
    assert re.search(r"Lp_Type_Fill_In_Line_Form\.Show\s+vbModeless", start)
    assert 'Sh_Box_Opened "fil"' in start and "Sh_Focus_Document" in start
    assert not any(re.search(r"^\s*End\s*$", c) for c in procs["Lp_Type_Fill_In_Line"]), (
        "an End in the button's macro would take down every box that stays open")
    done = procs["Lp_Fil_Done"]
    leave = next(n for n, c in enumerate(done) if 'Sh_Box_Closed "fil"' in c)
    assert leave == max(n for n, c in enumerate(done) if "Sh_" in c), "357 must leave the list last"

    code = "\n".join(form_code(repo_root, "Lp_Type_Fill_In_Line_Form"))
    assert re.search(r'CancelButton\.Caption\s*=\s*"Done"', code)
    assert not re.search(r"\.Hide\b|Unload Me", code), "the box hides or unloads itself again"
    query = re.search(r"Sub UserForm_QueryClose\(.*?\)(.*?)End Sub", code, re.S)
    assert query and "Lp_Fil_Done" in query.group(1) and "Cancel = True" in query.group(1)
    for n in range(1, 22):
        m = re.search(rf"Sub CommandButton{n}_Click\(\)(.*?)End Sub", code, re.S)
        assert m and re.search(rf"Lp_Fil_Type 1, {n}\b", m.group(1)), f"button {n} types the wrong length"
    m = re.search(r"Sub FillToRightMargin_Click\(\)(.*?)End Sub", code, re.S)
    assert m and "Lp_Fil_Type 2, SpinButton1.Value" in m.group(1)


def test_horizontal_to_vertical_stays_open(repo_root):
    """Jerry, 9/23/2026: 348 non-modal, F6 and Shift+F6, no range."""
    procs = procedures(module_text(repo_root))
    start = "\n".join(procs["Lp_Hvb_Start"])
    assert re.search(r"Lp_Horz_To_Vert_List_Form\.Show\s+vbModeless", start)
    assert 'Sh_Box_Opened "hv"' in start and "Sh_Focus_Document" in start
    button = procs["Lp_Horz_List_To_Vertical"]
    assert not any(re.search(r"^\s*End\s*$", c) for c in button)
    assert not any("Lp_Is_Text_Selected" in c for c in button), (
        "Lp_Is_Text_Selected ends with End, which takes down every box that stays open")
    okay = "\n".join(procs["Lp_Hvb_Okay"])
    assert "Lp_Horz_To_Vert_Hidden Selection.Range" in okay and "VistaType LP (392)" in okay
    done = procs["Lp_Hvb_Done"]
    leave = next(n for n, c in enumerate(done) if 'Sh_Box_Closed "hv"' in c)
    assert leave == max(n for n, c in enumerate(done) if "Sh_" in c), "348 must leave the list last"

    code = "\n".join(form_code(repo_root, "Lp_Horz_To_Vert_List_Form"))
    assert not re.search(r"\.Hide\b|Unload Me|^\s*End\s*$", code, re.M)
    query = re.search(r"Sub UserForm_QueryClose\(.*?\)(.*?)End Sub", code, re.S)
    assert query and "Lp_Hvb_Done" in query.group(1) and "Cancel = True" in query.group(1)
    ok = re.search(r"Sub Cmd_Ok_Click\(\)(.*?)End Sub", code, re.S).group(1)
    for kind in ("ORDERED", "SPACED", "TABBED"):
        assert f'"{kind}"' in ok
    assert "Lp_Hvb_Okay listKind, (AscendingOrderCheckbox = True)" in ok


def test_short_cut_key_hover_text_matches_the_label(repo_root):
    """Jerry, 9/23/2026: the labels on 348 and 356 were copied from 357 and kept its hover text,
    "Shift+Alt+Ctrl+_". The hover text is set to each label's own caption when the box opens."""
    for form in ("Lp_Horz_To_Vert_List_Form", "Lp_Table_Tools_Menu_Form"):
        code = "\n".join(form_code(repo_root, form))
        init = re.search(r"Sub UserForm_Initialize\(\)(.*?)End Sub", code, re.S)
        assert init and "ShortCutKeyLabel.ControlTipText = ShortCutKeyLabel.Caption" in init.group(1), form
