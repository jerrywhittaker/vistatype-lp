"""The converted book is saved BEFORE MathType formats the equations (issue #42).

Jerry's decision, 10/27/2026. MathType's whole-book Format Equations sometimes crashes Word in
MT6.DLL (three of six whole-book runs of a 1,358-equation book, 10/25-10/26/2026) and sometimes
hangs it (Jerry had to power off the build box, 10/27/2026). It ran while the converted book was
still an unsaved Document2, so the whole conversion was lost each time. A mitigation, not a cure:
the book is now on disk before Format Equations starts, so a crash or a hang costs only the
formatting - the user opens the saved book and clicks Format Equations on the MathType tab.

The order, when MathType equations are chosen and MathType is installed: the equations go in,
the font pass and the repaginate, box 401 (in place of 295) says the formatting comes after the
save, the Save As, Format Equations, a second repaginate, a plain doc.Save. Every other path is
unchanged. Word cannot run here; these are checks on the source.

Large print only, from 10/28/2026 (Jerry): box 402, how to set MathType up for large print, comes
after the Save As and before Format Equations, and STAYS OPEN (Sh_Ask_Modeless) so the
transcriber can open MathType and set it up while the conversion waits. Okay formats; Cancel or
the red X skips the formatting and the second save, and 313 says how to format later.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

CONVERTER = "Sh_Convert_XML_File_To_Word_Document"
FORMAT_CALL = 'if not sh_mathtype_format_equations(finaldoc, (sh_gp_string_2 = "brl")) then'


def _key(line):
    return " ".join(line.lower().split())


def _code(repo_root):
    """The converter's statements, comments dropped, continuation lines joined."""
    out, cur = [], ""
    for raw in procedures(module_text(repo_root))[CONVERTER]:
        s = raw.strip()
        if not s or s.startswith("'"):
            continue
        cur = cur + " " + s if cur else s
        if cur.endswith(" _"):
            cur = cur[:-2]
            continue
        out.append(_key(cur))
        cur = ""
    return out


def _first(code, pred, start=0, what=""):
    for i in range(start, len(code)):
        if pred(code[i]):
            return i
    raise AssertionError(f"{CONVERTER}: not found after statement {start}: {what}")


def test_format_equations_runs_after_the_save_and_a_save_follows_it(repo_root):
    code = _code(repo_root)
    calls = [i for i, l in enumerate(code) if l == FORMAT_CALL]
    assert len(calls) == 1, f"expected one Format Equations call in {CONVERTER}, found {len(calls)}"
    fmt = calls[0]
    # Both ways the book is first saved come before Format Equations.
    execute = _first(code, lambda l: l == "dlgsaveas.execute", what="dlgSaveAs.Execute")
    named = _first(code, lambda l: l == "doc.save", what="the named document's doc.Save")
    assert execute < fmt and named < fmt, (
        "Format Equations must run AFTER the book is saved, or a crash in MT6.DLL loses the whole "
        "conversion (issue #42)")
    # ...and it runs at FormatTheEquations, only when it is due.
    label = _first(code, lambda l: l == "formattheequations:", what="the FormatTheEquations label")
    gate = _first(code, lambda l: l == "if mtformat then", start=label, what="If mtFormat Then")
    assert label < gate < fmt
    # After it: the bar back, a repaginate (MathType redraws every equation at its true size,
    # which moves the breaks), then a plain save, and only then "Finished".
    show = _first(code, lambda l: l == "sh_progress_show", start=fmt + 1, what="Sh_Progress_Show")
    repag = _first(code, lambda l: l == "doc.repaginate", start=fmt + 1, what="doc.Repaginate")
    save = _first(code, lambda l: l == "doc.save", start=fmt + 1, what="the second doc.Save")
    done = _first(code, lambda l: l.startswith('sh_progress_say 100, "finished."'), what="Finished")
    assert fmt < show < repag < save < done
    # The second save is skipped when the user declined the Save As (the book has no file), and
    # when the formatting was skipped on 402 (nothing has changed since the first save).
    assert code[save - 3] == 'if doc.path <> "" and not mtskip then', code[save - 3:save + 1]


def test_only_the_mathtype_path_saves_early(repo_root):
    code = _code(repo_root)
    sets = [i for i, l in enumerate(code) if l == "mtformat = true"]
    assert len(sets) == 1, "mtFormat must be set in one place only"
    # Set in the Else of "If Not Sh_MathType_Is_Installed() Then", inside the MathType branch.
    mt = _first(code, lambda l: l == 'if mathmode = "mathtype" then', what='If mathMode = "MATHTYPE"')
    inst = _first(code, lambda l: l == "if not sh_mathtype_is_installed() then", start=mt,
                  what="the MathType-installed check")
    els = _first(code, lambda l: l == "else", start=inst, what="its Else")
    assert mt < inst < els < sets[0] and code[sets[0] - 1] == "else"
    assert "mtformat = false" not in code


def test_box_401_says_the_formatting_comes_after_the_save_before_the_save_as(repo_root):
    code = _code(repo_root)
    say = _first(code, lambda l: l.startswith("sh_say ") and '"vistatype lp (401)"' in l,
                 what="Sh_Say ... (401)")
    display = _first(code, lambda l: "dlgsaveas.display" in l, what="dlgSaveAs.Display")
    assert say < display
    # Straight inside the MathType branch - box 402 no longer comes before it (10/28/2026).
    assert code[say - 1] == "if mtformat then", "box 401 is for the MathType path only"
    text = code[say]
    for words in ("after the save, not before it", "several minutes on a large book",
                  "saved again", "open the saved book", "format equations on the mathtype tab"):
        assert words in text, f"box 401 does not say {words!r}"
    assert " she " not in text and " her " not in text
    # 295 is still there for every other path.
    assert any('"vistatype lp (295)"' in l for l in code)
    # 401 is a new number: nothing else in the module uses it.
    whole = module_text(repo_root)
    assert len(re.findall(r"\(401\)", whole)) == 1


def test_the_bar_says_the_book_is_saved_while_mathtype_formats(repo_root):
    code = _code(repo_root)
    fmt = code.index(FORMAT_CALL)
    before = [l for l in code[:fmt] if l.startswith("sh_progress_say 66,")]
    assert any('"the book is saved. mathtype is formatting the equations' in l for l in before)


def test_the_bar_says_it_will_not_move_not_that_it_waits(repo_root):
    """Jerry, 10/27/2026: the bar (373) said "...so the bar waits"; he wants "so the bar will
    not move". Every line the converter puts on the bar."""
    code = _code(repo_root)
    says = [l for l in code if "so the bar" in l]
    assert says, "no bar line says 'so the bar ...' any more"
    for l in says:
        assert "so the bar waits" not in l, l
        assert "so the bar will not move" in l, l
    fmt = code.index(FORMAT_CALL)
    assert any(l.startswith("sh_progress_say 66,") and
               l.endswith('one long step, so the bar will not move."') for l in code[:fmt])
    # Nowhere else in the module's code either (the changelog may quote the old words).
    code_lines = [l for l in module_text(repo_root).splitlines() if not l.strip().startswith("'")]
    assert not any("so the bar waits" in l for l in code_lines)


JERRYS_402 = (
    "This book has been identified as a large-print title, and MathType is installed on this "
    "computer. MathType can be set up to produce large-print equations by opening MathType and "
    "setting the font size to match the size used in the book's base font (Normal style) text. "
    "The equation font can also be changed to match the book's typeface, either VistaTypeLP Sans "
    "or Tahoma.\r"
    "\r"
    "Here is how to set up the font MathType uses:\r"
    '1. Open MathType and select "Inline".\r'
    '2. Select the "Style" tab followed by "Define" on the menu.\r'
    '3. In the "Define Styles" dialog, change the "Primary Font" to the font used in the book '
    "(VistaTypeLP Sans or Tahoma).\r"
    '4. Select the "Apply" button, then select the "OK" button.\r'
    "\r"
    "Here is how to set up the font sizes MathType uses:\r"
    '1. Select the "Size" tab.\r'
    '2. Change the "Full" setting to match the point size used in the document.\r'
    '3. Change the "Subscript/Superscript" setting to match the point size used in the document, '
    'and make sure that "%" is changed to "Points".\r'
    '4. Make the same changes (as above) for "Sub-Subscript/Superscript".\r'
    '5. Select the "Apply" button, then select the "OK" button.\r'
    "\r"
    'The current settings can be stored in a location of your choice by selecting the "Preferences" '
    'tab followed by "Equation Preferences" on the menu, then selecting "Save to File". It is '
    "suggested that the file be named using the base font size. The saved preferences can then be "
    "loaded from that file when needed."
)


def _vba_value(expr):
    """The value of a VBA string expression made of "literals" and vbCr joined by &."""
    out, i = [], 0
    while i < len(expr):
        c = expr[i]
        if c in " &":
            i += 1
        elif c == '"':
            j, buf = i + 1, []
            while True:
                if expr[j] == '"' and expr[j + 1:j + 2] == '"':
                    buf.append('"')
                    j += 2
                elif expr[j] == '"':
                    break
                else:
                    buf.append(expr[j])
                    j += 1
            out.append("".join(buf))
            i = j + 1
        elif expr.startswith("vbCr", i) and not expr.startswith("vbCrLf", i):
            out.append("\r")
            i += 4
        else:
            raise AssertionError(f"unexpected token in {expr!r} at {i}")
    return "".join(out)


def _box_402_text(repo_root):
    text = ""
    for raw in procedures(module_text(repo_root))[CONVERTER]:
        s = raw.strip()
        if s.startswith("lpNote = lpNote & "):
            text += _vba_value(s[len("lpNote = lpNote & "):])
        elif s.startswith("lpNote = "):
            text = _vba_value(s[len("lpNote = "):])
    return text


CLOSING_LINE_402 = "When MathType is set up, click Okay to format the equations."


def test_box_402_is_jerrys_text_exactly(repo_root):
    """Jerry's text, 10/27/2026, with only American spelling, the product's name, closing
    quotation marks and one garbled line put right, and one closing line added on 10/28/2026 when
    the box began to stay open and wait. Anything else is a change to his words."""
    text = _box_402_text(repo_root)
    assert text == JERRYS_402 + "\r\r" + CLOSING_LINE_402
    # Past what a MsgBox can show, which is why it goes through Sh_Say.
    assert len(text) > 1024
    assert " she " not in text and " her " not in text and "colour" not in text


def test_box_402_comes_after_the_save_and_before_format_equations(repo_root):
    """Jerry, 10/28/2026: "It should come after message 401 (after the file save) and before the
    equation formatting." Large print with MathType installed only, at FormatTheEquations - so a
    declined Save As (312, No), which jumps there, reaches it too."""
    code = _code(repo_root)
    asks = [i for i, l in enumerate(code) if '"vistatype lp (402)"' in l]
    assert len(asks) == 1, "expected one box 402"
    a402 = asks[0]
    assert code[a402] == 'mtskip = not sh_ask_modeless(lpnote, "vistatype lp (402)", doc)', code[a402]
    s401 = _first(code, lambda l: '"vistatype lp (401)"' in l, what="box 401")
    execute = _first(code, lambda l: l == "dlgsaveas.execute", what="dlgSaveAs.Execute")
    named = _first(code, lambda l: l == "doc.save", what="the named document's doc.Save")
    label = code.index("formattheequations:")
    fmt = code.index(FORMAT_CALL)
    assert s401 < execute < label < a402 < fmt and named < label
    # Inside "If Sh_GP_String_2 = "LP" Then", inside the "If mtFormat Then" at the label.
    gate = _first(code, lambda l: l == "if mtformat then", start=label, what="If mtFormat Then")
    assert code[gate + 1] == "mtskip = false"
    assert code[gate + 2] == 'if sh_gp_string_2 = "lp" then'
    assert all(l.startswith("lpnote = ") for l in code[gate + 3:a402 - 1]), code[gate + 3:a402 - 1]
    # The bar off the screen while 402 is up, so it covers neither MathType nor the box.
    assert code[a402 - 1] == "sh_progress_hide"
    # 402 is a new number: no other box in the module uses it (313's note names it, unquoted).
    assert len(re.findall(r'"VistaType LP \(402\)"', module_text(repo_root))) == 1


def test_box_402_stays_open_and_the_conversion_waits_on_it(repo_root):
    """Jerry, 10/28/2026: 402 "should be a non-modal message so that the user can start and
    configure MathType". Sh_Ask_Modeless shows Sh_Message_Form vbModeless - its own instance, so
    another macro's Sh_Say cannot take it down - and waits, lightly, until it is answered."""
    procs = procedures(_bar_module(repo_root))
    assert "Sh_Ask_Modeless" in procs
    head = procs["Sh_Ask_Modeless"][0] + procs["Sh_Ask_Modeless"][1]
    assert re.search(r"ByVal body As String, ByVal titleText As String", head, re.I), head
    assert re.search(r"ByVal d As Document\)\s*As Boolean", head, re.I), head
    body = [_key(l) for l in procs["Sh_Ask_Modeless"] if l.strip() and not l.strip().startswith("'")]
    new = body.index("set sh_askbox_form = new sh_message_form")
    show = body.index("if err.number = 0 then sh_askbox_form.show vbmodeless")
    opened = body.index('sh_box_opened "msg"')
    focus = body.index("sh_focus_document d")
    loop = body.index("do")
    end = body.index("loop while stillup")
    answer = body.index("sh_ask_modeless = sh_msg_answer")
    closed = body.index('sh_box_closed "msg"')
    assert new < show < opened < focus < loop < end < answer < closed
    # The book is in front BEFORE the box is shown, so the box belongs to the book's window.
    assert body.index("d.activate") < show
    wait = body[loop:end + 1]
    assert "doevents" in wait and "sh_sleepapi 50" in wait and "stillup = sh_askbox_form.visible" in wait
    # The words come from the variables the form reads, with Cancel offered and no answer yet.
    for need in ("sh_msg_body = body", "sh_msg_title = titletext", "sh_msg_offer_cancel = true",
                 "sh_msg_answer = false"):
        assert need in body[:new], need
    # Never "Sh_Message_Form.Show" on the shared instance - Sh_Say unloads that one.
    assert not any(l.startswith("sh_message_form.") for l in body)
    # Sleep declared for 64-bit and 32-bit Office alike.
    text = _bar_module(repo_root)
    assert 'Private Declare PtrSafe Sub Sh_SleepApi Lib "kernel32" Alias "Sleep"' in text
    assert 'Private Declare Sub Sh_SleepApi Lib "kernel32" Alias "Sleep"' in text


def test_box_402_follows_the_f6_rule(repo_root):
    """docs/UI-Conventions.md, 9/23/2026: a box that stays open joins the shared F6 list, F6 and
    Shift+F6 in the box send the keyboard back to the book, and it opens with the book focused."""
    procs = procedures(_bar_module(repo_root))
    toggle = [_key(l) for l in procs["Sh_Box_ToggleFocus"]]
    at = toggle.index('case "msg"')
    assert toggle[at + 1] == "sh_ask_modeless_togglefocus"
    tf = [_key(l) for l in procs["Sh_Ask_Modeless_ToggleFocus"]]
    assert 'sh_box_closed "msg"' in tf and "sh_focus_modeless_form sh_askbox_title" in tf
    kd = [_key(l) for l in procs["Sh_Ask_Modeless_KeyToDocument"]]
    # Only for the instance that is waiting - a modal message's F6 still does nothing.
    assert "if not sh_askbox_up then exit sub" in kd
    assert "if not f is sh_askbox_form then exit sub" in kd
    assert "sh_focus_document sh_askbox_doc" in kd
    form = (repo_root / "src" / "forms" / "Sh_Message_Form.frm").read_bytes().decode("latin-1")
    for c in ("Msg_Text", "Okay_Button", "Cancel_Button"):
        m = re.search(rf"Sub {c}_KeyDown\(.*?\)(.*?)End Sub", form, re.S)
        assert m and "vbKeyF6" in m.group(1) and "Sh_Ask_Modeless_KeyToDocument Me" in m.group(1), c
    # The red X is Cancel (Jerry, 10/11/2026): it answers False and hides, ending the wait.
    q = re.search(r"Sub UserForm_QueryClose\(.*?\)(.*?)End Sub", form, re.S).group(1)
    assert "Sh_Msg_Answer = False" in q and "Me.Hide" in q


def test_cancel_on_402_skips_the_formatting_and_says_so(repo_root):
    """Cancel or the red X on 402: no Format Equations, no second repaginate or save (nothing
    changed since the first save), Word's options back, and 313 tells the user how to format the
    equations later."""
    code = _code(repo_root)
    label = code.index("formattheequations:")
    skip = _first(code, lambda l: l == "if mtskip then", start=label, what="If mtSkip Then")
    els = _first(code, lambda l: l == "else", start=skip, what="its Else")
    fmt = code.index(FORMAT_CALL)
    assert skip < els < fmt
    note = " ".join(code[skip + 1:els])
    for words in ("mathnote = mathnote & ", "the equations were not formatted",
                  "vistatype lp (402) was canceled", "click format equations on the mathtype tab",
                  "then save the book"):
        assert words in note, words
    assert "doc.repaginate" not in code[skip:els] and "doc.save" not in code[skip:els]
    # The re-home comes before the split, so on both ways the Close brings the book forward.
    assert code[skip - 1] == "sh_progress_rehome doc"
    # Word's options go back on both ways, after the split.
    end = _first(code, lambda l: l == "doc.undoclear", start=fmt)
    assert code[end + 1] == "end if"
    assert code[end + 2:end + 5] == ["application.options.checkgrammarasyoutype = gram_prev",
                                     "application.options.checkspellingasyoutype = spell_prev",
                                     "application.options.pagination = pag_prev"]
    # 313 carries mathNote.
    s313 = _first(code, lambda l: '"operation complete (313)"' in l, what="313")
    assert "mathnote" in code[s313]


def test_a_book_closed_while_402_is_up_ends_the_run_quietly(repo_root):
    code = _code(repo_root)
    a402 = _first(code, lambda l: '"vistatype lp (402)"' in l, what="402")
    assert code[a402 + 1] == "if not sh_doc_is_open(doc) then"
    block = code[a402 + 2:code.index("exit sub", a402) + 1]
    assert block == ["application.options.checkgrammarasyoutype = gram_prev",
                     "application.options.checkspellingasyoutype = spell_prev",
                     "application.options.pagination = pag_prev",
                     "sh_progress_close", "exit sub"], block


def test_the_book_not_the_blank_document_is_in_front_after_the_second_save(repo_root):
    """Jerry, 10/27/2026: after the automatic save that follows Format Equations, the blank
    document was left in front. The bar's home moves to the book before the bar is shown for
    Format Equations, so no later Show or Close brings the blank document forward; and the book
    is activated after the bar has closed, the last thing to choose a window."""
    code = _code(repo_root)
    label = code.index("formattheequations:")
    gate = _first(code, lambda l: l == "if mtformat then", start=label, what="If mtFormat Then")
    rehome = _first(code, lambda l: l == "sh_progress_rehome doc", start=gate, what="the re-home")
    show = _first(code, lambda l: l == "sh_progress_show", start=gate, what="Sh_Progress_Show")
    # After box 402 (which hides the bar), before the bar is shown again for Format Equations.
    a402 = _first(code, lambda l: '"vistatype lp (402)"' in l, start=gate, what="402")
    assert a402 < rehome < show < code.index(FORMAT_CALL)
    assert code[rehome + 1] == "if mtskip then"
    fmt = code.index(FORMAT_CALL)
    save = _first(code, lambda l: l == "doc.save", start=fmt + 1, what="the second doc.Save")
    done = _first(code, lambda l: l.startswith('sh_progress_say 100, "finished."'), what="Finished")
    close = _first(code, lambda l: l == "sh_progress_close", start=done, what="Sh_Progress_Close")
    assert code[close + 1:close + 5] == ["sh_spin_doevents", "application.activate", "doc.activate",
                                         "activewindow.view.type = wdprintview"]
    # Nothing between the second save and the end activates the bar's old home or re-shows the bar.
    end = code.index("exit sub", close)
    for l in code[save:end]:
        assert l not in ("currentdoc.activate", "sh_progress_show", "finaldoc.activate"), l
    # The declined save, on the same path, has the book activated after the close too.
    leave = _first(code, lambda l: l == "if savedeclined then", start=fmt, what="If saveDeclined")
    assert code[leave + 1:leave + 4] == ["sh_progress_close", "application.activate", "doc.activate"]


def _bar_module(repo_root):
    return (repo_root / "src" / "vba" / "ShNonModalMessage.bas").read_bytes().decode("utf-8")


def test_sh_progress_rehome_moves_the_bars_home_and_nothing_else(repo_root):
    procs = procedures(_bar_module(repo_root))
    assert "Sh_Progress_Rehome" in procs, "Sh_Progress_Rehome is missing from ShNonModalMessage.bas"
    head = procs["Sh_Progress_Rehome"][0]
    assert re.search(r"\(\s*ByVal\s+d\s+As\s+Document\s*\)", head, re.I), head
    body = [_key(l) for l in procs["Sh_Progress_Rehome"][1:] if l.strip()]
    assert body[0] == "if not sh_progress_up then exit sub"
    assert "set sh_progress_doc = d" in body and "sh_progress_over d" in body
    assert body.index("set sh_progress_doc = d") < body.index("sh_progress_over d")
    for bad in ("activate", "sh_focus_document", ".show", "unload"):
        assert not any(bad in l for l in body), f"Sh_Progress_Rehome must not {bad}"
    # And Show and Close use the home it sets.
    show = [_key(l) for l in procs["Sh_Progress_Show"]]
    close = [_key(l) for l in procs["Sh_Progress_Close"]]
    assert "sh_focus_document sh_progress_doc" in show
    assert "sh_focus_document sh_progress_doc" in close


def test_word_options_stay_off_through_format_equations(repo_root):
    code = _code(repo_root)
    fmt = code.index(FORMAT_CALL)
    restores = [i for i, l in enumerate(code) if l == "application.options.pagination = pag_prev"]
    # One in the stabilize block, gated on Not mtFormat; one after Format Equations; and the
    # error handler's.
    # (Not counting the one for a book closed while box 402 was up, which ends the run.)
    gated = [i for i in restores if i < fmt
             and "if not sh_doc_is_open(doc) then" not in code[i - 3:i]]
    assert len(gated) == 1 and "if not mtformat then" in code[gated[0] - 3:gated[0]], (
        "background pagination must stay off until after Format Equations on the MathType path")
    assert any(fmt < i < code.index("doc.repaginate", fmt) + 6 for i in restores)


def test_a_declined_save_as_still_formats_then_leaves_unsaved(repo_root):
    code = _code(repo_root)
    no = _first(code, lambda l: l == "if userchoice = vbno then", what="If userChoice = vbNo")
    assert code[no + 1] == "savedeclined = true" and code[no + 2] == "goto formattheequations"
    fmt = code.index(FORMAT_CALL)
    leave = _first(code, lambda l: l == "if savedeclined then", start=fmt, what="If saveDeclined")
    done = _first(code, lambda l: l.startswith('sh_progress_say 100,'), what="Finished")
    assert leave < done
    assert code[leave + 1] == "sh_progress_close"
    assert "exit sub" in code[leave:done]
