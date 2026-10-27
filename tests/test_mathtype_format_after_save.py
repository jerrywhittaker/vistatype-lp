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
    # The second save is skipped when the user declined the Save As (the book has no file).
    assert code[save - 3] == 'if doc.path <> "" then', code[save - 3:save + 1]


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


def test_word_options_stay_off_through_format_equations(repo_root):
    code = _code(repo_root)
    fmt = code.index(FORMAT_CALL)
    restores = [i for i, l in enumerate(code) if l == "application.options.pagination = pag_prev"]
    # One in the stabilize block, gated on Not mtFormat; one after Format Equations; and the
    # error handler's.
    gated = [i for i in restores if i < fmt]
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
