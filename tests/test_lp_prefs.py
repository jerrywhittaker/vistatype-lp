"""Spacing and Italics Settings - the module side (Jerry, 11/24/2026).

Three large print choices - line spacing, letter spacing and the italics marker - are kept IN
THE BOOK (document variables LpLineMultiple, LpLetterSpacing, LpItalicsMarker), else as the
transcriber's SAVED DEFAULT ([LpBookDefaults] in VistaType.ini), else BUILT IN. Jerry: "the
book's own settings win". Lp_Attach_The_Template writes them; Normalize Styles and the italics
pass read them.

What this checks, on the source (all of it reaches a Document or VistaType.ini, which the
headless VBA runner cannot; the pure helpers are in tests/vba/TestLpPrefs.bas):

  1. Normalize Styles reads the book's values through the getters, not the constants.
  2. Each getter asks the book, then VistaType.ini's [LpBookDefaults], then the constant - and
     the Pick functions hold that order.
  3. Lp_Attach_Lp_Template clears Lp_Pref_Chosen before 338 opens, so a choice never leaks.
  4. The attach writes the three variables after Media and before File Cleanup, saves the INI
     only for "Save as my default", clears the choice once used, and converts a changed marker
     on a book that was already large print, before Normalize Styles.
  5. The italics pass asks for the book's marker and colors the dashed runs on every way out.
  6. The settings folder is still never deleted by the uninstaller.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

GETTERS = {
    "Lp_Pref_Line_Multiple": ("Lp_Pref_Pick_Line", "LP_PREF_VAR_LINE", "LP_PREF_KEY_LINE"),
    "Lp_Pref_Letter_Spacing": ("Lp_Pref_Pick_Letter", "LP_PREF_VAR_LETTER", "LP_PREF_KEY_LETTER"),
}
PICKS = {
    "Lp_Pref_Pick_Line": ("Lp_Pref_Line_Valid", "LP_LINE_MULTIPLE"),
    "Lp_Pref_Pick_Letter": ("Lp_Pref_Letter_Valid", "LP_LETTER_SPACING"),
}


def procs(repo_root):
    return procedures(module_text(repo_root))


def body(repo_root, name):
    p = procs(repo_root)
    assert name in p, f"{name} is missing from LPandBrlMacros.bas"
    return p[name]


def joined(repo_root, name):
    """The procedure's code with line continuations joined."""
    return re.sub(r"\s*_\s*\n\s*", " ", "\n".join(body(repo_root, name)))


def first(lines, pattern):
    return next((i for i, c in enumerate(lines) if re.search(pattern, c)), None)


def test_the_names_are_as_planned(repo_root):
    text = module_text(repo_root)
    for const, value in (("LP_ITALICS_DEFAULT", "Dash"), ("LP_PREF_INI_SECTION", "LpBookDefaults"),
                         ("LP_PREF_KEY_LINE", "LineMultiple"), ("LP_PREF_KEY_LETTER", "LetterSpacing"),
                         ("LP_PREF_KEY_MARKER", "ItalicsMarker"), ("LP_PREF_VAR_LINE", "LpLineMultiple"),
                         ("LP_PREF_VAR_LETTER", "LpLetterSpacing"), ("LP_PREF_VAR_MARKER", "LpItalicsMarker")):
        assert re.search(rf'^Private Const {const} As String = "{value}"\r?$', text, re.M), const
    for public in ("Lp_Pref_Chosen As Boolean", "Lp_Pref_Chosen_Line As Single",
                   "Lp_Pref_Chosen_Letter As Single", "Lp_Pref_Chosen_Marker As String",
                   "Lp_Pref_Chosen_AsDefault As Boolean"):
        assert re.search(rf"^Public {public}\r?$", text, re.M), public


def test_normalize_uses_the_getters_not_the_constants(repo_root):
    lines = body(repo_root, "Lp_Normalize_Styles")
    code = "\n".join(lines)
    assert "LP_LINE_MULTIPLE" not in code and "LP_LETTER_SPACING" not in code, (
        "Normalize Styles must read the book's spacing, not the built-in constants")
    line = first(lines, r"^\s*nsLine = Lp_Pref_Line_Multiple\(ActiveDocument\)\s*$")
    letter = first(lines, r"^\s*nsLetter = Lp_Pref_Letter_Spacing\(ActiveDocument\)\s*$")
    assert line is not None and letter is not None
    uses = [i for i, c in enumerate(lines) if re.search(r"\bns(Line|Letter)\b", c) and "= Lp_Pref_" not in c
            and not re.match(r"^\s*Dim ", c)]
    assert len(uses) == 4, f"four uses (Normal, body, headings, the other styles), got {len(uses)}"
    assert min(uses) > max(line, letter), "read before they are used"


def test_each_getter_asks_the_book_then_the_ini_then_the_constant(repo_root):
    for getter, (pick, var, key) in GETTERS.items():
        code = joined(repo_root, getter)
        assert re.search(rf"\(ByVal doc As Document\)", code), f"{getter} must take ByVal doc"
        m = re.search(rf"{getter} = {pick}\(Lp_Pref_Doc_Var\(doc, {var}\), "
                      rf"Sh_Setting_Read\(LP_PREF_INI_SECTION, {key}, \"\"\)\)", code)
        assert m, f"{getter}: the book's variable first, then [LpBookDefaults] {key}"
    for pick, (valid, const) in PICKS.items():
        code = joined(repo_root, pick)
        doc_ = code.index(f"If {valid}(fromDoc)")
        ini = code.index(f"ElseIf {valid}(fromIni)")
        builtin = code.index(f"{pick} = {const}")
        assert doc_ < ini < builtin, f"{pick}: book, then saved default, then {const}"
    code = joined(repo_root, "Lp_Pref_Pick_Marker")
    assert (code.index("Lp_Pref_Marker_Canonical(fromDoc)") < code.index("Lp_Pref_Marker_Canonical(fromIni)")
            < code.index("Lp_Pref_Pick_Marker = LP_ITALICS_DEFAULT"))


def test_the_saved_default_marker_is_for_new_books_only(repo_root):
    """Jerry, 11/24/2026: an old LP book with no marker recorded counts as Dash; the saved default
    marker applies only to a NEW book, on its first attach. So File Cleanup and Normalize Styles
    (through Lp_Pref_Italics_Marker) never read VistaType.ini for the marker."""
    code = joined(repo_root, "Lp_Pref_Italics_Marker")
    assert re.search(r'Lp_Pref_Italics_Marker = Lp_Pref_Pick_Marker\(Lp_Pref_Doc_Var\(doc, LP_PREF_VAR_MARKER\), ""\)',
                     code), "the book's own marker, else Dash"
    assert "Sh_Setting_Read" not in code
    code = joined(repo_root, "Lp_Pref_Attach_Marker")
    assert re.search(r"If alreadyLp Then\s+Lp_Pref_Attach_Marker = Lp_Pref_Italics_Marker\(doc\)\s+Else\s+"
                     r"Lp_Pref_Attach_Marker = Lp_Pref_Pick_Marker\(Lp_Pref_Doc_Var\(doc, LP_PREF_VAR_MARKER\), "
                     r'Sh_Setting_Read\(LP_PREF_INI_SECTION, LP_PREF_KEY_MARKER, ""\)\)\s+End If', code)
    # nothing else reads the saved marker
    readers = [n for n, v in procs(repo_root).items()
               if "LP_PREF_KEY_MARKER" in "\n".join(v) and "Sh_Setting_Read" in "\n".join(v)]
    assert readers == ["Lp_Pref_Attach_Marker"], readers


def test_the_book_variable_is_read_directly(repo_root):
    """Like Lp_Doc_Is_Screen: not through Sh_Read_Document_Variables and its shared string."""
    code = "\n".join(body(repo_root, "Lp_Pref_Doc_Var"))
    assert "On Error Resume Next" in code and "doc.Variables(varName)" in code
    assert "Sh_Read_Document_Variables" not in code and "Sh_GP_String" not in code


def test_numbers_are_stored_with_a_period(repo_root):
    """A book or an INI written where the decimal separator is a comma must read back anywhere:
    no CStr/Format/CDbl on a fraction, and Val (period only) to read."""
    text = "\n".join(body(repo_root, "Lp_Pref_Number_Text"))
    assert "Format" not in text and "CDbl" not in text
    assert re.search(r'CStr\(n \\ 100\) & "\." & Right\$\("0" & CStr\(n Mod 100\), 2\)', text)
    assert "Val(s)" in "\n".join(body(repo_root, "Lp_Pref_Parse"))
    for writer in ("Lp_Pref_Set_For_Doc", "Lp_Pref_Save_Default"):
        code = joined(repo_root, writer)
        assert code.count("Lp_Pref_Number_Text(Lp_Pref_Pick_") == 2, writer


def test_every_parameter_is_byval(repo_root):
    """The fault that shipped in 3.0.274 was a Variant passed to a ByRef String."""
    for name, lines in procs(repo_root).items():
        if not (name.startswith("Lp_Pref_") or name in ("Lp_Italics_Convert_Marker", "Lp_Italics_Marker_Walk",
                                                         "Lp_Italics_Mark_Run", "Lp_Italics_Mark_Uniform")):
            continue
        header = re.sub(r"\s*_\s*\n\s*", " ", "\n".join(lines[:3])).split(")")[0]
        params = [p for p in header.split("(", 1)[1].split(",") if p.strip()] if "(" in header else []
        assert all(p.strip().startswith("ByVal ") for p in params), f"{name}: {params}"


def test_lp_attach_lp_template_clears_the_choice_before_338(repo_root):
    lines = body(repo_root, "Lp_Attach_Lp_Template")
    reset = first(lines, r"^\s*Lp_Pref_Chosen = False\s*$")
    warn = first(lines, r"Lp_ReAttachWarning_Form\.Show")
    show = first(lines, r"Lp_Attach_An_Lp_Template_Form\.Show")
    assert reset is not None, "Lp_Attach_Lp_Template must clear Lp_Pref_Chosen"
    assert reset < warn and reset < show, "before either way into 338"


def test_the_attach_writes_the_three_variables(repo_root):
    lines = body(repo_root, "Lp_Attach_The_Template")
    media = first(lines, r'Sh_Write_Document_Variables "Media", DM')
    old = first(lines, r"^\s*prefOldMarker = Lp_Pref_Italics_Marker\(ActiveDocument\)")
    write_old = first(lines, r"^\s*Lp_Pref_Set_For_Doc ActiveDocument, prefLine, prefLetter, prefOldMarker\s*$")
    write_new = first(lines, r"^\s*Lp_Pref_Set_For_Doc ActiveDocument, prefLine, prefLetter, prefMarker\s*$")
    cleanup = first(lines, r"^\s*Lp_Fix_Common_File_Errors\s*$")
    assert None not in (media, old, write_old, write_new, cleanup)
    assert media < old < write_old < write_new < cleanup, (
        "the old marker is read before anything is written, and both before File Cleanup, "
        "which marks the italics and asks the book which marker")
    block = "\n".join(lines[old:write_new + 8])
    assert re.search(r"If Lp_Pref_Chosen Then\s+prefLine = Lp_Pref_Chosen_Line\s+"
                     r"prefLetter = Lp_Pref_Chosen_Letter\s+prefMarker = Lp_Pref_Chosen_Marker\s+Else\s+"
                     r"prefLine = Lp_Pref_Line_Multiple\(ActiveDocument\)\s+"
                     r"prefLetter = Lp_Pref_Letter_Spacing\(ActiveDocument\)\s+"
                     r"prefMarker = Lp_Pref_Attach_Marker\(ActiveDocument, Lp_Doc_Was_Already_LP\)\s+End If", block)
    # a book already LP keeps its OLD marker in the book until the conversion is done (review 5a)
    assert re.search(r"If Lp_Doc_Was_Already_LP Then\s+"
                     r"Lp_Pref_Set_For_Doc ActiveDocument, prefLine, prefLetter, prefOldMarker\s+Else\s+"
                     r"Lp_Pref_Set_For_Doc ActiveDocument, prefLine, prefLetter, prefMarker\s+End If", block)
    assert re.search(r"If Lp_Pref_Chosen And Lp_Pref_Chosen_AsDefault Then\s+"
                     r"Lp_Pref_Save_Default prefLine, prefLetter, prefMarker", block), (
        "VistaType.ini only for \"Save as my default\"")
    assert re.search(r"^\s*Lp_Pref_Chosen = False\b", block, re.M), "the choice is used once"


def test_a_changed_marker_is_converted_before_normalize(repo_root):
    lines = body(repo_root, "Lp_Attach_The_Template")
    template = first(lines, r"\.AttachedTemplate = TemplatePathandName")
    cond = first(lines, r"^\s*If prefOldMarker <> prefMarker Then\s*$")
    conv = first(lines, r"^\s*Lp_Italics_Runs_Left = Lp_Italics_Convert_Marker\(prefOldMarker, prefMarker\)\s*$")
    marker = first(lines, r"^\s*Lp_Pref_Set_Marker_For_Doc ActiveDocument, prefMarker\s*$")
    normalize = first(lines, r"^\s*Lp_Normalize_Styles\s*$")
    assert None not in (template, cond, conv, marker, normalize)
    assert template < cond < conv < marker < normalize, (
        "after the template brings the Text styles; the book is told its new marker only once "
        "the conversion is done, and before Normalize Styles marks any italic left")
    outer = "\n".join(lines[cond - 1:marker + 2])
    assert re.match(r"\s*If Lp_Doc_Was_Already_LP Then\s+If prefOldMarker <> prefMarker Then", outer)


def test_the_italics_pass_marks_in_the_books_color_on_every_way_out(repo_root):
    lines = body(repo_root, "Lp_Italics_To_Dashed_Underline")
    assert first(lines, r"^\s*marker = Lp_Pref_Italics_Marker\(ActiveDocument\)\s*$") is not None
    exits = [i for i, c in enumerate(lines) if re.match(r"^\s*Exit Sub\s*$", c)]
    assert len(exits) == 3
    for i in exits:
        assert re.search(r"^\s*Lp_Italics_Marker_Walk \S+, LP_ITALICS_DEFAULT, marker\s*$", lines[i - 1]), (
            f"the way out at line {i} of the sub does not mark in color first")


def test_the_color_walk_steps_and_keeps_bold(repo_root):
    walk = "\n".join(body(repo_root, "Lp_Italics_Marker_Walk"))
    assert "wdReplaceAll" not in walk and ".Replacement" not in walk, "a Find that steps (issue #56 suspect)"
    assert re.search(r"\.Wrap = wdFindStop", walk)
    assert re.search(r"If hit\.OMaths\.count > 0 Then", walk), "equations are stepped over (issue #48)"
    run = "\n".join(body(repo_root, "Lp_Italics_Mark_Uniform"))
    remember = run.index("emph = Lp_Remember_Bold_Italic(rr)")
    styled = run.index("rr.Style = toStyle")
    restore = run.index("Lp_Restore_Bold_Italic rr, emph")
    assert remember < styled < restore, "a character style strips direct bold (issue #17)"


def test_the_settings_folder_is_never_uninstalled(repo_root):
    iss = (repo_root / "installer" / "vistatype.iss").read_text(encoding="utf-8", errors="replace")
    m = re.search(r"^\[UninstallDelete\]\s*$(.*?)^\[", iss, re.M | re.S)
    assert m, "no [UninstallDelete] section"
    entries = [ln for ln in m.group(1).splitlines() if ln.strip() and not ln.strip().startswith(";")]
    assert not [e for e in entries if "VistaType LP Settings" in e or "VistaType.ini" in e], (
        "the uninstaller must not delete the transcriber's saved defaults")


def test_the_color_marker_is_a_tag_style_of_its_own(repo_root):
    """Jerry, 11/24/2026: going back to Dash converts only what VistaType LP colored, never Text
    Red put on by hand - so the marker is "Italics Red", based on "Text Red"."""
    code = "\n".join(body(repo_root, "Lp_Pref_Marker_Style_Name"))
    assert re.search(r'Lp_Pref_Marker_Style_Name = "Italics " & Mid\$\(marker, 5\)', code)
    code = "\n".join(body(repo_root, "Lp_Pref_Marker_Look_Name"))
    assert re.search(r'Lp_Pref_Marker_Look_Name = "Text " & Mid\$\(marker, 5\)', code)
    make = "\n".join(body(repo_root, "Lp_Pref_Italics_Style"))
    assert "doc.Styles.Add(Name:=nm, Type:=wdStyleTypeCharacter)" in make
    assert "s.BaseStyle = lookName" in make
    walk = "\n".join(body(repo_root, "Lp_Italics_Marker_Walk"))
    assert "Set fromStyle = Lp_Pref_Italics_Style(doc, fromMarker, False)" in walk, "never made to be found"
    assert "Set toStyle = Lp_Pref_Italics_Style(doc, toMarker, True)" in walk


def test_the_tag_styles_survive_a_reattach(repo_root):
    code = re.sub(r'"\s*_\s*\n\s*&\s*"', "", "\n".join(body(repo_root, "Lp_Remove_All_Styles_Except_Lp_Styles")))
    keep = re.search(r'LpStyles = "([^"]*)"', code).group(1)
    for color in ("Blue", "Green", "Orange", "Red", "Violet"):
        assert f",Italics {color}," in keep, f"Italics {color} is not on the keep-list"


def test_a_run_in_another_character_style_is_left_alone(repo_root):
    """Review 11/24/2026: setting a style replaced a footnote reference, a Words TOC bar or a
    Hyperlink. Only a run with no character style, or in ours, is changed; super/subscript and
    small caps are kept; a failure skips the run (its callers use Application.Run)."""
    run = "\n".join(body(repo_root, "Lp_Italics_Mark_Uniform"))
    resume = run.index("On Error Resume Next")
    cs = run.index("csName = rr.CharacterStyle")
    dpf = run.index("okStyle = (csName = rr.Document.Styles(wdStyleDefaultParagraphFont).NameLocal)")
    leave = run.index("If Err.Number <> 0 Or Not okStyle Then GoTo Done")
    styled = run.index("rr.Style = toStyle")
    assert resume < cs < dpf < leave < styled
    for prop in ("Superscript", "Subscript", "SmallCaps"):
        assert re.search(rf"If rr\.Font\.{prop} <> \w+ Then rr\.Font\.{prop} = \w+", run), prop
    assert "Lp_Italics_Mark_Uniform = (Err.Number = 0)" in run
    assert re.search(r"Done:\s+Err\.Clear\s+On Error GoTo 0", run)
    walk = "\n".join(body(repo_root, "Lp_Italics_Marker_Walk"))
    assert "leftAlone = leftAlone + Lp_Italics_Mark_Run(hit, fromStyle, toStyle)" in walk


def test_a_mixed_hit_is_split_into_even_stretches(repo_root):
    """Jerry, 11/24/2026, after a vba-test of 3.0.716: "a note" + its footnote reference, and
    "ital" + superscript "2" + "ic", were each one dashed hit and stayed dashed whole. A hit that
    mixes a character style, superscript, subscript or small caps is now cut into stretches where
    all four are the same, and each stretch is marked on its own. A plain hit stays one call."""
    run = "\n".join(body(repo_root, "Lp_Italics_Mark_Run"))
    flat = re.sub(r"\s*_\s*\n\s*", " ", run)
    assert re.search(r"\) As Long", flat.split("\n")[0]), "answers how many pieces were left"
    plain = flat.index("If Not mixed Then\n        If Not Lp_Italics_Mark_Uniform(rr, fromStyle, toStyle) Then")
    split = flat.index("For Each c In rr.Characters")
    assert plain < split, "a plain hit is one call; only a mixed one is read character by character"
    assert re.search(r"mixed = \(rr\.Font\.Superscript = wdUndefined Or rr\.Font\.Subscript = wdUndefined "
                     r"Or rr\.Font\.SmallCaps = wdUndefined\)", flat)
    assert re.search(r'key = csName & "\|" & CStr\(c\.Font\.Superscript\) & "\|" & CStr\(c\.Font\.Subscript\) '
                     r'& "\|" & CStr\(c\.Font\.SmallCaps\)', flat)
    # VBA's And reads both sides: ends(0) must never be touched (it would fall into the Then
    # under Resume Next).
    assert not re.search(r"If n > 0 And", flat)
    assert re.search(r"If n > 0 Then\s+If key = prevKey And c\.Start = ends\(n\) Then joinIt = True", flat)
    assert re.search(r"If Not Lp_Italics_Mark_Uniform\(part, fromStyle, toStyle\) Then leftAlone = leftAlone \+ 1",
                     flat)
    assert "wdReplaceAll" not in run and "DoEvents" not in run


def test_no_doevents_in_the_walk(repo_root):
    """Issue #56's crashes come at DoEvents points; the walk must not add any."""
    for name in ("Lp_Italics_Marker_Walk", "Lp_Italics_Mark_Run", "Lp_Italics_Mark_Uniform",
                 "Lp_Pref_Italics_Style"):
        code = "\n".join(body(repo_root, name))
        assert "DoEvents" not in code, name


def test_the_spacing_buttons_on_424(repo_root):
    """Jerry, 11/25/2026: each spacing row has a W3C WCAG button (LineReset / LetterReset, Reset
    until then) giving 1.25 and 0.12, and a VTLP V2 (APH) button (LineAph / LetterAph) giving the
    built-in 1.00 and 0.07. They only change the list; nothing is kept until Okay. About line/letter
    spacing says its text is to come, through Sh_Say with message 429."""
    import pathlib
    frm = (pathlib.Path(repo_root) / "src/forms/Lp_Spacing_Italics_Form.frm").read_bytes().decode("latin-1")
    subs = dict(re.findall(r"Private Sub (\w+)_Click\(\)\r\n(.*?)End Sub", frm, re.S))
    assert re.fullmatch(r"\s*LineList\.ListIndex = Lp_Pref_Line_Index\(Lp_Pref_Wcag_Line_Multiple\(\)\)\s*",
                        subs["LineReset"])
    assert re.fullmatch(r"\s*LetterList\.ListIndex = Lp_Pref_Letter_Index\(Lp_Pref_Wcag_Letter_Spacing\(\)\)\s*",
                        subs["LetterReset"])
    assert re.fullmatch(r"\s*LineList\.ListIndex = Lp_Pref_Line_Index\(Lp_Pref_Builtin_Line_Multiple\(\)\)\s*",
                        subs["LineAph"])
    assert re.fullmatch(r"\s*LetterList\.ListIndex = Lp_Pref_Letter_Index\(Lp_Pref_Builtin_Letter_Spacing\(\)\)\s*",
                        subs["LetterAph"])
    assert '"VistaType LP (429)"' in subs["AboutSpacing"] and "Sh_Say" in subs["AboutSpacing"]
    assert "MsgBox" not in frm
    text = module_text(repo_root)
    for name, const in (("Lp_Pref_Wcag_Line_Multiple", "LP_WCAG_LINE_MULTIPLE"),
                        ("Lp_Pref_Wcag_Letter_Spacing", "LP_WCAG_LETTER_SPACING")):
        assert re.search(rf"^Public Function {name}\(\) As Single\r?\n\s+{name} = {const}\r?$", text, re.M), name
    # the Alt keys, one letter each, none used twice on the box
    keys = re.findall(r'^\s+(\w+)\.Accelerator = "(\w)"\r?$', frm, re.M)
    assert dict(keys) == {"LineReset": "W", "LetterReset": "G", "LineAph": "V", "LetterAph": "P",
                          "AboutSpacing": "A"}
