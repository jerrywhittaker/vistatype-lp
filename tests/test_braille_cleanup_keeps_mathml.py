"""Braille cleanup leaves the MathML alone (issue #37, Jerry, 10/18/2026).

A braille conversion of a NIMAS or DAISY book puts each equation in as MathML text,
<math xmlns="http://www.w3.org/1998/Math/MathML" ...>...</math>, for DBT. The braille cleanup
passes damaged it: Dx_Fix_Primes made id="m1" into id="m1" with a double prime, the quote pass
curled the attribute quotes, the paragraph spacing pass made a middle dot a space, the Nemeth
passes took out U+2061, Compress Linear Math made x a times sign and took out the spaces.

The cure is ONE mechanism, not a test in every pass: Dx_Math_Park sets each MathML element
aside - its exact text in a document variable, a short mark in its place - before the passes,
and Dx_Math_Restore puts it back after them, error paths included.

The passes cannot run headlessly here (the braille configuration hangs an invisible Word), so
these hold the source: every entry point parks before its passes and restores after them and on
its error path, nothing else runs the dangerous passes, and the mark itself is one no pass can
change. That last one is a Python model of the passes' own patterns, not Word's Find. The pure
helpers are tested for real in tests/vba/TestMathML.bas.
"""
import re

from test_toc_box_stays_open import form_code
from test_toc_never_deletes_a_line import module_text, procedures

PARK = r"Dx_Math_Park\b"
RESTORE = r"Dx_Math_Restore\b"


def _procs(repo_root):
    return procedures(module_text(repo_root))


def _at(lines, pattern, start=0):
    hits = [i for i, l in enumerate(lines) if i >= start and re.search(pattern, l)]
    assert hits, f"no line matching {pattern!r}"
    return hits


# --- the mark, against the passes ---------------------------------------------------------

MARK = "\u27e6VTMATH000012\u27e7"
MATHML = ('<math xmlns="http://www.w3.org/1998/Math/MathML" id="m1" alttext=\'x > 2\'>'
          '<mi>x</mi><mo>\u00b7</mo><mn>2</mn><mo>\u2061</mo><mo>\u2013</mo><mtext>a  b</mtext></math>')

# The braille passes' Find/Replace patterns, written as Python regular expressions. Each is the
# pattern in the named sub, translated by hand from Word's wildcards.
PASSES = [
    ("Dx_Fix_Primes", r"([0-9]+)\u2019", "\\1\u2032"),
    ("Dx_Fix_Primes", r"([0-9]+)\u201d", "\\1\u2033"),
    ("Dx_Fix_Primes", r"([0-9]+)'", "\\1\u2032"),
    ("Dx_Fix_Primes", r'([0-9]+)"', "\\1\u2033"),
    ("Dx_Replace_Straight_Quotes_With_Smart_Quotes", r'"', "\u201c"),
    ("Dx_Replace_Straight_Quotes_With_Smart_Quotes", r"'", "\u2019"),
    ("Dx_Fix_Para_Space_Errors", "\u00b7+", " "),
    ("Dx_Fix_Para_Space_Errors", r" +\r", "\r"),
    ("Dx_Fix_Para_Space_Errors", r"\r[ \t]+", "\r"),
    ("Dx_Replace_Function_Application_With_Space", "\u2061", " "),
    ("Dx_Fix_En_Dash_Errors", "\u2013+", "-"),
    ("Dx_Fix_Em_Dash_Space_Errors", " +\u2014| *\u2014 +", "\u2014"),
    ("Sh_Remove_Multi_Spaces", " {2,}", " "),
    ("Sh_Remove_Spaces_Before_Punctuation", r" +([?.!;:\u201d`\]}])", "\\1"),
    ("Sh_Remove_Spaces_Before_Punctuation", r"([\u201c`\[{]) +", "\\1"),
    ("Dx_Replace_Underscore_With_Single_Underscore", "_{2,}", "_"),
    ("Dx_Replace_Fraction_Text_With_Compact_Fractions", r"(?<![0-9/])1/2(?![0-9/])", "\u00bd"),
    ("Dx_Fix_Equals_Before_Para_Mark", r"=\r", "=?\r"),
    ("Dx_Convert_Auto_List_To_Text", r"([0-9]+\.)\t", "\\1 "),
    ("Dx_Convert_Auto_List_To_Text", r"([A-Za-z]+\))\t", "\\1 "),
    ("Dx_Convert_Auto_List_To_Text", r"([0-9]+\))\t", "\\1 "),
    ("Dx_Compress_Linear_Math", "(?i)x", "\u00d7"),
    ("Dx_Compress_Linear_Math", " +", ""),
    ("Dx_Compress_Linear_Math", "[=<>\u2265\u2264\u2260\u2248]", " \\g<0> "),
    ("Sh_Para_Before_Dollar", r"\$pg", "\r$pg"),
]

# What can stand either side of an equation in a book.
NEIGHBORS = ["", " ", "5", "a", "x", '"', "'", ")\t", ".\t", " ?", "\u201c ", " \u201d", "1/2", "=\r", "\r"]


def _run_all(text, skip=()):
    for name, pat, repl in PASSES:
        if (name, pat) not in skip:
            text = re.sub(pat, repl, text)
    return text


def test_the_mark_is_the_one_in_the_source(repo_root):
    body = "\n".join(_procs(repo_root)["Dx_Math_Park_Mark"])
    assert re.search(r'ChrW\(&H27E6\) & "VTMATH" & Format\$\(n, "000000"\) & ChrW\(&H27E7\)', body), (
        "Dx_Math_Park_Mark is not the mark these checks were made against - check the new one "
        "against every pass before changing this")


def test_the_passes_damage_unprotected_mathml():
    """The reason for all this: left in the book, the MathML is changed - by the primes pass
    alone (id="m1"), and by the passes together."""
    assert re.sub(r'([0-9]+)"', "\\1\u2033", MATHML) != MATHML
    assert _run_all(MATHML) != MATHML


def test_no_pass_changes_the_mark_wherever_it_stands():
    for left in NEIGHBORS:
        for right in NEIGHBORS:
            out = _run_all(left + MARK + right)
            assert MARK in out, f"the mark did not survive between {left!r} and {right!r}: {out!r}"


def test_the_mark_leaves_its_neighbors_as_the_tags_did():
    """What is BESIDE the equation must come out as it would beside the element's own "<" and
    ">": a quote after it must not become a prime, a letter and ")" plus a tab must not become a
    list marker. "<q>" stands in for an element no pass touches; Compress Linear Math's padding
    of "<" and ">" is left out, since it would pad the stand-in's own brackets."""
    stand_in = "<q>"
    skip = {("Dx_Compress_Linear_Math", "[=<>\u2265\u2264\u2260\u2248]")}
    for left in NEIGHBORS:
        for right in NEIGHBORS:
            a = _run_all(left + MARK + right, skip).split(MARK)
            b = _run_all(left + stand_in + right, skip).split(stand_in)
            assert len(a) == 2 and len(b) == 2
            assert a == b, f"beside the mark {left!r}/{right!r} came out {a!r}; beside tags {b!r}"


# --- the park and the restore ---------------------------------------------------------------

def _first_code(lines):
    return next(l for l in lines[1:] if l.strip() and not l.strip().startswith("Dim "))


def test_park_stores_before_it_replaces_and_recovers_first(repo_root):
    park = _procs(repo_root)["Dx_Math_Park"]
    text = "\n".join(park)
    # Nested: a no-op that only counts.
    assert re.match(r"\s*If Dx_Math_Depth > 0 Then", _first_code(park))
    # Then: put back what is left, drop what nothing needs, and only then claim the depth.
    back = _at(park, r"^\s*Dx_Math_Put_Back doc")[0]
    prune = _at(park, r"top = Dx_Math_Prune\(doc\)")[0]
    depth = _at(park, r"^\s*Dx_Math_Depth = 1")[0]
    armed = _at(park, r"On Error GoTo ParkFailed")[0]
    assert back < prune < depth < armed
    # Numbered above any stored text that had to be kept, so a leftover never blocks an Add.
    add = _at(park, r"doc\.Variables\.Add Name:=Dx_Math_Park_Var\(top \+ n \+ 1\), Value:=txt")[0]
    swap = _at(park, r"el\.Text = Dx_Math_Park_Mark\(top \+ n\)")[0]
    stored = _at(park, r"If stored Then")[0]
    assert add < stored < swap, "the text is replaced before it is safely stored"
    assert _at(park, r"If txt = Left\$\(win, L\) Then"), "the element is not checked to read back exactly"
    assert _at(park, r"If el\.InlineShapes\.count = 0 And el\.Fields\.count = 0 Then"), (
        "a span holding a picture or a field could be parked, and Range.Text would lose it")
    # Only within the target, and the first Find is the whole cost when there is no MathML.
    assert "pos = target.Paragraphs(1).Range.Start" in text
    assert _at(park, r"If a\.Start >= target\.End Then Exit Do")
    fail = _at(park, r"^\s*ParkFailed:")[0]
    tail = "\n".join(park[fail:])
    assert re.search(r"Dx_Math_Restore doc, True", tail) and "Err.Raise" in tail


def test_restore_keeps_the_stored_texts_for_undo(repo_root):
    """Selected Cleanup and Compress Linear Math keep the undo list, so Ctrl+Z past the restore
    brings the marks back. Their texts must still be there: nothing but the next park's prune
    deletes a stored text."""
    procs = _procs(repo_root)
    for name in ("Dx_Math_Restore", "Dx_Math_Put_Back", "Dx_Math_Before_Save", "Dx_Math_Park",
                 "Dx_Math_On_Open", "Dx_Math_Marks_Waiting"):
        assert not any(".Delete" in l for l in procs[name]), f"{name} deletes a stored text"
    prune = procs["Dx_Math_Prune"]
    assert any(re.search(r"doc\.Variables\(nm\)\.Delete", l) for l in prune)
    # it deletes only those whose mark is NOT still in the book
    walk = _at(prune, r'Sh_Find_Text_Forward\(doc, ChrW\(&H27E6\) & "VTMATH", pos\)')[0]
    dele = _at(prune, r"doc\.Variables\(nm\)\.Delete")[0]
    assert walk < dele and re.search(r"probe = marks\(", "\n".join(prune[walk:dele]))
    assert re.search(r"Dx_Math_Prune = top", "\n".join(prune))

    back = procs["Dx_Math_Put_Back"]
    assert re.match(r"\s*On Error Resume Next", _first_code(back)), "the put-back can raise"
    assert _at(back, r"num = Dx_Math_Mark_Number\(m\.Text\)") and _at(back, r"m\.Text = original")
    assert _at(back, r"If doc\.Variables\.count = 0 Then Exit Function"), "a book never parked pays for a walk"

    rest = "\n".join(procs["Dx_Math_Restore"])
    assert re.search(r"If Dx_Math_Depth > 1 And Not unwinding Then", rest)
    assert re.search(r"Dx_Math_Restore = Dx_Math_Put_Back\(doc\)", rest)


def test_a_save_never_keeps_the_marks(repo_root):
    cls = (repo_root / "src" / "vba" / "VtEvents.cls").read_bytes().decode("latin-1")
    assert re.search(r"Private Sub App_DocumentBeforeSave\(ByVal Doc As Document, SaveAsUI As Boolean, "
                     r"Cancel As Boolean\)\r?\n\s*Dx_Math_Before_Save Doc\r?\n", cls), (
        "nothing puts the MathML back before a save")
    procs = _procs(repo_root)
    body = procs["Dx_Math_Before_Save"]
    assert re.match(r"\s*On Error Resume Next", _first_code(body))
    depth = _at(body, r"If Dx_Math_Depth > 0 Then Exit Sub")[0]
    put = _at(body, r"If Dx_Math_Marks_Waiting\(doc\) Then Dx_Math_Put_Back doc")[0]
    assert depth < put, "a save in the middle of a cleanup puts the MathML back while passes still run"
    assert not any("Sh_Say" in l or "MsgBox" in l or "Cancel" in l for l in body)


def test_opening_a_book_puts_back_what_a_save_kept(repo_root):
    """An End or an error in Word's own box resets the project and unhooks the events for the
    rest of the session, so the save net is gone just when it is needed. Opening the book again
    - later in VtEvents, or already open when AutoExec runs - puts the marks back."""
    cls = (repo_root / "src" / "vba" / "VtEvents.cls").read_bytes().decode("latin-1")
    start = cls.index("Private Sub App_DocumentOpen(ByVal Doc As Document)")
    block = cls[start:cls.index("End Sub", start)]
    assert "Dx_Math_On_Open Doc" in block, "opening a book does not put back its MathML"
    procs = _procs(repo_root)
    auto = procs["AutoExec"]
    loop = _at(auto, r"For Each openDoc In Documents")[0]
    assert re.search(r"Dx_Math_On_Open openDoc", auto[loop + 1])
    body = procs["Dx_Math_On_Open"]
    assert re.match(r"\s*On Error Resume Next", _first_code(body))
    assert _at(body, r"If Dx_Math_Marks_Waiting\(doc\) Then Dx_Math_Put_Back doc")
    assert not any("Sh_Say" in l or "MsgBox" in l for l in body)


def test_looking_for_marks_is_cheap(repo_root):
    """The stored texts stay until the next park, so every save and open of a parked book asks;
    it must not read the whole text of a book to answer."""
    body = _procs(repo_root)["Dx_Math_Marks_Waiting"]
    text = "\n".join(body)
    order = [_at(body, p)[0] for p in (r"If doc\.Variables\.count = 0 Then Exit Function",
                                       r'Left\$\(v\.Name, 10\) = "VtMathPark"',
                                       r'Sh_Find_Text_Forward\(doc, ChrW\(&H27E6\) & "VTMATH", doc\.Content\.Start\)',
                                       r"Sh_Reset_Find_Parameters")]
    assert order == sorted(order), "the cheap tests do not come first, or Find is not put back"
    assert ".Content.Text" not in text and "Selection" not in text


def _between(lines, park_pat, restore_pat, inner_pats):
    p = _at(lines, park_pat)[0]
    r = _at(lines, restore_pat, p)[0]
    for pat in inner_pats:
        for i in _at(lines, pat):
            assert p < i < r, f"{pat!r} at line {i} runs outside the park ({p}) and restore ({r})"
    return p, r


def test_full_file_cleanup_parks_round_every_pass(repo_root):
    """Full File Cleanup - and Attach BANA Template's first cleanup, which calls the same sub."""
    body = _procs(repo_root)["Dx_Fix_Common_File_Errors_Run"]
    p, r = _between(body, PARK + r" mathDoc\.Content", RESTORE + r" mathDoc$",
                    [r'Application\.Run MacroName:="(?!MS_Clear_F_and_R)', r"^\s*Sh_Replace_Manual_Line_Break "])
    assert _at(body, r"On Error GoTo FfcFailed")[0] < p
    fail = _at(body, r"^\s*FfcFailed:")[0]
    tail = "\n".join(body[fail:])
    assert re.search(RESTORE + r" mathDoc, True", tail) and "Err.Raise" in tail, (
        "an error in Full File Cleanup leaves the marks in the book")
    attach = "\n".join(_procs(repo_root)["Dx_Attach_BANA_Template_Run"])
    assert "Dx_Fix_Common_File_Errors_Run True" in attach


def test_selected_cleanup_parks_the_selection(repo_root):
    code = form_code(repo_root, "Dx_Selected_Cleanup_Form")
    start = _at(code, r"^Private Sub Okay_Click")[0]
    end = _at(code, r"^End Sub", start)[0]
    body = code[start:end + 1]
    p, r = _between(body, PARK + r" keep$", RESTORE + r" mathDoc$", [r"Application\.Run "])
    assert _at(body, r"Set keep = Selection\.Range\.Duplicate")[0] < p
    assert _at(body, r"On Error GoTo CleanupFailed")[0] < p
    fail = _at(body, r"^\s*CleanupFailed:")[0]
    rr = _at(body, RESTORE + r" mathDoc, True", fail)[0]
    assert rr < _at(body, r"Sh_Report_Error", fail)[0], "an error in Selected Cleanup leaves the marks"


def test_compress_linear_math_stays_in_the_selection(repo_root):
    """Found by reading, 10/18/2026: the x-to-times-sign pass ran with wdFindContinue, so with a
    selection Replace All went on through the whole book - every x became a times sign."""
    body = _procs(repo_root)["Dx_Compress_Linear_Math"]
    assert not any("wdFindContinue" in l for l in body), "a Compress Linear Math pass still runs past the selection"
    x = _at(body, r'^\s*\.Text = "x"')[0]
    assert re.search(r"\.Wrap = wdFindStop", "\n".join(body[x:x + 4]))
    p, r = _between(body, PARK + r" Selection\.Range", RESTORE + r" mathDoc$", [r"Selection\.Find\.Execute Replace:="])
    assert _at(body, r"On Error GoTo CompressFailed")[0] < p
    fail = _at(body, r"^\s*CompressFailed:")[0]
    assert _at(body, RESTORE + r" mathDoc, True", fail)[0] < _at(body, r"Sh_Report_Error", fail)[0]


def test_autotag_parks_round_its_spacing_pass(repo_root):
    body = _procs(repo_root)["Dx_AutoTag_Page_Numbers"]
    run = _at(body, r'Application\.Run MacroName:="Dx_Fix_Para_Space_Errors"')
    assert len(run) == 1
    assert re.search(PARK + r" ActiveDocument\.Content", body[run[0] - 1])
    assert re.search(RESTORE + r" ActiveDocument$", body[run[0] + 1])
    screen = _at(body, r"Application\.ScreenUpdating = False")[0]
    armed = _at(body, r"On Error GoTo MathFailed")[0]
    assert screen < armed < run[0] - 1, "the park is not after the screen goes off, or has no error path"
    assert re.match(r"\s*On Error GoTo 0", body[run[0] + 2])
    fail = _at(body, r"^\s*MathFailed:")[0]
    tail = "\n".join(body[fail:])
    assert re.search(RESTORE + r" ActiveDocument, True", tail) and "Err.Raise" in tail


def test_exercise_levels_parks_in_the_scratch_document(repo_root):
    body = _procs(repo_root)["Dx_Exercise_Levels_Hidden"]
    p, r = _between(body, PARK + r" tempDoc\.Content", RESTORE + r" tempDoc$", [r"^\s*Dx_Ex_Passes tempDoc"])
    assert r < _at(body, r"work\.FormattedText = body\.FormattedText")[0], "the list goes home still holding marks"
    fail = _at(body, r"^\s*eom:")[0]
    assert _at(body, RESTORE + r" tempDoc, True", fail), "an error leaves the park depth set"


# Passes that would change MathML text, and the only places allowed to run them: each of these
# parks round the call (checked above), or is itself reached only from one that does.
DANGEROUS = ("Dx_Fix_Primes", "Dx_Replace_Straight_Quotes_With_Smart_Quotes", "Dx_Fix_Para_Space_Errors",
             "Dx_Replace_Function_Application_With_Space", "Dx_Fix_En_Dash_Errors",
             "Dx_Fix_Em_Dash_Space_Errors", "Dx_Replace_Spaces_Before_Punctuation",
             "Dx_Replace_Fraction_Text_With_Compact_Fractions", "Dx_Fix_Equals_Before_Para_Mark",
             "Dx_Add_Qmark_To_Incomplete_Equations", "Dx_Ex_Passes", "Dx_Remove_Bullets")
PROTECTED = {"Dx_Fix_Common_File_Errors_Run", "Dx_AutoTag_Page_Numbers", "Dx_Exercise_Levels_Hidden",
             "Dx_Remove_Bullets", "Dx_Ex_Passes", "Dx_Selected_Cleanup_Form"}


def test_nothing_else_runs_the_dangerous_passes(repo_root):
    procs = _procs(repo_root)
    callers = {}
    for name, lines in procs.items():
        for l in lines[1:]:
            for d in DANGEROUS:
                if d != name and re.search(r'(?<!\w)' + d + r'(?!\w)', l):
                    callers.setdefault(d, set()).add(name)
    for formfile in (repo_root / "src" / "forms").glob("*.frm"):
        code = "\n".join(form_code(repo_root, formfile.stem))
        for d in DANGEROUS:
            if re.search(r'(?<!\w)' + d + r'(?!\w)', code):
                callers.setdefault(d, set()).add(formfile.stem)
    for d, who in callers.items():
        assert who <= PROTECTED, f"{d} is run from {sorted(who - PROTECTED)}, which does not set the MathML aside"
    assert callers["Dx_Remove_Bullets"] == {"Dx_Selected_Cleanup_Form"}
    assert callers["Dx_Ex_Passes"] == {"Dx_Exercise_Levels_Hidden"}


def test_the_helpers_take_byval_and_have_examples(repo_root):
    procs = _procs(repo_root)
    for name in ("Dx_Math_Park_Mark", "Dx_Math_Park_Var", "Dx_Math_Mark_Number", "Dx_Math_Element_Length",
                 "Dx_Math_Park", "Dx_Math_Restore", "Dx_Math_Put_Back", "Dx_Math_Prune",
                 "Dx_Math_Before_Save", "Dx_Math_On_Open", "Dx_Math_Marks_Waiting"):
        head = procs[name][0]
        params = head[head.index("(") + 1:head.rindex(")")].split(")")[0]
        for prm in filter(None, (x.strip() for x in params.split(","))):
            assert prm.replace("Optional ", "", 1).startswith("ByVal "), f"{name}: {prm!r} is not ByVal"
        assert not any("MsgBox" in l for l in procs[name])
    vba = (repo_root / "tests" / "vba" / "TestMathML.bas").read_bytes().decode("latin-1")
    for name in ("Dx_Math_Park_Mark", "Dx_Math_Mark_Number", "Dx_Math_Element_Length"):
        assert f"'@uses {name}" in vba
