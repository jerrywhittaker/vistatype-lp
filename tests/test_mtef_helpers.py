"""MathType equations made from the book's MathML: the MTEF encoder and the OLE file (10/20/2026).

Steps 1 and 2 of the MathType plan: Sh_MTEF_From_MathML writes an equation as MathType's MTEF,
Sh_MTEF_Ole_File wraps it in the OLE compound file Word keeps a MathType equation in, and
Sh_Base64 turns that into the text a Flat-OPC package carries. Nothing calls them yet - putting
them in the book is Step 3.

What they write was measured against MathType itself on VistaBuild: Section D of the Imagine
book (453 of 470 equations byte for byte; the other 17 differ only by issue #38's fault in the
target) and 355 more equations through MathType's own Convert Equations. The real checks of the
bytes are in tests/vba/TestMTEF.bas, which runs the VBA. These are checks on the source: every
helper Private, every parameter ByVal (the rule from the 3.0.274 compile error), a Version line
and a worked example in every header, no message boxes, and the VBA tests wired in.

11/5/2026 (Jerry, Basic Math Pre-Algebra for Dummies: 503 of 1,861 equations fell back to Word
equations): mspace, no-break spaces, plain-text mi runs, % and $, infinity and aleph, the Euclid
fonts, borrowing marks (mmultiscripts with mprescripts) and long division (issue #50). The checks
at the end hold the measured values in the source, and keep tools/compare_mtef.py - the Python
copy of the walker that measures a whole book - in step with the VBA.
"""
import ast
import re

from test_toc_never_deletes_a_line import module_text, procedures

HELPERS = ("Sh_MTEF_From_MathML", "Sh_MTEF_Not_Handled", "Sh_MTEF_Walk", "Sh_MTEF_Kids",
           "Sh_MTEF_One", "Sh_MTEF_Collapse", "Sh_MTEF_Mo_Text", "Sh_MTEF_Match_Fence",
           "Sh_MTEF_Fail", "Sh_MTEF_Add", "Sh_MTEF_Set_Size", "Sh_MTEF_Put_Char",
           "Sh_MTEF_Put_Template", "Sh_MTEF_Put_Line", "Sh_MTEF_Put_Fence", "Sh_MTEF_Put_Row",
           "Sh_MTEF_Put_Node", "Sh_MTEF_Script_Size", "Sh_MTEF_One_Char", "Sh_MTEF_Put_Scripts",
           "Sh_MTEF_Put_Fraction", "Sh_MTEF_Put_Root", "Sh_MTEF_Put_Enclose", "Sh_MTEF_Put_Fenced",
           "Sh_MTEF_Put_Accent", "Sh_MTEF_Put_Table", "Sh_MTEF_Symbol_Place",
           "Sh_MTEF_Extra_Place", "Sh_MTEF_Greek_Place", "Sh_MTEF_Put_Token",
           "Sh_MTEF_Style_Index", "Sh_MTEF_Space_Codes", "Sh_MTEF_Put_Space",
           "Sh_MTEF_Put_Prescripts", "Sh_MTEF_Euclid_Place",
           "Sh_MTEF_Preamble_Hex", "Sh_MTEF_Hex_Text", "Sh_MTEF_Text_Bytes", "Sh_MTEF_U16",
           "Sh_MTEF_U32", "Sh_MTEF_Dir_Entry", "Sh_MTEF_Ole_File", "Sh_Base64")


def _procs(repo_root):
    return procedures(module_text(repo_root))


def _signature(lines):
    head, n = lines[0], 1
    while head.rstrip().endswith("_"):
        head = head.rstrip()[:-1] + " " + lines[n]
        n += 1
    return head, n


def _header_comment(repo_root, name):
    """The comment block right after the signature, up to the first line that is not one.

    procedures() takes the comments off, so this reads the module text itself.
    """
    lines = module_text(repo_root).splitlines()
    start = next(i for i, l in enumerate(lines)
                 if re.match(r"Private (Sub|Function) %s\(" % name, l))
    n = start
    while lines[n].rstrip().endswith("_"):
        n += 1
    block = []
    for line in lines[n + 1:]:
        if not line.strip().startswith("'"):
            break
        block.append(line.strip())
    return block


def test_every_helper_is_there_and_private(repo_root):
    procs = _procs(repo_root)
    for name in HELPERS:
        assert name in procs, f"{name} is missing"
        assert procs[name][0].startswith("Private "), f"{name} is not Private"


def test_every_helper_takes_byval(repo_root):
    procs = _procs(repo_root)
    for name in HELPERS:
        head, _ = _signature(procs[name])
        params = head[head.index("(") + 1:head.rindex(")")].split(")")[0]
        for p in filter(None, (x.strip() for x in params.split(","))):
            assert p.startswith("ByVal "), f"{name}: parameter {p!r} is not ByVal"


def test_every_helper_has_a_version_and_an_example(repo_root):
    """Like its neighbours, each header says its version and shows the helper in use."""
    for name in HELPERS:
        block = _header_comment(repo_root, name)
        assert any(re.match(r"' Version: \d+\.\d+  Date: \d{1,2}/\d{1,2}/20\d\d", l) for l in block), \
            f"{name} has no Version line"
        examples = [l for l in block if l.startswith("'   ") and name in l]
        assert examples, f"{name} has no example line naming it in its header"


def test_no_message_boxes(repo_root):
    procs = _procs(repo_root)
    for name in HELPERS:
        assert not any("MsgBox" in l for l in procs[name]), f"{name} uses a bare MsgBox"
        assert not any("Sh_Say" in l or "Sh_Ask" in l for l in procs[name]), \
            f"{name} talks to the transcriber - these are pure helpers"


def test_no_word_objects(repo_root):
    """Pure helpers: no document, range or selection - that is what lets tests/vba run them."""
    procs = _procs(repo_root)
    for name in HELPERS:
        for line in procs[name]:
            code = line.split("'")[0]
            assert not re.search(r"\b(ActiveDocument|Selection|As Range|As Document)\b", code), \
                f"{name} reaches into Word: {line.strip()}"


def test_only_the_mathtype_route_calls_them(repo_root):
    """Step 3 put them in the book through Sh_MathML_Put_MathType and nothing else - the Word
    equations and pictures routes do not touch them (and Sh_MathML_Put_In asks Sh_MTEF_Not_Handled
    nothing: a failed equation simply falls back)."""
    procs = _procs(repo_root)
    callers = set()
    for name, lines in procs.items():
        if name in HELPERS:
            continue
        for line in lines:
            if re.search(r"\bSh_MTEF_|\bSh_Base64\b", line.split("'")[0]):
                callers.add(name)
    assert callers == {"Sh_MathML_Put_MathType"}, callers


def test_the_preamble_is_193_bytes(repo_root):
    body = "".join(_procs(repo_root)["Sh_MTEF_Preamble_Hex"])
    hexes = "".join(re.findall(r'"([0-9A-F]+)"', body))
    assert len(hexes) == 386
    assert hexes.startswith("050100070C44534D543700")   # MTEF 5, Windows, MathType 7, "DSMT7"
    assert hexes.endswith("0A0100")                     # FULL, then the LINE the equation goes in


def test_the_vba_tests_are_wired_in(repo_root):
    vba = (repo_root / "tests" / "vba" / "TestMTEF.bas").read_bytes().decode("latin-1")
    for name in ("Sh_MTEF_From_MathML", "Sh_MTEF_Not_Handled", "Sh_MTEF_Ole_File", "Sh_Base64"):
        assert f"'@uses {name}" in vba
    assert vba.count("With Suite.Test(") >= 30
    runner = (repo_root / "tests" / "vba" / "VtRunTests.bas").read_bytes().decode("latin-1")
    assert "TestMTEF_Suite()" in runner


# --- 11/5/2026: what Basic Math sent to Word equations ------------------------------------------

def _code(repo_root, name):
    return "\n".join(_procs(repo_root)[name])


def _vba_space_codes(repo_root):
    """Sh_MTEF_Space_Codes as {width: codes}."""
    table = {}
    for line in _procs(repo_root)["Sh_MTEF_Space_Codes"]:
        m = re.match(r'\s*Case (.+?):?\s*(?:Sh_MTEF_Space_Codes = "(\w+)")?\s*$', line)
        if m and m.group(1) != "Else":
            widths = re.findall(r'"([^"]*)"', m.group(1).split(":")[0])
            code = m.group(2) or None
            table.update({w: code for w in widths})
        m = re.match(r'\s*Sh_MTEF_Space_Codes = "(\w+)"', line)
        if m:
            for w, c in table.items():
                if c is None:
                    table[w] = m.group(1)
    return table


def _tool_table(repo_root, name):
    tree = ast.parse((repo_root / "tools" / "compare_mtef.py").read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(getattr(t, "id", "") == name for t in node.targets):
            return ast.literal_eval(node.value)
    raise AssertionError(name + " is not in tools/compare_mtef.py")


def test_mspace_widths_are_the_measured_ones(repo_root):
    """Measured 11/5/2026 through MathType's own MathML import: Basic Math's two widths are EF03."""
    vba = _vba_space_codes(repo_root)
    assert vba["thickmathspace"] == "03" and vba["0.25em"] == "03"
    assert vba["1em"] == "05" and vba["0.5em"] == "0403" and vba[""] == "01"
    assert "3pt" not in vba, "a width in points depends on the font size - not measured"
    tool = {w: "".join("%02X" % (c & 0xFF) for c in codes)
            for w, codes in _tool_table(repo_root, "SPACES").items()}
    assert vba == tool, "Sh_MTEF_Space_Codes and tools/compare_mtef.py SPACES differ"


def test_mspace_is_written_and_not_refused(repo_root):
    node = _code(repo_root, "Sh_MTEF_Put_Node")
    assert '"mspace"' in node and "Sh_MTEF_Put_Space st, node, kids, size" in node
    assert "Sh_MTEF_Put_Char st, 24, &HEF00& + " in _code(repo_root, "Sh_MTEF_Put_Space")


def test_long_division_is_template_26(repo_root):
    """Issue #50: as a Word equation the dividend is lost; here it is template 26, the dividend,
    and an absent slot - measured through MathType's own MathML import."""
    code = _code(repo_root, "Sh_MTEF_Put_Enclose")
    block = code[code.index('Case "longdiv"'):]
    block = block[:block.index("Exit Sub")]
    assert "Sh_MTEF_Put_Template st, 26, 0, size" in block
    assert "Sh_MTEF_Put_Line st, kids, 1, kids.count, size" in block
    assert "Sh_MTEF_Put_Line st, Nothing, 0, 0, size" in block


def test_borrowing_marks_are_template_29_variation_1(repo_root):
    code = _code(repo_root, "Sh_MTEF_Put_Prescripts")
    assert "Sh_MTEF_Put_Template st, 29, 1, size" in code
    assert code.index("Sh_MTEF_Put_Template st, 29, 1") < code.index("Sh_MTEF_Put_Node st, kids(1), size"), \
        "the base goes AFTER the template, as MathType wrote it"


def test_font_style_records_are_numbered_in_order(repo_root):
    """The old one-flag "styled" is gone: bold italic and the Euclid fonts share one numbering."""
    for name in HELPERS:
        assert not any('"styled"' in l for l in _procs(repo_root)[name]), name
    assert 'st("styles") = ""' in _code(repo_root, "Sh_MTEF_Walk")
    code = _code(repo_root, "Sh_MTEF_Style_Index")
    assert "6 + euclids" in code and "5 + euclids" in code
    char = _code(repo_root, "Sh_MTEF_Put_Char")
    for key in ('"B"', '"E1"', '"E2"'):
        assert "Sh_MTEF_Style_Index(st, %s)" % key in char


def test_euclid_places_match_the_tool(repo_root):
    vba = {}
    for line in _procs(repo_root)["Sh_MTEF_Euclid_Place"]:
        m = re.match(r"\s*Case &H([0-9A-F]+): Sh_MTEF_Euclid_Place = &H([0-9A-F]+)", line)
        if m:
            vba[int(m.group(1), 16)] = int(m.group(2), 16)
    tool = {k: (0x100 if f == -11 else 0x200) + p for k, (f, p) in _tool_table(repo_root, "EUCLID_PLACE").items()}
    assert vba == tool and len(vba) == 6


def test_no_break_space_and_aleph(repo_root):
    token = _code(repo_root, "Sh_MTEF_Put_Token")
    assert "ElseIf code = &H20 Or code = &HA0 Then" in token          # mi: TEXT
    assert "ElseIf code = &HA0 Or code = &H2019 Then" in token        # mo: FUNCTION
    assert "Sh_MTEF_Symbol_Place(code) >= 0 And code <> &H2135" in token, "aleph is measured in mi only"
    assert "Case &H2135: Sh_MTEF_Symbol_Place = &HC0" in _code(repo_root, "Sh_MTEF_Symbol_Place")


def test_the_comparison_tool_is_read_only(repo_root):
    """tools/compare_mtef.py measures; it must never write into the files it is given."""
    text = (repo_root / "tools" / "compare_mtef.py").read_text(encoding="utf-8")
    assert not re.search(r"open\([^)]*['\"][wa]b?['\"]", text)
    assert ".write(" not in text and ".save(" not in text
