"""DAISY/NIMAS conversion: the equations, and box 325 asking large print or braille (issue #37).

Jerry, 10/18/2026: none of the "EQ" pictures of the Imagine IM A1 7&8 NIMAS book reached the Word
file. Every one is named only in the altimg of a MathML equation, and Word's HTML importer drops
MathML. His rules, the same day:

- Box 325 asks LARGE PRINT OR BRAILLE ("LP" / "BRL") instead of keep or omit images. Both keep
  the pictures (braille drops them elsewhere), so the loop that deleted them is gone.
- The XML is looked at for MathML first, and only a large print book that HAS some is asked
  Word equations or EQ pictures ("Equations (398)"). Cancel stops with a message (399) before
  anything is done to the book.
- Each equation is parked as a VTEQnnnnnn placeholder for the import and put back afterwards:
  a Word equation (or its picture if that fails) or the picture for large print. (Braille got
  the MathML as text, for DBT, until 10/20/2026.)

Jerry, 10/20/2026: braille gets Word equations (OMML) instead, exactly as large print's "Word
equations (OMML)" choice does, and box 398 says the book holds equations before it asks. The
same day he had the MathML-as-text code removed - Sh_MathML_Put_Text, Sh_MathML_One_Line,
Sh_MathML_Unprefixed and the "BRL" mode of Sh_MathML_Put_In - and with it the braille cleanup
park (Dx_Math_Park and its helpers) that kept the cleanup passes off that text. The last test
here checks that none of it is defined or called anywhere in src/.

The question forms and the Word side cannot run headlessly (tests/vba/README.md), so these are
checks on the source. The string helpers are tested for real in tests/vba/TestMathML.bas.
"""
import re

from test_toc_box_stays_open import form_code
from test_toc_never_deletes_a_line import module_text, procedures

CONVERTER = "Sh_Convert_XML_File_To_Word_Document"
HELPERS = ("Sh_MathML_Namespace", "Sh_MathML_Prefix", "Sh_MathML_Pattern", "Sh_MathML_Element_Count",
           "Sh_MathML_Placeholder", "Sh_MathML_Fragments", "Sh_MathML_To_Placeholders",
           "Sh_MathML_Altimg", "Sh_MathML_Picture_Path", "Sh_MathML_With_Namespace",
           "Sh_Strip_Xml_Declaration", "Sh_OMML_Package",
           "Sh_Find_Text_From", "Sh_MathML_Put_Equation", "Sh_MathML_Put_Picture",
           "Sh_MathML_Put_In",
           # the review, 10/18/2026
           "Sh_Xml_Is_Space", "Sh_Xml_Start_Tag_Length", "Sh_Xml_Attribute", "Sh_Digits_Value",
           "Sh_Chr_From_Code", "Sh_Xml_Entity_Char", "Sh_Xml_Unescape", "Sh_Url_Unescape",
           "Sh_Find_Text_Forward",
           # the speed fix, 10/18/2026
           "Sh_MathML_Placeholder_Positions",
           # brackets MML2OMML.XSL split, issue #38, 10/20/2026
           "Sh_MathML_Wrap_Fences", "Sh_MathML_Wrap_Fences_In",
           # equations hidden in comments, issue #40, 10/25/2026
           "Sh_MathML_Uncomment")


def _procs(repo_root):
    return procedures(module_text(repo_root))


def _converter(repo_root):
    return _procs(repo_root)[CONVERTER]


def _first(lines, pattern):
    hits = [i for i, l in enumerate(lines) if re.search(pattern, l)]
    assert hits, f"{CONVERTER} has no line matching {pattern!r}"
    return hits[0]


def test_325_answers_large_print_or_braille(repo_root):
    code = "\n".join(form_code(repo_root, "DN_Keep_Or_Omit_Images_Form"))
    assert re.search(r'Sh_GP_String_2\s*=\s*"LP"', code) and re.search(r'Sh_GP_String_2\s*=\s*"BRL"', code)
    assert '"KEEP"' not in code and '"OMIT"' not in code
    frm = (repo_root / "src" / "forms" / "DN_Keep_Or_Omit_Images_Form.frm").read_bytes().decode("latin-1")
    assert 'Caption         =   "Large Print or Braille (325)"' in frm

    body = "\n".join(_converter(repo_root))
    assert re.search(r'If Sh_GP_String_2 <> "LP" And Sh_GP_String_2 <> "BRL" Then', body)


def test_no_omit_branch(repo_root):
    body = "\n".join(_converter(repo_root))
    assert '"OMIT"' not in body and '"KEEP"' not in body, "the converter still knows keep/omit images"
    assert not re.search(r"InlineShapes\(\w+\)\.Delete", body), "the converter still deletes pictures"
    assert "BreakLink" in body, "the converter no longer embeds the pictures"


def test_math_question_only_with_mathml_for_both_outputs(repo_root):
    """Jerry, 10/22/2026: Equations (398) is asked for braille as well as large print, whenever
    the book has MathML - no longer inside a large print branch."""
    lines = _converter(repo_root)
    found = _first(lines, r'mathPrefix = Sh_MathML_Prefix\(fileContent\)')
    has = _first(lines, r'If mathPrefix <> "#" Then')
    ask = _first(lines, r"DN_Equations_Form\.Show")
    assert found < has < ask, "the equation question is not inside 'has MathML'"
    between = "\n".join(lines[has:ask])
    assert "Sh_GP_String_2" not in between, "the equation question still depends on 325's answer"
    assert sum("DN_Equations_Form.Show" in l for l in lines) == 1


def test_braille_grays_out_the_pictures(repo_root):
    """Jerry, 10/22/2026: on a braille conversion 398 has "Pictures of the equations" grayed out,
    and a grayed-out choice can never be the answer."""
    code = "\n".join(form_code(repo_root, "DN_Equations_Form"))
    init = code[code.index("Private Sub UserForm_Initialize()"):]
    assert re.search(r'If Sh_GP_String_2 = "BRL" Then EquationPictures\.Enabled = False', init)
    assert re.search(r'If EquationPictures\.Value = True And EquationPictures\.Enabled Then\s+'
                     r'Sh_GP_String_3 = "PICTURE"', code)
    lines = _converter(repo_root)
    assert 'mathMode = "BRL"' not in "\n".join(lines), "braille still gets the MathML as text"


def test_398_is_gone_from_the_screen_before_the_bar(repo_root):
    """Jerry, testing 3.0.600: after "yes" on 400, the bar appeared with 398 still drawn. Word
    gets a DoEvents to redraw between the answer and the bar."""
    lines = _converter(repo_root)
    ask = _first(lines, r"DN_Equations_Form\.Show")
    answer = _first(lines, r"mathMode = Sh_GP_String_3")
    bar = _first(lines, r"Sh_Progress_Open")
    assert ask < answer < bar
    assert any(re.match(r"\s*DoEvents\s*$", l) for l in lines[answer:bar])


def test_398_says_the_book_has_equations(repo_root):
    frx = (repo_root / "src" / "forms" / "DN_Equations_Form.frx").read_bytes().decode("latin-1")
    assert "This file contains math equations. How do you want them presented in the output file?" in frx
    assert "Word equations (OMML)" in frx


def test_every_question_comes_before_anything_changes(repo_root):
    """The book is read first (so MathML can be found) and every question is asked before the
    bar opens, before a file is written and before the Word document exists."""
    lines = _converter(repo_root)
    read = _first(lines, r"fileContent = tStream\.ReadText")
    ask = _first(lines, r"DN_Equations_Form\.Show")
    q325 = _first(lines, r"DN_Keep_Or_Omit_Images_Form\.Show")
    for later in (r"Sh_Progress_Open", r"\.SaveToFile", r"Documents\.Add", r"InsertFile"):
        at = _first(lines, later)
        assert q325 < read < ask < at, f"{later} comes before a question or before the book is read"


def test_cancel_stops_with_a_message(repo_root):
    lines = _converter(repo_root)
    ask = _first(lines, r"DN_Equations_Form\.Show")
    near = "\n".join(lines[ask:ask + 8])
    assert re.search(r'If Sh_GP_String_3 <> "OMML" And Sh_GP_String_3 <> "MATHTYPE" And Sh_GP_String_3 <> "PICTURE" Then', near)
    assert re.search(r'Sh_Say "[^"]+",\s*_?\s*"VistaType LP \(399\)"', near.replace("\n", " "))
    assert re.search(r"Exit Sub", near)

    code = "\n".join(form_code(repo_root, "DN_Equations_Form"))
    assert re.search(r'Sh_GP_String_3\s*=\s*"OMML"', code) and re.search(r'Sh_GP_String_3\s*=\s*"PICTURE"', code)
    assert re.search(r'Sh_GP_String_3\s*=\s*"MATHTYPE"', code)
    frm = (repo_root / "src" / "forms" / "DN_Equations_Form.frm").read_bytes().decode("latin-1")
    assert 'Caption         =   "Equations (398)"' in frm


def test_the_equations_are_parked_and_put_back(repo_root):
    lines = _converter(repo_root)
    # parked from the import text, its hidden equations brought out of their comments (issue #40)
    park = _first(lines, r"Sh_MathML_To_Placeholders\(importText, mathPrefix\)")
    imp = _first(lines, r"InsertFile fileName:=importPath")
    pics = _first(lines, r"Next shp")
    back = _first(lines, r"Sh_MathML_Put_In\(finalDoc, mathFragments, mathPrefix, mathMode, folderPath")
    font = _first(lines, r"Sh_Set_Whole_Document_Font doc")
    quotes = _first(lines, r"Sh_Strip_Prodnote_Enclosing_Quotes finalDoc")
    assert park < imp < pics < back < font, "the equations are not parked before the import and put back after"
    assert quotes < back, "the equations go back before the prodnote quote pass, which deletes by position"
    # The .txt and .html beside the book keep their MathML: the placeholders go in a file of their own.
    assert _first(lines, r'importPath = folderPath & "~VistaType-import\.html"') > _first(lines, r"SaveToFile txtCopy")


def test_each_mode_puts_in_what_it_should(repo_root):
    procs = _procs(repo_root)
    put_in = "\n".join(procs["Sh_MathML_Put_In"])
    # Braille comes as "OMML", like large print's Word equations (Jerry, 10/20/2026); "MATHTYPE"
    # is large print's third choice (Steps 3 and 4 of the MathType plan, 10/20/2026). "PICTURE"
    # is the Case Else; there is no "BRL" (MathML text) case any more.
    cases = re.findall(r"^\s*Case\s+(.+?)\s*$", put_in, re.M)
    assert cases == ['"OMML"', '"MATHTYPE"', "Else"], f"Sh_MathML_Put_In's modes are {cases}"
    assert re.search(r"Sh_MathML_Put_Equation\(r, xslt, xsl,\s*_?\s*Sh_MathML_With_Namespace\(", put_in.replace("\n", " "))
    assert re.search(r"If Not done Then\s+done = Sh_MathML_Put_Picture\(r, picPath\)\s+If done Then asPicture",
                     put_in), "a Word equation that fails does not fall back to its picture"
    assert 'MML2OMML.XSL' in put_in

    pic = "\n".join(procs["Sh_MathML_Put_Picture"])
    assert re.search(r"AddPicture\W.*LinkToFile:=False, SaveWithDocument:=True", pic), "the picture is not embedded"

    eq = "\n".join(procs["Sh_MathML_Put_Equation"])
    assert "MSXML2.DOMDocument.6.0" in eq and "parseError" in eq and "transformNode" in eq
    assert re.search(r"r\.InsertXML Sh_OMML_Package\(", eq)

    # The count of equations that did not go in as asked reaches the closing message.
    body = "\n".join(_converter(repo_root))
    assert re.search(r'MsgBox "Conversion is complete[^\n]*& mathNote, vbInformation, "Operation Complete \(313\)"', body)


def test_no_new_msgbox(repo_root):
    procs = _procs(repo_root)
    for name in HELPERS:
        assert name in procs, f"{name} is missing"
        assert not any("MsgBox" in l for l in procs[name]), f"{name} uses a bare MsgBox"
    assert not any("MsgBox" in l for l in form_code(repo_root, "DN_Equations_Form"))
    assert not any("MsgBox" in l for l in form_code(repo_root, "DN_Keep_Or_Omit_Images_Form"))
    # The converter's own, as before: 308, 309, 310, 311, 295, 312 and 313, plus the line that
    # declares 312's answer As VbMsgBoxResult. None added.
    assert sum("MsgBox" in l for l in _converter(repo_root)) == 8


def test_every_helper_takes_byval(repo_root):
    procs = _procs(repo_root)
    for name in HELPERS:
        head = procs[name][0]
        n = 1
        while head.rstrip().endswith("_"):
            head = head.rstrip()[:-1] + " " + procs[name][n]
            n += 1
        params = head[head.index("(") + 1:head.rindex(")")].split(")")[0]
        for p in filter(None, (x.strip() for x in params.split(","))):
            assert p.startswith("ByVal "), f"{name}: parameter {p!r} is not ByVal"


def test_the_helpers_have_vba_examples(repo_root):
    vba = (repo_root / "tests" / "vba" / "TestMathML.bas").read_bytes().decode("latin-1")
    for name in ("Sh_MathML_Prefix", "Sh_MathML_To_Placeholders", "Sh_MathML_Picture_Path",
                 "Sh_MathML_With_Namespace", "Sh_MathML_Placeholder_Positions",
                 "Sh_MathML_Wrap_Fences"):
        assert f"'@uses {name}" in vba
    runner = (repo_root / "tests" / "vba" / "VtRunTests.bas").read_bytes().decode("latin-1")
    assert "TestMathML_Suite()" in runner


# --- The review of 10/18/2026 ---------------------------------------------------------------

def _body(repo_root, name):
    return _procs(repo_root)[name]


def test_log_line_keeps_every_character(repo_root):
    """Sh_One_Log_Line's "below 32" test read AscW, which is NEGATIVE from U+8000 up, so every
    such character - both halves of a math letter like U+1D465, every Private Use character -
    became a space. Found with the braille MathML text (removed 10/20/2026); the log-line fix
    stays."""
    log = "\n".join(_body(repo_root, "Sh_One_Log_Line"))
    assert re.search(r"code = AscW\(ch\) And &HFFFF&", log), "Sh_One_Log_Line still reads AscW signed"


def test_the_stylesheet_is_compiled_once(repo_root):
    put_in = _body(repo_root, "Sh_MathML_Put_In")
    text = "\n".join(put_in)
    loop = next(i for i, l in enumerate(put_in) if re.match(r"\s*For i = 1 To n\b", l))
    tpl = next(i for i, l in enumerate(put_in) if '"MSXML2.XSLTemplate.6.0"' in l)
    ft = next(i for i, l in enumerate(put_in) if '"MSXML2.FreeThreadedDOMDocument.6.0"' in l)
    assert ft < tpl < loop, "the template is not made once, before the walk"
    assert re.search(r"Set xslt\.stylesheet = xsl", text)
    assert '"MSXML2.DOMDocument.6.0"' in text, "no fallback stylesheet if the free-threaded one fails"
    assert not any("XSLTemplate" in l or "MML2OMML" in l for l in put_in[loop:]), "something is loaded per equation"

    eq = "\n".join(_body(repo_root, "Sh_MathML_Put_Equation"))
    assert "xslt.createProcessor()" in eq and "proc.transform" in eq and "proc.output" in eq
    assert re.search(r"If Len\(omml\) = 0 Then omml = src\.transformNode\(xsl\)", eq), (
        "no fallback to transformNode when the template route fails")
    assert "CreateObject(\"MSXML2.XSLTemplate" not in eq, "the template is made per equation"


def test_altimg_reads_quoted_values(repo_root):
    alt = "\n".join(_body(repo_root, "Sh_MathML_Altimg"))
    assert "Sh_Xml_Start_Tag_Length(fragment)" in alt and 'Sh_Xml_Attribute(Left$(fragment, tagLen), "altimg")' in alt
    assert "[^>]" not in alt, "altimg still stops at the first >"
    pattern = "\n".join(_body(repo_root, "Sh_MathML_Pattern"))
    assert "[^>]*/>" not in pattern, "an empty math element still stops at the first >"
    body = "\n".join(_body(repo_root, "Sh_MathML_With_Namespace"))
    assert "Sh_Xml_Start_Tag_Length(" in body and 'InStr(1, fragment, ">")' not in body

    path = "\n".join(_body(repo_root, "Sh_MathML_Picture_Path"))
    assert "Sh_Url_Unescape(Sh_Xml_Unescape(p))" in path, "altimg is not unescaped, XML first, then URL"
    assert 'Replace(p, "%20", " ")' not in path


def test_find_is_reset_after_the_walk(repo_root):
    put_in = _body(repo_root, "Sh_MathML_Put_In")
    nxt = next(i for i, l in enumerate(put_in) if re.match(r"\s*Next i\b", l))
    assert any("Sh_Reset_Find_Parameters" in l for l in put_in[nxt:nxt + 3]), "Find is not reset after the walk"
    fail = next(i for i, l in enumerate(put_in) if re.match(r"\s*PutInFailed:", l))
    assert "On Error GoTo PutInFailed" in "\n".join(put_in[:nxt])
    tail = "\n".join(put_in[fail:])
    assert "Sh_Reset_Find_Parameters" in tail and "Err.Raise" in tail, "the error path does not reset Find"


def test_picture_goes_in_before_the_placeholder_comes_out(repo_root):
    pic = _body(repo_root, "Sh_MathML_Put_Picture")
    add = next(i for i, l in enumerate(pic) if "AddPicture" in l)
    out = next(i for i, l in enumerate(pic) if re.match(r'\s*ph\.Text = ""', l))
    assert add < out, "the placeholder is taken out before the picture is in"
    assert not any(re.match(r'\s*r\.Text = ""', l) for l in pic), "the placeholder is still cleared first"
    assert re.search(r"If ph\.Text <> phText Then", "\n".join(pic)), "the placeholder is not checked before it goes"


# --- The speed fix, 10/18/2026 (measured on the build box, 2,139 equations: 556 s to 186 s) ----

def test_undo_list_is_cleared_once_before_the_walk(repo_root):
    """The undo list made each InsertXML cost 0.2 s instead of 0.06 s. The converted book is
    brand new, so it is cleared once, before the walk - not per equation."""
    put_in = _body(repo_root, "Sh_MathML_Put_In")
    loop = next(i for i, l in enumerate(put_in) if re.match(r"\s*For i = 1 To n\b", l))
    clears = [i for i, l in enumerate(put_in) if re.match(r"\s*targetDoc\.UndoClear\b", l)]
    assert len(clears) == 1, "the undo list is not cleared exactly once"
    assert clears[0] < loop, "the undo list is cleared inside the walk, not before it"
    on_err = next(i for i, l in enumerate(put_in) if re.match(r"\s*On Error GoTo PutInFailed", l))
    assert on_err < clears[0], "UndoClear runs before the error handler is on"


def test_placeholder_places_come_from_the_text_once(repo_root):
    """The places are read once from Content.Text, checked against the placeholder before use,
    and Find is still there for one that does not match."""
    put_in = _body(repo_root, "Sh_MathML_Put_In")
    loop = next(i for i, l in enumerate(put_in) if re.match(r"\s*For i = 1 To n\b", l))
    nxt = next(i for i, l in enumerate(put_in) if re.match(r"\s*Next i\b", l))
    reads = [i for i, l in enumerate(put_in)
             if re.search(r"Sh_MathML_Placeholder_Positions\(targetDoc\.Content\.Text, n\)", l)]
    assert len(reads) == 1 and reads[0] < loop, "the places are not read once, before the walk"
    walk = "\n".join(put_in[loop:nxt])
    assert "Content.Text" not in walk, "the book's text is read per equation"
    assert re.search(r"If r\.Text <> ph Then Set r = Nothing", walk), "a place is used without checking it"
    assert re.search(r"Set r = Sh_Find_Text_From\(targetDoc, ph, lastEnd\)", walk), "no Find for a place that does not match"
    assert re.search(r"shift = shift \+ grown", walk), "the places after an insertion are not moved by it"
    assert re.search(r"shift = r\.Start - placeAt\(i\)", walk), "a Find does not correct the places after it"

    pos = "\n".join(_body(repo_root, "Sh_MathML_Placeholder_Positions"))
    assert "Chr$(7)" in pos, "the table cell end (CR and BEL) is not allowed for"
    assert "Find" not in pos and "Range" not in pos, "Sh_MathML_Placeholder_Positions touches the document"


# --- Equation pictures keep their shape (Jerry, 10/18/2026) ----------------------------------

def test_every_equation_picture_has_its_aspect_ratio_locked(repo_root):
    """"When equations are presented as pictures, the pictures must have the aspect ratio
    locked." Every AddPicture in the math helpers is followed straight away by
    LockAspectRatio = msoTrue on the picture it returned."""
    found = 0
    for name in HELPERS + (CONVERTER,):
        body = _body(repo_root, name)
        for i, line in enumerate(body):
            if line.lstrip().startswith("'") or "AddPicture" not in line:
                continue
            found += 1
            m = re.match(r"\s*Set (\w+) = .*\.AddPicture\(", line)
            assert m, f"{name}: the AddPicture result is not kept, so it cannot be locked"
            nxt = next(l for l in body[i + 1:] if l.strip() and not l.lstrip().startswith("'"))
            assert re.match(rf"\s*{m.group(1)}\.LockAspectRatio = msoTrue\b", nxt), \
                f"{name}: the picture is not aspect-locked straight after AddPicture"
    assert found >= 1, "no AddPicture found in the math helpers - the test is looking in the wrong place"


def test_the_converters_picture_loop_locks_every_pictures_aspect_ratio(repo_root):
    """Jerry, 10/18/2026: the ordinary pictures of the book - the <img> pictures, not only the
    EQ equation pictures - must have their aspect ratio locked too. The converter's picture loop
    over finalDoc.InlineShapes sets LockAspectRatio = msoTrue on every picture it walks, linked
    or not, so it sits outside the LinkFormat test."""
    body = _body(repo_root, CONVERTER)
    code = [l for l in body if l.strip() and not l.lstrip().startswith("'")]
    start = next((i for i, l in enumerate(code)
                  if re.match(r"\s*For Each (\w+) In finalDoc\.InlineShapes\b", l)), None)
    assert start is not None, "the converter's picture loop over finalDoc.InlineShapes is gone"
    var = re.match(r"\s*For Each (\w+) In", code[start]).group(1)
    end = next(i for i in range(start + 1, len(code)) if re.match(rf"\s*Next {var}\b", code[i]))
    loop = code[start + 1:end]
    depth = 0
    locked = False
    for l in loop:
        s = l.strip()
        if re.match(rf"{var}\.LockAspectRatio = msoTrue\b", s) and depth == 0:
            locked = True
        if re.match(r"If\b.*\bThen$", s):
            depth += 1
        elif s == "End If":
            depth -= 1
    assert locked, "the picture loop does not lock every picture's aspect ratio (outside any If)"


# --- The MathML text and the braille cleanup park are gone (Jerry, 10/20/2026) ---------------

REMOVED = ("Sh_MathML_Put_Text", "Sh_MathML_One_Line", "Sh_MathML_Unprefixed",
           "Dx_Math_Park_Mark", "Dx_Math_Park_Var", "Dx_Math_Mark_Number", "Dx_Math_Element_Length",
           "Dx_Math_Park", "Dx_Math_Restore", "Dx_Math_Put_Back", "Dx_Math_Prune",
           "Dx_Math_Before_Save", "Dx_Math_On_Open", "Dx_Math_Marks_Waiting", "Dx_Math_Depth")


def _code_only(line):
    """The line without a trailing comment - a ' outside a string starts one."""
    inside = False
    for i, ch in enumerate(line):
        if ch == '"':
            inside = not inside
        elif ch == "'" and not inside:
            return line[:i]
    return line


def test_mathml_text_and_braille_park_are_gone(repo_root):
    """Braille gets Word equations (OMML) now, so the MathML-as-text path and the park that kept
    braille cleanup off it were removed completely. Nothing in src/ may define or call any of
    it, and no cleanup may look for the park's VTMATH marks. Comments are allowed to name them -
    the module's dated changelog is the history."""
    src = repo_root / "src"
    files = [p for p in src.rglob("*") if p.suffix.lower() in (".bas", ".cls", ".frm")]
    assert files, "no VBA found under src/"
    hits = []
    for p in files:
        for n, line in enumerate(p.read_bytes().decode("latin-1").splitlines(), 1):
            code = _code_only(line)
            for name in REMOVED + ("VTMATH",):
                if re.search(rf"\b{name}\b", code):
                    hits.append(f"{p.relative_to(repo_root)}:{n}: {name}")
    assert not hits, "removed MathML text / park code is still in use:\n" + "\n".join(hits)


# --- Issue #38, 10/20/2026 ------------------------------------------------------------------

def test_brackets_are_wrapped_before_mml2omml(repo_root):
    """Word's MML2OMML.XSL took the one operator inside (-5) or (P+b) for the bracket's
    separator, and (-5) came out "(|5)": 17 of Section D's 470 equations, 62 of the Imagine
    book's 2,139, measured on VistaBuild 10/20/2026. Sh_MathML_Put_Equation must hand the
    stylesheet the MathML with each such bracket's inside wrapped - and load nothing else, so
    no route reaches the stylesheet unwrapped. What the wrap does is tested for real, Word's
    stylesheet included, in tests/vba/TestMathML.bas."""
    eq = _body(repo_root, "Sh_MathML_Put_Equation")
    loads = [l for l in eq if ".loadXML(" in l]
    assert loads, "Sh_MathML_Put_Equation loads no MathML"
    assert all("loadXML(Sh_MathML_Wrap_Fences(mathml))" in l for l in loads), (
        "Sh_MathML_Put_Equation loads the MathML without Sh_MathML_Wrap_Fences:\n" + "\n".join(loads))
    assert not any(re.search(r"\.load\(", l) for l in eq)

    # The wrap is on the copy for the stylesheet only: the parked fragments are not touched.
    put_in = "\n".join(_body(repo_root, "Sh_MathML_Put_In"))
    assert "Sh_MathML_Wrap_Fences" not in put_in
    assert not re.search(r"fragments\([^)]*\)\s*=", put_in), "Sh_MathML_Put_In rewrites a parked equation"

    walk = "\n".join(_body(repo_root, "Sh_MathML_Wrap_Fences_In"))
    # A list keeps its separators: comma, semicolon, bar.
    for sep in ('sep <> ","', 'sep <> ";"', 'sep <> "|"'):
        assert sep in walk, f"the walk wraps a list too ({sep} missing)"
    # The fence pairs are MML2OMML.XSL's own, in its order.
    assert 'openChars = "([{<"' in walk and 'closeChars = ")]}>"' in walk
