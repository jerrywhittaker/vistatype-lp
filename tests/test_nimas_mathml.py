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
  a Word equation (or its picture if that fails) or the picture for large print, and the MathML
  as text for braille, for DBT.

Jerry, 10/20/2026: braille gets Word equations (OMML) instead, exactly as large print's "Word
equations (OMML)" choice does, and box 398 says the book holds equations before it asks.

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
           "Sh_MathML_Unprefixed", "Sh_Strip_Xml_Declaration", "Sh_OMML_Package",
           "Sh_Find_Text_From", "Sh_MathML_Put_Equation", "Sh_MathML_Put_Picture",
           "Sh_MathML_Put_Text", "Sh_MathML_Put_In",
           # the review, 10/18/2026
           "Sh_Xml_Is_Space", "Sh_Xml_Start_Tag_Length", "Sh_Xml_Attribute", "Sh_Digits_Value",
           "Sh_Chr_From_Code", "Sh_Xml_Entity_Char", "Sh_Xml_Unescape", "Sh_Url_Unescape",
           "Sh_Find_Text_Forward", "Sh_MathML_One_Line",
           # the speed fix, 10/18/2026
           "Sh_MathML_Placeholder_Positions")


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


def test_math_question_only_for_large_print_and_only_with_mathml(repo_root):
    lines = _converter(repo_root)
    found = _first(lines, r'mathPrefix = Sh_MathML_Prefix\(fileContent\)')
    has = _first(lines, r'If mathPrefix <> "#" Then')
    brl = _first(lines, r'If Sh_GP_String_2 = "BRL" Then')
    ask = _first(lines, r"DN_Equations_Form\.Show")
    assert found < has < brl < ask, "the equation question is not inside 'has MathML' and 'not braille'"
    assert any(re.match(r"\s*Else\s*$", l) for l in lines[brl:ask]), (
        "the equation question is not in the large print branch")
    assert sum("DN_Equations_Form.Show" in l for l in lines) == 1


def test_braille_gets_word_equations(repo_root):
    """Jerry, 10/20/2026: a braille conversion puts the equations in as Word equations, the same
    mode large print's "Word equations (OMML)" uses - not the MathML text, and not asked."""
    lines = _converter(repo_root)
    brl = _first(lines, r'If Sh_GP_String_2 = "BRL" Then')
    ask = _first(lines, r"DN_Equations_Form\.Show")
    branch = "\n".join(lines[brl:ask])
    assert re.search(r'mathMode = "OMML"', branch), "braille does not get Word equations"
    assert 'mathMode = "BRL"' not in "\n".join(lines), "braille still gets the MathML as text"


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
    assert re.search(r'If Sh_GP_String_3 <> "OMML" And Sh_GP_String_3 <> "PICTURE" Then', near)
    assert re.search(r'Sh_Say "[^"]+",\s*_?\s*"VistaType LP \(399\)"', near.replace("\n", " "))
    assert re.search(r"Exit Sub", near)

    code = "\n".join(form_code(repo_root, "DN_Equations_Form"))
    assert re.search(r'Sh_GP_String_3\s*=\s*"OMML"', code) and re.search(r'Sh_GP_String_3\s*=\s*"PICTURE"', code)
    frm = (repo_root / "src" / "forms" / "DN_Equations_Form.frm").read_bytes().decode("latin-1")
    assert 'Caption         =   "Equations (398)"' in frm


def test_the_equations_are_parked_and_put_back(repo_root):
    lines = _converter(repo_root)
    park = _first(lines, r"Sh_MathML_To_Placeholders\(fileContent, mathPrefix\)")
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
    for mode in ('"BRL"', '"OMML"'):
        assert f"Case {mode}" in put_in
    assert re.search(r"Sh_MathML_Put_Text\(r, Sh_MathML_Unprefixed\(", put_in), "braille is not given the MathML text"
    assert re.search(r"Sh_MathML_Put_Equation\(r, xslt, xsl,\s*_?\s*Sh_MathML_With_Namespace\(", put_in.replace("\n", " "))
    assert re.search(r"If Not done Then\s+done = Sh_MathML_Put_Picture\(r, picPath\)\s+If done Then asPicture",
                     put_in), "a Word equation that fails does not fall back to its picture"
    assert 'MML2OMML.XSL' in put_in

    text = "\n".join(procs["Sh_MathML_Put_Text"])
    assert re.search(r"r\.Text = ", text) and "InsertFile" not in text and "InsertXML" not in text

    pic = "\n".join(procs["Sh_MathML_Put_Picture"])
    assert re.search(r"AddPicture\W.*LinkToFile:=False, SaveWithDocument:=True", pic), "the picture is not embedded"

    eq = "\n".join(procs["Sh_MathML_Put_Equation"])
    assert "MSXML2.DOMDocument.6.0" in eq and "parseError" in eq and "transformNode" in eq
    assert re.search(r"r\.InsertXML Sh_OMML_Package\(", eq)

    # The count of equations that did not go in as asked reaches the closing message.
    body = "\n".join(_converter(repo_root))
    assert re.search(r'MsgBox "Conversion is complete[^\n]*& mathNote, vbInformation, "Operation Complete \(313\)"', body)


def test_braille_puts_text_not_a_new_style(repo_root):
    """Jerry, 10/18/2026: ordinary paragraph font, no style of its own and nothing hidden, until
    he has checked it in DBT."""
    text = "\n".join(_procs(repo_root)["Sh_MathML_Put_Text"])
    assert not re.search(r"\.Style\s*=|\.Hidden|\.Font\.", text)


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
                 "Sh_MathML_With_Namespace", "Sh_MathML_Unprefixed", "Sh_MathML_Placeholder_Positions"):
        assert f"'@uses {name}" in vba
    runner = (repo_root / "tests" / "vba" / "VtRunTests.bas").read_bytes().decode("latin-1")
    assert "TestMathML_Suite()" in runner


# --- The review of 10/18/2026 ---------------------------------------------------------------

def _body(repo_root, name):
    return _procs(repo_root)[name]


def test_braille_mathml_keeps_every_character(repo_root):
    """Sh_One_Log_Line's "below 32" test read AscW, which is NEGATIVE from U+8000 up, so every
    such character - both halves of a math letter like U+1D465, every Private Use character -
    became a space. The braille MathML must not go through it, and the log line must not lose
    them either."""
    text = "\n".join(_body(repo_root, "Sh_MathML_Put_Text"))
    assert "Sh_One_Log_Line" not in text, "the braille MathML still goes through the log-line cleaner"
    assert re.search(r"r\.Text = Sh_MathML_One_Line\(mathml\)", text)
    one = "\n".join(_body(repo_root, "Sh_MathML_One_Line"))
    assert "Replace(" in one and "Mid$(" not in one and "AscW" not in one, (
        "Sh_MathML_One_Line must only Replace line breaks and tabs, not walk the characters")
    for brk in ("vbCrLf", "vbCr", "vbLf", "vbTab"):
        assert brk in one

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
    for name in ("Sh_MathML_With_Namespace", "Sh_MathML_Unprefixed"):
        body = "\n".join(_body(repo_root, name))
        assert "Sh_Xml_Start_Tag_Length(" in body and 'InStr(1, ' + ('fragment' if 'Namespace' in name else 's') + ', ">")' not in body, name

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
