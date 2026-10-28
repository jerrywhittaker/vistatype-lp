"""DAISY/NIMAS conversion: tables of equations taken apart before the import (issue #43).

Jerry, 10/27/2026: Bookshare's "Basic Math & Pre-Algebra For Dummies" sets its equations two to
four side by side in one-row tables (`<table class="Dummies-Mini-Table">`, one `<td>` per
equation, each a picture with its MathML in a comment - issue #40). For large print MathType
enlarges the equations, so side by side they no longer fit. His rules:

- Any table that holds an equation is taken apart entirely, words included. Every cell's
  content goes on its own line, in reading order - row by row, left to right - so words to the
  right of an equation follow below it. The table is then gone.
- Done in the .xml, before the import.

So `Sh_MathML_Untable` runs on the import text straight after `Sh_MathML_Uncomment` (which is
when an equation hidden in a comment becomes live math that can be seen) and before the
equations are counted and parked. Empty cells give nothing; a caption comes first; a table
with no equation is not touched; large print and braille alike.

These are checks on the source, with its patterns lifted out and run here. The helper itself is
tested for real in Word by tests/vba/TestMathMLTables.bas.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

CONVERTER = "Sh_Convert_XML_File_To_Word_Document"
HELPER = "Sh_MathML_Untable"
CELL = "Sh_MathML_Cell_Paragraph"
LEFTOVER = "Sh_MathML_Table_Leftover"
TYPED = "Sh_MathML_Typed_Equation"

VBA_STRING = re.compile(r'"((?:[^"]|"")*)"')


def _procs(repo_root):
    return procedures(module_text(repo_root))


def _index(lines, text):
    hits = [i for i, l in enumerate(lines) if text in l]
    assert len(hits) == 1, f"expected one line holding {text!r}, found {len(hits)}"
    return hits[0]


def _patterns(lines, lead="re.Pattern ="):
    """Every VBA string literal assigned on a line starting with lead, in order, "" read as "."""
    found = []
    for l in lines:
        s = l.strip()
        if s.startswith(lead):
            m = VBA_STRING.search(s)
            assert m, l
            found.append(m.group(1).replace('""', '"'))
    return found


def test_tables_are_taken_apart_after_the_uncommenting(repo_root):
    """After Sh_MathML_Uncomment, so the equations hidden in comments are seen; before the
    equations are counted and parked; on the import text only, never the .txt and .html."""
    lines = _procs(repo_root)[CONVERTER]
    un = _index(lines, "importText = Sh_MathML_Uncomment(fileContent, mathPrefix)")
    ut = _index(lines, f"importText = {HELPER}(importText, mathPrefix)")
    frag = _index(lines, "mathFragments = Sh_MathML_Fragments(importText, mathPrefix)")
    park = _index(lines, "Sh_MathML_To_Placeholders(importText, mathPrefix)")
    assert un + 1 == ut and ut < frag and ut < park
    assert not any(f"{HELPER}(fileContent" in l for l in lines)
    assert _index(lines, "outStream.SaveToFile txtCopy, 2") < ut
    # large print and braille alike: nothing between the uncommenting and the untabling asks which
    assert not any("Sh_GP_String_2" in l for l in lines[un:ut + 1])


def test_the_table_tag_pattern(repo_root):
    pat = re.compile(_patterns(_procs(repo_root)[HELPER])[0], re.I)
    for tag in ("<table>", '<table class="Dummies-Mini-Table" id="table_0000000">', "</table>",
                "<tr>", '<tr id="tr_1">', "</tr>", "<td>", '<td class="Mini-TB-Body" rowspan="1">',
                "</td>", "<th>", "</th>", "<TD>", "<td/>", "<caption>", "</caption>"):
        m = pat.fullmatch(tag)
        assert m, tag
    # a ">" inside a quoted value does not end the tag
    assert pat.match('<td title="a > b">x').group(0) == '<td title="a > b">'
    # not table tags: thead, tbody, tfoot, MathML's own mtd/mtr/mtable, track, a longer name
    for text in ("<thead>", "<tbody>", "<tfoot>", "<m:mtd>", "<mtd>", "<m:mtr>", "<mtr>",
                 "<m:mtable>", "<mtable>", "<track>", "<tdx>", "<tables>"):
        assert not pat.search(text), text


def test_math_is_seen_only_with_the_books_prefix(repo_root):
    lines = _procs(repo_root)[HELPER]
    assert any('mathRe.Pattern = "<" & prefix & "math(?=[\\s>/])"' in l for l in lines)
    pat = re.compile("<" + "m:" + r"math(?=[\s>/])")
    assert pat.search("<m:math>") and pat.search('<m:math altimg="a.png">') and pat.search("<m:math/>")
    assert not pat.search("<m:mathx>") and not pat.search("<mml:math>")


def test_an_empty_cell_gives_nothing_and_a_picture_or_math_is_never_empty(repo_root):
    trim, keep, strip, block = _patterns(_procs(repo_root)[CELL])
    keep, strip, block = re.compile(keep, re.I), re.compile(strip, re.I), re.compile(block, re.I)

    def empty(cell):
        cell = re.sub(trim, "", cell)
        if keep.search(cell):
            return False
        return strip.sub("", cell).replace("\xa0", "") == ""

    for cell in ("", " ", "\n  \t", '<p id="x" class="Mini-TB-Body" xml:space="preserve">  </p>',
                 "<p>&#160;</p>", "<p>&nbsp;</p>", "<p>&#xA0;</p>", "<p>\xa0</p>", "<br/>"):
        assert empty(cell), repr(cell)
    for cell in ("x", "<p>Threes:</p>", '<img src="a.png" />', "<m:math><m:mn>1</m:mn></m:math>",
                 "<math><mn>1</mn></math>", "<mml:math/>", "<p>$pg5</p>"):
        assert not empty(cell), repr(cell)

    # already a paragraph: as it is; anything else: wrapped in a <p>
    for cell in ('<p class="Mini-TB-Body">x</p>', "<h2>x</h2>", "<list><li>x</li></list>",
                 "<table><tr><td>x</td></tr></table>", '<div class="x">y</div>'):
        assert block.search(cell), cell
    for cell in ("x", "<pagenum>5</pagenum>", "<prodnote>x</prodnote>", "<span>x</span>",
                 '<img src="a.png"/>', "<m:math><m:mn>1</m:mn></m:math>"):
        assert not block.search(cell), cell


def test_outside_the_cells_only_the_structure_goes(repo_root):
    """thead, tbody, tfoot, colgroup and col go with the table; anything else between rows -
    a page number - is kept."""
    strip, trim = _patterns(_procs(repo_root)[LEFTOVER])
    strip = re.compile(strip, re.I)

    def left(text):
        return re.sub(trim, "", strip.sub("", text))

    assert left('\n  <tbody id="tbody_0000000">\n   ') == ""
    assert left("</thead><TBODY></tbody><tfoot/><colgroup><col width='2'/></colgroup>") == ""
    assert left("\n <p>$pg5</p>\n") == "<p>$pg5</p>"
    assert left("<p>x</p>") == "<p>x</p>"


def test_a_table_is_kept_whole_unless_it_holds_an_equation(repo_root):
    body = "\n".join(_procs(repo_root)[HELPER])
    assert "takeApart = mathRe.Test(fWhole(d))" in body
    # Jerry, 10/28/2026: or one row, every cell that is not empty an equation, typed or not
    assert "If fRows(d) = 1 And fCells(d) > 0 And fAllEq(d) Then takeApart = True" in body
    assert 'If tagName = "tr" Then fRows(d) = fRows(d) + 1' in body
    assert "If Not (mathRe.Test(fCell(d)) Or Sh_MathML_Typed_Equation(fCell(d))) Then fAllEq(d) = False" in body
    assert "rep = fCap(d) & fOut(d)" in body            # the caption first
    assert "rep = fWhole(d)" in body
    # a table that never closes leaves the text as it was
    assert "If d <> 0 Then Exit Function" in body
    assert body.index(f"{HELPER} = xml") < body.index("If d <> 0 Then Exit Function")


def test_the_helpers_take_byval_and_say_nothing(repo_root):
    procs = _procs(repo_root)
    assert "(ByVal xml As String, ByVal prefix As String) As String" in procs[HELPER][0]
    assert "(ByVal cellText As String) As String" in procs[CELL][0]
    assert "(ByVal between As String) As String" in procs[LEFTOVER][0]
    assert "(ByVal cellText As String) As Boolean" in procs[TYPED][0]
    for name in (HELPER, CELL, LEFTOVER, TYPED):
        assert not any("MsgBox" in l or "Sh_Say" in l or "Sh_Ask" in l for l in procs[name])


def test_the_vba_suite_is_registered(repo_root):
    vba = (repo_root / "tests" / "vba" / "TestMathMLTables.bas").read_bytes().decode("latin-1")
    assert f"'@uses {HELPER}" in vba
    runner = (repo_root / "tests" / "vba" / "VtRunTests.bas").read_bytes().decode("latin-1")
    assert "TestMathMLTables_Suite()" in runner


def test_a_typed_equation_is_an_equals_sign_outside_the_tags(repo_root):
    """Jerry, 10/28/2026: Basic Math's one-row tables of typed equations come apart too."""
    (tags,) = _patterns(_procs(repo_root)[TYPED])
    typed = lambda c: "=" in re.sub(tags, "", c)
    assert typed('<p class="Mini-TB-Body">25% = 0.25</p>')
    assert typed('<em id="em_1">A</em><span class="text"> = (2, 3) </span>')
    assert not typed('<p class="a" id="b">words</p>')
