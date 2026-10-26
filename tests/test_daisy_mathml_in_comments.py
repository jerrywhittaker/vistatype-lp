"""DAISY/NIMAS conversion: equations hidden in XML comments, after the book's picture (issue #40).

Jerry, 10/25/2026: a DAISY book made from an EPUB (Bookshare's "Basic Math & Pre-Algebra For
Dummies") writes each of its 1,861 equations as the book's picture with the MathML after it in a
comment,

    <img src="images/eq01036.png" alt="5 times 2" class="center" ... id="img_00001" />
    <!--<m:math xmlns:m="http://www.w3.org/1998/Math/MathML"><m:mrow>...</m:mrow></m:math>-->

`Sh_MathML_To_Placeholders` parked the math and left `<!--VTEQ000001-->`; Word's HTML import
drops comments, so not one equation went in with any choice on Equations (398), while the bar
said 1,861 and 313 told the user to search for VTEQ. His rules:

- The hidden equations are wanted. MathType, then a Word equation, then the picture, as before.
- Where an equation goes in, the book's picture beside it goes. It stays only where the equation
  fails.

So `Sh_MathML_Uncomment` takes each comment that holds one math element and nothing else away,
and when the book's picture is straight before it, takes the picture out and makes its src the
math's altimg - the converter's own fallback then puts the picture in only where the equation
cannot go in. On the import file only.

Also: four equations declare `xmlns1:m="..."`, which MSXML refused ("Reference to undeclared
namespace prefix: 'xmlns1'"); `Sh_MathML_With_Namespace` takes it off.

These are checks on the source, with the two new patterns lifted out of it and run here. The
helpers are tested for real in Word by tests/vba/TestMathMLComments.bas.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

CONVERTER = "Sh_Convert_XML_File_To_Word_Document"
HELPER = "Sh_MathML_Uncomment"

VBA_STRING = re.compile(r'"((?:[^"]|"")*)"')

PAIR = ('<p>FIGURE 1-12</p>\n                      '
        '<img src="images/eq01036.png" alt="5 times 2" class="center" data-ocr-confidence="1.0" '
        'data-math-confidence="0.9999982" id="img_00001" />\n                      '
        '<!--<m:math xmlns1:m="http://www.w3.org/1998/Math/MathML"><m:mrow><m:mn>5</m:mn>'
        '<m:mo>&#215;</m:mo><m:mn>2</m:mn></m:mrow></m:math>-->\n                      '
        '<span class="text" id="span_00015">next</span>')


def _procs(repo_root):
    return procedures(module_text(repo_root))


def _index(lines, text):
    hits = [i for i, l in enumerate(lines) if text in l]
    assert len(hits) == 1, f"expected one line holding {text!r}, found {len(hits)}"
    return hits[0]


def _pattern(lines, lead):
    """The VBA string literal assigned on the one line starting with lead, "" read as "."""
    i = _index(lines, lead)
    m = VBA_STRING.search(lines[i][lines[i].index(lead):])
    assert m, lines[i]
    return m.group(1).replace('""', '"')


def test_the_import_text_is_uncommented_first(repo_root):
    """The equations are counted and parked from the uncommented text, and the .txt and .html
    beside the book are written before it, from the book as it is."""
    lines = _procs(repo_root)[CONVERTER]
    un = _index(lines, f"importText = {HELPER}(fileContent, mathPrefix)")
    frag = _index(lines, "mathFragments = Sh_MathML_Fragments(importText, mathPrefix)")
    park = _index(lines, "Sh_MathML_To_Placeholders(importText, mathPrefix)")
    assert un < frag and un < park
    assert not any("Sh_MathML_Fragments(fileContent" in l for l in lines)
    assert not any("Sh_MathML_To_Placeholders(fileContent" in l for l in lines)
    # the files beside the book, and the empty page number pass, come first and are untouched
    assert _index(lines, "outStream.SaveToFile txtCopy, 2") < un
    assert _index(lines, "fileContent = Sh_Tag_Empty_Pagenums(fileContent)") < un


def test_the_comment_pattern_on_the_measured_shapes(repo_root):
    pat = re.compile(_pattern(_procs(repo_root)[HELPER], "re.Pattern ="))

    m = pat.search(PAIR)
    assert m.group(1).startswith('<img src="images/eq01036.png"') and m.group(1).endswith("/>")
    assert m.group(2) == "\n                      "
    assert m.group(3).startswith("<m:math xmlns1:m=") and m.group(3).endswith("</m:math>")

    # a ">" inside a quoted value does not end the picture's tag
    m = pat.search('<img alt="x > 2" src="a.png"/> <!--<m:math/>-->')
    assert m.group(1) == '<img alt="x > 2" src="a.png"/>'

    # text between the picture and the comment: the comment alone
    m = pat.search('<img src="a.png"/> and <!--<m:math/>-->')
    assert m.group(1) is None and m.group(0) == "<!--<m:math/>-->"

    # a comment ends at its first "-->"
    ms = list(pat.finditer("<!-- a note --><p>x</p><!--<m:math/>-->"))
    assert [x.group(3) for x in ms] == [" a note ", "<m:math/>"]

    # an <image> or <imgx> is not a picture
    assert pat.search("<imgx src='a.png'/><!--<m:math/>-->").group(1) is None


def test_only_a_comment_holding_one_math_element_is_brought_out(repo_root):
    lines = _procs(repo_root)[HELPER]
    assert _pattern(lines, "one.Pattern =").startswith("^\\s*(")
    assert any('& Sh_MathML_Pattern(prefix) & ")\\s*$"' in l for l in lines)
    # two math elements in one comment read as one by that pattern, and are left alone
    assert any("If InStr(2, mathml, tagName) > 0 Then mathml = \"\"" in l for l in lines)
    # every other comment is taken out (the picture before it stays), so no math is left inside
    # one to be counted and parked and then dropped by the import with its placeholder (review)
    assert any("parts(2 * i + 1) = imgTag & gap" == l.strip() for l in lines)
    assert not any("m.Value" in l for l in lines)


def test_no_comment_is_left_in_the_import_text(repo_root):
    """Run the helper's own comment pattern as the helper uses it: whatever a comment held, the
    text handed back has no "<!--" in it, so Sh_MathML_Fragments counts only live math."""
    pat = re.compile(_pattern(_procs(repo_root)[HELPER], "re.Pattern ="))
    book = ('<img src="a.png"/> <!-- a note --><p>x</p>'
            "<!--<m:math><m:mn>1</m:mn></m:math><m:math><m:mn>2</m:mn></m:math>-->"
            "<!-- see <m:math><m:mn>3</m:mn></m:math> -->")
    left = pat.sub(lambda m: (m.group(1) or "") + (m.group(2) or ""), book)
    assert "<!--" not in left and "<m:math" not in left
    assert left == '<img src="a.png"/> <p>x</p>'


def test_the_picture_becomes_the_altimg_only_when_there_is_none(repo_root):
    lines = _procs(repo_root)[HELPER]
    body = "\n".join(lines)
    assert 'If imgTag <> "" And Sh_MathML_Altimg(mathml) = "" Then' in body
    assert 'src = Sh_Xml_Attribute(imgTag, "src")' in body
    assert re.search(r'InStr\(1, src, "://"\) > 0 Or LCase\$\(Left\$\(Trim\$\(src\), 5\)\) = "data:"', body)
    # with an altimg the picture goes; without one it stays where it was
    assert 'parts(2 * i + 1) = tagName & " altimg=" & q & src & q & Mid$(mathml, Len(tagName) + 1)' in body
    assert "parts(2 * i + 1) = imgTag & gap & mathml" in body


def test_a_stray_xmlns_digits_declaration_is_taken_off(repo_root):
    lines = _procs(repo_root)["Sh_MathML_With_Namespace"]
    pat = re.compile(_pattern(lines, "re.Pattern ="))
    ns = "http://www.w3.org/1998/Math/MathML"
    assert pat.sub("", f'<m:math xmlns1:m="{ns}">') == "<m:math>"
    assert pat.sub("", f"<m:math altimg='a.png' xmlns12:m='{ns}'>") == "<m:math altimg='a.png'>"
    # a real declaration stays
    for tag in (f'<m:math xmlns:m="{ns}">', f'<math xmlns="{ns}">', f'<m:math xmlnsx:m="{ns}">'):
        assert pat.sub("", tag) == tag
    # on the opening tag only, before the real declaration is looked for
    body = "\n".join(lines)
    assert "re.Replace(Left$(fragment, gt), \"\") & Mid$(fragment, gt + 1)" in body
    assert body.index("re.Replace(") < body.index("If Sh_Xml_Attribute(Left$(fragment, gt), declName)")


def test_the_helper_takes_byval_and_says_nothing(repo_root):
    head = _procs(repo_root)[HELPER][0]
    assert "(ByVal xml As String, ByVal prefix As String) As String" in head
    assert not any("MsgBox" in l or "Sh_Say" in l for l in _procs(repo_root)[HELPER])


def test_the_vba_suite_is_registered(repo_root):
    vba = (repo_root / "tests" / "vba" / "TestMathMLComments.bas").read_bytes().decode("latin-1")
    assert f"'@uses {HELPER}" in vba
    assert "'@uses Sh_MathML_With_Namespace" in vba
    runner = (repo_root / "tests" / "vba" / "VtRunTests.bas").read_bytes().decode("latin-1")
    assert "TestMathMLComments_Suite()" in runner
