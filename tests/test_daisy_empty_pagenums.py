"""DAISY/NIMAS conversion: page numbers written as EMPTY <pagenum> elements get $pg tags too.

Jerry, 10/25/2026: a DAISY book made from an EPUB (Bookshare's "Basic Math & Pre-Algebra For
Dummies") came out of Convert DAISY/NIMAS to Word with not one $pg tag. Every one of its 345
page numbers is an empty element with the number only in the id,

    <pagenum epub:type="pagebreak" page="normal" id="pagenum_0000163" smilref="...#pagenum_0000163" />

and the converter's two patterns read the number from the text, >163</pagenum>. His rules:

- The two text patterns stay exactly as they are.
- A second pass, for DAISY and NIMAS alike, tags only the EMPTY ones - self-closing, or an
  opening and closing tag with nothing but white space between.
- The number is the run of digits at the end of the id, leading zeros dropped: $pg163. No id,
  no digits at the end of it, or digits that are all zeros: $pg? for the user to fix.

These are checks on the source, plus the helper's own two patterns lifted out of the source and
run here, with VBScript's replacement rules, on the shapes measured in that book. The helper is
tested for real in Word by tests/vba/TestEmptyPagenums.bas.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

CONVERTER = "Sh_Convert_XML_File_To_Word_Document"
HELPER = "Sh_Tag_Empty_Pagenums"

DAISY_PATTERN = 'reTag.Pattern = ">([0-9A-z]{1,})</pagenum>"'
DAISY_REPLACE = 'reTag.Replace(fileContent, "<p>$pg$1<p></p></pagenum><p></p>")'
NIMAS_PATTERN = 'reTag.Pattern = ">([0-9A-z-]{1,}</pagenum>)"'
NIMAS_REPLACE = 'reTag.Replace(fileContent, "<p>$pg$1<p></p>")'

VBA_STRING = re.compile(r'"((?:[^"]|"")*)"')


def _procs(repo_root):
    return procedures(module_text(repo_root))


def _index(lines, text):
    hits = [i for i, l in enumerate(lines) if text in l]
    assert len(hits) == 1, f"expected one line holding {text!r}, found {len(hits)}"
    return hits[0]


def _literal(line):
    """The first VBA string literal on a line, with "" read as one quotation mark."""
    m = VBA_STRING.search(line)
    assert m, f"no string on {line!r}"
    return m.group(1).replace('""', '"')


def _vbscript_replace(pattern, replacement, text):
    """RegExp.Replace with Global = True, as VBScript reads the replacement: $1..$9 are groups,
    $$ is one dollar sign, and a $ before anything else is an ordinary character."""
    def expand(m):
        out = []
        i = 0
        while i < len(replacement):
            ch = replacement[i]
            nxt = replacement[i + 1] if i + 1 < len(replacement) else ""
            if ch == "$" and nxt.isdigit():
                out.append(m.group(int(nxt)) or "")
                i += 2
            elif ch == "$" and nxt == "$":
                out.append("$")
                i += 2
            else:
                out.append(ch)
                i += 1
        return "".join(out)
    return re.sub(pattern, expand, text)


def _helper_passes(repo_root):
    """The helper's (pattern, replacement) pairs, in order, read from the source."""
    lines = _procs(repo_root)[HELPER]
    end = _literal(lines[_index(lines, "Const EMPTY_END As String")])
    pats = [l for l in lines if re.search(r"\bre\.Pattern\s*=", l)]
    reps = [l for l in lines if re.search(r"\bre\.Replace\(", l)]
    assert len(pats) == 2 and len(reps) == 2, "the helper should make exactly two passes"
    passes = []
    for p, r in zip(pats, reps):
        assert "& EMPTY_END" in p, f"a pass that does not end on an empty element: {p.strip()}"
        passes.append((_literal(p) + end, _literal(r)))
    return passes


def _tag(repo_root, text):
    for pattern, replacement in _helper_passes(repo_root):
        text = _vbscript_replace(pattern, replacement, text)
    return text


def _tagged(open_tag, pg):
    return open_tag + "<p>$pg" + pg + "<p></p></pagenum><p></p>"


def test_the_text_patterns_are_unchanged(repo_root):
    body = "\n".join(_procs(repo_root)[CONVERTER])
    for text in (DAISY_PATTERN, DAISY_REPLACE, NIMAS_PATTERN, NIMAS_REPLACE):
        assert body.count(text) == 1, f"{CONVERTER} no longer has {text!r} exactly once"


def test_the_empty_pass_runs_after_them_for_daisy_and_nimas(repo_root):
    lines = _procs(repo_root)[CONVERTER]
    call = _index(lines, f"fileContent = {HELPER}(fileContent)")
    nimas = _index(lines, NIMAS_REPLACE)
    daisy = _index(lines, DAISY_REPLACE)
    assert call > nimas and call > daisy, "the empty pass must come after the text patterns"
    between = [l.strip() for l in lines[nimas + 1:call] if l.strip()]
    assert between == ["End If"], (
        f"{HELPER} must run straight after the DAISY/NIMAS If, for both: {between}")
    prod = _index(lines, "Sh_Tag_Prodnotes_As_Prodnote_Style(fileContent)")
    assert call < prod, "the $pg tags go in before the prodnotes are marked, as before"


def test_the_helper_takes_its_text_byval(repo_root):
    first = _procs(repo_root)[HELPER][0]
    assert re.search(rf"Function {HELPER}\(ByVal \w+ As String\) As String", first), first


def test_numbers_from_the_id(repo_root):
    t = ('<pagenum epub:type="pagebreak" page="normal" id="pagenum_0000163" '
         'smilref="Book00011.smil#pagenum_0000163"')
    assert _tag(repo_root, t + " />") == _tagged(t, "163")
    t = '<pagenum epub:type="pagebreak" id="pagenum_0000012" page="normal"'
    assert _tag(repo_root, t + "/>") == _tagged(t, "12")
    for padded, n in (("00007", "7"), ("000070", "70"), ("0000700", "700"), ("ch3_page_0045", "45")):
        t = f'<pagenum page="normal" id="{padded}"'
        assert _tag(repo_root, t + " />") == _tagged(t, n), padded


def test_no_number_is_a_question_mark(repo_root):
    for t in ('<pagenum epub:type="pagebreak" id="p" page="normal" smilref="Book00001.smil#p"',
              '<pagenum page="normal" id="pagenum_00000"',
              '<pagenum page="normal" id="pagenum_0000000"',
              '<pagenum id="page_10_000"',
              '<pagenum page="normal"',
              '<pagenum smilref="a.smil#pagenum_0000009"',
              "<pagenum"):
        assert _tag(repo_root, t + " />") == _tagged(t, "?"), t


def test_open_and_close_with_only_white_space(repo_root):
    t = '<pagenum id="pagenum_0000042" page="normal"'
    assert _tag(repo_root, t + "></pagenum>") == _tagged(t, "42")
    assert _tag(repo_root, t + ">  \r\n\t</pagenum>") == _tagged(t, "42")


def test_mid_sentence_and_between_list_items(repo_root):
    t = '<pagenum page="normal" id="pagenum_0000099"'
    assert (_tag(repo_root, "<p>Some <strong>bold</span>" + t + " /> text</strong> here.</p>")
            == "<p>Some <strong>bold</span>" + _tagged(t, "99") + " text</strong> here.</p>")
    assert (_tag(repo_root, '<li>x</li><pagenum id="pagenum_01" /><li>y</li>')
            == "<li>x</li>" + _tagged('<pagenum id="pagenum_01"', "1") + "<li>y</li>")


def test_page_numbers_with_text_are_left_alone(repo_root):
    for t in ('<pagenum id="p12" page="normal">12</pagenum>',
              '<pagenum id="p13">13-14</pagenum>',
              # as the converter's DAISY and NIMAS patterns leave them
              '<pagenum id="p12" page="normal"<p>$pg12<p></p></pagenum><p></p>',
              '<pagenum id="p12" page="normal"<p>$pg12-13</pagenum><p></p>',
              '<p>Cost: $5 and $10.</p><pagenumber id="p1" /><prodnote id="n2" />'):
        assert _tag(repo_root, t) == t, t


def test_the_text_pattern_then_the_empty_pass(repo_root):
    """A book holding both kinds: each page number tagged once, by the pass meant for it."""
    book = ('<dtbook><p>a</p><pagenum id="p1" page="normal">1</pagenum><p>b</p>'
            '<pagenum page="normal" id="pagenum_0000002" /><p>c</p><pagenum id="p" /></dtbook>')
    after_text = _vbscript_replace(r">([0-9A-z]{1,})</pagenum>",
                                   "<p>$pg$1<p></p></pagenum><p></p>", book)
    out = _tag(repo_root, after_text)
    assert re.findall(r"\$pg([0-9]+|\?)", out) == ["1", "2", "?"]


def test_the_vba_suite_is_registered(repo_root):
    vba = (repo_root / "tests" / "vba" / "TestEmptyPagenums.bas").read_bytes().decode("latin-1")
    assert f"'@uses {HELPER}" in vba
    runner = (repo_root / "tests" / "vba" / "VtRunTests.bas").read_bytes().decode("latin-1")
    assert "TestEmptyPagenums_Suite()" in runner
