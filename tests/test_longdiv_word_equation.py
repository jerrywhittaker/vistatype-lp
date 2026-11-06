"""Issue #50: long division as a Word equation keeps its dividend (Jerry, 11/5/2026).

Word's MML2OMML.XSL writes nothing for <menclose notation="longdiv">, so on the "Word equations"
route (and the fallback from MathType) 5 with 2 under the bracket came out as "5".
Sh_MathML_Longdiv_For_Word reshapes each longdiv into ")" and a top-bar menclose, which the
stylesheet carries, before Sh_MathML_Put_Equation hands the MathML over. The behavior itself is
tested against Word's own stylesheet in tests/vba/TestMathML.bas.
"""
import re
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src" / "vba" / "LPandBrlMacros.bas"


def _func(name):
    text = SRC.read_text(encoding="latin-1")
    start = text.index("Private Function %s(" % name)
    end = text.index("End Function", start)
    return text[start:end]


def test_word_equation_route_reshapes_long_division():
    body = _func("Sh_MathML_Put_Equation")
    assert re.search(r"src\.loadXML\(Sh_MathML_Longdiv_For_Word\(Sh_MathML_Wrap_Fences\(mathml\)\)\)", body)


def test_reshape_uses_a_top_bar_and_a_parenthesis():
    body = _func("Sh_MathML_Longdiv_For_Word")
    assert 'setAttribute "notation", "top"' in body
    assert 'mo.Text = ")"' in body
    assert '= "longdiv"' in body


def test_reshape_leaves_other_mathml_alone():
    body = _func("Sh_MathML_Longdiv_For_Word")
    # the fast way out, and the error path, both hand back what came in
    assert re.search(r'InStr\(1, mathml, "longdiv", vbTextCompare\) = 0 Then Exit Function', body)
    assert re.search(r"Unchanged:\s*\n\s*Err\.Clear\s*\n\s*Sh_MathML_Longdiv_For_Word = mathml", body)
