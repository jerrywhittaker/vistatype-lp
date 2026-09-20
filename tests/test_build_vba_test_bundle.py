"""build_vba_test_bundle.py -- gathering the VBA procedures a test exercises.

Most of the helpers worth testing are Private, so a test document that merely referenced the
built add-in could not call them at all -- and would be testing the last build rather than the
source. This lifts the source instead. If it lifts the wrong thing, the VBA tests are testing
nothing, and there is no compiler here to say so.
"""
import pytest

import build_vba_test_bundle as mod

from conftest import write_crlf

MODULE = """\
Option Explicit

Private Const LP_FONT_SANS As String = "VistaTypeLP Sans"
Private Const LP_UNUSED As String = "nobody wants this"

Public Function Sh_Top(ByVal s As String) As String
    Sh_Top = Sh_Helper(s)
End Function

Private Function Sh_Helper(ByVal s As String) As String
    ' mentions Sh_Never_Called, but only in a comment
    Sh_Helper = UCase$(s)
End Function

Private Function Sh_Never_Called(ByVal s As String) As String
    Sh_Never_Called = s
End Function

Function Lp_Font_Name() As String
    Lp_Font_Name = LP_FONT_SANS
End Function

Private Sub Sh_Deep()
    Sh_Never_Called "x"
End Sub
"""


def repo(tmp_path, uses=("Sh_Top", "Lp_Font_Name"), module=MODULE):
    write_crlf(tmp_path / "src" / "vba" / "LPandBrlMacros.bas", module)
    write_crlf(tmp_path / "tests" / "vba" / "TestThing.bas",
               'Attribute VB_Name = "TestThing"\n'
               + "".join("'@uses %s\n" % u for u in uses))
    return tmp_path


def bundle(tmp_path, **kw):
    out, names, decls = mod.build(root=repo(tmp_path, **kw),
                                  out_path=tmp_path / "out.bas")
    return out.read_bytes().decode("cp1252"), names, decls


# ---------------------------------------------------------------- the pieces

def test_strip_comment_keeps_an_apostrophe_inside_a_string():
    assert "Sh_A" in mod.strip_comment('Call Sh_A   \' Sh_B')
    assert "Sh_B" not in mod.strip_comment('Call Sh_A   \' Sh_B')
    assert "Sh_C" not in mod.strip_comment('x = "don\'t call Sh_C"')


def test_make_public_promotes_only_a_private_procedure():
    assert mod.make_public("Private Function Sh_X() As String").startswith("Public Function")
    assert mod.make_public("Private Sub Sh_X()").startswith("Public Sub")
    assert mod.make_public("Public Function Sh_X()") == "Public Function Sh_X()"
    assert mod.make_public("Function Sh_X()") == "Function Sh_X()"


def test_make_public_leaves_a_private_variable_alone():
    """Only procedures. A Private Const stays private and still resolves in the module."""
    line = "Private Const LP_FONT_SANS As String = \"x\""
    assert mod.make_public(line) == line


# ---------------------------------------------------------------- what gets gathered

def test_the_named_procedures_are_gathered(tmp_path):
    text, names, _ = bundle(tmp_path)
    assert "Sh_Top" in names
    assert "Lp_Font_Name" in names
    assert "Public Function Sh_Top" in text


def test_what_they_call_is_gathered_too(tmp_path):
    """Otherwise the module does not compile and the run hangs on a dialog nobody can answer."""
    _, names, _ = bundle(tmp_path)
    assert "Sh_Helper" in names


def test_a_private_helper_becomes_callable(tmp_path):
    text, _, _ = bundle(tmp_path)
    assert "Public Function Sh_Helper" in text
    assert "Private Function Sh_Helper" not in text


def test_a_name_mentioned_only_in_a_comment_is_not_gathered(tmp_path):
    _, names, _ = bundle(tmp_path)
    assert "Sh_Never_Called" not in names


def test_nothing_unrelated_is_dragged_in(tmp_path):
    text, names, _ = bundle(tmp_path)
    assert "Sh_Deep" not in names
    assert "Sh_Deep" not in text


def test_a_module_level_constant_a_gathered_procedure_uses_comes_with_it(tmp_path):
    """Lp_Indent_Factor_For_Font reads LP_FONT_SANS; without it the module will not compile."""
    text, _, decls = bundle(tmp_path)
    assert "LP_FONT_SANS" in decls
    assert 'Private Const LP_FONT_SANS As String = "VistaTypeLP Sans"' in text


def test_a_constant_nobody_uses_is_left_behind(tmp_path):
    _, _, decls = bundle(tmp_path)
    assert "LP_UNUSED" not in decls


def test_a_helper_without_a_project_prefix_is_still_gathered(tmp_path):
    """Everything here is namespaced Lp_/Dx_/Sh_/MS_/DN_, but the gather must not DEPEND on
    that -- one that is not would be missed, and a missing helper is a hang, not an error."""
    module = MODULE.replace("    Sh_Helper = UCase$(s)", "    Sh_Helper = PlainHelper(s)")
    module += "\nPrivate Function PlainHelper(ByVal s As String) As String\n" \
              "    PlainHelper = s\n" \
              "End Function\n"
    text, names, _ = bundle(tmp_path, module=module)
    assert "PlainHelper" in names
    assert "Public Function PlainHelper" in text


def test_a_procedure_reached_two_steps_away_is_gathered(tmp_path):
    module = MODULE.replace("    Sh_Helper = UCase$(s)",
                            "    Sh_Helper = Sh_Never_Called(s)")
    _, names, _ = bundle(tmp_path, module=module)
    assert "Sh_Never_Called" in names


def test_the_output_is_an_importable_module(tmp_path):
    text, _, _ = bundle(tmp_path)
    assert text.startswith('Attribute VB_Name = "VtFunctionsUnderTest"')
    assert "Option Explicit" in text


def test_the_output_is_crlf(tmp_path):
    """Word imports it; the rest of the VBA source here is CRLF."""
    out, _, _ = mod.build(root=repo(tmp_path), out_path=tmp_path / "out.bas")
    raw = out.read_bytes()
    assert raw.count(b"\n") == raw.count(b"\r\n")


# ---------------------------------------------------------------- refusals

def test_a_uses_line_naming_something_that_does_not_exist_is_refused(tmp_path):
    """Silently dropping it would leave the test calling a name that is not there, which in an
    invisible Word is a compile error, a modal dialog, and a hang."""
    with pytest.raises(SystemExit) as exc:
        mod.build(root=repo(tmp_path, uses=("Sh_Top", "Sh_Was_Renamed")),
                  out_path=tmp_path / "out.bas")
    assert "Sh_Was_Renamed" in str(exc.value)


def test_no_uses_lines_at_all_is_refused(tmp_path):
    with pytest.raises(SystemExit) as exc:
        mod.build(root=repo(tmp_path, uses=()), out_path=tmp_path / "out.bas")
    assert "nothing to bundle" in str(exc.value)


# ---------------------------------------------------------------- the real source

def test_the_real_tests_still_bundle(repo_root, tmp_path):
    """Every '@uses line in tests/vba names a procedure that is still in src/vba."""
    out, names, _ = mod.build(root=repo_root, out_path=tmp_path / "out.bas")
    for expected in ("Sh_IsValidRomanNumeral", "Lp_Merged_Pg_Number", "Sh_AC_Unescape"):
        assert expected in names
    assert "Private Function" not in out.read_bytes().decode("cp1252")
