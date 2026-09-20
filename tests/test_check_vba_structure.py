"""check_vba_structure.py -- the three mistakes that ship as "Compile error in hidden module".

Nothing in this build compiles VBA, so these are caught here or on a transcriber's machine.
The 8/18/2026 build this check was written for is the first case below.
"""
import check_vba_structure as mod

from conftest import write_crlf


def bas(tmp_path, text, name="Mod.bas"):
    return write_crlf(tmp_path / name, text)


def test_a_clean_module_has_no_problems(tmp_path):
    path = bas(tmp_path, """\
Option Explicit

Private Const LIMIT As Long = 10

Sub Lp_One()
    With Selection
        .Text = "x"
    End With
End Sub

Function Sh_Two(ByVal s As String) As String
    Sh_Two = s
End Function
""")
    assert mod.check(str(path)) == []


def test_a_const_written_beside_its_procedure_is_caught(tmp_path):
    """The 8/18/2026 build: three Private Const lines below the first Sub."""
    path = bas(tmp_path, """\
Sub Lp_One()
End Sub

Private Const LIMIT As Long = 10

Sub Lp_Two()
End Sub
""")
    problems = mod.check(str(path))
    assert len(problems) == 1
    assert "module-level declaration below the first procedure" in problems[0]
    assert "Private Const LIMIT As Long = 10" in problems[0]


def test_a_dim_inside_a_procedure_is_not_a_problem(tmp_path):
    path = bas(tmp_path, """\
Sub Lp_One()
    Dim i As Long
    Const LOCAL As Long = 2
End Sub
""")
    assert mod.check(str(path)) == []


def test_two_procedures_with_the_same_name(tmp_path):
    path = bas(tmp_path, """\
Sub Lp_One()
End Sub

Sub Lp_One()
End Sub
""")
    problems = mod.check(str(path))
    assert len(problems) == 1
    assert "procedure Lp_One is declared 2 times" in problems[0]


def test_a_with_that_is_never_closed(tmp_path):
    path = bas(tmp_path, """\
Sub Lp_One()
    With Selection
        .Text = "x"
End Sub
""")
    problems = mod.check(str(path))
    assert len(problems) == 1
    assert "has 1 unclosed With" in problems[0]
    assert "Lp_One" in problems[0]


def test_nested_withs_that_balance_are_fine(tmp_path):
    path = bas(tmp_path, """\
Sub Lp_One()
    With Selection
        With .Font
            .Bold = True
        End With
    End With
End Sub
""")
    assert mod.check(str(path)) == []


def test_comments_and_blank_lines_are_ignored(tmp_path):
    path = bas(tmp_path, """\
Sub Lp_One()
End Sub

' Private Const LIMIT As Long = 10

Sub Lp_Two()
End Sub
""")
    assert mod.check(str(path)) == []


def test_a_form_designer_header_is_not_treated_as_code(tmp_path):
    """The header above the Attribute lines is Word's, and it is full of Begin/End blocks.

    Reading it as VBA would report a declaration below a procedure on every single form.
    """
    path = write_crlf(tmp_path / "Lp_Some_Form.frm", """\
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Some_Form
   Caption         =   "VistaType LP"
   ClientHeight    =   3000
End
Attribute VB_Name = "Lp_Some_Form"
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
End Sub
""")
    assert mod.check(str(path)) == []


def test_the_real_repository_passes_its_own_check(repo_root):
    """A canary. If this fails, `make build` is already refusing."""
    problems = []
    for pattern in ("src/vba/*.bas", "src/vba/*.cls", "src/forms/*.frm"):
        for path in sorted(repo_root.glob(pattern)):
            problems.extend(mod.check(str(path)))
    assert problems == []
