"""check_form_calls.py -- a UserForm calling a macro that no longer exists.

8/3/2026: Sh_About_License_Text was removed from LPandBrlMacros and the two About dialogs
were left calling it. The build was clean; it failed on the transcriber's machine as
"Compile error in hidden module: Lp_About_Title_And_Agreement".
"""
import check_form_calls as mod

from conftest import write_crlf

FORM_HEADER = """\
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Lp_Some_Form
   Caption         =   "VistaType LP"
End
Attribute VB_Name = "Lp_Some_Form"
Attribute VB_Exposed = False
"""


def tree(tmp_path, form_body, module_body='Sub Lp_Known()\nEnd Sub\n'):
    write_crlf(tmp_path / "src" / "vba" / "LPandBrlMacros.bas", module_body)
    write_crlf(tmp_path / "src" / "forms" / "Lp_Some_Form.frm", FORM_HEADER + form_body)
    return tmp_path


def test_code_of_form_drops_the_designer_header():
    code = mod.code_of_form(FORM_HEADER + "\nOption Explicit\nSub X()\nEnd Sub\n")
    assert "VERSION 5.00" not in code
    assert "Caption" not in code
    assert "Option Explicit" in code


def test_code_of_form_keeps_everything_after_the_last_attribute_line():
    """A form can carry several Attribute lines; the code starts after the LAST one."""
    text = FORM_HEADER + "Attribute VB_Creatable = False\nSub X()\nEnd Sub\n"
    code = mod.code_of_form(text)
    assert "Attribute" not in code
    assert "Sub X()" in code


def test_a_form_calling_a_missing_macro_stops_the_build(tmp_path, run_tool):
    tree(tmp_path, "\nPrivate Sub Btn_Click()\n    Lp_Gone_Away\nEnd Sub\n")
    r = run_tool("check_form_calls.py", tmp_path)
    assert r.returncode == 1
    assert "Lp_Gone_Away" in r.stdout
    assert "Compile error in hidden module" in r.stdout


def test_a_form_calling_a_macro_that_exists_is_fine(tmp_path, run_tool):
    tree(tmp_path, "\nPrivate Sub Btn_Click()\n    Lp_Known\nEnd Sub\n")
    r = run_tool("check_form_calls.py", tmp_path)
    assert r.returncode == 0
    assert "form calls OK" in r.stdout


def test_application_run_of_a_missing_macro_warns_but_does_not_stop_the_build(tmp_path, run_tool):
    """Late-bound: it compiles, and fails when the transcriber presses the button."""
    tree(tmp_path, '\nPrivate Sub Btn_Click()\n    Application.Run "Lp_Gone_Away"\nEnd Sub\n')
    r = run_tool("check_form_calls.py", tmp_path)
    assert r.returncode == 0
    assert "WARNING" in r.stdout
    assert "Lp_Gone_Away" in r.stdout
    assert "cannot be found or has been disabled" in r.stdout


def test_another_form_is_a_name_a_form_may_use(tmp_path, run_tool):
    """Forms are referenced as objects, so a form's own name counts as defined."""
    write_crlf(tmp_path / "src" / "vba" / "LPandBrlMacros.bas", "Sub Lp_Known()\nEnd Sub\n")
    write_crlf(tmp_path / "src" / "forms" / "Lp_Other_Form.frm", FORM_HEADER)
    write_crlf(tmp_path / "src" / "forms" / "Lp_Some_Form.frm",
               FORM_HEADER + "\nPrivate Sub Btn_Click()\n    Lp_Other_Form.Show\nEnd Sub\n")
    r = run_tool("check_form_calls.py", tmp_path)
    assert r.returncode == 0


def test_a_name_without_a_project_prefix_is_not_checked(tmp_path, run_tool):
    """Deliberately narrow: only Lp_ Dx_ Sh_ MS_ DN_ Vt_ names are the project's own."""
    tree(tmp_path, "\nPrivate Sub Btn_Click()\n    SomeWordThing.Refresh\nEnd Sub\n")
    r = run_tool("check_form_calls.py", tmp_path)
    assert r.returncode == 0


def test_a_commented_out_call_is_ignored(tmp_path, run_tool):
    tree(tmp_path, "\nPrivate Sub Btn_Click()\n    ' Lp_Gone_Away\nEnd Sub\n")
    r = run_tool("check_form_calls.py", tmp_path)
    assert r.returncode == 0


def test_the_real_repository_passes_its_own_check(repo_root, run_tool):
    r = run_tool("check_form_calls.py", repo_root)
    assert r.returncode == 0, r.stdout
