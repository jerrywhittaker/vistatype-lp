"""check_frm_eol.py -- a .frm must stay CRLF.

With bare LF, Word does not recognize the designer header, treats those lines as VBA and
drops them into the form's code module. The build succeeds; it fails only on the
transcriber's machine, as "Compile error in hidden module: <FormName>".
"""
import pytest


def form(tmp_path, name, data):
    path = tmp_path / "src" / "forms" / name
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return path


def test_crlf_endings_pass(tmp_path, run_tool):
    form(tmp_path, "Lp_A_Form.frm", b'Attribute VB_Name = "Lp_A_Form"\r\nSub X()\r\nEnd Sub\r\n')
    r = run_tool("check_frm_eol.py", tmp_path)
    assert r.returncode == 0


def test_bare_lf_stops_the_build(tmp_path, run_tool):
    form(tmp_path, "Lp_A_Form.frm", b'Attribute VB_Name = "Lp_A_Form"\nSub X()\nEnd Sub\n')
    r = run_tool("check_frm_eol.py", tmp_path)
    assert r.returncode == 1
    assert "bare LF" in r.stderr
    assert "Lp_A_Form.frm" in r.stderr


def test_a_single_stray_lf_in_an_otherwise_crlf_file_is_caught(tmp_path, run_tool):
    """What a Linux editor does to one line. The count in the message is that one line."""
    form(tmp_path, "Lp_A_Form.frm",
         b'Attribute VB_Name = "Lp_A_Form"\r\nSub X()\nEnd Sub\r\n')
    r = run_tool("check_frm_eol.py", tmp_path)
    assert r.returncode == 1
    assert "(1 bare LF)" in r.stderr


def test_the_real_forms_are_all_crlf(repo_root, run_tool):
    r = run_tool("check_frm_eol.py", repo_root)
    assert r.returncode == 0, r.stderr
