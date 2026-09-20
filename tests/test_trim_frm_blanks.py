"""trim_frm_blanks.py -- the blank lines Word piles up on every .frm export.

The two About dialogs are re-exported on EVERY build because the version stamper rewrites
their captions, so by 8/1/2026 they had grown to 50 and 43 blank lines. Left alone they make
a one-word caption change look like a real edit.
"""
import pytest

import trim_frm_blanks as mod

HEAD = 'VERSION 5.00\r\nAttribute VB_Name = "Lp_A_Form"\r\n'
BODY = "Option Explicit\r\n\r\nSub X()\r\nEnd Sub\r\n"


def frm(tmp_path, data):
    path = tmp_path / "Lp_A_Form.frm"
    path.write_bytes(data if isinstance(data, bytes) else data.encode("cp1252"))
    return path


def test_the_leading_run_collapses_to_one_blank_line(tmp_path):
    path = frm(tmp_path, HEAD + "\r\n" * 6 + BODY)
    _, removed = mod.trim(str(path))
    assert removed == 5
    assert path.read_bytes() == (HEAD + "\r\n" + BODY).encode("cp1252")


def test_the_trailing_run_collapses_to_one_blank_line(tmp_path):
    path = frm(tmp_path, HEAD + "\r\n" + BODY + "\r\n" * 5)
    mod.trim(str(path))
    assert path.read_bytes().endswith(b"End Sub\r\n")
    assert not path.read_bytes().endswith(b"End Sub\r\n\r\n")


def test_blank_lines_between_subs_are_the_authors_and_are_kept(tmp_path):
    """Export noise is only at the two ends. Everything in between is Jerry's formatting."""
    body = "Option Explicit\r\n\r\n\r\n\r\nSub X()\r\nEnd Sub\r\n\r\n\r\nSub Y()\r\nEnd Sub\r\n"
    path = frm(tmp_path, HEAD + "\r\n" + body)
    mod.trim(str(path))
    text = path.read_bytes().decode("cp1252")
    assert "Option Explicit\r\n\r\n\r\n\r\nSub X()" in text
    assert "End Sub\r\n\r\n\r\nSub Y()" in text


def test_crlf_survives(tmp_path):
    """The load-bearing property. A .frm rewritten with bare LF fails on the user's machine."""
    path = frm(tmp_path, HEAD + "\r\n" * 4 + BODY)
    mod.trim(str(path))
    raw = path.read_bytes()
    assert raw.count(b"\n") == raw.count(b"\r\n")


def test_trimming_twice_changes_nothing_the_second_time(tmp_path):
    path = frm(tmp_path, HEAD + "\r\n" * 6 + BODY + "\r\n" * 4)
    mod.trim(str(path))
    after_first = path.read_bytes()
    _, removed = mod.trim(str(path))
    assert removed == 0
    assert path.read_bytes() == after_first


def test_an_already_tidy_form_is_left_alone(tmp_path):
    path = frm(tmp_path, HEAD + "\r\n" + BODY)
    before = path.read_bytes()
    _, removed = mod.trim(str(path))
    assert removed == 0
    assert path.read_bytes() == before


def test_a_file_with_no_attribute_line_is_not_a_form_and_is_untouched(tmp_path):
    path = frm(tmp_path, "Sub X()\r\n\r\n\r\n\r\nEnd Sub\r\n")
    before = path.read_bytes()
    _, removed = mod.trim(str(path))
    assert removed == 0
    assert path.read_bytes() == before


def test_a_bare_lf_file_is_refused_rather_than_rewritten(tmp_path):
    """Rewriting it would hide the very fault check_frm_eol.py exists to catch."""
    path = frm(tmp_path, b'Attribute VB_Name = "X"\nSub X()\nEnd Sub\n')
    before = path.read_bytes()
    with pytest.raises(SystemExit):
        mod.trim(str(path))
    assert path.read_bytes() == before
