"""check_vba_line_length.py -- over 1023 characters and `make build` HANGS.

7/30/2026: a 1110-character changelog line in the LPandBrlMacros header cost three hung
builds. Word gives no error and no dialog; the stale lock file it leaves behind then looks
like the cause.
"""


def bas(tmp_path, line):
    path = tmp_path / "src" / "vba" / "Mod.bas"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes((line + "\r\n").encode("cp1252"))
    return path


def test_a_line_at_the_limit_is_allowed(tmp_path, run_tool):
    bas(tmp_path, "' " + "x" * 1021)          # exactly 1023
    r = run_tool("check_vba_line_length.py", tmp_path)
    assert r.returncode == 0
    assert "VBA line lengths OK" in r.stdout


def test_one_character_over_the_limit_stops_the_build(tmp_path, run_tool):
    bas(tmp_path, "' " + "x" * 1022)          # exactly 1024
    r = run_tool("check_vba_line_length.py", tmp_path)
    assert r.returncode == 1
    assert "1024 chars" in r.stdout
    assert "HANGS" in r.stdout


def test_a_long_but_legal_line_is_only_a_note(tmp_path, run_tool):
    bas(tmp_path, "' " + "x" * 899)           # exactly 901, over the 900 warning mark
    r = run_tool("check_vba_line_length.py", tmp_path)
    assert r.returncode == 0
    assert "note:" in r.stdout
    assert "901 chars" in r.stdout


def test_running_outside_the_repository_is_refused(tmp_path, run_tool):
    """Otherwise an empty glob reads as "nothing too long" and the gate passes silently."""
    r = run_tool("check_vba_line_length.py", tmp_path)
    assert r.returncode == 1
    assert "run from the repo root" in r.stderr


def test_the_real_source_is_within_the_limit(repo_root, run_tool):
    r = run_tool("check_vba_line_length.py", repo_root)
    assert r.returncode == 0, r.stdout
