"""check_try_scope.py -- saying when `make try` cannot test what was changed.

Jerry, 8/23/2026: he should never have to work out for himself that what he is about to
test cannot be tested that way. The load-bearing part is only_the_version_bump: `make try`
bumps the version on every run, so without it the warning names the installer EVERY time,
and a warning that always fires is a warning nobody reads.
"""
import subprocess

import check_try_scope as mod

ISS = '''#define AppName "VistaType LP"
#define AppVer "3.0.460"
#define AppPublisher "Jerry Whittaker"
'''


def git(tmp_path, *args):
    return subprocess.run(["git", *args], cwd=str(tmp_path),
                          capture_output=True, text=True, check=True)


def repo(tmp_path, text=ISS):
    git(tmp_path, "init", "-q")
    git(tmp_path, "config", "user.email", "test@example.com")
    git(tmp_path, "config", "user.name", "Test")
    path = tmp_path / "installer" / "vistatype.iss"
    path.parent.mkdir(parents=True)
    path.write_text(text, encoding="utf-8")
    git(tmp_path, "add", "-A")
    git(tmp_path, "commit", "-q", "-m", "first")
    return path


REL = "installer/vistatype.iss"


def test_a_version_bump_alone_does_not_count_as_a_change(tmp_path, monkeypatch):
    path = repo(tmp_path)
    path.write_text(ISS.replace("3.0.460", "3.0.461"), encoding="utf-8")
    monkeypatch.chdir(tmp_path)
    assert mod.only_the_version_bump(REL) is True


def test_a_real_installer_edit_beside_the_bump_does_count(tmp_path, monkeypatch):
    path = repo(tmp_path)
    path.write_text(ISS.replace("3.0.460", "3.0.461")
                       .replace("Jerry Whittaker", "Jerry Whittaker (VistaType LP)"),
                    encoding="utf-8")
    monkeypatch.chdir(tmp_path)
    assert mod.only_the_version_bump(REL) is False


def test_a_real_installer_edit_on_its_own_does_count(tmp_path, monkeypatch):
    path = repo(tmp_path)
    path.write_text(ISS.replace("VistaType LP", "VistaType LP 3"), encoding="utf-8")
    monkeypatch.chdir(tmp_path)
    assert mod.only_the_version_bump(REL) is False


def test_an_unchanged_file_is_not_a_bump(tmp_path, monkeypatch):
    """No change at all must not read as "only the version", or nothing is ever reported."""
    repo(tmp_path)
    monkeypatch.chdir(tmp_path)
    assert mod.only_the_version_bump(REL) is False


def test_the_makefiles_own_version_line_counts_too(tmp_path, monkeypatch):
    """The other place the bump lands."""
    text = 'APPVER := 3.0.460\nOTHER := x\n'
    git(tmp_path, "init", "-q")
    git(tmp_path, "config", "user.email", "test@example.com")
    git(tmp_path, "config", "user.name", "Test")
    (tmp_path / "Makefile").write_text(text, encoding="utf-8")
    git(tmp_path, "add", "-A")
    git(tmp_path, "commit", "-q", "-m", "first")
    (tmp_path / "Makefile").write_text(text.replace("3.0.460", "3.0.461"), encoding="utf-8")
    monkeypatch.chdir(tmp_path)
    assert mod.only_the_version_bump("Makefile") is True


def test_every_path_that_needs_a_real_installer_is_matched():
    """The list itself. A pattern that stops matching costs an hour on the wrong question."""
    def why(path):
        for pattern, text in mod.NEEDS_INSTALLER:
            if pattern.search(path):
                return text
        return None

    assert why("installer/vistatype.iss") is not None
    assert why("src/ribbon/customUI14.xml") is not None
    assert why("src/keymap/lp-template-keymap.xml") is not None
    assert why("LargePrintTemplate.dotx") is not None
    assert why("assets/fonts/vistatypelp-sans/VistaTypeLPSans-Bold.ttf") is not None


def test_ordinary_vba_and_form_edits_are_a_fair_test_of_make_try():
    def matched(path):
        return any(p.search(path) for p, _ in mod.NEEDS_INSTALLER)

    assert not matched("src/vba/LPandBrlMacros.bas")
    assert not matched("src/forms/Lp_Some_Form.frm")
    assert not matched("docs/Reported-Errors.md")
