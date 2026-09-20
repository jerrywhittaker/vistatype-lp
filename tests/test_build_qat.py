"""build_qat.py -- the Quick Access Toolbar files.

The hidden ribbon tab is the source of truth for which VistaType buttons can appear on the
toolbar. The curated toolbar is hand-maintained and only checked; the icons-only file is
generated and is APPENDED to whatever the transcriber already has, so it must never be able
to move, hide or reorder a button they put there themselves.
"""
import xml.etree.ElementTree as ET

import pytest

import build_qat as mod

from ribbon_fixture import RIBBON

TOOLBAR = """<?xml version="1.0" encoding="utf-8"?>
<mso:customUI xmlns:x1="__VT_DOTM_PATH__" xmlns:msox="http://schemas.microsoft.com/office/2006/01/customui/special" xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui">
  <mso:ribbon>
    <mso:qat>
      <mso:sharedControls>
        <mso:control idQ="mso:FileSave" visible="true"/>
        <mso:separator idQ="msox:vtsep1" visible="true"/>
        <mso:control idQ="x1:btn_Lp_One" visible="true" imageMso="HappyFace"/>
        <mso:control idQ="x1:btn_Dx_One" visible="true" imageMso="Cat"/>
      </mso:sharedControls>
    </mso:qat>
  </mso:ribbon>
</mso:customUI>
"""


def toolbar(tmp_path, text=TOOLBAR):
    path = tmp_path / "qat-template.officeUI"
    path.write_text(text, encoding="utf-8")
    return str(path)


def test_the_hidden_tabs_buttons_are_read_with_their_icons():
    buttons = mod.hidden_tab_buttons(RIBBON)
    assert buttons == {"btn_Lp_One": ("HappyFace", "One"),
                       "btn_Dx_One": ("Cat", "One")}


def test_a_ribbon_without_the_hidden_tab_is_refused():
    with pytest.raises(SystemExit):
        mod.hidden_tab_buttons("<customUI><ribbon><tabs></tabs></ribbon></customUI>")


def test_a_toolbar_that_matches_the_ribbon_has_no_problems(tmp_path, monkeypatch):
    monkeypatch.setattr(mod, "FULL", toolbar(tmp_path))
    assert mod.validate_full(mod.hidden_tab_buttons(RIBBON)) == []


def test_a_toolbar_naming_a_button_that_is_not_on_the_hidden_tab(tmp_path, monkeypatch):
    """The entry would render blank on the transcriber's machine, with no error."""
    text = TOOLBAR.replace('idQ="x1:btn_Dx_One"', 'idQ="x1:btn_Never_Existed"')
    monkeypatch.setattr(mod, "FULL", toolbar(tmp_path, text))
    problems = mod.validate_full(mod.hidden_tab_buttons(RIBBON))
    assert len(problems) == 1
    assert "btn_Never_Existed" in problems[0]


def test_an_icon_that_disagrees_with_the_ribbon_is_reported(tmp_path, monkeypatch):
    text = TOOLBAR.replace('idQ="x1:btn_Dx_One" visible="true" imageMso="Cat"',
                           'idQ="x1:btn_Dx_One" visible="true" imageMso="Dog"')
    monkeypatch.setattr(mod, "FULL", toolbar(tmp_path, text))
    problems = mod.validate_full(mod.hidden_tab_buttons(RIBBON))
    assert len(problems) == 1
    assert "toolbar icon 'Dog' but ribbon icon 'Cat'" in problems[0]


def test_a_missing_toolbar_file_is_skipped_not_failed(tmp_path, monkeypatch):
    monkeypatch.setattr(mod, "FULL", str(tmp_path / "not-there.officeUI"))
    assert mod.validate_full(mod.hidden_tab_buttons(RIBBON)) == []


def test_the_generated_file_keeps_the_hidden_tabs_order(tmp_path, monkeypatch):
    out = tmp_path / "qat-icons-only.officeUI"
    monkeypatch.setattr(mod, "ICONS_ONLY", str(out))
    order = mod.generate_icons_only(mod.hidden_tab_buttons(RIBBON), RIBBON)
    assert order == ["btn_Lp_One", "btn_Dx_One"]
    text = out.read_text(encoding="utf-8")
    assert text.index("btn_Lp_One") < text.index("btn_Dx_One")


def test_the_generated_file_can_never_hide_a_users_own_button(tmp_path, monkeypatch):
    """It is APPENDED to their toolbar. No mso: entries, and no visible="false" anywhere."""
    out = tmp_path / "qat-icons-only.officeUI"
    monkeypatch.setattr(mod, "ICONS_ONLY", str(out))
    mod.generate_icons_only(mod.hidden_tab_buttons(RIBBON), RIBBON)
    text = out.read_text(encoding="utf-8")
    # The header comment explains this rule, so it names the thing it forbids. Look at the
    # controls only.
    controls = text[text.index("-->"):]
    assert 'visible="false"' not in controls
    assert 'idQ="mso:' not in controls
    assert 'idQ="x1:btn_Lp_One"' in controls


def test_the_generated_file_is_well_formed(tmp_path, monkeypatch):
    out = tmp_path / "qat-icons-only.officeUI"
    monkeypatch.setattr(mod, "ICONS_ONLY", str(out))
    mod.generate_icons_only(mod.hidden_tab_buttons(RIBBON), RIBBON)
    assert mod.validate_xml(str(out)) is True


def test_a_file_that_does_not_parse_is_refused(tmp_path):
    """A toolbar file Word cannot parse produces no error there -- the install just quietly
    leaves the toolbar alone -- so here is the only useful place to catch it."""
    bad = tmp_path / "broken.officeUI"
    bad.write_text("<mso:customUI><unclosed></mso:customUI>", encoding="utf-8")
    with pytest.raises(SystemExit):
        mod.validate_xml(str(bad))


def test_a_double_hyphen_in_the_comment_would_be_caught(tmp_path):
    """Word and PowerShell both refuse a file whose XML comment contains one."""
    bad = tmp_path / "hyphen.officeUI"
    bad.write_text('<?xml version="1.0"?>\n<!-- a -- b -->\n<root/>\n', encoding="utf-8")
    with pytest.raises(SystemExit):
        mod.validate_xml(str(bad))


def test_the_real_toolbar_and_ribbon_still_agree(repo_root, run_tool):
    r = run_tool("build_qat.py", repo_root, "--check-only")
    assert r.returncode == 0, r.stdout + r.stderr
