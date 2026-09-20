"""build_ribbon_tabs.py -- the tabs Word will let a transcriber hide and reorder.

Jerry, 7/28/2026: "Even I want to hide the braille macro tab when I'm producing large print,
and it is not unusual that I will rearrange the tabs." Tabs from the add-in's embedded
customUI do not appear in Customize the Ribbon; tabs in the user's own Word.officeUI do.
"""
import xml.etree.ElementTree as ET

import pytest

import build_ribbon_tabs as mod

from ribbon_fixture import RIBBON


def test_sanitize_makes_a_legal_id_out_of_a_label():
    assert mod.sanitize("VistaType LP") == "VistaType_LP"
    assert mod.sanitize("Braille Macros") == "Braille_Macros"
    assert mod.sanitize("Files & Pages") == "Files___Pages"


def test_attr_returns_the_literal_source_text():
    """Entities must pass through byte-identical; decoding and re-escaping is where an
    ampersand gets doubled and an arrow turns to mojibake."""
    assert mod.attr('<button label="Two &amp;&amp; More"/>', "label") == "Two &amp;&amp; More"
    assert mod.attr('<button label="a &#8594; b"/>', "label") == "a &#8594; b"
    assert mod.attr('<button id="x"/>', "label") is None


def test_the_hidden_and_retired_tabs_are_not_offered_to_the_user():
    """The QAT tab must stay embedded: every toolbar in the field names its button ids."""
    ids = [t[0] for t in mod.visible_tabs(RIBBON)]
    assert ids == ["tab_LP", "tab_BRL"]


def test_every_button_is_emitted_as_a_reference_not_a_standalone_macro_button():
    """A standalone onAction button calls the sub directly and skips RibbonAction, which
    clears Sh_Pos_Depth / Sh_Pos_Saved on every press."""
    body, _ = mod.generate(RIBBON)
    assert '<mso:control idQ="x1:btn_Lp_One" visible="true"/>' in body
    assert "onAction" not in body


def test_the_tabs_are_anchored_in_the_mso_namespace():
    """The legacy file anchored a tab after Adobe Acrobat's, which only resolved on a
    machine that had Acrobat."""
    body, _ = mod.generate(RIBBON)
    assert 'insertBeforeQ="mso:TabInsert"' in body
    assert "insertBeforeMso" not in body
    mod.check_anchors(body)


def test_an_anchor_outside_the_mso_namespace_is_refused():
    with pytest.raises(SystemExit) as exc:
        mod.check_anchors('<mso:tab insertBeforeQ="acrobat:Tab"/>')
    assert "would not resolve elsewhere" in str(exc.value)


def test_a_word_control_keeps_its_mso_prefix():
    body, _ = mod.generate(RIBBON)
    assert '<mso:control idQ="mso:Bold" visible="true"/>' in body


def test_the_counts_name_each_tab_and_its_buttons():
    _, counts = mod.generate(RIBBON)
    assert counts == [("VistaType LP", "vt_tab_VistaType_LP", 2),
                      ("Braille Macros", "vt_tab_Braille_Macros", 1)]


def test_the_whole_generated_file_is_well_formed(tmp_path):
    body, _ = mod.generate(RIBBON)
    text = mod.HEADER.format(ribbon="src/ribbon/customUI14.xml") + body + mod.FOOTER
    out = tmp_path / "ribbon-tabs.officeUI"
    out.write_text(text, encoding="utf-8")
    ET.parse(out)


def test_the_generated_comment_has_no_double_hyphen(tmp_path):
    """An XML comment may not contain one. Word and PowerShell refuse the whole file, and
    the only symptom is an install that silently leaves the ribbon alone."""
    text = mod.HEADER.format(ribbon="src/ribbon/customUI14.xml")
    comment = text[text.index("<!--") + 4:text.index("-->")]
    assert "--" not in comment


# ---------------------------------------------------------------- button ids are a contract

def ids_file(tmp_path, lines):
    path = tmp_path / "installer" / "ribbon-button-ids.txt"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("# shipped ids\n" + "".join(l + "\n" for l in lines), encoding="utf-8")
    return path


def test_a_shipped_button_id_that_disappears_stops_the_build(tmp_path, monkeypatch):
    """The hard rule. Toolbars in the field reference these by id; a removed one renders
    blank on the transcriber's machine with no error, and their file is not ours to fix."""
    monkeypatch.chdir(tmp_path)
    ids_file(tmp_path, ["btn_Lp_One", "btn_Lp_Two", "btn_Dx_One", "btn_Lp_Gone",
                        "btn_Lp_Retired_Long_Ago"])
    with pytest.raises(SystemExit) as exc:
        mod.check_ids_are_stable(RIBBON)
    assert "btn_Lp_Retired_Long_Ago" in str(exc.value)
    assert "already shipped" in str(exc.value)


def test_a_renamed_button_id_reads_as_a_removal(tmp_path, monkeypatch):
    """Which is the point: a rename is the silent way to break every installed toolbar."""
    monkeypatch.chdir(tmp_path)
    ids_file(tmp_path, ["btn_Lp_One_OldName", "btn_Lp_Two", "btn_Dx_One", "btn_Lp_Gone"])
    with pytest.raises(SystemExit) as exc:
        mod.check_ids_are_stable(RIBBON)
    assert "btn_Lp_One_OldName" in str(exc.value)


def test_a_new_button_id_is_recorded_without_complaint(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    path = ids_file(tmp_path, ["btn_Lp_One", "btn_Lp_Two", "btn_Lp_Gone"])
    mod.check_ids_are_stable(RIBBON)
    assert "btn_Dx_One" in path.read_text()


def test_a_missing_ledger_is_seeded_from_the_ribbon(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "installer").mkdir()
    mod.check_ids_are_stable(RIBBON)
    written = (tmp_path / "installer" / "ribbon-button-ids.txt").read_text()
    for bid in ("btn_Lp_One", "btn_Lp_Two", "btn_Dx_One", "btn_Lp_Gone"):
        assert bid in written


def test_the_retired_tabs_ids_are_still_tracked(tmp_path, monkeypatch):
    """Retiring a button means moving it to the hidden retired tab, not deleting the id."""
    monkeypatch.chdir(tmp_path)
    (tmp_path / "installer").mkdir()
    mod.check_ids_are_stable(RIBBON)
    written = (tmp_path / "installer" / "ribbon-button-ids.txt").read_text()
    assert "btn_Lp_Gone" in written


def test_the_real_ribbon_still_carries_every_id_that_has_shipped(repo_root, run_tool, tmp_path):
    """A canary against the tracked ribbon -- run on a COPY.

    check_ids_are_stable APPENDS a newly added id to the ledger, and that happens even with
    --check-only. A test must never write into the repository, so the two files it reads are
    copied out first. A removed id -- the failure that matters -- still fails here.
    """
    import shutil
    for rel in ("src/ribbon/customUI14.xml", "installer/ribbon-button-ids.txt"):
        dst = tmp_path / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(repo_root / rel, dst)
    r = run_tool("build_ribbon_tabs.py", tmp_path, "--check-only")
    assert r.returncode == 0, r.stdout + r.stderr
    assert (tmp_path / "installer/ribbon-button-ids.txt").read_text() == \
        (repo_root / "installer/ribbon-button-ids.txt").read_text(), \
        "the ribbon has button ids that are not yet recorded in the shipped-ids ledger"
