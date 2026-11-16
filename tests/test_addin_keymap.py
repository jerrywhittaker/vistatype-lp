"""The add-in's own keyboard shortcuts -- src/keymap/addin-keymap.xml.

Word honors a key assignment stored in a loaded global add-in in every document, whatever
template it is attached to. So Ctrl+Alt+Shift+I -> Sh_Doc_Info is stored in LPandBRL.dotm, not
only in LargePrintTemplate.dotx, where it worked only in books attached to that template
(Jerry, 11/16/2026).

inject_keymap.py REPLACES word/customizations.xml in the .dotm. The shell .dotm already had one
(seven MathType keys and a reference to its attached toolbars), so the file must keep carrying
them, or the build would silently drop them.
"""
import pathlib
import re
import zipfile

import pytest

import inject_keymap as mod

ROOT = pathlib.Path(__file__).resolve().parents[1]
KEYMAP = ROOT / "src" / "keymap" / "addin-keymap.xml"
MAKEFILE = ROOT / "Makefile"
BUILT = ROOT / "dist" / "LPandBRL.dotm"
SHELL = ROOT / "LPandBRL.dotm"

# Ctrl (0x200) + Alt (0x400) + Shift (0x100) + I (0x49), as Word writes it.
DOC_INFO_KEY = "0749"
DOC_INFO_MACRO = "LPANDBRL.LPANDBRLMACROS.SH_DOC_INFO"


def bindings(xml):
    """kcmPrimary -> macro name, for the single-key macro bindings."""
    return dict(re.findall(
        r'<wne:keymap wne:kcmPrimary="([0-9A-F]{4})"><wne:macro wne:macroName="([^"]+)"/>',
        xml))


def test_ctrl_alt_shift_i_runs_doc_info():
    assert bindings(KEYMAP.read_text(encoding="utf-8")).get(DOC_INFO_KEY) == DOC_INFO_MACRO


def test_the_build_puts_it_into_the_addin():
    text = MAKEFILE.read_text(encoding="utf-8")
    build = text.split("\nbuild:", 1)[1].split("\n\n", 1)[0]
    assert "inject_keymap.py dist/$(DOTM) src/keymap/addin-keymap.xml" in build
    # After the ribbon goes in, so nothing later in the recipe rewrites the package.
    assert build.index("inject_customui.py dist/$(DOTM)") < build.index("addin-keymap.xml")


def test_it_keeps_every_binding_the_shell_addin_already_had():
    """The shell .dotm is the base every build starts from; whatever keys it carries would be
    lost the moment the keymap part is replaced, unless this file carries them too."""
    if not SHELL.exists():
        pytest.skip("no shell LPandBRL.dotm in the repo root")
    with zipfile.ZipFile(SHELL) as z:
        if mod.PART not in z.namelist():
            return
        before = z.read(mod.PART).decode("utf-8")
    after = KEYMAP.read_text(encoding="utf-8")
    for key in re.findall(r"<wne:keymap\b[^>]*>.*?</wne:keymap>", before):
        assert key in after, "dropped from the add-in keymap: " + key
    # The toolbar reference points at word/_rels/customizations.xml.rels -> attachedToolbars.bin.
    if "<wne:toolbars>" in before:
        assert before[before.index("<wne:toolbars>"):] == after[after.index("<wne:toolbars>"):]


def test_the_built_addin_carries_it():
    """Only when a build is lying around; a fresh clone has no dist/."""
    if not BUILT.exists():
        pytest.skip("no dist/LPandBRL.dotm - run make build")
    with zipfile.ZipFile(BUILT) as z:
        built = z.read(mod.PART).decode("utf-8")
    if built != KEYMAP.read_text(encoding="utf-8"):
        pytest.skip("dist/LPandBRL.dotm predates this keymap - run make build")
    assert bindings(built).get(DOC_INFO_KEY) == DOC_INFO_MACRO
