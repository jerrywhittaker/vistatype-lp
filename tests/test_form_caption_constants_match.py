"""A window title held in the VBA must match its form's caption, character for character.

Found 9/22/2026, issue 16: the two $pg constants in ShNonModalMessage.bas still read
"Validate $pg Tags" and "Delete/change/add $pg" after the forms gained their dialog numbers,
(364) and (366).

That matters because the only way to bring a modeless box to the front is to look its window
up by title, and Windows matches the WHOLE title, not the start of it. Sh_PgVal_FocusMenu gets
0 back, reads that as "no menu on screen", and gives F6 up. The keyboard route dies silently,
and only a transcriber working without a mouse ever finds out.

So every constant that names a form's window is locked to that form's caption here. Renaming a
dialog, or giving it a number, now fails the build instead of the keyboard.
"""
import re

import pytest

# (module under src/vba, constant name, form under src/forms)
PAIRS = [
    ("ShNonModalMessage.bas", "SH_PGVAL_TITLE_LIST", "Sh_Valid_Ref_Pg_No_2_Form.frm"),
    ("ShNonModalMessage.bas", "SH_PGVAL_TITLE_DOC", "Sh_Valid_Ref_Pg_No_4_Form.frm"),
    ("LPandBrlMacros.bas", "LP_RST_TITLE", "Lp_Same_Pic_Range_Form.frm"),
    ("LPandBrlMacros.bas", "LP_TOCB_TITLE", "Lp_TOC_Format_And_Color_Form.frm"),
    ("LPandBrlMacros.bas", "LP_FIL_TITLE", "Lp_Type_Fill_In_Line_Form.frm"),
    ("LPandBrlMacros.bas", "LP_HVB_TITLE", "Lp_Horz_To_Vert_List_Form.frm"),
]

CONST = r'Const\s+{name}\s+As\s+String\s*=\s*"([^"]*)"'
CAPTION = r'^\s*Caption\s*=\s*"([^"]*)"'


def constant_value(repo_root, module, name):
    text = (repo_root / "src" / "vba" / module).read_text(encoding="utf-8")
    m = re.search(CONST.format(name=re.escape(name)), text)
    assert m, f"{name} is no longer declared in {module}"
    return m.group(1)


def caption(repo_root, form):
    path = repo_root / "src" / "forms" / form
    assert path.exists(), f"{form} is missing from src/forms"
    # The caption is in the designer header, which is the first handful of lines.
    for line in path.read_text(encoding="latin-1").splitlines()[:20]:
        m = re.match(CAPTION, line)
        if m:
            return m.group(1)
    pytest.fail(f"{form} has no Caption line in its designer header")


@pytest.mark.parametrize("module,name,form", PAIRS, ids=[p[1] for p in PAIRS])
def test_the_constant_is_the_whole_caption(repo_root, module, name, form):
    assert constant_value(repo_root, module, name) == caption(repo_root, form), (
        f"{name} in {module} is not {form}'s caption. Windows matches the whole window title, "
        f"so the keyboard cannot reach that box."
    )


def test_every_pair_is_listed(repo_root):
    """A new modeless form with a title constant must be added to PAIRS above.

    Finds `Const <NAME>_TITLE As String` in src/vba and insists each one is covered here.
    """
    listed = {name for _, name, _ in PAIRS}
    found = set()
    for path in sorted((repo_root / "src" / "vba").glob("*.bas")):
        for m in re.finditer(r'Const\s+(\w*TITLE\w*)\s+As\s+String\s*=\s*"', path.read_text(encoding="utf-8")):
            found.add(m.group(1))
    assert found <= listed, f"a window-title constant is not locked to a form: {sorted(found - listed)}"
