"""Two uncalled large-print picture macros stay deleted.

Issue #9, 10/6/2026. Nothing called Lp_Make_All_Pictures_In_Selected_Table_Inline or
Lp_Picture_Color_Change_Menu - no ribbon button, key, form or other macro - so Jerry had both
deleted. The background picture menu (Lp_Bakgrnd_Picture_Menu_Form) opens
Lp_Change_Image_Color_Form itself, so nothing was lost with the menu macro.

The other three macros issue #9 named were dealt with on 9/24/2026 in the issue #3 change; see
test_dx_page_breaks_removed.py.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

GONE = ("Lp_Make_All_Pictures_In_Selected_Table_Inline", "Lp_Picture_Color_Change_Menu")


def test_the_two_macros_are_gone(repo_root):
    procs = procedures(module_text(repo_root))
    back = [name for name in GONE if name in procs]
    assert not back, f"deleted for issue #9 because nothing called them, but back: {back}"


def test_nothing_calls_them(repo_root):
    pattern = re.compile(r"\b(" + "|".join(GONE) + r")\b")
    hits = []
    for folder in ("src/vba", "src/forms", "src/ribbon", "src/keymap", "installer"):
        for path in sorted((repo_root / folder).rglob("*")):
            if not path.is_file() or path.suffix.lower() in (".frx", ".dotm", ".dotx", ".exe"):
                continue
            for n, line in enumerate(path.read_bytes().decode("latin-1").splitlines(), 1):
                code = line.split("'", 1)[0] if path.suffix.lower() in (".bas", ".frm", ".cls") else line
                if pattern.search(code):
                    hits.append(f"{path.relative_to(repo_root)}:{n}")
    assert not hits, f"something calls a macro deleted for issue #9: {hits}"


def test_the_picture_menu_still_opens_the_color_form(repo_root):
    form = (repo_root / "src" / "forms" / "Lp_Bakgrnd_Picture_Menu_Form.frm").read_bytes().decode("latin-1")
    code = [line.split("'", 1)[0] for line in form.splitlines()]
    assert any(re.search(r"\bLp_Change_Image_Color_Form\.Show\b", c) for c in code), (
        "Lp_Bakgrnd_Picture_Menu_Form no longer opens Lp_Change_Image_Color_Form - that was the "
        "reason Lp_Picture_Color_Change_Menu could be deleted (issue #9).")
