"""Attach LP Template leaves Heading 1-5 at their style's size, not the base size (issue #19).

9/26/2026. Found by Jerry in an interactive test: the heading sizes were right before Attach LP
Template and all at the base size after it. Sh_Set_Whole_Document_Font lays the base size over
every character as DIRECT formatting, and direct formatting overrides the style, so the heading
sizes Lp_Normalize_Styles sets never showed. Reproduced on the build box: 28/26/24/22/20 before,
all 18 after, with the styles still saying 28-20.

The cure is Lp_Restore_Style_Sizes, run after Lp_Normalize_Styles. What this checks:

  1. The attach calls it, and AFTER both the whole-document sizing and Lp_Normalize_Styles -
     before either, it would restore sizes that are then flattened or not yet final.
  2. It covers Heading 1-5, "1 point" and "TOC Heading", which the same line flattens.
  3. It sets the SIZE only, by Find and Replace on the style - never Font.Reset or re-applying
     the style, which would strip bold, italics or dashed underline inside a heading.
  4. It reads each style's own size, so it follows whatever Lp_Normalize_Styles decided.

Not unit-testable in VBA: it reaches a Document, and the headless runner hangs on that
(tests/vba/README.md). So this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

HELPER = "Lp_Restore_Style_Sizes"
ATTACH = "Lp_Attach_The_Template"
STYLES = ("Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5", "1 point", "TOC Heading")


def the_procs(repo_root):
    procs = procedures(module_text(repo_root))
    for name in (HELPER, ATTACH):
        assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return procs


def first_line(lines, pattern):
    return next((i for i, c in enumerate(lines) if re.search(pattern, c)), None)


def test_the_attach_restores_after_the_sizing_and_normalize(repo_root):
    lines = the_procs(repo_root)[ATTACH]
    sizing = first_line(lines, r"Sh_Set_Whole_Document_Font ActiveDocument, Lp_Base_Font_Name")
    normalize = first_line(lines, r'MacroName:="Lp_Normalize_Styles"')
    restore = first_line(lines, rf"^\s*{HELPER}\s+ActiveDocument\s*$")
    assert sizing is not None and normalize is not None
    assert restore is not None, f"{ATTACH} does not call {HELPER} ActiveDocument"
    assert sizing < normalize < restore, "the restore must follow the sizing and Normalize Styles"


def test_the_helper_covers_the_flattened_styles(repo_root):
    body = "\n".join(the_procs(repo_root)[HELPER])
    quoted = set(re.findall(r'"([^"]*)"', body)) - {""}
    assert set(STYLES) <= quoted, f"{HELPER} misses {sorted(set(STYLES) - quoted)}"


def test_the_helper_sets_the_size_only_from_the_style(repo_root):
    lines = the_procs(repo_root)[HELPER]
    body = "\n".join(lines)
    assert re.search(r"\(ByVal Doc As Document\)", lines[0]), "the Document must be passed ByVal"
    assert re.search(r"\.Replacement\.Font\.Size = sty\.Font\.Size", body)
    assert re.search(r"\.Style = sty\b", body) and "Replace:=wdReplaceAll" in body
    for banned in (r"\.Font\.Reset", r"\.Style = .*\n.*\.Range\.Style", r"Replacement\.Font\.(Bold|Italic|Underline)",
                   r"\.Range\.Style\s*="):
        assert not re.search(banned, body), f"{HELPER} must set the size only ('{banned}')"
    assert "Doc.Content.Find" in body, "main text only, the story Sh_Set_Whole_Document_Font sizes"
