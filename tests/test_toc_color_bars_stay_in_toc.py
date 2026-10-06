"""Add Color Bars (353) takes the leader dots off the TOC only.

Issue #5, 10/6/2026. Before its remove-leaders loop, Lp_TOC_Color_Bars_Form selected the whole
book (ActiveDocument.Content.Select), so every TOC-styled paragraph anywhere in the book lost its
leader dots, not just the TOC being colored. The TOC box selects its held range
(Lp_Tocb_Range.Select) before it shows this form, so the selection already is the TOC; the line
was deleted.

Not unit-testable in VBA: a UserForm blocks the headless runner (tests/vba/README.md). So this
is a check on the source.
"""
import re

from test_toc_box_stays_open import form_code

FORM = "Lp_TOC_Color_Bars_Form"
WIDENS = re.compile(r"\bContent\s*\.\s*Select\b|\bWholeStory\b", re.IGNORECASE)


def test_the_form_never_selects_the_whole_book(repo_root):
    for line in form_code(repo_root, FORM):
        assert not WIDENS.search(line), (
            f"{FORM} selects the whole book again ({line.strip()}). The remove-leaders loop "
            "would strip the dots from every TOC-styled paragraph, not only the TOC (issue #5).")


def test_the_toc_box_selects_the_toc_before_showing_the_form(repo_root):
    """The form relies on the selection being the TOC; that is the caller's job."""
    from test_toc_never_deletes_a_line import module_text, strip_comment

    lines = module_text(repo_root).splitlines()
    show = next((i for i, c in enumerate(lines) if re.search(FORM + r"\.Show\b", c)), None)
    assert show is not None, f"Nothing shows {FORM} any more"
    start = max(i for i, c in enumerate(lines[:show]) if re.match(r"\s*(Private\s+|Public\s+)?Sub\s", c))
    select = [c for c in lines[start:show] if "Lp_Tocb_Range.Select" in strip_comment(c)]
    assert select, f"The TOC box no longer selects its held range before showing {FORM}"
