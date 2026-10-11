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


def test_remove_color_bars_leaves_each_entry_at_its_own_level(repo_root):
    """Issue #17, 10/11/2026, Jerry: each TOC entry keeps its own level. Lp_TOC_Remove_Color_Bars
    set TOC 1 on every paragraph that had a bar, then put back its saved ParagraphFormat - which
    put the old style back too. Both lines are gone; neither may come back."""
    from test_toc_never_deletes_a_line import module_text, procedures

    body = procedures(module_text(repo_root))["Lp_TOC_Remove_Color_Bars"]
    for c in body:
        assert not re.search(r"\.Style\s*=\s*doc\.Styles\(\s*\"TOC \d\"\s*\)", c), (
            f"Remove Color Bars sets a TOC level again: {c.strip()}")
        assert not re.search(r"\.Style\s*=\s*\"TOC \d\"", c), (
            f"Remove Color Bars sets a TOC level again: {c.strip()}")
        assert not re.search(r"ParagraphFormat\s*=", c), (
            f"Remove Color Bars puts a saved ParagraphFormat back again: {c.strip()}")
    # It still does its job: Words runs to Default Paragraph Font, and the leaders come back.
    assert any("wdStyleDefaultParagraphFont" in c for c in body)
    assert any("wdTabLeaderDots" in c for c in body)


# --- issue #17 again, 10/11/2026, Jerry on 3.0.559: the bars took the bold off ------------------
#
# A character style - a "Words..." bar, or Default Paragraph Font - strips the direct bold and
# italic of the range it goes on (measured on vistabuild, 10/11/2026). Add and Remove Color Bars
# each remember them with Lp_Remember_Bold_Italic just before the style line and put them back
# with Lp_Restore_Bold_Italic just after. Needs a real document, so this checks the source.

BAR_STYLE = re.compile(r"\.Style\s*=\s*BarColor\d\b", re.IGNORECASE)
DPF_STYLE = re.compile(r"\.Style\s*=\s*doc\.Styles\(\s*wdStyleDefaultParagraphFont\s*\)", re.IGNORECASE)


def _code_lines(lines):
    return [c for c in lines if c.strip()]


def _style_lines_are_wrapped(lines, pattern, where):
    code = _code_lines(lines)
    hits = [i for i, c in enumerate(code) if pattern.search(c)]
    assert hits, f"{where} no longer sets the style this test guards"
    for i in hits:
        target = code[i][:pattern.search(code[i]).start()].strip()
        before = code[i - 1] if i > 0 else ""
        after = code[i + 1] if i + 1 < len(code) else ""
        assert re.search(r"=\s*Lp_Remember_Bold_Italic\(\s*" + re.escape(target) + r"\s*\)", before), (
            f"{where}: '{code[i].strip()}' is not straight after Lp_Remember_Bold_Italic({target}) - "
            "the style strips the entry's bold and italic (issue #17)")
        assert re.search(r"^\s*Lp_Restore_Bold_Italic\s+" + re.escape(target) + r"\s*,", after), (
            f"{where}: '{code[i].strip()}' is not followed by Lp_Restore_Bold_Italic {target} - "
            "the entry's bold and italic are not put back (issue #17)")


def test_add_color_bars_keeps_bold_and_italic(repo_root):
    _style_lines_are_wrapped(form_code(repo_root, FORM), BAR_STYLE, FORM)


def test_remove_color_bars_keeps_bold_and_italic(repo_root):
    from test_toc_never_deletes_a_line import module_text, procedures

    body = procedures(module_text(repo_root))["Lp_TOC_Remove_Color_Bars"]
    _style_lines_are_wrapped(body, DPF_STYLE, "Lp_TOC_Remove_Color_Bars")


def test_the_bold_italic_helpers_exist_and_take_byval(repo_root):
    """The form calls them, so they are Public; ByVal, the rule since 8/26/2026."""
    from test_toc_never_deletes_a_line import module_text, procedures

    procs = procedures(module_text(repo_root))
    for name in ("Lp_Remember_Bold_Italic", "Lp_Restore_Bold_Italic"):
        assert name in procs, f"{name} is gone from LPandBrlMacros"
        header = procs[name][0]
        assert re.match(r"\s*Public\s", header), f"{name} must be Public - the form calls it"
        params = header[header.index("(") + 1:header.rindex(")")]
        for p in filter(None, (x.strip() for x in params.split(","))):
            assert p.startswith("ByVal "), f"{name}: parameter '{p}' is not ByVal"
