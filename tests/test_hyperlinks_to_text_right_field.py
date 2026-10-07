"""Hyperlinks to Text removes the internal link itself, not a field picked by the link's number.

Issue #14, found 9/20/2026 by reading: both Lp_Convert_Hyperliks_To_Text and
Dx_Convert_Hyperliks_To_Text walked the hyperlinks with i and then unlinked .Range.Fields(i). i
counts hyperlinks, not fields, so with any other field in front of a link - a page reference, a
date, a TOC - the wrong field was frozen to plain text and the link stayed.

The cure deletes the hyperlink itself, which keeps its text. The macros work on the active
document, which hangs the headless VBA runner (tests/vba/README.md), so this is a check on the
source; the behavior was measured separately on vistabuild.
"""
import re

import pytest

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

SUBS = ["Lp_Convert_Hyperliks_To_Text", "Dx_Convert_Hyperliks_To_Text"]


@pytest.mark.parametrize("sub", SUBS)
def test_no_field_is_picked_by_the_hyperlink_number(repo_root, sub):
    procs = procedures(module_text(repo_root))
    assert sub in procs, f"{sub} is missing from LPandBrlMacros.bas"
    code = [strip_comment(c) for c in procs[sub]]
    hits = [c.strip() for c in code if re.search(r"\.Fields\(\s*i\s*\)", c, re.IGNORECASE)]
    assert not hits, (
        f"{sub} picks a field by the hyperlink's number again - it freezes the wrong field "
        f"whenever another field comes before a link (issue #14): {hits}")
    assert any(re.search(r"\.Hyperlinks\(\s*i\s*\)\.Delete\b", c, re.IGNORECASE) for c in code), (
        f"{sub} must remove an internal link with .Hyperlinks(i).Delete.")
