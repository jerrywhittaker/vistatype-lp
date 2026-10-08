"""The hyperlink conversions touch web and email links only, and leave links within the book alone.

Issue #14, found 9/20/2026 by reading and measured 10/7/2026 on vistabuild: both
Lp_Convert_Hyperliks_To_Text and Dx_Convert_Hyperliks_To_Text unlinked .Range.Fields(i) where i
counted HYPERLINKS, so with any other field in front of an internal link the wrong field was frozen
and the link stayed - and the next loop then wrote the link's empty Address over it, erasing its
words ("See Chapter One for more." became "See  for more.").

Jerry's rule, 10/7/2026: NIMAS files have no internal links and DAISY's are lost in the conversion
to Word, so these macros deal only with URLs and email addresses and IGNORE links within the book.
The two Convert_Hyper_To_Addresses macros likewise no longer write "#bookmark" over them.

The macros work on the active document, which hangs the headless VBA runner
(tests/vba/README.md), so this is a check on the source.
"""
import re

import pytest

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

TO_TEXT = ["Lp_Convert_Hyperliks_To_Text", "Dx_Convert_Hyperliks_To_Text"]
TO_ADDRESSES = ["Lp_Convert_Hyper_To_Addresses", "Dx_Convert_Hyper_To_Addresses"]


def _code(repo_root, sub):
    procs = procedures(module_text(repo_root))
    assert sub in procs, f"{sub} is missing from LPandBrlMacros.bas"
    return [strip_comment(c) for c in procs[sub]]


@pytest.mark.parametrize("sub", TO_TEXT + TO_ADDRESSES)
def test_only_links_with_an_address_are_touched(repo_root, sub):
    code = _code(repo_root, sub)
    assert any(re.search(r"\.Address\s*<>\s*\"\"", c) for c in code), (
        f"{sub} must leave a link with no Address - a link within the book - alone (issue #14).")


@pytest.mark.parametrize("sub", TO_TEXT)
def test_internal_links_are_not_unlinked_or_deleted(repo_root, sub):
    code = _code(repo_root, sub)
    hits = [c.strip() for c in code
            if re.search(r"\.Fields\(\s*i\s*\)|\.Hyperlinks\(\s*i\s*\)\.Delete\b|\.SubAddress\b",
                         c, re.IGNORECASE)]
    assert not hits, (
        f"{sub} works on links within the book again - Jerry's rule is to ignore them, and the "
        f"old pass froze the wrong field (issue #14): {hits}")
