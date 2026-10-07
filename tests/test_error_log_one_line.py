"""Every error-log entry stays on one line.

Issue #24, found 10/6/2026 in VistaType-Errors.log on vistabuild (3.0.539): the MSForms
description "DataObject:GetFromClipboard OpenClipboard Failed" ends in CR LF NUL. Sh_Report_Error
wrote it as it came, so the entry was split across two lines, and the next rewrite of the log by
Sh_Log_Newest_First left the earlier entry's details as about 2,000 NUL characters.

The cleaning itself is unit-tested in VBA (tests/vba/TestTextHelpers.bas, Sh_One_Log_Line). This
checks that the logging path actually uses it - the part VBA cannot test, because Sh_Report_Error
shows a dialog and reads the active document.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment


def _code(repo_root, name):
    procs = procedures(module_text(repo_root))
    assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return [strip_comment(c) for c in procs[name]]


def test_report_error_cleans_the_description_and_the_log_line(repo_root):
    code = _code(repo_root, "Sh_Report_Error")
    for var in ("errText", "logLine"):
        assert any(re.search(rf"\b{var}\s*=\s*Sh_One_Log_Line\(\s*{var}\s*\)", c, re.IGNORECASE)
                   for c in code), (
            f"Sh_Report_Error must pass {var} through Sh_One_Log_Line - a description ending in "
            f"CR LF NUL splits its log entry and wipes the one before it (issue #24).")


def test_description_is_cleaned_before_anything_uses_it(repo_root):
    code = _code(repo_root, "Sh_Report_Error")
    clean = next(i for i, c in enumerate(code)
                 if re.search(r"\berrText\s*=\s*Sh_One_Log_Line", c, re.IGNORECASE))
    first_use = next(i for i, c in enumerate(code)
                     if re.search(r"&\s*errText\b", c, re.IGNORECASE))
    assert clean < first_use, "errText is used before Sh_One_Log_Line cleans it."


def test_log_drops_nuls_already_in_the_file(repo_root):
    code = _code(repo_root, "Sh_Log_Newest_First")
    assert any(re.search(r"Replace\(\s*whole\s*,\s*vbNullChar\s*,", c, re.IGNORECASE)
               for c in code), (
        "Sh_Log_Newest_First must drop the NULs a log damaged by issue #24 still holds.")
