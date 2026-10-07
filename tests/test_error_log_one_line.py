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


def test_log_is_not_read_with_readall(repo_root):
    # Measured 10/7/2026 on a copy of the damaged log: ReadAll on a file holding NULs returned
    # junk, the Write then raised after CreateTextFile had emptied the file, and the whole log
    # was lost. The file is read in binary instead, by Sh_Read_Log_Text.
    writer = _code(repo_root, "Sh_Log_Newest_First")
    reader = _code(repo_root, "Sh_Read_Log_Text")
    for name, code in (("Sh_Log_Newest_First", writer), ("Sh_Read_Log_Text", reader)):
        assert not any(re.search(r"\.ReadAll\b", c, re.IGNORECASE) for c in code), (
            f"{name} reads the log with ReadAll again - on a log holding NULs that wipes the "
            f"whole file (issue #24).")
    assert any(re.search(r"\bSh_Read_Log_Text\(", c) for c in writer), (
        "Sh_Log_Newest_First must read the old log through Sh_Read_Log_Text.")
    assert any(re.search(r"\bOpen\b.*\bFor\s+Binary\b", c, re.IGNORECASE) for c in reader), (
        "Sh_Read_Log_Text must read the log in binary.")


def test_log_is_never_emptied_before_it_is_written(repo_root):
    # Issue #25, reproduced 10/7/2026 on vistabuild: the log was emptied by CreateTextFile and
    # then written as ANSI, so a Greek document name failed the write with the file already
    # empty and every earlier entry was lost. It is written to a scratch file, and only a
    # successful write is copied over the log.
    code = _code(repo_root, "Sh_Log_Newest_First")
    creates = [c for c in code if re.search(r"\bCreateTextFile\(", c, re.IGNORECASE)]
    assert creates, "Sh_Log_Newest_First no longer writes the log at all."
    for c in creates:
        assert not re.search(r"CreateTextFile\(\s*logPath\b", c, re.IGNORECASE), (
            f"the log itself is opened for writing, which empties it before the write: {c.strip()}")
    # Plain ANSI on purpose (Jerry's Notepad showed the 3.0.546 Unicode log with the letters
    # spaced out), so what ANSI cannot hold is turned into "?" before the write.
    assert any(re.search(r"StrConv\(\s*StrConv\(\s*keep\s*,\s*vbFromUnicode\s*\)\s*,\s*vbUnicode\s*\)", c,
                         re.IGNORECASE) for c in code), (
        "Sh_Log_Newest_First must turn characters ANSI cannot hold into ? before writing.")


def test_report_error_knows_when_the_log_was_not_written(repo_root):
    # The old Err.Number test after the call never saw a failure: an error handled under
    # Resume Next inside the sub is cleared when the sub exits.
    code = _code(repo_root, "Sh_Report_Error")
    assert any(re.search(r"If\s+Not\s+Sh_Log_Newest_First\(", c, re.IGNORECASE) for c in code), (
        "Sh_Report_Error must use Sh_Log_Newest_First's answer to tell whether the log was written.")
