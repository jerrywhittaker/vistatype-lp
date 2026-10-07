"""A braille button on a book with no BANA template says so and stops.

Issue #8, Jerry's rule 10/6/2026. Every braille button ran Dx_Is_BANA_Template_Attached first,
and with no BANA template on the book that ran the whole attach - template list, translation
question, and an UndoClear - so pressing Dashes/Primes/Fractions could empty an hour's undo list
for a job that had nothing to do with attaching. Now the check says "The BANA template is not
attached." (394) and returns False, and the user presses Attach BANA Template.

Later the same day Jerry took the attach off the last three buttons too - AutoTag Ref Pages,
Validate $pg Tags and Manual Tag Ref Page - so the check never attaches anything, and the
attachIfMissing option that let them is gone. Prodnote to TN's own message 277 was retired into
394. Then, after testing 3.0.537, he ruled those three must run on a book with the Normal template
too, so they do not call the check at all. A book with no braille code recorded has no
BrailleType variable, and Manual Tag's direct read of it raised error 5825 on a roman numeral, so
they read it through Dx_Ensure_BrailleType(False) - the book's own setting, else the translation
table SWIFT recorded, asking and recording nothing - and "" is tagged the UEB way. (A first cut
read only the book's own setting, which tagged a SWIFT-attached EBAE book the UEB way; the
vba-review caught it before it shipped.) Delete Prodnotes, the Export/Import, DAISY/NIMAS/Text Tools and Help groups never called
the check, so they show no message either.

Not unit-testable in VBA: the check reads ActiveDocument, which hangs the headless runner
(tests/vba/README.md). So this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

CHECK = "Dx_Is_BANA_Template_Attached"
MESSAGE = 'Sh_Say "The BANA template is not attached.", "Braille Macros (394)"'
REF_PAGE_BUTTONS = ("Dx_AutoTag_Page_Numbers", "Dx_Ref_Pg_Number_Sequence_Menu",
                    "Dx_Manual_Tag_with_Dollar_pg")
PRODNOTE_TO_TN = "Dx_Change_Prodnotes_To_Transcriber_Notes"
NO_CHECK = ("Sh_Delete_Prodnote_Paragraphs", "Dx_ExportSelectionToNewFile",
            "Dx_Import_Exported_Selection_File", "DN_Menu_Starter",
            "DN_Remove_Para_Formatting_From_Text_Files", "Dx_Video_Links", "Dx_About")
CALL = re.compile(r"\b" + CHECK + r"\s*\(([^)]*)\)", re.IGNORECASE)


def _procs(repo_root):
    procs = procedures(module_text(repo_root))
    assert CHECK in procs, f"{CHECK} is missing from LPandBrlMacros.bas"
    return procs


def _calls(procs):
    """{macro: argument text} for every macro that calls the check."""
    found = {}
    for name, code in procs.items():
        if name == CHECK:
            continue
        for line in code:
            m = CALL.search(strip_comment(line))
            if m:
                found[name] = m.group(1).strip()
    return found


def test_the_message_and_no_parameter(repo_root):
    code = [strip_comment(c).strip() for c in _procs(repo_root)[CHECK]]
    assert re.search(r"Function\s+" + CHECK + r"\s*\(\s*\)", code[0], re.IGNORECASE), (
        f"{CHECK} takes no parameter - nothing may ask it to attach the template: {code[0]}")
    assert MESSAGE in code, f"{CHECK} must say exactly: {MESSAGE}"
    say = code.index(MESSAGE)
    assert code[say + 1] == "Exit Function", f"{CHECK}: the message must be followed by Exit Function."


def test_the_check_never_attaches(repo_root):
    code = [strip_comment(c) for c in _procs(repo_root)[CHECK]]
    attach = [c.strip() for c in code if re.search(r"\bDx_Attach_BANA_Template(_Run)?\b", c)]
    assert not attach, f"{CHECK} must never run the attach (issue #8): {attach}"


def test_every_button_uses_the_plain_check(repo_root):
    calls = _calls(_procs(repo_root))
    passing = {n: a for n, a in calls.items() if a}
    assert not passing, f"nothing may pass anything to {CHECK} (issue #8): {passing}"


def test_only_the_attach_button_reaches_the_attach(repo_root):
    procs = _procs(repo_root)
    callers = sorted(n for n, code in procs.items() if n != "Dx_Attach_BANA_Template_Run"
                     and any(re.search(r"\bDx_Attach_BANA_Template_Run\b", strip_comment(c))
                             for c in code))
    assert callers == ["Dx_Attach_BANA_Template"], (
        f"only the Attach BANA Template button may run the attach (issue #8): {callers}")


def test_prodnote_to_tn_uses_394_and_277_is_gone(repo_root):
    code = [strip_comment(c).strip() for c in _procs(repo_root)[PRODNOTE_TO_TN]]
    assert MESSAGE in code, f"{PRODNOTE_TO_TN} must say exactly: {MESSAGE}"
    used = [l.strip() for l in module_text(repo_root).splitlines() if "(277)" in strip_comment(l)]
    assert not used, f"message 277 was retired into 394 (issue #8): {used}"
    others = [l.strip() for l in module_text(repo_root).splitlines()
              if "(394)" in strip_comment(l) and MESSAGE not in l]
    assert not others, f"message 394 is used for something else too: {others}"


def test_no_check_on_the_buttons_jerry_left_out(repo_root):
    procs = _procs(repo_root)
    calls = _calls(procs)
    for name in NO_CHECK:
        assert name not in calls, (
            f"{name} calls {CHECK}; Jerry's rule, 10/6/2026, is no template message on it.")


OTHER_BRAILLE_BUTTONS = ("Dx_Format_Tagged_Page_Numbers", "Dx_Embed_Ref_Pg_No",
                         "Dx_UnEmbed_Ref_Pg_No", "Dx_Selected_File_CleanUp",
                         "Dx_Add_Color_To_Foreign_Language_Words", "Dx_Spelling_List",
                         "Dx_Format_Exercise_Lv_1_and_Lv_2", "Dx_Type_Dashes",
                         "Dx_File_Fix_Sequence", "Dx_Compress_Linear_Math")
DOC_OPEN = re.compile(r'Application\.Run\s+MacroName:="Sh_Is_Doc_Open"', re.IGNORECASE)
RAW_READ = re.compile(r'ActiveDocument\.Variables\(\s*"BrailleType"\s*\)', re.IGNORECASE)


def test_the_reference_page_buttons_run_without_the_template(repo_root):
    # Jerry, 10/6/2026, after testing 3.0.537: no template message on these three - they must
    # run on a book with the Normal template. A document must still be open.
    procs = _procs(repo_root)
    calls = _calls(procs)
    for name in REF_PAGE_BUTTONS:
        assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
        assert name not in calls, f"{name} calls {CHECK}; it must run without the BANA template."
        code = [strip_comment(c) for c in procs[name]]
        assert any(DOC_OPEN.search(c) for c in code), f"{name} must still run Sh_Is_Doc_Open."


def test_every_other_braille_button_still_checks(repo_root):
    calls = _calls(_procs(repo_root))
    missing = [n for n in OTHER_BRAILLE_BUTTONS if n not in calls]
    assert not missing, f"these braille buttons must still check for the template: {missing}"


def test_no_braille_code_recorded_reads_safely_as_ueb(repo_root):
    # A Normal-template book has no BrailleType variable; reading it directly raises 5825.
    procs = _procs(repo_root)
    for name in ("Dx_Manual_Tag_with_Dollar_pg", "Dx_AutoTag_Page_Numbers"):
        raw = [c.strip() for c in procs[name] if RAW_READ.search(strip_comment(c))]
        assert not raw, f"{name} reads BrailleType directly - error 5825 with none recorded: {raw}"
        assert any(re.search(r"Dx_Ensure_BrailleType\(\s*False\s*\)", strip_comment(c))
                   for c in procs[name]), (
            f"{name} must read the braille code through Dx_Ensure_BrailleType(False), which also "
            "reads what SWIFT recorded and asks nothing.")
    assert "Dx_Held_BrailleType" not in procs, (
        "Dx_Held_BrailleType read only the book's own setting and missed what SWIFT recorded.")
    auto = " ".join(strip_comment(c) for c in procs["Dx_AutoTag_Page_Numbers"])
    assert re.search(r'If\s+Not\s+\(\s*Dx_Ensure_BrailleType\(\s*False\s*\)\s+Like\s+"EBA\[TN\]"\s*\)\s+Then',
                     auto), (
        "AutoTag must tag every book the UEB way unless it is EBAE, the same test as Manual Tag.")
