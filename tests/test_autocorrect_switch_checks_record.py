"""The AutoCorrect switch checks which list Word really holds before it saves one (issue #21).

VistaType LP keeps three "Replace text as you type" lists - ordinary documents, large print,
braille - and Sh_AutoCorrect_Switch swaps them, recording in VistaType.ini which one is up.
Word stores its own list only on a CLEAN quit, so after Word ends any other way the record and
Word's list disagree. Measured on the build box, 3.0.517, 9/29/2026: the record said "DEF", Word
held the large print list, the switch trusted the record and did nothing, and the next large
print book wrote that list over AutoCorrect-DEF.txt. Every ordinary-only entry was lost for good.

The decision - which list is Word's closest to - is Sh_AC_Pick_Table, and it and the real switch
are tested in VBA by tests/vba/TestAutoCorrectTables.bas. That suite runs only on the build box
and only on `make installer`/`make vba-test`, so this pins the shape of the switch on every
build: the parts that, quietly taken out, bring the fault straight back.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

SWITCH = "Sh_AutoCorrect_Switch"


def the_procs(repo_root):
    procs = procedures(module_text(repo_root))
    for name in (SWITCH, "Sh_AC_Identify", "Sh_AC_Pick_Table", "Sh_AC_Remember"):
        assert name in procs, "%s is missing from LPandBrlMacros.bas" % name
    return procs


def first(lines, pattern):
    return next((i for i, c in enumerate(lines) if re.search(pattern, c, re.IGNORECASE)), None)


def test_the_no_op_exit_waits_for_the_check(repo_root):
    # The fault itself: "the record says this list is up, so do nothing" - on the first call of a
    # session, when the record can be wrong.
    for code in the_procs(repo_root)[SWITCH]:
        if re.search(r"loadedNow\s*=\s*cfgType.*Exit\s+Sub", code, re.IGNORECASE):
            assert re.search(r"\bSh_AC_Checked\b", code), (
                "Sh_AutoCorrect_Switch may skip the switch on the record's word only once this "
                "session has checked the record (Sh_AC_Checked): " + code.strip())


def test_the_first_call_asks_what_word_holds_before_saving(repo_root):
    lines = the_procs(repo_root)[SWITCH]
    identify = first(lines, r"If\s+Not\s+Sh_AC_Checked\s+Then\s+loadedNow\s*=\s*Sh_AC_Identify\(")
    save = first(lines, r"Sh_AC_Write\(\s*loadedNow\b")
    assert identify is not None, (
        "the first call of a session must set loadedNow from Sh_AC_Identify, not the record")
    assert save is not None and identify < save


def test_no_list_is_saved_into_a_file_it_cannot_be_told_belongs_to(repo_root):
    lines = the_procs(repo_root)[SWITCH]
    for i, code in enumerate(lines):
        if re.search(r"Sh_AC_Write\(\s*loadedNow\b", code, re.IGNORECASE):
            assert re.search(r"If\s+Len\(loadedNow\)\s*>\s*0\s+Then\s*$", lines[i - 1]), (
                "Word's list may be saved into loadedNow's file only when Sh_AC_Identify could "
                "name it (\"\" means it could not)")


def test_no_table_is_seeded_from_a_list_nobody_can_name(repo_root):
    lines = the_procs(repo_root)[SWITCH]
    seed = first(lines, r"Sh_AC_Write\(\s*cfgType\b")
    guard = first(lines, r"If\s+Len\(loadedNow\)\s*=\s*0\s+Then\s+Exit\s+Sub")
    assert seed is not None and guard is not None and guard < seed


def test_the_record_is_trusted_only_once_it_took(repo_root):
    # Sh_AC_Checked is set from Sh_AC_Remember's answer and from nothing else. A plain True would
    # trust a record that a read-only VistaType.ini refused.
    text = module_text(repo_root)
    sets = [c for c in re.findall(r"^[^'\r\n]*\bSh_AC_Checked\s*=\s*[^\r\n]*", text, re.MULTILINE)
            if not re.search(r"\bIf\b", c)]
    assert sets, "Sh_AC_Checked is never set"
    for c in sets:
        assert re.search(r"Sh_AC_Checked\s*=\s*Sh_AC_Remember\(", c), c.strip()


def test_a_tie_goes_to_the_record(repo_root):
    # LP and braille often hold the same list; only the record can tell them apart.
    lines = the_procs(repo_root)["Sh_AC_Pick_Table"]
    assert first(lines, r"If\s+recordedDiff\s*=\s*best\s+Then") is not None
