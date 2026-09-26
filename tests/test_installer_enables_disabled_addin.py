"""The installer takes LPandBRL.dotm back off Word's Disabled Items list.

9/26/2026, 3.0.509 on the build box. The box was reset with Word open and the add-in loaded.
On its next start Word asked "Word is running into problems with the lpandbrl.dotm add-in. Do
you want to disable it now?", the answer was Yes, and the add-in went onto
HKCU\\Software\\Microsoft\\Office\\16.0\\Word\\Resiliency\\DisabledItems. The VistaType tabs
still showed, because they live in the user's Word.officeUI, but every button in them points
into the add-in, so the tabs were empty. Reinstalling did not cure it: the entry outlives the
file, and each time Word saved its ribbon settings with the add-in still off it stripped the
VistaType buttons out of Word.officeUI.

EnableDisabledAddIn in installer/vistatype.iss now deletes that entry at ssPostInstall. It must
delete ONLY an entry naming lpandbrl.dotm - anything else on the list is another add-in's.

Not testable here: Inno's Pascal Script only runs inside a built Setup.exe on Windows. This is
a check on the source; that the entry really goes needs an install on the build box with one
planted.
"""
import re

ISS = "installer/vistatype.iss"
PROC = "EnableDisabledAddIn"


def iss_text(repo_root):
    return (repo_root / ISS).read_text(encoding="utf-8")


def pascal_proc(text, name):
    """The body of one procedure, from its header to the next top-level procedure/function."""
    m = re.search(rf"^procedure {name}\b.*?(?=^(?:procedure|function) )", text, re.S | re.M)
    assert m, f"{name} is missing from {ISS}"
    return m.group(0)


def test_the_routine_reads_words_disabled_items_list(repo_root):
    body = pascal_proc(iss_text(repo_root), PROC)
    assert r"\Word\Resiliency\DisabledItems" in body
    assert "RegGetValueNames(HKCU" in body
    assert "RegQueryBinaryValue(HKCU" in body


def test_it_deletes_only_an_entry_naming_the_addin(repo_root):
    body = pascal_proc(iss_text(repo_root), PROC)
    # The path inside the blob is UTF-16: the nulls must go before the name can be found.
    assert "StringChangeEx(Text, #0, ''" in body
    assert "Pos('lpandbrl.dotm', Lowercase(Text)) > 0" in body
    # Exactly one delete, and it sits inside that test - never an unconditional sweep.
    deletes = re.findall(r"RegDeleteValue\(", body)
    assert len(deletes) == 1
    assert body.index("Pos('lpandbrl.dotm'") < body.index("RegDeleteValue(")
    assert "RegDeleteKey" not in body


def test_it_runs_at_post_install(repo_root):
    text = iss_text(repo_root)
    step = pascal_proc(text, "CurStepChanged")
    post = step[step.index("CurStep = ssPostInstall"):]
    assert f"{PROC}();" in post


def test_setup_refuses_to_run_while_word_is_open(repo_root):
    """Why the order against the [Run] entries does not matter: Word is closed throughout."""
    text = iss_text(repo_root)
    m = re.search(r"^function InitializeSetup\(\): Boolean;\s*begin\s*Result := OfficeIsClear\(",
                  text, re.M)
    assert m, "InitializeSetup no longer starts by refusing to run while Word is open"
