"""Full File Cleanup always cleans the whole book.

Jerry's rule, 10/11/2026: Full File Cleanup cleans the whole book; only Selection Clean Up
honors a selection - "If the user selects the whole book, it is still a selection."

A dozen of the passes in both sequences work on the selection when there is one
(`Selection.Type = wdSelectionNormal`), and many passes leave text selected behind them, since a
Selection.Find that is not a replace-all selects what it finds. So text selected when the button
was pressed, or left selected by one pass, shrank every later pass to that text. The cure is in
the two step subs, Dx_Ffc_Step and Lp_Ffc_Step, which announce every pass on the progress bar
just before it runs: each collapses the selection first.

So this checks two things: that both step subs collapse, and that no pass in either sequence
runs without a step announced in front of it - a pass added later without one would get the
selection the pass before it left.

Neither sequence can be driven headlessly end to end (the configuration sub at the top of the
braille one hangs an invisible Word), so this holds the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

SEQUENCES = {
    "Dx_Fix_Common_File_Errors_Run": "Dx_Ffc_Step",
    "Lp_Fix_Common_File_Errors": "Lp_Ffc_Step",
}

# Calls that are not cleanup passes, so need no step in front of them.
NOT_A_PASS = re.compile(
    r"^(Sh_Save_User_Position|Sh_Return_User_To_Start_Position|Sh_Spin_DoEvents|Sh_Progress_\w+"
    r"|MS_Clear_F_and_R_Params_and_Clipboard)$")

COLLAPSE = re.compile(r"^\s*Selection\.Collapse\b")


def pass_called(line):
    """The name of the macro a line runs, or None."""
    m = re.search(r'Application\.Run\s+MacroName:="(\w+)"', line)
    if m:
        return m.group(1)
    m = re.match(r"^\s*(?:Call\s+)?((?:Lp|Dx|Sh)_\w+)\b(?!\s*=)", line)
    if m:
        return m.group(1)
    return None


def test_both_step_subs_collapse_the_selection(repo_root):
    procs = procedures(module_text(repo_root))
    for step in SEQUENCES.values():
        body = procs[step]
        collapse = next((n for n, c in enumerate(body) if COLLAPSE.match(c)), None)
        assert collapse is not None, (
            f"{step} no longer collapses the selection - a selection would shrink Full File "
            "Cleanup's passes")
        say = next(n for n, c in enumerate(body) if "Sh_Progress_Say" in c)
        assert collapse < say, f"{step} must collapse before it announces the pass"
        assert "wdCollapseStart" in body[collapse], f"{step} must collapse to the start"


def test_every_pass_has_a_step_in_front_of_it(repo_root):
    procs = procedures(module_text(repo_root))
    for sequence, step in SEQUENCES.items():
        announced = False
        passes = 0
        for line in procs[sequence][1:-1]:
            name = pass_called(line)
            if name is None:
                continue
            if name == step:
                announced = True
                continue
            if NOT_A_PASS.match(name):
                continue
            assert announced, (
                f"{sequence} runs {name} with no {step} in front of it, so it gets whatever "
                "selection the pass before it left")
            announced = False
            passes += 1
        assert passes > 30, f"{sequence}: only {passes} passes found - the check has lost its way"
