"""Attach LP Template catches errors all the way through (issue #20).

10/9/2026. Found by reading, 9/28/2026. Lp_Attach_The_Template sets On Error GoTo AttachFailed
near its top, but two places switched catching OFF again:

  1. The style-hiding loop was guarded by On Error GoTo AvoidCrash, and the AvoidCrash: label
     was followed by On Error GoTo 0. That left the rest of the sub - attaching the template, fonts,
     margins, pictures, Normalize Styles and the Save As - with no handler, so a failure showed
     Word's own Run-time error dialog over a frozen progress bar and logged nothing. The jump
     also fell through the label with no Resume.
  2. The portrait orientation line's On Error Resume Next was closed with On Error GoTo 0.

What this checks, in the body of Lp_Attach_The_Template (comments taken off), from the first
On Error GoTo AttachFailed down to the AttachFailed: label:

  - no On Error GoTo 0 at all, and no On Error GoTo to any label but AttachFailed;
  - no AvoidCrash label;
  - every On Error Resume Next is closed by On Error GoTo AttachFailed before the handler;
  - On Error GoTo AttachFailed is in force on the .AttachedTemplate = line and at the end.

Not unit-testable in VBA: forcing an error mid-attach needs a real document and template on the
build box. So this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

ATTACH = "Lp_Attach_The_Template"

ON_ERROR = re.compile(r"^\s*On\s+Error\s+(GoTo\s+(\w+)|Resume\s+Next)\s*$", re.IGNORECASE)
HANDLER_LABEL = re.compile(r"^AttachFailed:\s*$")


def handler_states(lines):
    """Walk the body and give, for each line, which On Error is in force when it runs.

    Returns (states, problems): states[i] is "AttachFailed", "Resume Next", "0", another label,
    or None before the first On Error. problems lists anything the rules above forbid.
    """
    states = []
    problems = []
    state = None
    seen_handler = False
    for i, code in enumerate(lines):
        if HANDLER_LABEL.match(code):
            break
        if re.match(r"^\s*AvoidCrash:", code, re.IGNORECASE):
            problems.append(f"line {i}: an AvoidCrash label is still there")
        m = ON_ERROR.match(code)
        if m:
            if m.group(1).lower().startswith("resume"):
                state = "Resume Next"
            else:
                label = m.group(2)
                if label.lower() == "attachfailed":
                    state = "AttachFailed"
                    seen_handler = True
                elif seen_handler:
                    state = label
                    problems.append(f"line {i}: On Error GoTo {label} after AttachFailed was set")
                else:
                    state = label
        states.append(state)
    if not seen_handler:
        problems.append("On Error GoTo AttachFailed is never set")
    elif state != "AttachFailed":
        problems.append(f"the body ends with On Error {state} in force, not AttachFailed")
    return states, problems


def test_the_checker_catches_the_old_shapes():
    old = [
        "Sub Lp_Attach_The_Template()",
        "    On Error GoTo AttachFailed",
        "    On Error GoTo AvoidCrash",
        "        .Styles(x).Visibility = True",
        "AvoidCrash:",
        "    On Error GoTo 0",
        "        .AttachedTemplate = TemplatePathandName",
        "    Exit Sub",
        "AttachFailed:",
    ]
    _, problems = handler_states(old)
    assert any("AvoidCrash" in p for p in problems)
    assert any("GoTo 0" in p for p in problems)

    portrait = [
        "    On Error GoTo AttachFailed",
        "    On Error Resume Next",
        "    Selection.PageSetup.Orientation = wdOrientPortrait",
        "    On Error GoTo 0",
        "    Exit Sub",
        "AttachFailed:",
    ]
    _, problems = handler_states(portrait)
    assert problems, "a Resume Next closed with GoTo 0 must be caught"

    good = [
        "    On Error GoTo AttachFailed",
        "    On Error Resume Next",
        "    Selection.PageSetup.Orientation = wdOrientPortrait",
        "    Err.Clear",
        "    On Error GoTo AttachFailed",
        "    Exit Sub",
        "AttachFailed:",
    ]
    assert handler_states(good)[1] == []


def test_attach_failed_is_in_force_all_the_way_through(repo_root):
    procs = procedures(module_text(repo_root))
    assert ATTACH in procs, f"{ATTACH} is missing from LPandBrlMacros.bas"
    lines = procs[ATTACH]
    assert any(HANDLER_LABEL.match(c) for c in lines), f"{ATTACH} has lost its AttachFailed: label"

    states, problems = handler_states(lines)
    assert problems == [], f"{ATTACH}: " + "; ".join(problems)

    attach = next((i for i, c in enumerate(lines) if re.search(r"\.AttachedTemplate\s*=", c)), None)
    assert attach is not None, f"{ATTACH} no longer sets .AttachedTemplate"
    assert attach < len(states) and states[attach] == "AttachFailed", \
        f"On Error GoTo AttachFailed must be in force on the .AttachedTemplate line, " \
        f"not {states[attach] if attach < len(states) else 'nothing'}"
