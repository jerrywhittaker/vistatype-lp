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

Also (issue #26): none of the attach's steps is started through Application.Run, which never
passes an error back to AttachFailed; each is called directly. See DIRECT_STEPS below.

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


# --- issue #26: the steps are called directly, not through Application.Run --------------------
#
# 10/9/2026. Word does not pass an error back out of Application.Run (the 3.0.256 row in
# docs/Reported-Errors.md): a failure inside a step run that way shows Word's own Run-time error
# dialog, and AttachFailed never sees it. These twelve were converted to direct calls. Each is a
# parameterless Public Sub in LPandBrlMacros, declared once.

DIRECT_STEPS = [
    "Lp_Remove_All_Styles_Except_Lp_Styles",
    "Lp_Fix_Abbyy_Text_and_Headers",
    "Lp_SetPicturesToInlineAndLockAspectRatio",
    "Lp_ResizePicturesAndShapesToFitPageWidthAndPageHeight",
    "Lp_Normalize_Styles",
    "Lp_Convert_Hyper_To_Addresses",
    "Lp_Set_TOC_and_Print_Page_Num_Tab_Stops",
    "Sh_Set_Prodnote_Style_Visibility",
    "Lp_Set_Display_For_Large_Print",
    "MS_Set_Word_Config_For_Large_Print",
    "MS_Clear_F_and_R_Params_and_Clipboard",
    "Lp_Turn_on_Styles_Pane",
    "Lp_Fix_Common_File_Errors",   # 9/28/2026, the first one done
]


def run_calls(lines, names):
    """The names in `names` that `lines` start through Application.Run (any spelling of it)."""
    found = []
    for code in lines:
        m = re.search(r"\bApplication\.Run\b(.*)$", code, re.IGNORECASE)
        if not m:
            continue
        for n in names:
            if re.search(r'"(\w+\.)?' + re.escape(n) + r'"', m.group(1), re.IGNORECASE):
                found.append(n)
    return found


def test_the_run_checker_catches_the_old_shapes():
    old = [
        '                .Application.Run MacroName:="Lp_Remove_All_Styles_Except_Lp_Styles"',
        '    Application.Run MacroName:="Lp_Normalize_Styles"',
        '    Application.Run "LPandBrlMacros.Lp_Turn_on_Styles_Pane"',
    ]
    assert run_calls(old, DIRECT_STEPS) == [
        "Lp_Remove_All_Styles_Except_Lp_Styles", "Lp_Normalize_Styles", "Lp_Turn_on_Styles_Pane"]
    assert run_calls(["    Lp_Normalize_Styles"], DIRECT_STEPS) == []


def test_attach_steps_are_not_run_through_application_run(repo_root):
    procs = procedures(module_text(repo_root))
    lines = procs[ATTACH]
    found = run_calls(lines, DIRECT_STEPS)
    assert found == [], (
        f"{ATTACH} calls {', '.join(sorted(set(found)))} through Application.Run, which never "
        f"passes an error back to AttachFailed - call it directly (issue #26)")
    for n in DIRECT_STEPS:
        assert any(re.match(r"^\s*(Call\s+)?" + re.escape(n) + r"\b", c) for c in lines), \
            f"{ATTACH} no longer calls {n} directly"


# --- issue #26: the error handlers clear the saved cursor position ----------------------------
#
# 10/9/2026, found by a review of #26, by reading. Sh_Save_User_Position / Sh_Return_User_To_
# Start_Position keep their state in two Publics, Sh_Pos_Depth and Sh_Pos_Saved, and only
# RibbonAction cleared them. A step that saved the position and then failed - File Cleanup,
# Hyperlinks to Addresses - now lands in a handler instead of Word's dialog and End (which wiped
# the Publics), so the state stayed at Depth 1 / Saved True, and the next macro started from a
# shortcut or a form button thought it was nested and never put the cursor back. Each handler
# must clear both, the same two lines RibbonAction uses, before it reports.

RESETS = {"Sh_Pos_Depth = 0": r"^\s*Sh_Pos_Depth\s*=\s*0\s*$",
          "Sh_Pos_Saved = False": r"^\s*Sh_Pos_Saved\s*=\s*False\s*$"}

# (where, sub, handler label, the name it reports under)
POSITION_HANDLERS = [
    ("src/vba/LPandBrlMacros.bas", ATTACH, "AttachFailed", "Lp_Attach_The_Template"),
    ("src/forms/Lp_File_Cleanup_Sub_Menu_Form.frm", "OkayButton_Click", "CleanupFailed",
     "Lp_Fix_Common_File_Errors"),
]


def handler_resets(lines, label):
    """What is missing from the handler `label` in `lines`: each reset must come before its
    Sh_Report_Error. Returns a list of problems, empty when both are there."""
    start = next((i for i, c in enumerate(lines) if re.match(r"^" + label + r":\s*$", c)), None)
    if start is None:
        return [f"no {label}: label"]
    body = lines[start + 1:]
    report = next((i for i, c in enumerate(body) if re.match(r"^\s*Sh_Report_Error\b", c)),
                  len(body))
    problems = []
    for line, pat in RESETS.items():
        if not any(re.match(pat, c) for c in body[:report]):
            problems.append(f"{label} does not run {line} before Sh_Report_Error")
    return problems


def test_the_reset_checker_catches_the_old_shape():
    old = ["AttachFailed:", "    failNumber = Err.Number", "    Sh_Progress_Close",
           '    Sh_Report_Error "x", failNumber, failText', "    Sh_Pos_Depth = 0"]
    assert len(handler_resets(old, "AttachFailed")) == 2
    good = ["AttachFailed:", "    failNumber = Err.Number", "    Sh_Pos_Depth = 0",
            "    Sh_Pos_Saved = False", '    Sh_Report_Error "x", failNumber, failText']
    assert handler_resets(good, "AttachFailed") == []


def test_ribbon_action_still_uses_the_same_two_lines(repo_root):
    text = (repo_root / "src/vba/RibbonCallbacks.bas").read_bytes().decode("latin-1")
    lines = procedures(text)["RibbonAction"]
    for line, pat in RESETS.items():
        assert any(re.match(pat, c) for c in lines), f"RibbonAction no longer has {line}"


def test_error_handlers_clear_the_saved_position(repo_root):
    for where, sub, label, reported in POSITION_HANDLERS:
        text = (repo_root / where).read_bytes().decode("latin-1")
        procs = procedures(text)
        assert sub in procs, f"{sub} is missing from {where}"
        lines = procs[sub]
        assert any(re.match(r'^\s*Sh_Report_Error\s+"' + reported + '"', c) for c in lines), \
            f"{where} {sub}: {label} no longer reports as {reported}"
        problems = handler_resets(lines, label)
        assert problems == [], f"{where} {sub}: " + "; ".join(problems) + " (issue #26)"
