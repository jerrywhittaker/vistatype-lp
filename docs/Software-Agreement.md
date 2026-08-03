# Software Agreement text (for the About dialogs)

This is the canonical wording for the **Software Agreement** shown in the
`Lp_About_Title_And_Agreement` and `Dx_About_Title_And_Agreement` UserForms, updated to
match the project's GPLv3 license. It replaces the older permissive wording ("…at no cost
to others…").

Keeping the text here (in git) means it can be reviewed and updated without editing the
form's binary layout — see **Implementation** below.

---

## Agreement text

> **Software License**
>
> VistaType LP is **free software**: you may run, copy, study, share, and modify it under
> the terms of the **GNU General Public License, version 3 (GPLv3)**, as published by the
> Free Software Foundation.
>
> **Your freedoms.** You may use the Software for any purpose, examine how it works, and
> redistribute copies. You may modify it and distribute your modified versions — but those
> versions must also be licensed under the GPLv3, and you must make their source code
> available. This notice and the license must be included with all copies.
>
> **No warranty.** This program is distributed in the hope that it will be useful, but
> WITHOUT ANY WARRANTY — without even the implied warranty of MERCHANTABILITY or FITNESS
> FOR A PARTICULAR PURPOSE. The author makes no guarantees regarding its performance,
> reliability, or suitability for any task, and is not liable for any loss or damage
> arising from its use.
>
> **Full terms.** A copy of the GNU General Public License is included with the Software
> (the `LICENSE` file). If you did not receive it, see <https://www.gnu.org/licenses/>.
>
> By installing, running, or otherwise using this Software you acknowledge these terms.
>
> © 2015–2026 Jerry Whittaker · jerry@thewhittakers.org

---

## Implementation — **DONE 8/2/2026**

Both forms now carry this wording. The steps below are kept as the record of what was done.

What actually shipped differs from the plan in two ways, both improvements:

- The wording is **not** duplicated into each form's `UserForm_Initialize`. It lives once in
  `Sh_Software_Agreement_Text` in `LPandBrlMacros`, and both dialogs call it — so the LP and
  Braille versions cannot drift apart. Edit the function, and this file, together.
- Each dialog gained a **View Full License** button (`Sh_Show_Full_License`) that opens the
  complete GPL the installer writes to `%AppData%\VistaType LP\LICENSE.txt`, falling back to
  a pointer at gnu.org. So the summary is what you read, and the licence itself is one click
  away.

The forms were reworked headlessly on the build box via the VBA object model and exported —
**not** with `make pull`, which would have wiped `src/vba` and `src/forms` and churned all 48
`.frx` files. The version label was left strictly alone: `Import-Vba.ps1` finds it by matching
the caption *text*, so disturbing it would silently stop every future build stamping the
version.

### The original plan (for reference)

The text above is longer than the old labels, so **make it scroll** rather than growing
the dialog. A `Label` cannot scroll; use a read-only multiline text box.

1. On each About form, delete the separate agreement labels
   (`SoftwareAgreementLabel`, `PermissionGrantLabel`, `DisclaimerLabel`,
   `ResponsibilityLabel`, `TerminationLabel`, `BindingAgreementLabel` and their text
   labels). **Keep** the title, copyright, version, and the View/Exit buttons.
2. Add one **TextBox** where the agreement was, sized to the available space, with:
   - `MultiLine = True`
   - `WordWrap = True`
   - `ScrollBars = 2` (`fmScrollBarsVertical`)
   - `Locked = True` (read-only, but still scrollable/selectable)
   - `TabStop = False`, a sunken `SpecialEffect` if you like the boxed look
   - Name it `AgreementTextBox`.
3. Fill it from code in `UserForm_Initialize`, so the wording stays in git-tracked source
   (no `.frx` edits needed to change it later):

```vba
Private Sub UserForm_Initialize()
    ' ...existing dual-monitor centering code stays...

    Dim s As String
    s = "Software License" & vbCrLf & vbCrLf
    s = s & "VistaType LP is free software: you may run, copy, study, share, and " _
          & "modify it under the terms of the GNU General Public License, version 3 " _
          & "(GPLv3), as published by the Free Software Foundation." & vbCrLf & vbCrLf
    s = s & "Your freedoms. You may use the Software for any purpose, examine how it " _
          & "works, and redistribute copies. You may modify it and distribute your " _
          & "modified versions - but those versions must also be licensed under the " _
          & "GPLv3, and you must make their source code available. This notice and the " _
          & "license must be included with all copies." & vbCrLf & vbCrLf
    s = s & "No warranty. This program is distributed in the hope that it will be " _
          & "useful, but WITHOUT ANY WARRANTY - without even the implied warranty of " _
          & "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. The author makes no " _
          & "guarantees regarding its performance, reliability, or suitability for any " _
          & "task, and is not liable for any loss or damage arising from its use." _
          & vbCrLf & vbCrLf
    s = s & "Full terms. A copy of the GNU General Public License is included with the " _
          & "Software (the LICENSE file). If you did not receive it, see " _
          & "https://www.gnu.org/licenses/." & vbCrLf & vbCrLf
    s = s & "By installing, running, or otherwise using this Software you acknowledge " _
          & "these terms." & vbCrLf & vbCrLf
    s = s & Chr(169) & " 2015-2026 Jerry Whittaker  -  jerry@thewhittakers.org"

    AgreementTextBox.Text = s
    AgreementTextBox.CurLine = 0   ' show the top, not the end
End Sub
```

4. Do the same on both `Lp_` and `Dx_` About forms (the Dx title line stays "…for the
   Duxbury Braille Translator BANA Braille Templates").
5. After the change, `make pull` to bring the updated forms into `src/forms/`, then
   `make build` and smoke-test.
