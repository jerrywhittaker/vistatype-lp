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
> VistaType LP and Braille Macros is **free software**: you may run, copy, study, share, and modify it under
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
> **Full terms.** A complete copy of the GNU General Public License is installed with the
> Software. Click **View Full License** below to read it, or see <https://www.gnu.org/licenses/>.
>
> By installing, running, or otherwise using this Software you acknowledge these terms.
>
> © 2015–2026 Jerry Whittaker · jerry@vistatypelp.org

---

## Implementation — **DONE 8/2/2026**

Both forms now carry this wording. The steps below are kept as the record of what was done.

What actually shipped differs from the plan in two ways, both improvements:

- The wording is **not** duplicated into each form's `UserForm_Initialize`. It lives once in
  `Sh_Software_Agreement_Text` in `LPandBrlMacros`, and both dialogs call it — so the LP and
  Braille versions cannot drift apart. Edit the function, and this file, together.
- Each dialog gained a **View Full License** button (`Sh_Show_Full_License`) that opens the
  complete GPL the installer writes to `%AppData%\VistaType LP\LICENSE.txt`, falling back to
  a pointer at gnu.org. So the summary is what you read, and the license itself is one click
  away.

The forms were reworked headlessly on the build box via the VBA object model and exported —
**not** with `make pull`, which would have wiped `src/vba` and `src/forms` and churned all 48
`.frx` files. The version label was left strictly alone: `Import-Vba.ps1` finds it by matching
the caption *text*, so disturbing it would silently stop every future build stamping the
version.

### The original plan — removed 9/20/2026

A superseded plan stood here, ending with a step that said to run `make pull` — which this same
file says two sections above must never be used for form work. It described a route that was not
taken. Removed; it is in the history if it is ever wanted.
