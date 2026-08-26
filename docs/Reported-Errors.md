# Reported errors, and what fixed them

**Check this file before investigating any reported error.** It exists so that a fault already
diagnosed is never diagnosed twice — and so a report from someone running an older build can be
answered without reading a line of code.

Jerry's call, 8/26/2026, in his words: *"I don't want to spend time and tokens on
re-investigating errors that have already been fixed."*

## How a report arrives

From 3.0.261 a failed ribbon or toolbar button writes one line into
`%AppData%\VistaType LP Settings\VistaType-Errors.log` on the user's machine and shows them
dialog **240**, which asks them to send it. A line looks like this:

```
2026-08-26 13:42:47  v3.0.261  Lp_Video_Links  err 5941 "The requested member of the collection
does not exist."  step "Setting table alternating color style"  Word 16.0  doc "Document1"
cfg DEF  tmpl "Normal.dotm"  8.5x11 P  Tahoma 12
```

Everything the table below is keyed on comes off that one line: the **version**, the **error
number**, the **macro** and the **step**.

## How to use it

1. Take the version, error number, macro and step from the reported line.
2. Look for a row matching the error number **and** the macro.
3. **A match is strong evidence, not proof.** VBA gives no call stack, so the macro named is the
   one the *button* ran — the fault may live in something it called, and the same error number
   in the same macro can have two causes. On a match, confirm against that sub's own
   `' Version:` block before answering.
4. Compare their version against the two columns:

| | means | what to tell them |
|---|---|---|
| their version **≥ Shipped in** | they already have the fix | it is a different fault — investigate |
| their version **< Shipped in**, and Shipped in is filled | fixed and released | update to that release |
| **Fixed in** filled, **Shipped in** empty | fixed on `dev`, not released yet | known and fixed; it goes out with the next release |
| no row at all | not seen before | investigate, then add a row |

**Fixed in and Shipped in are different on purpose.** `3.0.X` numbers are private build
counters — nothing reaches a user until the next `X.Y` release. Telling someone to update to a
build that was never released sends them looking for something that does not exist.

## The standing rule

**A row is added in the same change as the fix, never afterwards.** A register that lags is
worse than none, because it says "not fixed" about something that is.

Record a fault here whether it arrived through the error log or Jerry found it himself — the
question this file answers is "has this already been dealt with", and where it came from does
not change that.

---

## The register

| Reported | Version | Error | Macro | Step | Cause | Fixed in | Shipped in |
|---|---|---|---|---|---|---|---|
| 8/26/2026, Jerry | 3.0.248 | none — wrong values, no error raised | `LP_Attach_An_Lp_Template_Form` (Attach LP Template dialog) | choosing a paper size, ticking Customize, ticking Mirrored | every media button cleared `CustomizeCheckBox` part-way down its own run; that assignment raises `CustomizeCheckBox_Click`, whose revert put `PTM/PBM/PLM/PRM` back from the previous media's `Hold_` values. Binding width and the margins the book was actually set with ran one media button behind | 3.0.249 | — |
| 8/26/2026, review | 3.0.249 | none — wrong page size, no error raised | `LP_Attach_An_Lp_Template_Form` (orientation handlers) | picking a tablet, ticking Customize, picking another tablet, picking the first again | both orientation handlers arranged the size boxes by testing `Hold_Orientation` (the orientation when Customize was last ticked) instead of the media's own. On a match they took the swap arm, and only paper had an `If IsPaper` arm below to put the numbers back — so a tablet reached Word with width and height transposed | 3.0.250 | — |
| 8/26/2026, review | 3.0.250 | any — swallowed, never seen | `Lp_Attach_The_Template` | anywhere after "Set Page Size and Orientation" | version 1.8 (10/24/2023) opened `On Error Resume Next` around the portrait-orientation line, for a real reason — it raises when the document starts with a drop cap — but never closed it. It stayed in force for the remaining 210 lines: the page size, the styles, the table banding, the TOC tab stops and a dozen `Application.Run` calls. Every failure in them was discarded and the user got a half-formatted book and no message | 3.0.251 | — |
| 8/26/2026, Jerry | 3.0.256 | 5941 and every other | all 47 ribbon and toolbar buttons | any failing macro | `RibbonAction` wrapped `Application.Run control.Tag` in `On Error GoTo`, but **Word does not pass an error back out of `Application.Run`** — it shows its own Run-time error dialog, which offers Debug, and the handler is never entered. Measured on the build box: four real button presses wrote a marker immediately before the call and never the one after it. Fixed by dispatching directly through the generated `Sh_Dispatch` | 3.0.257 | — |
| 8/26/2026, review | 3.0.262 | any — swallowed, never seen | `Lp_AutoTag_Page_Numbers` | anywhere after the roman-numeral loop | `On Error Resume Next` set inside that loop (`'prevents crash in a table`, right for the short block it guards) was never turned off, so it covered the remaining **339** lines: the `$pg` colouring pass, the paragraph mark taken off the top of the file, the screen-updating restore and the loop that counts the tags. A count that failed reported "There are no tagged page numbers in this document", which reads as a correct answer | 3.0.263 | — |
| 8/26/2026, review | 3.0.262 | any — swallowed, never seen | `Dx_AutoTag_Page_Numbers` | anywhere after the roman-numeral loop | the same fault on the braille side, covering the remaining **96** lines. Closed after the `LoopEnd:` label in both, which is the only placing that runs on every path — there are `GoTo LoopEnd` jumps earlier in each loop that would skip an `On Error GoTo 0` put at the end of the `If` block | 3.0.263 | — |

*(Nothing in this register has shipped yet — the whole of it is on `dev`. Fill "Shipped in" when
the release goes out.)*
