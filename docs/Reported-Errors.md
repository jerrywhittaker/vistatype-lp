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
| 8/26/2026, Jerry | 3.0.264 | none — Word appears to hang, no error raised | `Dx_UnEmbed_Ref_Pg_No` | selecting something that is not a reference page number, or nothing to unembed | both early exits used `End`, which does **not** restore `Application.ScreenUpdating`. The screen was turned off eight lines above, so the user pressed Okay on the message and Word sat frozen. Now restores the screen and uses `Exit Sub` — provably equivalent, since the only caller is its own ribbon button — which also spares the module state `End` wipes (`gRibbon`, `Sh_ConfiguredAs`) | 3.0.265 | — |
| 8/26/2026, Jerry | 3.0.264 | none — Word appears to hang, no error raised | `LP_Picture_Alignment_Form` (`CmdOkay_Click`) | pressing Okay with nothing selected, or a picture that is not inline | the same fault: `End` after screen updating was turned off. The screen is restored before it now. `End` is deliberately KEPT here — this is a form's button handler, and `Exit Sub` would leave the dialog open and waiting, which is a decision about the dialog rather than a repair | 3.0.265 | — |
| 8/26/2026, Jerry | 3.0.265 | none — the list read wrongly, no error raised | `Sh_Copy_Ref_Pg_Tags_To_Temp_File` (Validate $pg Tags) | reading the validation list, braille side | the temp document was laid out in **2 columns** for braille and **3** for large print, so five separate paragraphs appeared on what looked like one line — only paragraph marks revealed it. That column width is also why a line was cut at 23 characters, and a braille entry (page number plus Duxbury code) runs to 37. Now one column, `SH_PGVAL_LINE_MAX` raised to 40. Applied to both sides: one constant, one builder | 3.0.266 | — |
| 8/26/2026, Jerry | 3.0.266 | none — Locate silently found nothing | `Sh_PgVal_TextOfCurrentTag` (Locate, in $pg validation) | pressing Locate on a braille entry | a tagged line is `$pg12-14[[*lec*]][[*i*]]14`, and the `[[…]]` Duxbury codes carry the **DBT Code** style, which is hidden. `Range.Text` returns hidden characters regardless of display, so Locate searched the book for a string that is not visible and matched nothing — as Word's own Navigation pane also fails to. Now takes the **leading visible run**, stopping at the first hidden character, which is contiguous in the book and is the tag plus page number. **Not cured:** the EBAE lower-roman shape `$pg[[*ii*]]v` puts the code ahead of the number, leaving only `$pg`, so Locate lands on the first tag. Curing that needs the list to remember where each entry came from rather than searching for its text | 3.0.267 | — |
| 8/26/2026, Jerry | 3.0.268 | none — reported as "not working", was **not a defect** | `Dx_Embed_Ref_Pg_No`, `Dx_UnEmbed_Ref_Pg_No` | pressing Embed or UnEmbed on a `$pg` page number | the numbers had not been through **Format $pg Tags**, so they carried none of the four reference-page styles and both macros correctly declined. Two things were investigated and DISPROVED along the way, recorded so nobody repeats them: `Selection.Style` **does** report a character style (measured on the build box — it takes it from the character to the LEFT of the insertion point, so inside the number or immediately to its right both work), and the detection code was never at fault. What changed is the message: a paragraph still showing a `$pg` tag is now told to run Format $pg Tags (241/243) instead of being told to select the number again (242/244) | 3.0.269 | — |
| 8/26/2026, Jerry | 3.0.270 | none — a dialog on every macro | `Dx_Is_BANA_Template_Attached` | any braille macro on a document prepared with Duxbury's SWIFT | SWIFT **does** attach the BANA template (`BANA Braille 2017.dot`, so the `InStr` test finds it) but does **not** set the `BrailleType` document variable — its only variable is `SWIFT.StyleMap.GUID`. The lookup raised, every braille macro jumped to Choose Translation, and nothing remembered the answer. The document had it all along: `DBTTemplate = "English (UEB) - BANA.dxt"` in its custom properties, which nothing read. Now derived and written once. **`PRE-UEB` is tested before `UEB`** — Duxbury names the EBAE tables `English (BANA Pre-UEB Textbook DE) - BANA.dxt`, which contains "UEB", so the other order would translate every EBAE book as UEB. SWIFT writes that property **on save, not on attach** (A/B tested), so a brand new unsaved document is still asked once | 3.0.272 | — |
| 8/26/2026, Jerry | 3.0.272 | none — refused a valid page number | `Dx_Manual_Tag_with_Dollar_pg` | typing `12` in its own paragraph and pressing Manual Tag Ref Page | it validated the line as a **roman numeral** and refused anything else, answering "This is not a valid roman numeral". Most page numbers are arabic, and the large print half of the same feature never validated at all. The roman test now only decides the EBAE `[[*ii*]]` code; anything holding a digit is tagged as it stands, and a line of prose gets dialog 245 | 3.0.273 | — |
| 8/26/2026, review | 3.0.272 | none — pages silently left untagged | `Sh_IsValidRomanNumeral`, used by **both** `Dx_AutoTag_Page_Numbers` and `Lp_AutoTag_Page_Numbers` | AutoTag Ref Pages on any book with roman front matter | a hand-typed list of 36 numerals that ran `I` to `X` then jumped to `XVI`. It accepted **28 of the first 100** — `xi`, `xii`, `xiii`, `xiv`, `xv`, `xxvi`–`xxix` were all rejected, which is exactly where front matter lives. Those pages were skipped in **braille and large print alike**, with nothing to say so, and the list is byte-identical in the repository's first commit. Replaced with a value round trip: read the numeral as a number, write it back, require a match. Now 100 of 100, with `IIII`, `VX`, `IC` still refused | 3.0.273 | — |
| 8/26/2026, Jerry | 3.0.272 | none — pages silently left untagged | both AutoTag macros (paragraph window) | roman page numbers of 388 or above | both loops skip any paragraph longer than **11** characters — 10 of text plus the paragraph mark — so `CCCLXXXVIII` (388) and every numeral above it was passed over even with a correct validator. 347 of the numbers below 4000 did not fit. Widened to **16**: `MMMDCCCLXXXVIII` (3888) is the longest below 4000 at 15 characters, so there is no gap from i to mmmcmxcix. Found because Jerry asked whether a wider range was possible | 3.0.273 | — |

*(Nothing in this register has shipped yet — the whole of it is on `dev`. Fill "Shipped in" when
the release goes out.)*
