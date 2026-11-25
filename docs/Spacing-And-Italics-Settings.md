# Spacing and Italics Settings (LP, dialog 424)

Three choices a transcriber can make for a large print book, from the **Spacing and Italics
Settings** button (Alt+T) on Attach LP Template & Select Output Media (338). They were set by
Jerry on 11/24/2026 and first built in 3.0.717.

| Setting | Choices | Built-in default |
|---|---|---|
| Line spacing | Multiple 1.00 to 2.00, in steps of 0.05 | 1.25 (1.5 × the type size in Tahoma) |
| Letter spacing | 0.00 to 0.25 × the type size, in steps of 0.01 | 0.12 (WCAG 1.4.12) |
| Italics shown as | Dashed underline, Blue, Green, Orange, Red, Violet | Dashed underline |

Each row has a **Reset** button that puts back the built-in default. At the bottom, **This
document only** or **Save as my default** decides where the choice is kept.

## Where the values live

- **The book's own settings:** document variables `LpLineMultiple`, `LpLetterSpacing` and
  `LpItalicsMarker`, written when the template is attached. A reopened book keeps them.
- **The user's saved defaults:** `%AppData%\VistaType LP Settings\VistaType.ini`, section
  `[LpBookDefaults]`. The uninstaller leaves this file alone.
- **Built-in:** the constants `LP_LINE_MULTIPLE`, `LP_LETTER_SPACING` and `LP_ITALICS_DEFAULT`
  in `LPandBrlMacros.bas`.

## Which value is used (Jerry's rules, 11/24/2026)

- **The book's own settings win** on a re-attach.
- **Spacing:** the book's own value, then the saved default, then the built-in.
- **Italics:** the book's own marker, otherwise **Dash**. The saved default marker applies only
  to a **new** book on its first attach (`Lp_Pref_Attach_Marker`). A book made before this
  feature counts as Dash until someone chooses otherwise for it in 424.
- **Importing a file** (`Lp_Import_Exported_Selection_File`) gives the imported text this
  book's three settings, and converts italics the file had marked another way.

## The color markers

- A color marker uses its own character styles, **Italics Blue / Green / Orange / Red /
  Violet**, each based on the matching **Text** color style, so it looks the same. The
  add-in creates them in the book when they're first needed, and re-attach keeps them
  (`Lp_Remove_All_Styles_Except_Lp_Styles`).
- Changing the marker converts what was marked before. Going from Dash converts **every**
  dash-underlined run, including any the publisher put there (Jerry accepted this). Going
  from a color converts **only** that Italics style, so text colored Text Red by hand stays
  red.
- Text in another character style is left as it was and stays dashed: a footnote number, a
  TOC color bar (Words …), a hyperlink. A word that mixes them, like a word before a footnote
  number or one with a superscript inside it, is split, so the plain letters get the color
  and the rest keeps its own look.
- If an italic phrase had its own color, the marker's color replaces it, and going back to
  Dash makes it automatic. The original color isn't brought back.
- The conversion is a Find that steps through the text, with no Replace All and no
  `DoEvents` (issue #56). The book's `LpItalicsMarker` is written only after the conversion.

## Measured

- **Color markers:** on vistabuild with 3.0.716 and 3.0.717, plain, bold and superscript
  italics turn into the Italics style with bold and superscript kept. Footnote references,
  Words bars, hyperlinks and hand-applied Text Red are left alone. Back to Dash and on to Red
  convert exactly as above. Each pass took well under a second on the sample.
- **Line spacing:** Multiple 1.25 at 18 to 28 pt gives 1.51 × the type size.

## Not yet measured

- Speed on a large book with many italics.
