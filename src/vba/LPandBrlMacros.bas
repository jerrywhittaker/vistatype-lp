Attribute VB_Name = "LPandBrlMacros"
' VistaType LP - Large Print Template and Macros and Macros for the BANA template of the Duxbury Braille Translator
' Copyright (C) 2015-2026 Jerry Whittaker
' jerry@thewhittakers.org
'
' This program is free software: you can redistribute it and/or modify it under
' the terms of the GNU General Public License as published by the Free Software
' Foundation, either version 3 of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT ANY
' WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A
' PARTICULAR PURPOSE. See the GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License along with this
' program. If not, see <https://www.gnu.org/licenses/>.
'
'
' Released 7/19/2026 - Version 3.0 - performance pass (ScreenUpdating discipline, O(n) loops, DoEvents throttle), save-once/stabilize, idempotent config, QAT installer fix
' This code changed 2/22/2026 12:20 AM - Not Released - Fixes for new Version 2.2.3
'
' Notes:    - MS - 8/21/2026 - Two faults in Piece 4, both found by Jerry testing 3.0.214 and 3.0.215, and neither of them
'           - MS - 8/21/2026 - in the machinery itself. FIRST: Sh_Note_Book_Settings was called ABOVE the spelling and grammar
'           - MS - 8/21/2026 - writes that both books deliberately keep at the very bottom, so it recorded HER values as the
'           - MS - 8/21/2026 - book's for two settings in large print and six in braille. The next ordinary document then found
'           - MS - 8/21/2026 - them different, decided she had changed them, and wrote the BOOK's switched-off values into her
'           - MS - 8/21/2026 - preferences - the 8/18 fault rebuilt by the thing meant to prevent it. The call is now genuinely
'           - MS - 8/21/2026 - last in both subs. SECOND, and it is a decision rather than a slip: Jerry opened a large print
'           - MS - 8/21/2026 - document and found boxes checked he had never checked. They are four of the fifteen settings
'           - MS - 8/21/2026 - dropped from all three configurations on 8/18 as "identical everywhere" - which was true, and
'           - MS - 8/21/2026 - which also took them away from the BOOKS, leaving Word's own value standing inside one. Smart
'           - MS - 8/21/2026 - quotes and hyperlinks, both forms, are back in the two book configurations and in the tracked
'           - MS - 8/21/2026 - list, 27 -> 31, store stamp 2 -> 3. Not in the ordinary configuration: that is the whole
'           - MS - 8/21/2026 - difference between a book setting and a preference imposed on her Word. The other eleven stay
'           - MS - 8/21/2026 - hers, including "*bold* and _italic_ with real formatting", which is the one Jerry unchecked.
' Notes:    - MS - 8/20/2026 - PIECE 4 of the automatic-configuration plan, with the typing half of Piece 3 that depends on it.
'           - MS - 8/20/2026 - The twenty Options and AutoCorrect writes in MS_Set_Word_Config_For_New_Install are GONE. They set
'           - MS - 8/20/2026 - Word's factory value on every ordinary document open, which LOOKED like restoring her settings and
'           - MS - 8/20/2026 - was overwriting them: Jerry, testing 3.0.210, cleared every box in AutoCorrect, AutoFormat and
'           - MS - 8/20/2026 - AutoFormat As You Type, closed Word, came back and found them checked again. Nothing had failed to
'           - MS - 8/20/2026 - save; that sub was rewriting them. AutoFormatAsYouTypeApplyFirstIndents went too - no book ever
'           - MS - 8/20/2026 - touched it, so the write could only ever overwrite her own choice, the 8/18/2026 reasoning again.
'           - MS - 8/20/2026 - What puts them back is the LEDGER, and it now covers TWENTY-SEVEN settings rather than six. The list
'           - MS - 8/20/2026 - is not guesswork: Sh_Tracked_Settings is the union of what the two book configurations actually
'           - MS - 8/20/2026 - write, read out of those subs. Twenty of them the ordinary configuration used to rewrite, which is
'           - MS - 8/20/2026 - why nobody noticed; the other seven it never put back at all - the six spelling and grammar ones
'           - MS - 8/20/2026 - found on 8/18, plus ShowStylePreviews and RestrictLinkedStyles, switched on by the books and off by
'           - MS - 8/20/2026 - nothing since 2021. ADDING A SETTING TO EITHER BOOK CONFIGURATION MEANS ADDING IT TO THAT LIST.
'           - MS - 8/20/2026 - Telling "she changed it" from "a book set it" is the point of the piece, and it is done with a
'           - MS - 8/20/2026 - record rather than a guess. New Sh_Note_Book_Settings runs at the END of each book configuration
'           - MS - 8/20/2026 - and writes what is live into a second section, [BookApplied]. Reading the live values at the end is
'           - MS - 8/20/2026 - deliberate: every write in those subs is guarded with If <>, so a setting already holding the book's
'           - MS - 8/20/2026 - value is never written, and recording what was INTENDED would have to know about writes that never
'           - MS - 8/20/2026 - happened. Sh_Restore_Transcriber_Settings (3.0) then compares live against that record - anything
'           - MS - 8/20/2026 - that no longer matches, she changed while working, and it becomes hers. Jerry, 8/20/2026: she does
'           - MS - 8/20/2026 - far more in Word than large print and braille and does not stop being an ordinary user when she
'           - MS - 8/20/2026 - opens a book, so a change made INSIDE one is a real preference and follows her out to her letters.
'           - MS - 8/20/2026 - Three paths, in the order tested. Store written by an older build: throw it away whole rather than
'           - MS - 8/20/2026 - half-read it, and clear the book half FIRST or the save declines and this path repeats for ever.
'           - MS - 8/20/2026 - No book in force since we were last here: the live values ARE hers - this is what the old sub got
'           - MS - 8/20/2026 - wrong, restoring over the top of a change she had just made in one letter and opened another.
'           - MS - 8/20/2026 - A book in force: learn, restore, then clear the record so the next letter does not compare against
'           - MS - 8/20/2026 - a book long gone.
'           - MS - 8/20/2026 - The guard that makes it survive Word CLOSING is in Sh_Save_Transcriber_Settings: it declines while a
'           - MS - 8/20/2026 - book record is outstanding. Quit Word inside a book and Word saves the BOOK's settings as its own and
'           - MS - 8/20/2026 - loads them next session, where AutoExec runs before anything can be configured and would write them
'           - MS - 8/20/2026 - straight into her preferences, losing the real ones for good. The book record outlives the session in
'           - MS - 8/20/2026 - the same file, which is the only reason the question can be answered at all - and the first letter she
'           - MS - 8/20/2026 - opens next session then learns and restores exactly as if she had never quit.
'           - MS - 8/20/2026 - The store carries a version stamp now (VT_STORE_STAMP_NOW), so a file of the old six-setting shape is
'           - MS - 8/20/2026 - discarded rather than half-read. Bump it whenever Sh_Tracked_Settings changes.
' Notes:    - Sh - 8/20/2026 - the Styles pane needs TWO handles written as well, for the same reason the navigation pane did, and
'           - Sh - 8/20/2026 - it is worth writing the reason down once properly. Application.TaskPanes(wdTaskPaneFormatting) needs a
'           - Sh - 8/20/2026 - DOCUMENT WINDOW to hang the pane on, and when Word is empty it quietly does nothing at all. Jerry,
'           - Sh - 8/20/2026 - testing 3.0.209: closing the last document put the navigation pane down and left the Styles pane
'           - Sh - 8/20/2026 - sitting there - the identical fault to 3.0.207 one build earlier, and the identical cause.
'           - Sh - 8/20/2026 - CommandBars("Styles") is the other handle on the same pane and is application-wide. Measured on the
'           - Sh - 8/20/2026 - build box 8/20/2026: with NO document open, CommandBars("Navigation") and CommandBars("Styles") both
'           - Sh - 8/20/2026 - still answer and can be written, while TaskPanes cannot. New Sh_Set_Styles_Pane writes both, and
'           - Sh - 8/20/2026 - Sh_Tidy_Panes_After_Close (1.1) is its ONLY caller and the only place that needs it - everywhere else
'           - Sh - 8/20/2026 - the pane is set while configuring a document that is open in front of us, where TaskPanes alone is
'           - Sh - 8/20/2026 - right. Do not sweep the other sites into it.
'           - Sh - 8/20/2026 - The general lesson, since it has now cost two builds: a task pane in Word has a per-document handle
'           - Sh - 8/20/2026 - and an application-wide one, they are not interchangeable, and the difference only shows itself when
'           - Sh - 8/20/2026 - there is no document.
' Notes:    - Sh - 8/20/2026 - the panes now come down AFTER a close, not during one, and the two entries below about recording a
'           - Sh - 8/20/2026 - hide and undoing it are superseded. Jerry, testing 3.0.207: he cancelled at the "save your changes?"
'           - Sh - 8/20/2026 - prompt and had to put both panes back by hand. The idea was wrong, not the implementation.
'           - Sh - 8/20/2026 - DocumentBeforeClose fires on close ATTEMPTS and Word says nothing afterwards about which button she
'           - Sh - 8/20/2026 - pressed, so three versions running tried to hide first and undo it on the next activation. Stop
'           - Sh - 8/20/2026 - guessing and LOOK instead: Sh_HandleDocumentClosing (1.6) now hides nothing at all and simply asks
'           - Sh - 8/20/2026 - Application.OnTime for a tick one second out. New Sh_Tidy_Panes_After_Close is that tick. OnTime
'           - Sh - 8/20/2026 - cannot run while a modal dialog is up, which is the whole trick - it cannot fire until she has
'           - Sh - 8/20/2026 - answered the prompt, however long she takes, and by then the document has either gone or it has not.
'           - Sh - 8/20/2026 - Its test is Documents.count = 0, and that is deliberate rather than lazy. Word empty: put the panes
'           - Sh - 8/20/2026 - down, so the next document does not inherit them, which is what Jerry asked for. Anything still open:
'           - Sh - 8/20/2026 - do NOTHING - a cancelled close, or one of two documents closed, leaves a document in front of her
'           - Sh - 8/20/2026 - whose panes are the ones she has been working with, and a close elsewhere is no reason to take them.
'           - Sh - 8/20/2026 - A cancelled close now costs her nothing at all: the panes never move, so there is not even a flicker.
'           - Sh - 8/20/2026 - Gone with it: Sh_PaneHiddenForDoc, Sh_PaneWasStyles, Sh_PaneWasNav, and the recovery block in
'           - Sh - 8/20/2026 - Sh_HandleDocumentActivated (1.4). Do not bring any of them back.
' Notes:    - Sh - 8/20/2026 - the navigation pane has TWO halves and both have to be written. ActiveWindow.DocumentMap is a WINDOW
'           - Sh - 8/20/2026 - property and CommandBars("Navigation").Visible is application-wide. Writing DocumentMap alone works
'           - Sh - 8/20/2026 - while a document is being CONFIGURED, because the window is there and stays - which is why the
'           - Sh - 8/20/2026 - ordinary configuration looked right in 3.0.206 - but it does NOTHING during DocumentBeforeClose,
'           - Sh - 8/20/2026 - where the window it names is on its way out and Word's own remembered pane state is the thing left
'           - Sh - 8/20/2026 - standing. Jerry, testing 3.0.207: the Styles pane went and the navigation pane stayed.
'           - Sh - 8/20/2026 - Dx_Attach_BANA_Template has written both halves since 2021, with the On Error added in 2023 (2.3)
'           - Sh - 8/20/2026 - because the CommandBars line raises on some Word versions - so the answer was already in the file.
'           - Sh - 8/20/2026 - New Sh_Set_Navigation_Pane and Sh_Navigation_Pane_Showing put the pair in one place, and the reading
'           - Sh - 8/20/2026 - half asks the application-wide setting FIRST, since that is the half that survives a close and is
'           - Sh - 8/20/2026 - therefore the half worth handing back. Used by Sh_HandleDocumentClosing (1.5),
'           - Sh - 8/20/2026 - Sh_HandleDocumentActivated (1.3), MS_Set_Word_Config_For_New_Install and Lp_Attach_The_Template -
'           - Sh - 8/20/2026 - that last one had written the window half alone since long before any of this.
' Notes:    - Sh - 8/20/2026 - the panes come down with ANY document that closes, not only a large print one, and the navigation
'           - Sh - 8/20/2026 - pane comes down with the Styles pane. Jerry, testing 3.0.206: he opened both panes in a letter, saved,
'           - Sh - 8/20/2026 - and closed it, and both were still sitting there. A letter is DEFINED as having no panes, so leaving
'           - Sh - 8/20/2026 - them behind was the odd case rather than the tidy-up. Sh_HandleDocumentClosing is 1.4.
'           - Sh - 8/20/2026 - The trap this walks into is the one pane visibility always sets: it is WORD-WIDE, so with two letters
'           - Sh - 8/20/2026 - open, closing one takes the panes off the other. Two new Publics, Sh_PaneWasStyles and Sh_PaneWasNav,
'           - Sh - 8/20/2026 - note what was on screen BEFORE anything is hidden - once those two lines have run there is no way left
'           - Sh - 8/20/2026 - to ask - and Sh_HandleDocumentActivated 1.2 hands it back. What it hands back depends on what is on
'           - Sh - 8/20/2026 - screen, not on what closed: a large print book gets the Styles pane and no navigation pane because
'           - Sh - 8/20/2026 - that is what its configuration asks for; braille gets neither, which is what the close already left;
'           - Sh - 8/20/2026 - and an ORDINARY document gets back exactly what was up. That last case is the point of recording it.
'           - Sh - 8/20/2026 - Her letter's panes are hers - she opened them by hand, and the ordinary configuration does not touch
'           - Sh - 8/20/2026 - them again once the document is open - so a close that happened to some OTHER document must not be
'           - Sh - 8/20/2026 - what takes them away. It also still covers the case the recording was invented for on the same day:
'           - Sh - 8/20/2026 - a close the transcriber CANCELLED at the "save your changes?" prompt, where she never left the
'           - Sh - 8/20/2026 - document at all. Closing the LAST document raises no activation, so nothing is handed back and Word
'           - Sh - 8/20/2026 - is left tidy, which is what was asked for.
' Notes:    - MS - 8/20/2026 - the SCREEN half of PIECE 3 of the automatic-configuration plan. What a plain letter looks like when it
'           - MS - 8/20/2026 - opens is now stated: no formatting marks, no Styles pane, no navigation pane, both rulers, print view.
'           - MS - 8/20/2026 - Two of those five changed. ShowAll was forced ON here and is now OFF - Jerry, after testing 3.0.205:
'           - MS - 8/20/2026 - forcing pilcrows on meant a transcriber who works without them switched them off again on EVERY letter
'           - MS - 8/20/2026 - she opened, not once. The rulers are the opposite case and stay forced, the vertical one especially,
'           - MS - 8/20/2026 - since it is buried in File > Options > Advanced > Display. And ActiveWindow.DocumentMap = False is new,
'           - MS - 8/20/2026 - closing the navigation pane; it is a WINDOW property so it cannot reach another document, unlike the
'           - MS - 8/20/2026 - other half of the navigation pane that Dx_Attach_BANA_Template writes - CommandBars("Navigation") is
'           - MS - 8/20/2026 - application-wide, and bringing the two together is still to do.
'           - MS - 8/20/2026 - Also gone from the ordinary configuration, and not to be put back: FormattingShowNextLevel,
'           - MS - 8/20/2026 - StyleSortMethod and FormattingShowFilter. All three are DOCUMENT properties saved inside her file, so
'           - MS - 8/20/2026 - writing them on every ordinary document open overwrote what that document was carrying - the 7/24/2026
'           - MS - 8/20/2026 - complaint again, in her letters this time rather than her books. Decision 6 of the plan.
'           - MS - 8/20/2026 - The TYPING half of Piece 3 is deliberately NOT in this: the twenty Options and AutoCorrect writes still
'           - MS - 8/20/2026 - stand, because nothing can replace them until the settings ledger of Piece 4 exists. Until then,
'           - MS - 8/20/2026 - clearing an AutoFormat box and reopening the document still puts it back, and that is not a saving
'           - MS - 8/20/2026 - fault - it is this sub rewriting it, exactly as it always has.
' Notes:    - LP - 8/20/2026 - PIECE 1 of the automatic-configuration plan (docs/Automatic-Configuration-Plan.md, agreed with the beta
'           - LP - 8/20/2026 - tester the same day): document type is read from the ATTACHED TEMPLATE and not from the style "Box Black".
'           - LP - 8/20/2026 - Jerry: the style test was written before he knew how to test for an attached template. The function that
'           - LP - 8/20/2026 - did it is named Lp_Is_The_Attached_Template_LP, so the name was right all along and only the implementation
'           - LP - 8/20/2026 - was the stopgap - it is now 2.0 and asks AttachedTemplate.Name against the new LP_TEMPLATE_FILE constant,
'           - LP - 8/20/2026 - by name and never by path, since the Templates folder differs between machines and Office versions.
'           - LP - 8/20/2026 - The style test is KEPT, as new Lp_Was_Made_As_An_Lp_Book, because it answers a different question that is
'           - LP - 8/20/2026 - still needed: not "is this a large print document" but "was this ever MADE as a large print book". The
'           - LP - 8/20/2026 - difference is a real document - a book on an OBSOLETE template still has Box Black, but the template on it
'           - LP - 8/20/2026 - is not ours. Test that book the new way and it reads as an ordinary document, is configured as one, and is
'           - LP - 8/20/2026 - never warned that it must be re-attached before she edits it. Of the eight callers, five wanted the type
'           - LP - 8/20/2026 - question and were left alone; three wanted the book question and were moved: Sh_HandleDocumentOpened's
'           - LP - 8/20/2026 - obsolete-template warning, Lp_Attach_Lp_Template's re-attach warning (a book on an old template is exactly
'           - LP - 8/20/2026 - the one that needs telling its text will re-flow), and LP_Attach_An_Lp_Template_Form's Initialize, which
'           - LP - 8/20/2026 - notes the face a legacy book is already set in - the FACE must not be rewritten whatever template it came on.
'           - LP - 8/20/2026 - Sh_HandleDocumentOpened's two tests became If/ElseIf rather than nested, because they are now independent
'           - LP - 8/20/2026 - questions rather than one refining the other. That also closes a hole the old shape had: a document carrying
'           - LP - 8/20/2026 - our template but WITHOUT the style could not be seen at all, and was configured as an ordinary document.
' Notes:    - LP - 8/20/2026 - the Styles pane follows the large print document now: open one and the pane is there, close one
'           - LP - 8/20/2026 - and it goes with it. Jerry, after attaching the template to a book, saving it, opening it again and
'           - LP - 8/20/2026 - finding no pane. This is a deliberate change to the 8/2/2026 rule that made the pane the user's own,
'           - LP - 8/20/2026 - and only half of it moves. The pane's VISIBILITY is what follows the document. The Recommended sort
'           - LP - 8/20/2026 - and the "Select styles to show" filter are NOT forced on open and Lp_Turn_on_Styles_Pane is NOT put
'           - LP - 8/20/2026 - back in either handler - those two are DOCUMENT properties saved into the file, so a book the attach
'           - LP - 8/20/2026 - set to Recommended already reopens that way, and a transcriber who chose "All Styles" in her own book
'           - LP - 8/20/2026 - keeps it. Forcing those on every open was the 7/24 and 8/2/2026 complaint and it stays fixed.
'           - LP - 8/20/2026 - Opening: MS_Set_Word_Config_For_Large_Print, inside its skip-display guard, which makes it the mirror
'           - LP - 8/20/2026 - of what the other two configurations have always done - braille and default both HIDE the pane there,
'           - LP - 8/20/2026 - and large print was the one of the three that said nothing, so it showed whatever the last document
'           - LP - 8/20/2026 - left behind. Being inside that guard also means merely SWITCHING to another book cannot move it.
'           - LP - 8/20/2026 - Closing: Sh_HandleDocumentClosing hides it again, which 8/2/2026 removed as a plain bug - and that bug
'           - LP - 8/20/2026 - was real, so the line is not simply back. DocumentBeforeClose fires on close ATTEMPTS and Word raises
'           - LP - 8/20/2026 - it BEFORE the "save your changes?" prompt, so X then Cancel took the pane and nothing put it back. The
'           - LP - 8/20/2026 - Cancel flag the event passes is no help: it says what the handler wants, not what she is about to
'           - LP - 8/20/2026 - answer. So the hide is RECORDED in new Public Sh_PaneHiddenForDoc, and Sh_HandleDocumentActivated (1.1)
'           - LP - 8/20/2026 - reconsiders it the next time a document is looked at - which a cancelled close makes immediate, since
'           - LP - 8/20/2026 - she never left the book. It tests what is ON SCREEN rather than whether the named document survived,
'           - LP - 8/20/2026 - which covers the second case too: close one of two open large print books and the pane belongs to the
'           - LP - 8/20/2026 - other one. That test sits BEFORE the handler's same-document early exit, or the cancelled close - same
'           - LP - 8/20/2026 - document, same configuration - would walk straight past it. Pane visibility is WORD-WIDE, which is why
'           - LP - 8/20/2026 - the recovery matters: the hide does not take the pane off one window, it takes it off all of them.
' Notes:    - Sh - 8/20/2026 - new Sh_Close_And_Reopen, called by Lp_Attach_The_Template right after the Save As. Word does not
'           - Sh - 8/20/2026 - repaint the Quick Style gallery on the Home tab when the attached template changes which styles belong
'           - Sh - 8/20/2026 - in it, and closing the file and opening it again is the only thing that does - Jerry's own finding.
'           - Sh - 8/20/2026 - He hit it after putting qFormat on the five List Bullet styles the same day. Two answers were tried
'           - Sh - 8/20/2026 - first and are written up at the macro so nobody spends the afternoon again: the document is ALREADY
'           - Sh - 8/20/2026 - correct, measured straight after an attach (List Bullet through List Bullet 5 each read
'           - Sh - 8/20/2026 - QuickStyle = True, Visibility = False), so nothing about the styles wants fixing; and toggling each
'           - Sh - 8/20/2026 - gallery style QuickStyle off and back on does NOT refresh it - that was 3.0.201, tested, stale.
'           - Sh - 8/20/2026 - Document.UpdateStyles does refresh it and MUST NEVER BE USED: it re-copies every style from the
'           - Sh - 8/20/2026 - template, throwing away the base font and size just applied and dragging a book set in the dropped
'           - Sh - 8/20/2026 - VistaTypeLP Legible back to Tahoma, moving every page break in it.
'           - Sh - 8/20/2026 - The open handler is deliberately NOT suppressed. Reopening makes a NEW window, so the per-window
'           - Sh - 8/20/2026 - settings the attach applied are gone with the old one, and Sh_HandleDocumentOpened putting them back
'           - Sh - 8/20/2026 - is the point. It cannot loop - that handler only re-opens the attach dialog for an OBSOLETE template,
'           - Sh - 8/20/2026 - and a document reaching here has the current one. It saves explicitly rather than closing with
'           - Sh - 8/20/2026 - wdDoNotSaveChanges, and does nothing at all to a document that has never been saved: a stale gallery
'           - Sh - 8/20/2026 - is far better than a lost book. If the reopen fails it returns NOTHING and the caller says where the
'           - Sh - 8/20/2026 - file is and stops - never the active document, which on a machine with a second book open handed back
'           - Sh - 8/20/2026 - the WRONG one and announced the save as finished over it. It reports a file that comes back READ-ONLY
'           - Sh - 8/20/2026 - (a sync program still holding the ~$ owner file), captures ConfirmConversions BEFORE arming its error
'           - Sh - 8/20/2026 - handler, and closes with screen updating OFF so the activate handler stands down and a neighboring
'           - Sh - 8/20/2026 - document's configuration is not switched in and straight back out, rewriting her AutoCorrect list twice.
'           - Sh - 8/20/2026 - The attach re-asserts the Styles pane afterwards: pane VISIBILITY is application state and does not
'           - Sh - 8/20/2026 - survive the last document closing, though the sort and filter do, being saved in the file.
'           - Sh - 8/20/2026 - What makes the reopen safe at all is one line elsewhere - Lp_Attach_The_Template sets
'           - Sh - 8/20/2026 - UpdateStylesOnOpen = False right after attaching, so the saved file does not re-copy styles from the
'           - Sh - 8/20/2026 - template on open. Do not tidy that line away.
'           - Sh - 8/20/2026 - Sh_HandleDocumentClosing got its On Error back, lost with the Styles-pane line on 8/2/2026: version 1.1
'           - Sh - 8/20/2026 - added it because that sub crashes if an image is selected when a document closes, and this change makes
'           - Sh - 8/20/2026 - the add-in close a document on every single attach.
' Notes:    - LP - 8/20/2026 - the bundled VistaTypeLP Legible typeface is GONE - dropped from the product, not merely un-defaulted.
'           - LP - 8/20/2026 - Jerry's call, and the reason is coverage: the face carries a Latin character set. English, French, German,
'           - LP - 8/20/2026 - Spanish and Italian set correctly in it; Latin, mathematics, the IPA and the Greek that runs through medical
'           - LP - 8/20/2026 - transcription do not. Word substitutes a character the face has not got SILENTLY, out of another typeface at
'           - LP - 8/20/2026 - another size, which breaks the one promise large print makes - every character at the size the reader asked
'           - LP - 8/20/2026 - for. Tahoma covers the lot, and is what a large print book was set in for years before 8/8/2026.
'           - LP - 8/20/2026 - The Typeface choice is off the attach dialog (FontChoiceFrame / FontTahoma / FontLegible are removed from
'           - LP - 8/20/2026 - LP_Attach_An_Lp_Template_Form, with their two Click handlers and the Initialize block). A new large print
'           - LP - 8/20/2026 - book is Tahoma, with no question asked, exactly as before the choice existed. The installer no longer ships
'           - LP - 8/20/2026 - the four .ttf files or their OFL license, and REMOVES a copy it finds already installed.
'           - LP - 8/20/2026 - A book ALREADY set in the face keeps it. LP_FONT_LEGIBLE became LP_FONT_LEGACY_LEGIBLE and now does one
'           - LP - 8/20/2026 - job: recognizing such a book in AttachOkay_Click so re-attaching leaves its typeface alone. Without that,
'           - LP - 8/20/2026 - opening one of Jerry's or the testers' books and pressing Attach would rewrite it to Tahoma and move every
'           - LP - 8/20/2026 - page break, silently, in a book that may already be in a reader's hands. The EmbedTrueTypeFonts line stays
'           - LP - 8/20/2026 - for the same reason - those books carry their own copy of the face and go on setting correctly after the
'           - LP - 8/20/2026 - font leaves the machine. Lp_Indent_Factor_For_Font keeps its 1.054 case for the same books.
'           - LP - 8/20/2026 - KEPT, because neither is about that typeface: Sh_Is_Font_Installed and Sh_Font_Status_Text answer the
'           - LP - 8/20/2026 - question for ANY face, and Sh_Doc_Info's "is this typeface here?" line is the only thing that ever says out
'           - LP - 8/20/2026 - loud that Word is substituting. Lp_Apply_Base_Font_To_Styles is kept too: Normal alone does not reach the
'           - LP - 8/20/2026 - 12 colored character styles or "No Spacing", so Tahoma still has to be put on all 14 by hand.
' Notes:    - BRL - 8/18/2026 - attaching a BANA template now puts the whole document into TIMES NEW ROMAN 14 pt. It was Courier New
'           - BRL - 8/18/2026 - 12 pt from 8/5/2026 until today. The reason is INTEROPERABILITY, not typography: Duxbury ships its own
'           - BRL - 8/18/2026 - Word add-in, SWIFT, and SWIFT sets Times New Roman 14 when it attaches a template. A transcriber may use
'           - BRL - 8/18/2026 - both, and a file whose face depends on which add-in touched it last is one more thing for her to think
'           - BRL - 8/18/2026 - about. Monospace was the right instinct for braille source and is overruled by this. Jerry, 8/18/2026.
'           - BRL - 8/18/2026 - The sub is renamed to match - it named the font -
'           - BRL - 8/18/2026 - and it is not a ribbon button or a toolbar id, so nothing in the field references it. It still runs LAST,
'           - BRL - 8/18/2026 - after the optional cleanups, for the reason it always did: those passes rewrite text, so a font set any
'           - BRL - 8/18/2026 - earlier is not the font the transcriber ends up reading.
' Notes:    - Dx - 8/21/2026 - both books now switch ON "Set left- and first-indent with tabs and backspaces" (Options.TabIndentKey) - Jerry, 3.0.220. Braille switched it OFF from 8/18/2026 until then; large print never wrote it, and inherited braille's OFF because it is one global Word setting.
' Notes:    - Dx - 8/18/2026 - braille switched OFF "Set left- and first-indent with tabs and backspaces" (Options.TabIndentKey).
'           - Dx - 8/18/2026 - Jerry, 8/18/2026. It had been one of the fifteen removed from all three configurations earlier the same
'           - Dx - 8/18/2026 - day, on the measurement that all three wrote it True - which was correct at the time and is now wrong.
'           - Dx - 8/18/2026 - It comes back as a SWITCHED setting, not a forced one: noted while the ordinary configuration is in force
'           - Dx - 8/18/2026 - and put back on the way to an ordinary document, through the same store as the five spelling ones. The two
'           - Dx - 8/18/2026 - helpers are renamed Sh_Save_/Sh_Restore_Transcriber_Settings, since they no longer cover spelling alone.
' Notes:    - Sh - 8/18/2026 - a small SETTINGS STORE, %AppData%\VistaType LP Settings\VistaType.ini, written with Word's own
'           - Sh - 8/18/2026 - System.PrivateProfileString (Sh_Settings_File / Sh_Setting_Read / Sh_Setting_Write). The five spelling and
'           - Sh - 8/18/2026 - grammar settings now live there instead of in module variables. That is not tidiness: VBA's End statement
'           - Sh - 8/18/2026 - resets every module-level variable, and this project runs End on ordinary paths - 34 times in this module,
'           - Sh - 8/18/2026 - plus dialog Cancel buttons and Sh_Is_Doc_Open when no document is open. Held in memory, one Cancel in a
'           - Sh - 8/18/2026 - braille file lost what she had, the next ordinary document put nothing back, and the switch after that
'           - Sh - 8/18/2026 - recorded braille's own switched-off values as her preference for good. The file also survives Word closing,
'           - Sh - 8/18/2026 - so her settings now hold across sessions, and it can be copied to a new machine. A SEPARATE folder from
'           - Sh - 8/18/2026 - %AppData%\VistaType LP, which the uninstaller wipes whole - see installer/vistatype.iss. Related: the five
'           - Sh - 8/18/2026 - writes in the braille sub, and the two in the large print one, now come AFTER the line recording what is in
'           - Sh - 8/18/2026 - force, so a configuration that raises part way through cannot leave them off while the record still says DEF.
' Notes:    - Dx - 8/18/2026 - Word's AutoFormat command is no longer run by Dx_Fix_Common_File_Errors, the last place in the project
'           - Dx - 8/18/2026 - that ran it. The large print side dropped the same call on 8/13/2026 because AutoFormat REMOVES EMPTY
'           - Dx - 8/18/2026 - PARAGRAPHS and no combination of its options stops it; the braille copy had the identical fault and nobody
'           - Dx - 8/18/2026 - had looked. Quotes, fractions and dashes are done by dedicated steps in the same sequence, and the live
'           - Dx - 8/18/2026 - hyperlinks it made were turned back into text a few steps later. Symbols and ordinals go with it, neither
'           - Dx - 8/18/2026 - wanted in a DBT source file. Jerry, 8/18/2026. It also removes the reason the on-demand AutoFormat settings
'           - Dx - 8/18/2026 - had to be written at all - they were arguments to that command, and are now the transcriber's like the rest.
' Notes:    - MS - 8/18/2026 - the three Word configurations, measured against each other and then cut back. FIFTEEN settings were being
'           - MS - 8/18/2026 - written IDENTICALLY by all three - replaceText ("Replace text as you type"), smart quotes, hyperlinks, two
'           - MS - 8/18/2026 - initial capitals, days of the week, cAPS LOCK, tab indents and the rest - so nothing was ever switching them
'           - MS - 8/18/2026 - and the 45 writes did one thing only: overwrite whatever the transcriber had chosen. That is what the beta
'           - MS - 8/18/2026 - tester on 3.0.135 reported. All 45 deleted; those settings are hers now, in every kind of document. Of what
'           - MS - 8/18/2026 - remains, every write in the large print and braille subs is guarded with If <> - not for switching, where the
'           - MS - 8/18/2026 - values genuinely differ, but because about a dozen places run those subs DIRECTLY while that configuration is
'           - MS - 8/18/2026 - already in force. Five spelling and grammar settings (grammar as you type, contextual speller, mixed digits,
'           - MS - 8/18/2026 - smart tags, ignore uppercase) were switched off by braille, two of them by large print, and put back by
'           - MS - 8/18/2026 - nothing: ONE braille file turned grammar checking off for every document she opened, that session and every
'           - MS - 8/18/2026 - session after. They are now noted while the ordinary configuration is genuinely in force (seeded by AutoExec
'           - MS - 8/18/2026 - before anything can configure) and put back on the way to an ordinary document. CorrectKeyboardSetting is no
'           - MS - 8/18/2026 - longer written by any of the three. Counts: default 35 -> 20, large print 37 -> 22, braille 40 -> 25.
' Notes:    - Sh - 8/18/2026 - Sh_HandleDocumentNew now skips itself while a MACRO is running, the same test Sh_HandleDocumentActivated
'           - Sh - 8/18/2026 - has had since 8/9/2026 and the only one of the handlers that was missing it. Word raises NewDocument for
'           - Sh - 8/18/2026 - every Documents.Add, so the scratch document that Lp_/Dx_Copy_To_Temp_Doc make - from about 35 places -
'           - Sh - 8/18/2026 - was reconfiguring Word for an ORDINARY document in the middle of a large print or braille job: Styles pane
'           - Sh - 8/18/2026 - closed and screen refreshed where the macro had deliberately turned refreshing off, the configuration
'           - Sh - 8/18/2026 - written twice, and the 18 fraction entries added and deleted again on every run. A document a macro makes
'           - Sh - 8/18/2026 - FOR the transcriber is configured when her cursor reaches it instead. Jerry, 8/18/2026.
' Notes:    - MS - 8/18/2026 - the DEFAULT Word configuration (MS_Set_Word_Config_For_New_Install, the one an ordinary document gets)
'           - MS - 8/18/2026 - was checked setting by setting against Word's factory defaults. Thirty-two of the thirty-four already agreed. Two did not:
'           - MS - 8/18/2026 - AutoFormatAsYouTypeApplyFirstIndents was never written and is now True; AutoFormatReplacePlainTextEmphasis
'           - MS - 8/18/2026 - was True where a clean install has it False; and CorrectKeyboardSetting, which follows the keyboard to the
'           - MS - 8/18/2026 - language being typed, is no longer written at all - it belongs to a multilingual transcriber and to nothing
'           - MS - 8/18/2026 - in large print or braille. The AutoCorrect block now also records what must NEVER be written there: the four
'           - MS - 8/18/2026 - AutoAdd properties and the four exception lists, which are her accumulated work and cannot be rebuilt.
'           - MS - 8/18/2026 - Large print and braille were NOT touched - Jerry scoped this to the default configuration only.
' Notes:    - MS - 8/18/2026 - the DEFAULT Word configuration no longer deletes the 18 compact fraction AutoCorrect entries. It adds them,
'           - MS - 8/18/2026 - exactly as the braille configuration does (Jerry, 8/18/2026). What that changes: Word's own half, quarter and
'           - MS - 8/18/2026 - three-quarters come from AutoFormatAsYouTypeReplaceFractions and never stopped working, but the other fifteen
'           - MS - 8/18/2026 - - 1/3, 2/3, 1/5 and the rest - exist only as these entries, and the deletes here (10/19/2021) undid braille's
'           - MS - 8/18/2026 - adds the same day. They now stay in the transcriber's own AutoCorrect list, in her own documents, for good.
'           - MS - 8/18/2026 - The eighteen live in ONE place, Sh_Add_Compact_Fractions, called by both, so the two lists cannot drift; it
'           - MS - 8/18/2026 - writes only an entry that is missing or holds something else, because the default sub runs on EVERY ordinary
'           - MS - 8/18/2026 - document and needless writes are what brought on the privacy notice of 7/18/2026. Large print still deletes
'           - MS - 8/18/2026 - them: a large print book keeps 1/2 as it was typed.
'           - Sh - 8/9/2026 - the Word configuration now FOLLOWS the document. Word keeps one set of Options and AutoCorrect entries
'           - Sh - 8/9/2026 - for the whole application, so an LP book and a braille file open together could not both have their
'           - Sh - 8/9/2026 - configuration in force: whichever was opened last won and went on winning, and Sh_Doc_Info reported the
'           - Sh - 8/9/2026 - configuration of a document the transcriber had clicked away from. VtEvents now takes Word's
'           - Sh - 8/9/2026 - WindowActivate as well, and Sh_HandleDocumentActivated puts the right one of the three back. That event
'           - Sh - 8/9/2026 - fires constantly, so only a change of TYPE costs anything: same type as the config already in force, do
'           - Sh - 8/9/2026 - nothing (the Options are already right, and the per-document and per-window settings travel with the
'           - Sh - 8/9/2026 - document). Four guards come first. It skips while a MACRO is running - Application.ScreenUpdating being
'           - Sh - 8/9/2026 - off is the marker - because Lp_/Dx_Copy_To_Temp_Doc activate a document and are called from about 35
'           - Sh - 8/9/2026 - places, and reconfiguring mid-macro would put back the ~74 Options writes taken out of fourteen dialogs
'           - Sh - 8/9/2026 - on 7/24/2026 and refresh the screen where the code says it must not. It skips the GPL window the add-in
'           - Sh - 8/9/2026 - opens itself, whose Sh_Skip_Open_Handler guard only spans the open. It has a re-entrancy flag, because
'           - Sh - 8/9/2026 - closing Reading view raises WindowActivate again. And a failed config puts ScreenUpdating back - the
'           - Sh - 8/9/2026 - braille one turns it off with no handler of its own, which on a click path would have frozen the screen
'           - Sh - 8/9/2026 - and done it again on the next click. There is deliberately NO "same document" shortcut: a document's type
'           - Sh - 8/9/2026 - changes while it stays active when a template is attached, and a name-based skip would never look again.
'           - Sh - 8/9/2026 - It never runs the rest of the open handler - that one can raise the obsolete-template message box and open
'           - Sh - 8/9/2026 - the attach dialog, which is right once per open and intolerable on every click.
'           - Sh - 8/9/2026 - Switching sets the TYPING side only - Options, AutoCorrect, spelling, fractions - and leaves the screen
'           - Sh - 8/9/2026 - alone: Styles pane, formatting marks, both rulers, style area, view type. Jerry's call. The typing side is
'           - Sh - 8/9/2026 - the part that is actually WRONG when it belongs to the other document; the screen is the part a transcriber
'           - Sh - 8/9/2026 - arranges for a session and would resent losing, and taking it back on every glance would be the 7/24 and
'           - Sh - 8/9/2026 - 8/2/2026 Styles-pane complaints returning. New Public Sh_Config_Skip_Display, set only by
'           - Sh - 8/9/2026 - Sh_Apply_Word_Config (new Optional DisplayToo) and always cleared, guards the display blocks inside the
'           - Sh - 8/9/2026 - three MS_ subs. OPENING a document, and attaching a template, still set the display exactly as before.
'           - Sh - 8/9/2026 - New: Sh_HandleDocumentActivated, Sh_Doc_Config_Type ("LP"/"BRL"/"DEF"), Sh_Apply_Word_Config,
'           - Sh - 8/9/2026 - Sh_Is_Addins_Own_Document, and the Public Sh_ConfiguredAs - written by the three MS_Set_Word_Config_*
'           - Sh - 8/9/2026 - subs themselves, beside MS_Word_Config, because about a dozen other places run those subs directly and
'           - Sh - 8/9/2026 - recording it in the callers would leave it describing a configuration Word was not actually carrying.
'           - Sh - 8/9/2026 - The three-way config in Sh_HandleDocumentOpened and Sh_HandleDocumentNew now goes through
'           - Sh - 8/9/2026 - Sh_Apply_Word_Config, so opening and switching can never drift apart.
'           - Sh - 8/8/2026 - American spellings throughout, in comments and in anything the transcriber reads. Jerry's call: the
'           - Sh - 8/8/2026 - product is for American transcribers. licence->license, colour->color, behaviour->behavior,
'           - Sh - 8/8/2026 - recognise->recognize, normalise->normalize, grey->gray, "dialogue window"->"dialog window". The three
'           - Sh - 8/8/2026 - legal texts are NOT touched (LICENSE, docs/Software-Agreement.md, and the font's OFL.txt, which is a
'           - Sh - 8/8/2026 - verified character-for-character transcription and must stay that way). No style name, ribbon id or
'           - Sh - 8/8/2026 - control id contains any of these words, so nothing a document or an installed toolbar points at moved.
'           - BRL - 8/8/2026 - the "Prodnote -> TN" button on the Braille Macros tab now reads "Prodnote to TN". The arrow was
'           - BRL - 8/8/2026 - ChrW(8594), outside Windows-1252, which forced Sh_Prodnote_Info_Form to build its copy of the name
'           - BRL - 8/8/2026 - with ChrW to survive export. Plain words on both sides now. Only the LABEL changed - the button id
'           - BRL - 8/8/2026 - btn_Dx_Change_Prodnotes_To_Transcriber_Notes is untouched, so every toolbar already in the field
'           - BRL - 8/8/2026 - still resolves it.
'           - Sh - 8/8/2026 - Sh_Prodnote_Info_Form's note rewritten and the form made taller to suit (Jerry). The attach dialog
'           - Sh - 8/8/2026 - LP_Attach_An_Lp_Template_Form was re-laid-out (Jerry) - layout only, no code change. The installer's
'           - Sh - 8/8/2026 - welcome page says what the add-ins are and that the typeface carries its own separate license.
'           - LP - 8/8/2026 - the transcriber now CHOOSES the typeface on the attach dialog, beside the point size: Tahoma as
'           - LP - 8/8/2026 - before, or the bundled VistaTypeLP Legible, which is Atkinson Hyperlegible rescaled so that a point
'           - LP - 8/8/2026 - size set in Word matches the PRINTED letter size (18 pt Atkinson measured about 16.5 on Jerry's
'           - LP - 8/8/2026 - font ruler, 8/8/2026). Nothing new stores the choice: like the point size, it is read back off the
'           - LP - 8/8/2026 - document's Normal style, so it survives closing and reopening and a book made before this feature
'           - LP - 8/8/2026 - existed correctly answers Tahoma. LargePrintTemplate.dotx is deliberately NOT touched - it stays
'           - LP - 8/8/2026 - the Tahoma baseline and the macros override per document, which is what makes that true.
'           - LP - 8/8/2026 - New: Lp_Apply_Base_Font_To_Styles (Normal, plus the 13 styles that name a face of their own and so
'           - LP - 8/8/2026 - ignore Normal - the 12 colored character styles and "No Spacing", which has no basedOn at all),
'           - LP - 8/8/2026 - Sh_Is_Font_Installed, Sh_Font_Status_Text, Lp_Indent_Factor_For_Font. Lp_Attach_The_Template also
'           - LP - 8/8/2026 - EMBEDS the font when it is ours, so a book still sets correctly on a machine without it; Tahoma is
'           - LP - 8/8/2026 - never embedded. The Tahoma-missing warning left Lp_Attach_Lp_Template and is now the dialog graying
'           - LP - 8/8/2026 - out a face that is not installed - a better answer, and it takes an "End" statement out of the path
'           - LP - 8/8/2026 - that used to wipe every Public on the way past.
' Notes:    - Sh - 8/6/2026 - INSTALLER FIX: upgrading over an older version left the ribbon tabs stale (Jerry, 8/6/2026).
'           - Sh - 8/6/2026 - Both tabs lost the delete-prodnote icons and the braille tab lost change-prodnote-to-TN. The tabs the
'           - Sh - 8/6/2026 - installer writes into the user's own Word.officeUI were only touched when the ribbon-tabs task was
'           - Sh - 8/6/2026 - ticked - and it is unchecked by default - so an upgrade kept the OLD tabs while UserRibbonTabs stayed
'           - Sh - 8/6/2026 - 1, which keeps the add-in's own correct copies hidden. Merge-Qat.ps1 has a new Tabs mode, Refresh:
'           - Sh - 8/6/2026 - update tabs the user already has, never add one they declined. Nothing in this module changed.
' Notes:    - BRL - 8/5/2026 - Dx_AutoTag_Page_Numbers now CONSOLIDATES back-to-back reference page numbers into one range
'           - BRL - 8/5/2026 - (Jerry, 8/5/2026), the braille answer to the large print merge of 8/2/2026. It runs at the seam
'           - BRL - 8/5/2026 - between the two halves of that macro: after everything that strips an earlier tagging, before the
'           - BRL - 8/5/2026 - tagging passes. That is the one place where the numbers are BARE - no $pg, no DBT codes, styles back
'           - BRL - 8/5/2026 - to Normal - so consolidating there means 12/13/14 becomes 12-14 and the existing hyphen pass turns
'           - BRL - 8/5/2026 - it into $pg12-14[[*lec*]][[*i*]]14, DBT code and all. Merging AFTER tagging, which is how it was
'           - BRL - 8/5/2026 - first written, would have had to build that code by hand. Only numbers carrying a DIGIT merge:
'           - BRL - 8/5/2026 - a pure roman pair is left alone because the [[*ii*]] lower-case marker is applied by the roman
'           - BRL - 8/5/2026 - numeral loop further down, which cannot recognize a range. Tables are skipped and a section break
'           - BRL - 8/5/2026 - ends a run, both as on the LP side.
' Notes:    - Sh - 8/5/2026 - the DAISY/NIMAS conversion now ASKS whether to keep the images, straight after the DAISY/NIMAS
'           - Sh - 8/5/2026 - choice (Jerry, 8/5/2026). New DN_Keep_Or_Omit_Images_Form: "Keep images (Large Print)" is the default,
'           - Sh - 8/5/2026 - "Omit images (braille)" the alternative, answered through Sh_GP_String_2 as "KEEP" or "OMIT". Braille
'           - Sh - 8/5/2026 - almost never wants the pictures and large print always does. It is also the biggest lever on how long
'           - Sh - 8/5/2026 - the job takes. Jerry timed the same book on the same VM both ways: 5 min 30 sec keeping the images,
'           - Sh - 8/5/2026 - 2 min 30 sec omitting them. The delete loop runs BACKWARDS - deleting a shape renumbers
'           - Sh - 8/5/2026 - the rest, so forwards would skip every second picture.
' Notes:    - Sh - 8/5/2026 - EVERY ActiveDocument.Styles("name") lookup in this module is now guarded (Jerry asked, 8/5/2026).
'           - Sh - 8/5/2026 - There were 158 of them and they all raise run-time error 5941 on a document that does not carry the
'           - Sh - 8/5/2026 - style - which is normal: RefPageNemeth comes from the Nemeth templates, Print Pg Num from the LP one.
'           - Sh - 8/5/2026 - 76 whole Find passes are wrapped in Sh_Style_Exists, since a pass can do nothing without its style;
'           - Sh - 8/5/2026 - the rest were done by hand. Two shapes needed care. In Dx_Remove_Bullets five Styles() calls sit in
'           - Sh - 8/5/2026 - mutually exclusive branches, so each branch got its own test - one guard round the pass would have
'           - Sh - 8/5/2026 - skipped it whenever ANY of the five was missing. In Lp_TOC_CleanAndFormat_TOC the lookup sat on the
'           - Sh - 8/5/2026 - right of an Or; VBA does not short-circuit, so it ran even when the left side already matched - that
'           - Sh - 8/5/2026 - one compares by NAME now, via Sh_Para_Style_Is. New Sh_Style_In_Use covers the .InUse tests.
'           - Sh - 8/5/2026 - tools/lib/check_style_guards.py runs in `make build` and refuses to build if a new one appears.
' Notes:    - Sh - 8/5/2026 - the temp-file validation no longer leaves a Find All MULTIPLE SELECTION behind (Jerry, 8/5/2026).
'           - Sh - 8/5/2026 - It lists the tags by driving Find and Replace with SendKeys and picking Find In > Main Document, which
'           - Sh - 8/5/2026 - selects every $pg paragraph at once. That selection outlived the validation and the next macro to touch
'           - Sh - 8/5/2026 - Selection died with run-time error 4605 - AutoTag, in Jerry's case. The user cannot dodge it: the
'           - Sh - 8/5/2026 - Navigation pane refuses the job once there are more tags than it can hold. New Sh_Clear_Multi_Selection
'           - Sh - 8/5/2026 - now runs when validation finishes AND at the top of both AutoTag macros.
' Notes:    - BRL - 8/5/2026 - Dx_AutoTag_Page_Numbers no longer fails with run-time error 5941 on a document whose $pg tags were
'           - BRL - 8/5/2026 - already tagged (Jerry, 8/5/2026). Its first act is to STRIP the earlier tagging, and three of those
'           - BRL - 8/5/2026 - passes ask for RefPageNumber or RefPageNemeth by name. Styles(name) raises 5941 when the document does
'           - BRL - 8/5/2026 - not carry the style, and it need not - RefPageNemeth belongs to the Nemeth math templates. The whole
'           - BRL - 8/5/2026 - macro died before a single tag was placed. No style means nothing is formatted with it, so those passes
'           - BRL - 8/5/2026 - are now skipped. Same guard on the DBT code restore. NOTE: about 80 other Styles("name") lookups in this
'           - BRL - 8/5/2026 - module are still unguarded and can fail the same way; only the reported one was changed.
' Notes:    - Sh - 8/5/2026 - repaired the end marker in BOTH Horz_To_Vert_List forms. It is the OE ligature, what Alt+0140 types,
'           - Sh - 8/5/2026 - and it was the byte 8C until 7/26/2026, when a lossy read turned all five of them into U+FFFD - which
'           - Sh - 8/5/2026 - then shipped as three garbage characters. Now ChrW(338), which no re-encoding can damage. The macro
'           - Sh - 8/5/2026 - inserts the marker and strips it again, so all five spots were mangled alike and it may well have kept
'           - Sh - 8/5/2026 - working by luck; it needs a run on a real list all the same.
' Notes:    - BRL - 8/5/2026 - deleted Dx_MathSymbolsForm and Dx_MathSymbolsAgain. Jerry: an unfinished project, never wired up.
'           - BRL - 8/5/2026 - They referenced only each other - Again shows Form, Form unloads Again - so there was no way in from
'           - BRL - 8/5/2026 - the ribbon, the keymap, the QAT or any macro. Nothing else called anything in them either. The one
'           - BRL - 8/5/2026 - outside name they still mentioned was Dx_ShowMathSymbolsAgain, and only in the comment recording its
'           - BRL - 8/5/2026 - own deletion on 8/3/2026.
' Notes:    - Sh - 8/5/2026 - the DAISY/NIMAS conversion now puts the whole document into Tahoma 12 pt, just before the repaginate
'           - Sh - 8/5/2026 - and save at the end - a font change moves every line and page break, so it has to precede both. The new
'           - Sh - 8/5/2026 - Sh_Set_Whole_Document_Font does the work for it and for the braille attach's Courier New 12. Face and
'           - Sh - 8/5/2026 - size only: color is left alone, so the red $pg tags and red prodnotes come through unchanged.
' Notes:    - BRL - 8/5/2026 - Dx_Fix_Common_File_Errors now ends by coloring the $pg tags red, via the new Dx_Color_Dollar_PG_Red.
'           - BRL - 8/5/2026 - Deliberately NOT the LP side's Sh_Color_Dollar_PG_Red: that also sets the replacement style to Normal,
'           - BRL - 8/5/2026 - which Word applies to the whole paragraph and would strip the BANA style off every paragraph with a tag.
' Notes:    - BRL - 8/5/2026 - Dx_Attach_BANA_Template now ends by putting the whole document into Courier New 12 pt, via the new
'           - BRL - 8/5/2026 - Dx_Set_Whole_Document_To_Times_New_Roman_14. It runs LAST, after the optional fix-errors and para-mark
'           - BRL - 8/5/2026 - cleanups, because both of those rewrite text and would otherwise be the last word on the font.
' Notes:    - LP - 8/5/2026 - Lp_Add_Para_After_Image now EXITS at once if the LP template is attached. It repairs raw DAISY and
'           - LP - 8/5/2026 - NIMAS files, which is a before-the-template job; run afterwards from File Cleanup it added stray
'           - LP - 8/5/2026 - paragraph marks to a document that was already formatted.
' Notes:    - Sh - 8/5/2026 - Sh_ReplaceNonBreakingSpacesWithNormalSpace now SKIPS any non-breaking space in a paragraph styled
'           - Sh - 8/5/2026 - "Print Pg Num". The bar built by Lp_Format_Page_Numbers begins and ends with one, and the color pass
'           - Sh - 8/5/2026 - in that macro finds those ends by searching for ^s, so flattening them took the bar apart. A document
'           - Sh - 8/5/2026 - with no such paragraph still gets the single whole-document replace, which is the usual case.
'           - Sh - 8/3/2026 - the progress bar is weighted to where the time ACTUALLY goes, measured end to end on a real NIMAS book
'                             with 1,236 images (301 s total): read/tag/write 0.03 s, InsertFile 13.95 s, the image loop 213.27 s at
'                             ~173 ms each, everything else under 3 s, SaveAs 70.74 s for 43.7 MB. So the loop is 71 per cent of the
'                             job and the save 23. The first weights had the bar resting at 20% through a 14-second import and jumping
'                             to 88% before a three-and-a-half-minute stage - close to backwards
'           - Sh - 8/3/2026 - and the bar now STEPS INSIDE the image loop, every 10 images. That loop is the only long stage that is a
'                             loop, so it is the only one that can show progress at all: InsertFile and SaveAs are single blocking
'                             calls and nothing can move during them. It turns 71 per cent of the run from a frozen screen into a bar
'                             that visibly advances, which is the whole point - Jerry: "many will think the computer is frozen and
'                             well... reboot time". Every 10 rather than every 1 because a repaint per image would add its own cost to
'                             the loop being measured
'           - Sh - 8/3/2026 - two ways to make that loop faster were tried and BOTH FAILED, recorded so they are not tried again.
'                             Fields.Unlink does nothing for these pictures: they arrive as linked InlineShapes, not INCLUDEPICTURE
'                             fields - 143 fields against 1,236 shapes - and the saved file came to 0.3 MB against 43.7, so nothing was
'                             embedded. Setting SavePictureWithDocument without BreakLink was no faster. The 213 seconds is the price
'                             of putting 86 MB of JPEGs inside the document, and that is what makes the file standalone
'           - Sh - 8/3/2026 - the conversion shows a PROGRESS BAR now, not the old non-modal box. Jerry: without a visible activity
'                             indicator "many will think the computer is frozen and well... reboot time". New
'                             Sh_Convert_Progress_Form (51 forms) - MSForms has no progress bar control, so it is two labels, a sunken
'                             track and a colored fill whose Width the code drives; no ActiveX to fail to register on a transcriber's
'                             machine. Stage weights come from the timings below, so the bar is honest about where the work is
'           - Sh - 8/3/2026 - the old box carried a spinner that had NEVER turned: Sh_ShowNonModalMessage only shows the form and
'                             nothing in the converter ever called StartSpinner, so the flag stayed False and every tick exited at once.
'                             Its message told the user not to touch the keyboard and to "Wait for the BEEP!", written when a conversion
'                             took minutes rather than seconds. Both gone
'           - Sh - 8/3/2026 - one thing no indicator can fix, so the bar SAYS it: InsertFile is a single blocking call and VBA is
'                             single-threaded, so nothing moves during the HTML import. The bar stops at 20% and the message explains
'                             that it will, because an unexplained freeze is what causes the reboot
'           - Sh - 8/3/2026 - CAUTION on the timings below: they are the stages that were MEASURED, and they do not account for the
'                             whole conversion. Jerry reports 5 min 30 s end to end for a 900KB file on the VM, against about 2.7 s of
'                             measured stages. Roughly five minutes is therefore somewhere not yet timed - the InlineShapes BreakLink
'                             loop that embeds every image is the first place to look, then Repaginate and the save. The stage weights
'                             the progress bar uses are derived from the measured stages only, so they will be wrong in the same way
'           - Sh - 8/3/2026 - Sh_Document_Has_Prodnotes now WALKS THE PARAGRAPHS instead of using a style-aware Find. The note
'                             explaining prodnotes stopped appearing after a real conversion even though the document plainly had red
'                             prodnotes in it. Word's Find settings are STICKY - the ones a caller does not set explicitly carry over
'                             from whatever used Find last, and by that point Sh_Color_Dollar_PG_Red and others have been setting them
'                             all the way through. In isolation the Find answered True at every stage; in the real macro it did not
'           - Sh - 8/3/2026 - and an honest note: the Step 6 rewrite below removed the opening Show along with the temp-document code
'                             it was sitting inside, so for one build the box did not appear until the repaginate stage
'           - Sh - 8/3/2026 - DAISY/NIMAS conversion is faster, and the two slow parts were not where anyone would guess.
'                             Timed stage by stage on a real 900KB NIMAS book (403 page numbers, 782 prodnotes) on the build box:
'                               read+clean 0.00   $pg tagging 0.48   prodnote tagging 3.77   write html 0.00
'                               InsertFile 2.39   Fields.Unlink 0.04   font 0.27
'                             So more than half the time went on prodnote tagging - more than Word's own HTML import
'           - Sh - 8/3/2026 - prodnote tagging 3.79 s -> 0.00 s. The cost was InStr with vbTextCompare walking the whole file
'                             looking for <prodnote and </prodnote>, 1564 times: locale-aware matching a character at a time.
'                             It now lowercases a copy once and searches that with vbBinaryCompare, slicing from the original -
'                             same case-insensitive behavior, and LCase$ does not change the string's length so positions still
'                             line up. Output verified IDENTICAL to the old routine, character for character, on that book.
'                             Accumulation also changed from outStr = outStr & ... to Join, but measurement says that was worth
'                             0.09 s of the 3.79 - it is kept as the better shape, not as the fix
'           - Sh - 8/3/2026 - $pg tagging 0.48 s -> 0.01 s. It opened a hidden Word document, pushed the entire XML into it, ran a
'                             wildcard Find/Replace and read it all back, to perform ONE substitution. Now a VBScript.RegExp over
'                             the string. Output identical apart from the trailing paragraph mark Word appends to any document.
'                             Word wildcards escape < and > as \< and \> because they mean word boundaries; a regex does not
'           - Sh - 8/3/2026 - untouched, deliberately: the red. Prodnote color is set after the import (Styles("Prodnote").Font
'                             .Color) and Sh_Color_Dollar_PG_Red runs after that, so neither depends on any of this. Jerry's two
'                             conditions for trying a faster converter
'           - Sh - 8/3/2026 - dead code cleared, 311 lines and eight macros. Nothing referenced any of them - not the ribbon, not the
'                             keymap, not a form, not each other. MS_HandleTemplateChange (its comment claimed a class module,
'                             clsAppEvnets, that has never existed - the real event sink is VtEvents, which wires three events);
'                             Lp_Add_Hidden_PN_to_Page_Number_Bar, 121 lines superseded by Lp_Format_Page_Numbers;
'                             Lp_ReplaceNBSP_ExcludePrintPgNumbAndTables, whose only call site was commented out and which carried a typo
'                             ("Print Pg Numb") that meant its exclusion never fired anyway; the three TempPlaceholder bookmark helpers,
'                             left unreferenced by the 7/26/2026 cursor-position work; Dx_Set_Display_For_Braille; and Non_modal_Test.
'                             Module NewMacros.bas deleted with them - it held only ForceSaveAsDialog, a relic of the Normal.dotm days.
'                             The two commented-out calls that named removed macros went too. History in this changelog stays.
' Notes:    - Sh - 8/3/2026 - "View Full License" opens the GPL as a READ-ONLY WORD DOCUMENT. It went Notepad -> a box on a form ->
'                             this, and the middle step is the instructive one: an MSForms text box does not respond to the mouse wheel.
'                             The control has no wheel handling and no property to switch it on; the only way to add it is a Windows
'                             mouse hook, and a VBA callback from a system hook is a well known way to crash Word. Reading 674 lines by
'                             PageDown is not reasonable in a LARGE PRINT product. Word gives the wheel, Ctrl+scroll zoom, Find and
'                             screen reader support, in the application the transcriber is already in. Sh_License_Form is gone (49 forms)
'           - Sh - 8/3/2026 - two things opening a document would otherwise have done to the user. ConfirmConversions is held off across
'                             the open, or Word stops on its "Convert File" question for a .txt. And new Sh_Skip_Open_Handler makes
'                             Sh_HandleDocumentOpened leave the license completely alone - without it the license counts as an ordinary
'                             document, MS_Set_Word_Config_For_New_Install runs over it, and clicking a button to READ THE LICENSE would
'                             reset the transcriber's Styles pane. Measured on the build box 8/3/2026: an MSForms text box holds at least
'                             65,000 characters, so the 32,000 ceiling that shaped the earlier attempt is folklore here
'           - Sh - 8/2/2026 - both About dialogs now carry the GPLv3, not the old permissive agreement. The wording they had granted use
'                             "at no cost to others" and read like an MIT license - it never matched what this software actually ships
'                             under, which the installer has shown correctly all along. The thirteen old agreement labels are gone from
'                             each form, replaced by ONE scrollable read-only text box filled at run time from new
'                             Sh_Software_Agreement_Text. One copy of the wording, shared by both dialogs, so they cannot drift, and it
'                             lives in git-tracked text instead of inside a binary .frx. Canonical source is docs/Software-Agreement.md
'           - Sh - 8/2/2026 - the dialogs show a plain-English summary, not the license itself: a new "View Full License" button opens
'                             the complete GPL that the installer writes to %AppData%\VistaType LP\LICENSE.txt (Sh_Show_Full_License),
'                             falling back to a pointer at gnu.org if the add-in was copied into STARTUP by hand rather than installed.
'                             The version label was deliberately left untouched on both forms - Import-Vba.ps1 finds it by matching the
'                             caption TEXT, not the control name, so disturbing it would silently stop every future build stamping the
'                             version and the only symptom would be a stale About box
'           - Sh - 8/2/2026 - AutoTag Ref Pages was silently missing page numbers, and had been for years. The patterns are shaped
'                             "^013(a page number)^013" and a Word replace consumes BOTH paragraph marks, so where two numbers sit in
'                             consecutive paragraphs the mark AFTER the first is the very mark the second needs in FRONT of it - already
'                             eaten. One Execute therefore tagged ALTERNATE numbers: Jerry's sample of 8/2/2026 tagged 15 and 17 and
'                             walked past 16, and missed 14, G3 and G5 the same way. Roman numerals came out right only because they are
'                             tagged by a separate paragraph loop that never uses Find
'           - Sh - 8/2/2026 - the fix is new Sh_Replace_All_Until_Done: run the replace the caller has already set up until it stops
'                             changing anything, since on a fresh Execute every paragraph mark is available again. Two or three rounds
'                             converge. Safe to repeat - a tagged paragraph reads "$pg16", which cannot match a pattern wanting only
'                             digits or only letters between the marks, so nothing is tagged twice. Applied to all 22 affected passes,
'                             12 in Lp_AutoTag_Page_Numbers and 10 in Dx_AutoTag_Page_Numbers - the braille twin had the identical bug
'           - LP - 8/2/2026 - two reference page numbers back to back now make ONE pink bar instead of two stacked ones. Real print books
'                             contain blank pages and every page must still carry a number, so DAISY and NIMAS coders emit "$pg12" then
'                             "$pg13" - almost always at a chapter change. The $pg validation never showed this: it is an entirely visual
'                             review with no adjacency checking in it, so the doubled tag survived to Format $pg Tags and the two bars
'                             stacked into a mess. New Lp_Merge_Adjacent_Pg_Tags, called from Lp_Format_Page_Numbers after the passes that
'                             delete an empty "$pg" and before the wildcard that builds the bar, joins a run first to last: $pg12 + $pg13
'                             -> $pg12-13, three in a row -> $pg12-14. An already-hyphenated tag does not grow a second hyphen
'                             ($pg12-13 + $pg14 -> $pg12-14)
'           - LP - 8/2/2026 - ANY two adjacent tags merge whether or not the numbers run in sequence, so $pg12 beside $pg99 gives
'                             $pg12-99. Jerry's call: the "number" is often not a number - roman numerals, "12a", "A12b" all occur and
'                             "the next one" means nothing for them. A run never crosses a section break, a page break or a tab, and
'                             Jerry, 8/2/2026: any $pg in a TABLE is ignored outright. That last rule is worth more than it looks - it
'                             removes merging across a cell boundary, and the end-of-cell marker Chr(7) that Word refuses to delete,
'                             which would raise error 4605 and abort the macro with the screen still frozen and half the bars built.
'                             Tags in tables are still FORMATTED into bars as before; only merging skips them. Silent, and LARGE PRINT
'                             ONLY - the braille twin Dx_Format_Tagged_Page_Numbers is deliberately untouched
'           - LP - 8/2/2026 - the Styles pane belongs to the user again. Jerry: sophisticated users should keep their preferred pane
'                             settings from session to session. The large print macros forced wdStyleSortRecommended and
'                             wdShowFilterFormattingRecommended, and forced the pane open, in a dozen places - every LP document OPEN did
'                             it, four times over. Now exactly TWO things force it: attaching the LP template (new or re-attached, both of
'                             which arrive at Lp_Attach_The_Template), and the new QAT button. The braille side is untouched
'           - LP - 8/2/2026 - what changed: Lp_Set_Display_For_Large_Print and MS_Set_Word_Config_For_Large_Print no longer call
'                             Lp_Turn_on_Styles_Pane, which now has exactly one caller, the attach. CAREFUL - FormattingShowNextLevel was
'                             riding inside that call and exists nowhere else in either sub, so both now write it explicitly rather than
'                             lose it. Also dropped: the pane hide when an LP document closes, and the one in Lp_Copy_To_Temp_Doc. The
'                             latter was the worst of them - eighteen callers, and it never put the pane back, so the pane died on the
'                             first Table Tools or Fill-In Line and stayed dead all session. Pane visibility is WORD-WIDE, so it could
'                             never have tidied the temp window without closing the pane in the user's book too
'           - LP - 8/2/2026 - the close-time hide was also a plain bug: DocumentBeforeClose fires on close ATTEMPTS and nothing read its
'                             Cancel flag, so X then "Cancel" at the save prompt left you in the document with the pane gone
'           - Sh - 8/2/2026 - new Sh_Show_Recommended_Styles_Pane on the Quick Access Toolbar (btn_Sh_Show_Recommended_Styles_Pane, 7 QAT
'                             icons now) puts the pane back on demand. Its icon is QuickStylesGallery, NOT StylesPane: Word's own Styles
'                             Pane button is already on that toolbar, and two identical icons nine slots apart is a misclick waiting to
'                             happen. Word's is kept because it TOGGLES - it can close the pane, which ours deliberately cannot.
'                             Shared, not Lp_: one toolbar serves both tabs. Sets only the three
'                             things it advertises - it does NOT call Lp_Turn_on_Styles_Pane, which also writes the Word-wide
'                             RestrictLinkedStyles. Worth knowing: sort and filter are DOCUMENT properties saved into the file, so they
'                             follow the book; only pane visibility is a Word-wide setting that persists across documents
'           - LP - 8/2/2026 - Lp_Horz_To_Vert_List_Form finally drops its MS_Set_Word_Config_For_Large_Print call, the one the 7/24/2026
'                             round missed. Saves ~40 Options/AutoCorrect writes and 19 AutoCorrect deletes on every open of that dialog
'           - Sh - 8/1/2026 - DAISY/NIMAS -> Word: a converted book that carries prodnotes now ends with a note explaining what they are,
'                             why they are red, and what to do with them for braille and for large print. Shown after the file is saved and
'                             ONLY when the document actually has prodnotes, since it opens by saying it contains them. New UserForm
'                             Sh_Prodnote_Info_Form (47 forms -> 48) rather than a MsgBox: the text runs to about 1270 characters and VBA
'                             truncates a MsgBox at roughly 1024, which would have silently dropped the paragraph telling the user how to
'                             delete prodnotes. The arrow in "Prodnote -> TN" is ChrW(8594) - the ribbon button's own character, but not in
'                             the Windows-1252 set the module exports as, so a literal one would not survive a build. The Styles pane opens
'                             behind the note - every style listed, sorted As Recommended - so the red Prodnote entry is visible in the pane
'                             while the note explains it. Jerry, 8/1/2026. Guarded by On Error: the pane is a courtesy and must not cost the
'                             user the note. Note this is deliberate, and narrower than the 7/24/2026 change that stopped 14 LP dialogs
'                             resetting the pane on every open - it happens once, at the end of a conversion, not on every dialog
'           - Sh - 8/1/2026 - "does this document USE the Prodnote style" is now one function, Sh_Document_Has_Prodnotes. The style merely
'                             EXISTING proves nothing - every LP document defines Prodnote - so the test has to be a style-aware Find, and
'                             there is no reason to have two of those. Sh_Set_Prodnote_Style_Visibility now calls it instead of carrying its
'                             own copy, and it is also the cheap early exit for Sh_Strip_Prodnote_Enclosing_Quotes, which would otherwise
'                             walk every paragraph in the book to discover there were no prodnotes
'           - Sh - 8/1/2026 - DAISY/NIMAS -> Word: quotation marks that ENCLOSE a prodnote are now removed at the close of the conversion.
'                             Some books wrap the note in them - "Illustration of a barn." - where the marks belong to the source markup and
'                             not to the note. New Sh_Strip_Prodnote_Enclosing_Quotes, called from Sh_Convert_XML_File_To_Word_Document after
'                             the prodnotes are styled and before the document is stabilized and saved. Straight " " and ' ' and both curly
'                             pairs count; BOTH ends must match, so a stray mark at one end only is left alone, as is a curly open with a
'                             straight close. Each Prodnote paragraph is tested on its own first, then any RUN of consecutive Prodnote
'                             paragraphs in which that found nothing is tested as one note - a DAISY prodnote can arrive as several
'                             paragraphs with the opening mark on the first and the closing mark on the last. Paragraph-first is what keeps
'                             two separately-quoted notes sitting back to back from being read as a single quoted block
'           - Lp - 7/30/2026 - border weights now come from ONE place, Lp_Border_Weight_For_Base_Font. There were four rules and they disagreed.
'                             Lp_Set_Table_Border_Weights and Lp_SetSelectedTableBorderWeight matched the base size as TEXT against fifteen literals
'                             ("14","16",..."42"), so any ODD size and anything ABOVE 42 fell through every branch and the border was left untouched -
'                             a Table Tools color button appeared to do nothing. ApplyTableBorders had its own Select Case that disagreed on 14 sizes:
'                             18-21 came out at 3 pt where the box and reference-page borders were 2 1/4, and anything above 42 hit Case Else and got
'                             the THINNEST border of all, exactly where it should have been heaviest. Jerry: the weights should match everywhere, and
'                             Print Pg Num is the one that is right - so its thresholds are now the shared rule. Sizes 14-42 EVEN are unchanged.
'                             249 lines went, four near-identical 48-line blocks collapsing into one loop each. Dx_Red_Border_Images keeps its fixed
'                             6 pt, which is a red border on images and nothing to do with font size
'           - Sh - 7/26/2026 - macros now return the user to where the cursor was when they started, instead of dumping them at the top of the document. New Sh_Save_User_Position / Sh_Return_User_To_Start_Position pair records a CHARACTER OFFSET (Sh_Start_Pos / Sh_Start_Doc), not a bookmark; 43 macros and 6 UserForms converted. Sh_Pos_Depth makes only the OUTERMOST macro return the user, so the File_Fix_Sequence orchestrators scroll once instead of twenty times; Sh_Pos_Saved refuses a return that had no matching save; RibbonAction clears both per button press so a macro that stops on an error cannot wedge them
'           - Sh - 7/26/2026 - why the old TempPlaceholder bookmark could not do that job, three ways: (1) the copy-to-temp-doc macros paste back over the range the bookmark spans, so Word discards it and Sh_Move_To_And_Delete_Placeholder_Bookmark silently found nothing - its On Error GoTo ExitSub hid the failure; (2) one shared bookmark name, so any macro calling another had its mark deleted and recreated by the inner one; (3) Sh_Create_Temp_Bookmark uses ActiveDocument, so it landed in the temp file whenever one was active. The three TempPlaceholder helpers were left unreferenced, and deleted on 8/3/2026
'           - Sh - 7/26/2026 - also fixed: the restore now happens AFTER ScreenUpdating goes back on - 13 macros (10 of them Dx_) moved the cursor while the screen was frozen, leaving the insertion point correct but the window still showing the top of the document. Lp_Fix_Common_File_Errors had its ScreenUpdating restore commented out entirely, and marked its spot only after Selection.Collapse had already moved it; Lp_Format_Exercise_Lv_1_and_Lv_2 called the restore twice
'           - Sh - 7/26/2026 - LP_Picture_Alignment_Form called Sh_Create_Temp_Bookmark where it meant the move-and-delete, so it never returned the user AND left a stale bookmark behind - which Lp_Bakgrnd_Picture_Menu_Form had been jumping to, since that form called the return without ever creating a mark. The two forms were accidentally coupled through the one shared bookmark name
'           - Sh - 7/26/2026 - deliberately NOT converted, they are meant to leave you where they finish: Lp_Validate_Dollar_PG, the LP/Dx export and import selection macros, Sh_Move_Paragraph_To_Next_Page, Sh_Copy_Ref_Pg_Tags_To_Temp_File (parks you in a temp document on purpose and says so), and Lp_Attach_The_Template / Lp_Attach_Lp_Template (they Save As, so the document is no longer the one the position was measured in)
'           - Sh - 7/26/2026 - KEEP COMMENT LINES UNDER ~650 CHARS. VBA's limit is 1023 per line; a 2252-char changelog line here made Word hang forever on $doc.Save() during the headless build - no error, no event-log entry, Word idle at ~2s CPU. Cost most of a day on 7/26/2026
'           - Sh - 7/24/2026 - deleted two unreferenced UserForms (45 -> 43): Lp_Columns_Wanted_Formx (a stale duplicate of Lp_Columns_Wanted_Form, differing only in ClientHeight) and Dx_Create_Brl_Bullets_Form (orphaned - no braille bullet-CREATION sub exists any more, and nothing reads the Dx_GP_Counter_1 / Dx_GP_String_2 "Keep_Bullets"/"Remove_Bullets" values it set; the surviving Dx_Remove_Bullets uses Dx_Bullet_Removal_Form). Neither was shown, loaded, or named anywhere in the VBA or the ribbon
'           - LP - 7/24/2026 - Styles pane no longer gets reset every time an LP dialog opens: the 14 LP UserForms dropped the "MS_Set_Word_Config_For_Large_Print" call from UserForm_Initialize (opening the document already runs it). That call reached Lp_Turn_on_Styles_Pane, which forces StyleSortMethod = wdStyleSortRecommended / FormattingShowFilter = wdShowFilterFormattingRecommended -- so a user working with "Select styles to show: All Styles" had the pane sort knocked off alphabetical on every Fill-In Line, File Cleanup, Table Tools, etc. Also drops ~40 Options/AutoCorrect writes and 19 AutoCorrect.Entries deletes from each dialog open
'           - Sh - 7/21/2026 - DAISY/NIMAS -> Word: <prodnote> content (body text and inside tables) is now emitted as <p class="Prodnote"> with an mso-style-name rule so Word's HTML import applies the "Prodnote" paragraph style. New helper Sh_Tag_Prodnotes_As_Prodnote_Style handles both real shapes (NIMAS bare-text prodnotes, DAISY prodnotes containing <p> children, which span lines and so are string-parsed rather than wildcard-matched). "Prodnote" added to the LpStyles keep-list in Lp_Remove_All_Styles_Except_Lp_Styles so attaching the LP template does not flatten it back to Normal. EXPERIMENTAL - Prodnote style must exist in LargePrintTemplate.dotx
'           - Sh - 7/21/2026 - Sh_Convert_XML_File_To_Word_Document (DAISY/NIMAS -> Word) perf: ScreenUpdating now stays off through the whole import/repaginate region; live spell/grammar check and background pagination are silenced during it (captured + restored before the Save As UI); the HTML imports in Draft view; the fixed DoEvents pauses are trimmed (~18s -> ~2s); the redundant post-Unlink Fields.Update pass is dropped; and images embed via BreakLink without a per-image .Update disk re-fetch
'           - LP - 7/20/2026 - Lp_Remove_All_Styles_Except_Lp_Styles now converts in-use non-LP custom paragraph/linked styles to Normal before deleting them (foreign OCR/import/web-paste body styles are neutralized, not just dropped from the styles pane); built-in styles stay protected by the BuiltIn=False guard; whitelist membership is now an exact comma-delimited match instead of an InStr substring
'           - Perf (Tier 1) - 7/18/2026 - ScreenUpdating discipline: 48 chained cleanup subs (Lp_/Dx_/Sh_/MS_ Fix/Replace/Convert/Remove/Format/AutoTag families) now CAPTURE the prior ScreenUpdating state on entry and RESTORE it on exit (su_Prev) instead of unconditionally forcing True. When run inside a screen-off orchestrator (Lp_Attach_The_Template, Sh_Convert_XML_File_To_Word_Document, the cleanup forms) they no longer each force a full repaint mid-sequence; standalone behavior is identical. Lp_File_Cleanup_Sub_Menu_Form holds updating off across its selected cleanups. No logic change.
'           - LP - 7/18/2026 - Attach performance on large files: Lp_Normalize_Styles sets space-after once at the story level (was a per-paragraph loop), and Lp_Replace_Multiple_Para_Marks_No_Warning walks paragraphs via .Previous instead of indexed paras(i) (~O(n) vs ~O(n^2)); behavior unchanged
'           - LP - 7/18/2026 - Lp_Attach_The_Template / Sh_Convert_XML_File_To_Word_Document: Save As now uses a single Word Dialog object for .Display + .Execute so the file saves under the name the user types (two separate Dialogs() references lost the typed name); also removed the "template has been attached" prompt from the LP attach
'           - LP - 7/18/2026 - Lp_Attach_The_Template and Sh_Convert_XML_File_To_Word_Document now stabilize the document BEFORE saving, so each writes the file only once (attach/convert -> stabilize -> save) instead of save -> stabilize -> save
'           - LP - 7/18/2026 - MS_Set_Word_Config_For_New_Install now writes Options/AutoCorrect only when they differ (idempotent), so it no longer triggers Office's "restart to apply privacy settings" notice on new docs
'           - LP - 7/2/2026 - revision to file cleanup and normalization
'           - LP - 6/26/2026 - Revison of non-modal messaging - template normalizion optimized
'           - LP - 6/22/2026 - Revison of non-modal messaging
'           - LP - 6/22/2026 - ajusted timeing in Lp_Table_Convert_Options_Form using multiple "do events" to make sure that the proper doc is displayed at the end of the procedure
'           - LP - 6/22/2026 - make "Okay" the default button for retun key - Fixed label problems - improved directions for customization
'           - LP - 5/18/26 - Fixed bug in replacing all consecutive Para Marks  - Lp_Replace_Multiple_Para_Marks_No_Warning
'
'           - Braille - added a "?" to open-ended equations in the cleanup - 3+4=(para mark) becomes 3+4=?(para mark)
'
' versions and dates are displayed in:
'                 Dx_About_Title_And_Agreement_Form
'                      and
'                 Lp_About_Title_And_Agreement_Form
'
'**************************************************************************************************************************
' DBT template translation properties (from SWIFT) are store in the word document header with path "docProps/custom.xml"
' The .docx file is a zip file - change the extention from .docx to .zip and extract all to see contents
' Valid values are:  "English (UEB) - BANA with Nemeth.dxt"
'                    "English (UEB) - BANA.dxt"
'                    "English (BANA Pre-UEB Textbook DE) - BANA.dxt"
'                    "English (BANA Pre-UEB Textbook DE) - BANA Nemeth.dxt"
'**************************************************************************************************************************
'
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
'------------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------------
' set public variables
'------------------------------------------------------------------------------------
Option Explicit

Public MS_Word_Config As String
Public Dx_GP_String_1 As String
Public Dx_GP_String_2 As String
Public Dx_GP_Counter_1 As Integer
Public Dx_GP_Boolean_1 As Boolean

Public Dx_BANA_Template_Name As String
Public Dx_Attached_BANA_Template As String

Public Dx_UEB_EBAE_Boolean As Boolean
Public Dx_UEB_EBAE_String As String

Public Lp_GP_String_1 As String
Public Lp_GP_String_2 As String
' Set by Lp_Attach_Lp_Template on EVERY run, read by Lp_Attach_The_Template. It has its own
' variable because the general-purpose Lp_GP_String_1 it used to live in is written by other
' features and never cleared - see the note at the top of Lp_Attach_Lp_Template. 8/12/2026
Public Lp_Doc_Was_Already_LP As Boolean

' The markers Remove_Txt_Bxs_And_Frames leaves around rescued text, so a transcriber can find
' where it came from. Sh_Text_Frame_Warning_To_Red colors these, so all three macros have to
' agree on the wording character for character - which is why they are constants. 8/13/2026
Public Const SH_TBX_BELOW As String = "<CONTENT FROM A TEXT BOX OR FRAME IS BELOW>"
Public Const SH_TBX_ABOVE As String = "<CONTENT FROM A TEXT BOX OR FRAME IS ABOVE>"
Public Lp_GP_String_3 As String
Public Lp_GP_Boolean_1 As Boolean
Public Lp_GP_Counter_1 As Integer

Public Lp_Pic_Percent As String
Public Lp_Pic_All_Selectd As String

Public Sh_GP_String_1 As String
Public Sh_GP_String_2 As String   ' the conversion's image choice, "KEEP" or "OMIT"
Public Sh_GP_Boolean_1 As Boolean
Public Sh_GP_Counter_1 As Integer

' Where the user's cursor was when a macro started, so the macro can put them back.
' A character offset, NOT a bookmark: the TempPlaceholder bookmark cannot survive the
' copy-to-temp-doc / paste-back macros (the paste replaces the range the bookmark spans),
' is a single shared name that nested macros overwrite, and lands in whichever document
' happens to be active. A number has none of those failure modes.
Public Sh_Start_Pos As Long        ' Selection.Start when the outermost macro began
Public Sh_Start_Doc As String      ' the document it was measured in
Public Sh_Pos_Depth As Long        ' nesting level; only the outermost macro returns the user
Public Sh_Pos_Saved As Boolean     ' guards against a return with no matching save (would jump to a stale spot)

' Large Print Page and Font Settings
Public Lp_Base_Font_Size As String

' The typeface a large print book is set in. It is NOT stored anywhere of its own: it is read
' back off the document's Normal style by Lp_Get_Doc_Setup_Params, so it survives closing and
' reopening for free. 8/8/2026.
'
' From 8/20/2026 it is no longer a choice. New books are Tahoma; this still exists because an
' older book may be set in something else and is allowed to keep it - see
' LP_FONT_LEGACY_LEGIBLE below.
Public Lp_Base_Font_Name As String

' Deliberately UPPERCASE, and it is not a style choice. tools/lib/check_form_calls.py finds
' declarations with ^\s*(Public|Private|Dim|Global)\s+(\w+) -- on "Public Const Lp_Font_Legible"
' that captures the word "Const", not the name, so a form mentioning an Lp_-prefixed constant
' would fail the build as an undefined macro. LP_ is not one of the prefixes it checks.
Public Const LP_FONT_TAHOMA As String = "Tahoma"

' The bundled typeface VistaType LP shipped from 3.0.101 (8/8/2026) to 3.0.196 (8/20/2026), and
' no longer ships. It was Atkinson Hyperlegible rescaled so that a point size set in Word matched
' the printed letter size, and for those twelve days it was the default face for a large print
' book.
'
' It was dropped because its character set is a Latin one. English, French, German, Spanish and
' Italian are fine; Latin, mathematics, the IPA, and the Greek that runs through medical work are
' not - and a character the face has not got is substituted SILENTLY, out of some other typeface,
' at some other size. In large print that is unacceptable: the whole promise is that every
' character comes out at the size the reader asked for. Tahoma covers all of it. (Jerry,
' 8/20/2026.)
'
' The NAME stays because books were produced in it - Jerry's and two testers'. This constant now
' does one job: recognize such a book so that attaching the template again LEAVES ITS TYPEFACE
' ALONE, instead of rewriting it to Tahoma and moving every page break in a book that may already
' be in a reader's hands. Nothing offers this face and nothing installs it. Do not add either
' back.
Public Const LP_FONT_LEGACY_LEGIBLE As String = "VistaTypeLP Legible"

' The bundled typeface from 8/22/2026, and the answer to why the last one was dropped. It is
' Noto Sans, rescaled the same way Legible was so that a point size set in Word matches the
' printed letter size - Jerry measured 37.5 point against the ruler's 36, and 37.5/36 is exactly
' 25/24, so the face is built at 960 units per em instead of 1000 and nothing is rounded.
'
' What it fixes is the coverage that killed Legible. Greek is complete, all 49 letters including
' the accented ones, so the Greek running through medical transcription sets at the size the
' reader asked for instead of being substituted silently out of some other face. U+025E and the
' rest of the phonetic characters are there. So is a SLASHED ZERO, which no face we have shipped
' before had: Word cannot switch a slashed zero on - its Advanced font tab offers ligatures,
' number spacing, number forms and stylistic sets and nothing else - so the build makes the
' slashed shape the ordinary zero in the character map, where no setting can reach it.
'
' Four faces ship, not two. Word does not decline to italicize a face it has no italic for, it
' slants the upright one, and a slanted copy of a rescaled face is no longer the size the ruler
' says - so the emphasized words in a book would print at the wrong size, which is the exact
' fault the rescaling exists to prevent.
'
' Its line height was retuned to Tahoma's on the way through. Left alone, Noto answers 1.582
' times the point size when Word asks how tall a line is, against Tahoma's 1.207 - a quarter
' fewer lines on a page and about a third more paper, with Single spacing selected and nothing
' on screen to explain it. Measured in Word afterwards, both faces now give 21.76pt at 18 point,
' 28.96 at 24 and 43.51 at 36.
'
' Built by tools/lib/build_vistatypelp_sans.py from the Google release of Noto Sans, under the
' SIL Open Font License. Renamed because the OFL requires it.
Public Const LP_FONT_SANS As String = "VistaTypeLP Sans"

' Punctuation that must sit hard against a fill-in line: "____." never "____ ." So no space is
' put after the fill when one of these follows, and a fill to the right margin stops one
' character short to leave the punctuation somewhere to stand (Jerry, 8/10/2026).
'
' Quotation marks are deliberately absent: a closing one wants no space in front and an opening
' one does, and in straight-quote form they are the same character.
Public Const LP_TIGHT_PUNCTUATION As String = ".,;:!?)]}"

' Set while the add-in itself opens a document that is none of the user's business -- currently
' only the license, from Sh_Show_Full_License. Sh_HandleDocumentOpened checks it and leaves
' such a document completely alone, so reading the license cannot reconfigure Word or reset the
' transcriber's Styles pane. 8/3/2026.
Public Sh_Skip_Open_Handler As Boolean
Public TOCTabSetting As String
Public PPH As String  ' Print Page Height
Public PPW As String  ' Print Page Width
Public PTM As String  ' Print Page Top Margin
Public PBM As String  ' Print Page Bottom Margin
Public PLM As String  ' Print Page Left Margin
Public PMM As Boolean ' Is Print Page Mirrored
Public PPG As String  ' Print Page Gutter Size (inches)
Public PRM As String  ' Print Page Right Margin
Public PPO As String  ' Print Page Orintation (L=Landscape, P=Portrait)
Public DM As String   ' Document Media - can be "Paper" (P) or "Screen" (S)
Public MirrorString As String ' yes or no

' --- Application-event wiring -------------------------------------------------------------
' The add-in ships in Word's STARTUP folder, where AutoOpen/AutoNew/AutoClose do NOT fire per
' document (only AutoExec fires from a STARTUP global template). So document-type detection is
' driven by Word application events (see the VtEvents class), hooked once in AutoExec.
' gEvents is module-level so the event sink survives for the whole Word session. If hooking
' fails, gEvents stays Nothing and the Auto* stubs below run the same logic - keeping behavior
' correct if the add-in is instead loaded as Normal.dotm (the old deployment).
Dim gEvents As VtEvents
Public Sh_LastDocEvent As String   ' diagnostic breadcrumb: last document event handled

' --- The spelling and grammar settings braille and large print switch OFF ---------------
' Five settings belong here, and until 8/18/2026 nothing ever put any of them back. Braille
' switches off all five - correctly: the grammar checker and the contextual speller flag
' braille formatting endlessly - and large print switches off the last two. The DEFAULT
' configuration switched none of them on, so ONE braille file left grammar checking off in
' that file, in the transcriber's letters, in every document for the rest of the session, and
' in every session afterwards, because Word remembers them. She had no reason to connect it
' to VistaType. Jerry, 8/18/2026: note what she had and put it back, the same way everything
' else follows the document.
'
' Seeded by AutoExec, before any configuration has run, so the values captured are the ones
' Word loaded from her own settings. Re-captured whenever a book or braille configuration is
' applied while the ORDINARY one is genuinely in force (Sh_ConfiguredAs = "DEF"), which is how
' a change she makes mid-session is picked up. Put back by MS_Set_Word_Config_For_New_Install.
'
' The known limit, and it is unavoidable: if the last session ended with a braille file open,
' Word starts with braille's values and that is what gets captured. She sets them once more in
' an ordinary document and they hold from then on.
' They are kept in a FILE, not in variables here. VBA's End statement resets every module-level
' variable in the project, and this project runs End on perfectly ordinary paths - 34 times in
' this module alone, plus the dialogs' Cancel buttons, and Sh_Is_Doc_Open runs it whenever no
' document is open. Held in memory, one Cancel in a braille file would lose what she had; the
' next return to an ordinary document would put nothing back, and the switch after that would
' record braille's own switched-off values as her preference, for good. A file survives End,
' survives Word closing, and can be carried to a new machine. See Sh_Settings_File.
Private Const VT_STORE_FOLDER As String = "VistaType LP Settings"
Private Const VT_STORE_FILE As String = "VistaType.ini"
Private Const VT_STORE_MINE As String = "TranscriberSettings"

' The second half of the ledger: what the large print or braille configuration most recently
' WROTE. It is what makes "she changed it" tellable from "a book set it" - on the way back to an
' ordinary document, a tracked setting that no longer matches this is one she changed while
' working, and that is a preference. See Sh_Restore_Transcriber_Settings.
Private Const VT_STORE_BOOK As String = "BookApplied"

' The shape of the store, so a file written by an older build is thrown away whole rather than
' half-read. Bump it whenever Sh_Tracked_Settings changes. "1" was the six spelling and grammar
' settings of 8/18/2026; "2" the twenty-seven of 8/20/2026; "3" the thirty-one of 8/21/2026,
' when smart quotes and hyperlinks came back as book settings; "4" the thirty-four of 8/21/2026,
' when the three AutoFormat As You Type boxes large print had never written were added; "5" the
' thirty-six of 8/21/2026, adding first-indents and the spelling-checker suggestions setting; "6"
' the thirty-five of 8/21/2026, when first-indents came back out as the wrong property; "7" the
' thirty-six of 8/21/2026, adding CheckSpellingAsYouType; "8" the thirty-four of 8/22/2026,
' when the two spelling-checker settings left the ledger - the ordinary configuration forces
' them now as well, so there is no value of hers left to give back.
Private Const VT_STORE_STAMP As String = "StoreVersion"
Private Const VT_STORE_STAMP_NOW As String = "8"

' The large print template's file name, and as of 8/20/2026 the ONE fact that decides whether a
' document is a large print document - see Lp_Is_The_Attached_Template_LP. Compared by NAME and
' never by path: the template lives in the user's Templates folder, and that folder differs
' between machines and between Office versions. It is the only large print template name the
' add-in has ever shipped.
Private Const LP_TEMPLATE_FILE As String = "LargePrintTemplate.dotx"

' Which of the three configurations is in force. Word's Options and AutoCorrect entries are
' per-APPLICATION, not per-document, so with an LP book and a braille file both open only one
' of them can be active at a time. This remembers which, so that switching documents can put
' the right one back and Sh_Doc_Info's "Word is configured for..." line (MS_Word_Config)
' describes the document actually in front of the transcriber. See Sh_HandleDocumentActivated.
'
' Written by the three MS_Set_Word_Config_* subs THEMSELVES, beside MS_Word_Config, and not by
' the callers - about a dozen places run those subs directly (the braille and LP cleanup
' sequences among them). Recording it anywhere else would leave this saying "DEF" while Word
' actually carried the braille AutoCorrect entries, and the "same type, nothing to do" test
' below would then believe a lie.
Public Sh_ConfiguredAs As String       ' "LP", "BRL" or "DEF" - empty until the first config

' See Lp_Split_Ordered_List_Sequence.
Public Lp_ListSplitError As String

' What Lp_Horz_To_Vert_List_Form's OK button chose. The form records these and closes; the work
' happens in Lp_Horz_List_To_Vertical afterwards, never inside the form's own event handler.
Public Lp_Hv_Kind As String
Public Lp_Hv_Sort_Wanted As Boolean
Public Lp_Hv_Go As Boolean

' Which round trip Sh_Copy_To_Temp_Doc took, so Sh_Copy_From_Temp_Doc can take the same one
' back. Asking the document again would be wrong - by then the active document IS the temp one.
Public Sh_TempDocRoute As String

' True while a configuration is being applied because the transcriber SWITCHED documents, as
' opposed to opening one or attaching a template. The three MS_Set_Word_Config_* subs then set
' Word's typing behavior - Options, AutoCorrect, spelling, fractions, the things that are wrong
' when they belong to the other document - and leave the screen alone: the Styles pane, the
' formatting marks, both rulers, the style area and the view type.
'
' Jerry's call, 8/9/2026. Once a transcriber has arranged their screen for a session it is
' theirs, and having the pane close and the pilcrows come back on every time they glanced at the
' other document would be the 7/24 and 8/2/2026 complaints all over again. Opening a document,
' or attaching a template, still sets the display as it always did.
'
' Only Sh_Apply_Word_Config sets this, and it always clears it again, including when the
' configuration raises.
Public Sh_Config_Skip_Display As Boolean

' How many times Word has raised WindowActivate since this session began - counted before
' any of Sh_HandleDocumentActivated's guards, purely so the event's arrival can be told
' apart from a guard turning it away. Read it with Sh_GetActivateCount.
Public Sh_ActivateCount As Long

' The last document Sh_HandleDocumentActivated looked at, and what was in force when it did.
' WindowSelectionChange fires on every cursor movement, so the handler needs a first test that
' costs one string comparison. Pairing the name WITH Sh_ConfiguredAs is what makes it safe to
' skip: attaching a template changes a document's type while it stays active, and that also
' changes Sh_ConfiguredAs, so the pair stops matching and the document is examined afresh.
Public Sh_LastSeenDoc As String
Public Sh_LastSeenAs As String

' 8/20/2026 - three module variables lived here, Sh_PaneHiddenForDoc, Sh_PaneWasStyles and
' Sh_PaneWasNav. They existed to undo a pane hide that turned out not to be a close, because
' DocumentBeforeClose fires on close ATTEMPTS and nothing says afterwards whether the attempt
' succeeded. Recording what was hidden and reconsidering it on the next activation LOOKED right
' and did not work: Jerry, testing 3.0.207, cancelled at the "save your changes?" prompt and had
' to put both panes back by hand. Sh_Tidy_Panes_After_Close replaced the whole idea - it does not
' guess whether the close happened, it waits a second and looks. See it, and do not bring these
' variables back.

Sub AutoExec()
    ' Runs once when Word starts (fires even from a STARTUP global template, unlike AutoOpen).
    On Error Resume Next

    ' FIRST, before a document can be configured: note the settings a book would take away while
    ' they are still the transcriber's own. See the declarations at the top of this module.
    Sh_Save_Transcriber_Settings

    Set gEvents = New VtEvents
    Set gEvents.App = Application
    If Not gEvents Is Nothing Then
        If gEvents.App Is Nothing Then Set gEvents = Nothing   ' event hook failed
    End If
    ' A document opened as part of Word startup opens before this hook exists; handle it now.
    If (Not gEvents Is Nothing) And Documents.count > 0 Then Sh_HandleDocumentOpened
End Sub

Public Function Sh_GetLastDocEvent() As String
    Sh_GetLastDocEvent = Sh_LastDocEvent
End Function

Public Function Sh_GetActivateCount() As Long
    Sh_GetActivateCount = Sh_ActivateCount
End Function

Sub AutoNew()
    ' Back-compat stub: only acts if app events aren't hooked (i.e. loaded as Normal.dotm).
    If gEvents Is Nothing Then Sh_HandleDocumentNew
End Sub

' Version: 1.1  Date: 8/18/2026 - skip while a macro is running (Jerry, 8/18/2026)
' Version: 1.0  Date: 8/9/2026
Sub Sh_HandleDocumentNew()
    ' A MACRO is running, not a transcriber. The same test, for the same reason, as the one at
    ' the top of Sh_HandleDocumentActivated - see the long note there - but it was missing here,
    ' and Word raises NewDocument for EVERY Documents.Add, including the ones macros make.
    ' Lp_Copy_To_Temp_Doc and Dx_Copy_To_Temp_Doc create the scratch document from about 35
    ' places, so a book being cleaned up had Word reconfigured for an ORDINARY document half way
    ' through the job. Three costs, none of them visible: the Styles pane closed and the screen
    ' refreshed at the exact points the macro had turned refreshing off - the 7/24 and 8/2/2026
    ' complaints arriving by another door; the whole configuration written twice, once on the way
    ' in and once when the book came back; and, from 8/18/2026, the eighteen compact fraction
    ' entries added and then deleted again, 36 writes to Word's AutoCorrect store per macro run,
    ' which is the churn behind Office's "restart to apply your privacy settings" notice.
    ' Screen updating being off is the marker: every one of those macros turns it off, and a
    ' transcriber creating a document has it on.
    '
    ' Which transcriber-facing documents this actually reaches was checked site by site on
    ' 8/18/2026, because guessing got it wrong first time. The DAISY/NIMAS conversion creates its
    ' document with screen updating still ON (it goes off 22 lines later), and the two export
    ' macros make theirs Visible:=False and save and close it inside the macro - so none of those
    ' three changes at all. The ONE that does is the blank document Dx_Attach_BANA_Template makes
    ' when nothing is open, and that one is fine: it runs MS_Set_Word_Config_For_Braille directly
    ' a few lines later.
    '
    ' One thing to know before adding another Documents.Add for the transcriber inside a screen-off
    ' macro: the later pick-up by Sh_HandleDocumentActivated sets the TYPING side only - it passes
    ' DisplayToo False - so such a document would never get formatting marks, the rulers, Print view
    ' or the Styles pane. Nothing hits that today. Configure it in the macro itself, as
    ' Dx_Attach_BANA_Template does.
    If Not Application.ScreenUpdating Then Exit Sub

    ' Written only when work actually happens - see the breadcrumb note in Sh_HandleDocumentActivated.
    Sh_LastDocEvent = "NewDocument"
    Sh_Apply_Word_Config "DEF"
End Sub

' --- Switching between open documents -------------------------------------------------------
' Word carries ONE set of Options and AutoCorrect entries for the whole application, so a large
' print book and a braille file open side by side cannot both have their configuration in force.
' Until 8/9/2026 whichever document was opened LAST won, and went on winning: clicking back to
' the other one left Word set up for its neighbor, and Sh_Doc_Info kept reporting - truthfully,
' but uselessly - the configuration of a document the transcriber was no longer looking at.
'
' VtEvents.App_WindowActivate calls this whenever a document window is activated. That is often:
' clicking another window, Ctrl+F6, View > Switch Windows, and returning to Word from another
' program all raise it, as does opening or creating a document - and so does every macro that
' works in a temporary document. So the whole job is four cheap tests and, only when the answer
' has actually changed, one configuration:
'
'   a macro is running          -> nothing to do. See the ScreenUpdating test below; this is the
'                                  one that keeps the change off the macros' backs
'   the add-in's own window     -> nothing to do (the GPL text opened by Sh_Show_Full_License)
'   same TYPE as the config     -> nothing to do. The Options are already right, and the
'   already in force               per-document settings (style sort, formatting filter) and the
'                                  per-window ones (rulers, show/hide, view) travel with the
'                                  document and its own window. This covers two windows on one
'                                  book, Word merely regaining focus, and clicking between two
'                                  documents of the same kind
'   a different type            -> put that document's configuration in force
'
' There is deliberately NO "same document as last time" shortcut. It would be faster, but a
' document's type can CHANGE while it stays active - attaching the LP template, or a BANA one -
' and a name-based shortcut would then never look at it again. Asking the document what it is
' costs one style lookup, which is nothing beside what Word itself does on a click.
'
' What this must NEVER do is the rest of Sh_HandleDocumentOpened's work. That one can raise the
' obsolete-template message box and open the attach dialog - right once, when a document is
' opened, and intolerable on every click between two documents.
'
' Nothing here may raise: an error inside a Word application event can stop Word calling back
' for the rest of the session, which would silently kill document-type detection outright.
'
' Version: 1.4  Date: 8/20/2026 - the Styles pane recovery of 1.1 to 1.3 is GONE. Nothing on the close
'                               path hides a pane any more - see Sh_Tidy_Panes_After_Close
' Version: 1.0  Date: 8/9/2026 (Jerry asked for it: the configuration should follow the document)
Sub Sh_HandleDocumentActivated()
    Static busy As Boolean
    Dim newType As String
    Dim docName As String

    On Error GoTo eom

    ' Diagnostic breadcrumb, incremented before every guard below, so that "did the event even
    ' arrive?" and "did a guard turn it away?" can be told apart from outside. Sh_LastDocEvent
    ' cannot answer that: it is only written when work actually happens.
    Sh_ActivateCount = Sh_ActivateCount + 1

    ' Applying a configuration touches the window itself - closing Reading view is the first
    ' thing all three do - which can raise WindowActivate again before this one has finished.
    ' Without this the second pass would start the whole job over.
    If busy Then Exit Sub

    ' A document the add-in opened for its own reasons is not the user's working document.
    If Sh_Skip_Open_Handler Then Exit Sub
    If Documents.count = 0 Then Exit Sub

    ' A MACRO is running, not a transcriber. Word's own macros activate documents constantly -
    ' every Lp_Copy_To_Temp_Doc and Dx_Copy_To_Temp_Doc creates a document and activates it, and
    ' those two are called from about 35 places. Reconfiguring Word in the middle of one would
    ' put back the ~74 Options and AutoCorrect writes that were taken OUT of fourteen dialogs on
    ' 7/24/2026, and would refresh the screen at the exact points the code turns refreshing off.
    ' Screen updating being off is the marker: those macros all turn it off, and a transcriber
    ' clicking a window has it on.
    If Not Application.ScreenUpdating Then Exit Sub

    ' The GPL text stays on screen after Sh_Show_Full_License opens it, so its Sh_Skip_Open_Handler
    ' guard - which only spans the Documents.Open call - does not cover clicking back to it later.
    If Sh_Is_Addins_Own_Document() Then Exit Sub

    ' 8/20/2026 - a Styles pane recovery lived here and is gone. Nothing on the close path hides
    ' anything any more, so there is nothing to put back; Sh_Tidy_Panes_After_Close waits until
    ' the close has settled and only then decides. See it before adding anything of the kind here.
    docName = ActiveDocument.FullName

    ' One string comparison in the overwhelmingly common case - the cursor moved, or the same
    ' document was clicked again. See Sh_LastSeenDoc.
    If docName = Sh_LastSeenDoc And Sh_ConfiguredAs = Sh_LastSeenAs Then Exit Sub

    newType = Sh_Doc_Config_Type()
    If newType <> Sh_ConfiguredAs Then
        busy = True
        Sh_LastDocEvent = "WindowActivate"
        ' The typing side only. The screen the transcriber has arranged stays as they left it -
        ' see Sh_Config_Skip_Display.
        Sh_Apply_Word_Config newType, False
        busy = False
    End If

    Sh_LastSeenDoc = docName
    Sh_LastSeenAs = Sh_ConfiguredAs
    Exit Sub

eom:
    busy = False
End Sub

' True for a document the add-in opened for its own purposes rather than the transcriber's -
' today that means the GPL text under %AppData%\VistaType LP. Sh_Skip_Open_Handler covers only
' the instant such a document is OPENED; the window then stays on screen and can be clicked back
' to at any time, which is why the folder is tested here as well.
'
' Version: 1.0  Date: 8/9/2026
Public Function Sh_Is_Addins_Own_Document() As Boolean
    On Error GoTo eom
    Sh_Is_Addins_Own_Document = _
        (InStr(1, ActiveDocument.Path, Environ$("APPDATA") & "\VistaType LP", vbTextCompare) = 1)
eom:
End Function

' Which of the three configurations the ACTIVE document wants: "LP" large print, "BRL" braille,
' "DEF" everything else. Reads the attached template only - no document body - so it is cheap
' enough to run on every window activation.
'
' That "no document body" claim became true on 8/20/2026 rather than 8/9: until then the large
' print half went through a style lookup. Both halves now ask the attached template's name and
' nothing else, which is the rule as Jerry and the beta tester wrote it.
'
' Version: 1.1  Date: 8/20/2026 - no code change; Lp_Is_The_Attached_Template_LP 2.0 made the
'                               large print half a template test, which is what this always said
' Version: 1.0  Date: 8/9/2026
Public Function Sh_Doc_Config_Type() As String
    Dim attached As String

    Sh_Doc_Config_Type = "DEF"
    On Error GoTo eom

    If Lp_Is_The_Attached_Template_LP = True Then
        Sh_Doc_Config_Type = "LP"
    Else
        attached = ActiveDocument.AttachedTemplate
        If InStr(attached, "BANA Braille") > 0 Then Sh_Doc_Config_Type = "BRL"
    End If

eom:
End Function

' Puts one of the three configurations in force. The three MS_ subs record WHICH, into
' Sh_ConfiguredAs and MS_Word_Config, so this sub does not - see the note on Sh_ConfiguredAs.
'
' Configuration ONLY. No message boxes and no dialogs: this runs when someone merely clicks
' another window, and must never interrupt them.
'
' DisplayToo says whether the on-screen settings go with it. True when a document is OPENED or
' created - that is the moment to set the screen up for the kind of document it is. False when
' the transcriber merely SWITCHED to an already-open document, where the screen they have
' arranged is theirs to keep. See Sh_Config_Skip_Display.
'
' Version: 1.1  Date: 8/9/2026 - DisplayToo (Jerry: switching sets the typing, not the screen)
' Version: 1.0  Date: 8/9/2026
Sub Sh_Apply_Word_Config(ByVal cfgType As String, Optional ByVal DisplayToo As Boolean = True)
    Dim su_Prev As Boolean

    su_Prev = Application.ScreenUpdating
    On Error GoTo failed

    Sh_Config_Skip_Display = Not DisplayToo

    Select Case cfgType
        Case "LP"
            Application.Run MacroName:="MS_Set_Word_Config_For_Large_Print"
            ' Lp_Set_Display_For_Large_Print is display from top to bottom - show/hide, both
            ' rulers, print view, the formatting flags - so on a switch it is skipped whole.
            If DisplayToo Then Application.Run MacroName:="Lp_Set_Display_For_Large_Print"
        Case "BRL"
            Application.Run MacroName:="MS_Set_Word_Config_For_Braille"
            If DisplayToo Then Application.TaskPanes(wdTaskPaneFormatting).Visible = False
            Dx_GP_String_1 = "Doc_Is_Already_Brl"
        Case Else
            Application.Run MacroName:="MS_Set_Word_Config_For_New_Install"
    End Select

    Sh_Config_Skip_Display = False
    Exit Sub

failed:
    On Error Resume Next
    Sh_Config_Skip_Display = False
    ' MS_Set_Word_Config_For_Braille turns screen updating off near its top and back on near its
    ' bottom and has no error handler of its own, so anything raising in between used to leave
    ' Word not repainting at all. On this path that would be far worse than before: the next
    ' click would run the same failing config and freeze the screen again, and again. Put it
    ' back before anything else.
    Application.ScreenUpdating = su_Prev
    ' Say nothing is in force rather than claim a configuration that never finished, so the next
    ' switch tries again instead of believing the job is done.
    Sh_ConfiguredAs = ""
End Sub

Sub AutoOpen()
    ' Back-compat stub: only acts if app events aren't hooked (i.e. loaded as Normal.dotm).
    If gEvents Is Nothing Then Sh_HandleDocumentOpened
End Sub

Sub Sh_HandleDocumentOpened()
    ' A document the add-in opened for its own reasons - the license - is not the user's
    ' working document and must not be configured as one. See Sh_Show_Full_License. 8/3/2026.
    If Sh_Skip_Open_Handler Then Exit Sub

    Sh_LastDocEvent = "DocumentOpen"
    '
    ' Runs for every opened document - via VtEvents.App_DocumentOpen (STARTUP), or AutoOpen
    ' when loaded as Normal.dotm.
    ' If the document is a large print document then setting for Large Print are made - if doc is braille then brille settings are made
    '   otherwise the settings for a normal document are made.
    '
    ' Version 1.7  Date: 8/9/2026 - the three-way config now goes through Sh_Apply_Word_Config, shared with Sh_HandleDocumentNew and with switching between open documents (Sh_HandleDocumentActivated), so the three can never drift apart. Same calls in the same order. One behavior change: a configuration that raises is now caught inside Sh_Apply_Word_Config, so this sub carries on to Sh_Set_Prodnote_Style_Visibility instead of jumping to eom and skipping it
    ' Version 1.6  Date: 7/24/2026 - Sh_Set_Prodnote_Style_Visibility now runs for EVERY opened document (any template), not just large print, so the Prodnote style is removed from the Styles pane whenever the document contains no prodnotes regardless of the attached template
    ' Version 1.5  Date: 7/23/2026 - LP documents now called the Prodnote visibility helper on open (superseded by 1.6)
    ' Version 1.4  Date: 2/16/2026 - added call to p_CheckAndAssistDocumentState to check block and read only status
    ' Version 1.3  Date: 2/17/2024 - added Dx_GP_String_1 = "Doc_Is_Already_Brl" to bypass cleanup questions
    ' Version 1.2  Date: 10/22/2021 - Added section to determine if doc has obsolete lp template attached
    ' Version 1.1  Date: 11/9/2020 - added "Application.Run MacroName:="Lp_Set_Display_For_Large_Print"
    ' Version 1.0  Date: 10/28/2020
    '
    ' Author: Jerry Whittaker - jerry@thewhittakers.org
    '
    On Error GoTo eom
    ActiveDocument.ActiveWindow.View.ReadingLayout = False
    
    '************** start automatic word configuration ***************************
    On Error GoTo 0

    ActiveDocument.ActiveWindow.View.ReadingLayout = False

    ' Three cases, and the ORDER of the first two is what makes the obsolete-template warning
    ' possible. 8/20/2026: document type is now read from the attached template, so the first
    ' test is simply "is this a large print document". A book made on an OBSOLETE template is
    ' not one - its template is not ours - and would fall straight through to the ordinary
    ' configuration with no warning, which is why the second test exists and why it asks the
    ' looser question. See Lp_Was_Made_As_An_Lp_Book.
    '
    ' Written as ElseIf rather than nested, because the two tests are now independent questions
    ' about the document rather than one refining the other. The old shape asked Box Black first
    ' and the template second, which could not see a document carrying our template WITHOUT the
    ' style - rare, but it read as an ordinary document and was configured as one.
    If Lp_Is_The_Attached_Template_LP = True Then
        ' The current large print template is attached. The ordinary case, and much the commonest.
        Sh_Apply_Word_Config "LP"

    ElseIf Lp_Was_Made_As_An_Lp_Book() = True Then
        ' Made as a large print book, but not on the template we ship today - an older template,
        ' or one Word reset to Normal when the file was mailed or the template was moved. She
        ' must re-attach before editing, because character and line spacing differ between
        ' templates and the text will re-flow.
            
        Application.Run MacroName:="Lp_Get_Doc_Setup_Params"
        Dim OrientName As String
            
        If PPO = "L" Then 'discovered by Lp_Get_Doc_Setup_Params
            OrientName = "Landscape"
        Else
            OrientName = "Portrait"
        End If
                
        MsgBox " This document is using an obsolete large print template." & vbCr & vbCr _
                 & "YOU MUST ATTACH THE LATEST TEMPLATE IN ORDER TO CONTINUE EDITING THIS DOCUMENT!" & vbCr & vbCr _
                 & "When this message is closed, the 'Attach Lp Template & Select Output Media' form will be displayed. " _
                 & "Make note of the following information regarding this document to assist you in your media and font size selections on that form." & vbCr & vbCr _
                 & "      Font Size                           = " + Trim(Lp_Base_Font_Size) & vbCr _
                 & "      Paper/Screen Height        = " + Trim(Str(PPH)) & vbCr _
                 & "      Paper/ScreenWidth          = " + Trim(Str(PPW)) & vbCr _
                 & "      Top Margin                        = " + Trim(Str(PTM)) & vbCr _
                 & "      Bottom Margin                  = " + Trim(Str(PBM)) & vbCr _
                 & "      Left Margin                        = " + Trim(Str(PLM)) & vbCr _
                 & "      Right Margin                     = " + Trim(Str(PRM)) & vbCr _
                 & "      Mirrored Margins              = " + MirrorString & vbCr _
                 & "      Binding Width                    = " + PPG & vbCr _
                 & "      Orientation                        = " + OrientName & vbCr _
                 & "      Output Media Type           = " + DM & vbCr & vbCr _
                 & "Because of the differences in character and line spacing between the templates, attaching the latest template may " _
                 & "result in text flow changes which will require editing.", , "VistaType LP (123)"
                            
            Sh_Apply_Word_Config "LP"
            Lp_GP_String_3 = "Bypass Cleanup Checks"
            Application.Run MacroName:="Lp_Attach_Lp_Template"
            Exit Sub

    Else ' check if it is a braille document
        If InStr(ActiveDocument.AttachedTemplate, "BANA Braille") > 0 Then
            Sh_Apply_Word_Config "BRL"
        Else   'document is not Braille and not Large Print
            Sh_Apply_Word_Config "DEF"
         End If
    End If

    ' Prodnote in the Styles pane only when the document uses it -- for EVERY template
    ' (large print, braille, or Normal). The attached-template name must not matter here, so
    ' this runs after the per-template config above rather than inside the large-print branch.
    Application.Run MacroName:="Sh_Set_Prodnote_Style_Visibility"

eom: 'End of Macro

End Sub   '*** end of AutoOpen() macro ***

Sub AutoClose()
    ' Back-compat stub: only acts if app events aren't hooked (i.e. loaded as Normal.dotm).
    If gEvents Is Nothing Then Sh_HandleDocumentClosing
End Sub

Sub Sh_HandleDocumentClosing()
    Sh_LastDocEvent = "DocumentClose"
    '
    ' Runs when a document is closing - via VtEvents.App_DocumentBeforeClose (STARTUP), or
    ' AutoClose when loaded as Normal.dotm.
    ' If the document is a large print document then print view is set. (The styles pane and
    ' crop marks used to be turned off here too; see 1.2.)
    '
    ' Version 1.6  Date: 8/20/2026 - hides nothing itself. It asks Sh_Tidy_Panes_After_Close to look, a second
    '                                from now, at whether the close actually happened. 1.3 to 1.5 hid the panes
    '                                here and tried to undo it if the transcriber cancelled, which did not work
    ' Version 1.5  Date: 8/20/2026 - writes BOTH halves of the navigation pane. 1.4 wrote only
    '                                ActiveWindow.DocumentMap, which is a window property and does nothing
    '                                on a window that is closing - Jerry, testing 3.0.207
    ' Version 1.4  Date: 8/20/2026 - the panes come down for ANY document, not only a large print one, and the
    '                                navigation pane comes down with the Styles pane (Jerry, after testing 3.0.206:
    '                                closing a letter left both up). What was on screen is noted first, so the
    '                                recovery can hand it back - see Sh_PaneWasStyles
    ' Version 1.3  Date: 8/20/2026 - hides the Styles pane again (Jerry: the pane follows the large print document). 1.2 removed it as a plain bug and that bug is real, so it is not simply back - the document is recorded in Sh_PaneHiddenForDoc and Sh_HandleDocumentActivated puts the pane straight back if the close turns out not to have happened
    ' Version 1.2  Date: 8/2/2026 - no longer hides the Styles pane. It is the user's now, and this fired on close ATTEMPTS - cancelling the "save your changes?" prompt left you in the document with the pane gone
    ' Version 1.1  Date:  12/16/2021 - set on error - crashes if image is selected when document is closed    Application.TaskPanes(wdTaskPaneFormatting).Visible = True
    ' Version 1.0  Date: 10/17/2020
    ' Author: Jerry Whittaker - jerry@thewhittakers.org
    '
    ' Version 1.1 put an On Error here because this line "crashes if image is selected when
    ' document is closed", and the guard was lost with the Styles-pane line in 1.2 - the eom label
    ' below is what it left behind. Restored 8/20/2026, when Lp_Attach_The_Template started
    ' closing the document ITSELF after every attach: what used to need the transcriber to close a
    ' document with a picture selected now happens automatically, on every job. An error raised
    ' inside a Word application event can also stop Word calling the sink back for the rest of the
    ' session, which would kill document-type detection silently - see Sh_HandleDocumentActivated.
    On Error GoTo eom
    If Lp_Is_The_Attached_Template_LP = True Then ' is the document being closed an lp doc
        ActiveWindow.ActivePane.View.Type = wdPrintView
        'Application.Options.ShowCropMarks = False
    End If

    ' Ask, a second from now, whether that close actually happened - and hide nothing here.
    '
    ' This sub runs on close ATTEMPTS. Word raises DocumentBeforeClose BEFORE the "save your
    ' changes?" prompt, and nothing afterwards says which button she pressed. Versions 1.3 to 1.5
    ' hid the panes here and tried to undo it on the next activation if the document turned out to
    ' still be there. It read well and it did not work: Jerry, testing 3.0.207, cancelled at the
    ' prompt and had to put both panes back by hand.
    '
    ' So stop guessing and look instead. Application.OnTime cannot run while a modal dialog is up,
    ' which is the whole trick - the tick cannot fire until she has answered the prompt, and by
    ' then the document has either gone or it has not. Sh_Tidy_Panes_After_Close simply looks.
    '
    ' Nothing is hidden on this path at all, so a cancelled close costs her nothing: the panes
    ' never moved, and there is no flicker to explain either.
    On Error Resume Next
    Application.OnTime When:=Now + TimeSerial(0, 0, 1), Name:="Sh_Tidy_Panes_After_Close"
    Err.Clear

eom: 'End of Macro

End Sub  '*** end of AutoClose macro ***

Sub Dx_Attach_BANA_Template()
'
'    Attaches BANA Template
'    Changes view to draft
'    Converts word foreign language tags to BANA template styles
'    Turns show-all status on
'
'  Version: 3.4  Date: 8/18/2026 - that font is now Times New Roman 14 pt, not Courier New 12 pt - matching what Duxbury's own SWIFT add-in sets when IT attaches a template (Jerry)
'  Version: 3.3  Date: 8/5/2026 - puts the whole document into Courier New 12 pt as the last thing it does (Jerry)
'  Version: 3.2  Date: 7/24/2026 - repaint (ScreenUpdating on + ScreenRefresh) before the template-choice and translation-choice forms; they were shown while ScreenUpdating was off, so the Word workspace behind them rendered black instead of the normal gray
'  Version: 3.1  Date: 6/29/2025 - added Copy BANA Braille Template From Word Startup Folder to Templates Folder - Duxbury began
'                                  to place the BANA Braille 2025.dotx in the Word startup folder - need to copy to templates folder
'                                  or this attachment routine will not work - the template file will reside in both folders
'  Version: 3.0  Date: 3/4/2024 - removed "End" before end of macro - doc info now shows Word Configuration
'  Version: 2.9  Date: 2/29/2024 - removed check to see if BANA template was installed - redundant code
'  Version: 2.8  Date: 2/24/2024 - added "end" at end of macro to keep from looping into the File Cleanup Form
'  Version: 2.7  Date: 2/19/2024 - added Ms config for word
'  Version: 2.6  Date: 2/17/2024 - added choices for fix common file errors and multiple para marks
'  Version: 2.5  Date: 1/4/2024 - added Unload Dx_Choose_BANA_Template_Form
'  Version: 2.4  Date: 3/7/2023 - removed question on foreign language - extract text from boxes and frames
'  Version: 2.3  Date: 1/9/2023 - "added On Error Resume Next" to bypass error at "CommandBars("Navigation").Visible = False"
'  Version: 2.2  Date: 2/28/2022 - added fixes for foreign language content
'  Version: 2.1  Date: 4/16/2021 - more improvemets to attach template code
'  Version: 2.0  Date: 10/17/2020 - improved attach template code
'  Version: 1.9  Date: 9/9/2018 - removed fix body text styles - wiped out auto numbers
'  Version :1.8  Date: 1/10/2017
'
' Author: Jerry Whittaker - jerry@thewhittakers.org

    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' If no document is active then create a new blank document
    If Documents.count = 0 Then
        Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
    End If

    'attching the BANA template will wipe out Word's language tags - change
    '     the foreign language tags of Word to DBT style names before attaching the BANA Template
    Application.Run MacroName:="Dx_Fix_Foreign_Languages"

    ' Begin Copy BANA Braille Template From Word Startup Folder to Templates Folder
    ' Copilot prompt: "Copy all "BANA Braille*.dot*" files from Word's Startup folder to the User Templates folder,
    '                  overwriting any existing files with the same name."

    Dim fso As Object
    Dim sourceFolder As String, targetFolder As String
    Dim file As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    sourceFolder = Application.startupPath
    targetFolder = Application.Options.DefaultFilePath(wdUserTemplatesPath)

    If Right(sourceFolder, 1) <> "\" Then sourceFolder = sourceFolder & "\"
    If Right(targetFolder, 1) <> "\" Then targetFolder = targetFolder & "\"

    For Each file In fso.GetFolder(sourceFolder).Files
        If file.Name Like "BANA Braille*.dot*" Then
            fso.CopyFile file.Path, targetFolder & file.Name, True ' True = overwrite
        End If
    Next
    ' End Copy BANA Braille Template From Word Startup Folder to Templates Folder
    
    ' Get BANA Template Choice from user
    Dx_BANA_Template_Name = ""

    ' Repaint before the interactive form/dialogs below. ScreenUpdating has been off since the
    ' top of the macro, so a form shown now would sit over an unpainted (black) workspace
    ' instead of the normal gray. Turn it back off after the choice for the attach that follows.
    Application.ScreenUpdating = True
    Application.ScreenRefresh

    ' If more than one template is found then place the names of all BANA Braille
    ' templates in a list box for the user to select
    If Dx_BANA_Template_Name = "" Then
        Dx_Choose_BANA_Template_Form.Show 'present list box for choice
    End If

    Unload Dx_Choose_BANA_Template_Form

    If Dx_BANA_Template_Name = "" Then ' likely that the user hit the close-window X
        MsgBox "No template selected... attachment canceled", , "Braille Macros"
        End
    End If

    Application.ScreenUpdating = False

    ' attach BANA template
    Dim TemplatePathandName As String
    TemplatePathandName = Options.DefaultFilePath(wdUserTemplatesPath) + "\" + Dx_BANA_Template_Name
    
    With ActiveDocument
        .UpdateStylesOnOpen = True
        .AttachedTemplate = TemplatePathandName
        .UpdateStylesOnOpen = False  ' supresses any further style updates
    End With

    ' configure word settings for braille
    Application.Run MacroName:="MS_Set_Word_Config_For_Braille"
      
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdNormalView
    Else
        ActiveWindow.View.Type = wdNormalView
    End If
    
    ActiveWindow.DocumentMap = False
    On Error Resume Next
    CommandBars("Navigation").Visible = False
    ActiveWindow.ActivePane.View.ShowAll = True
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    ActiveDocument.UndoClear

    ' Repaint again before the translation-type form (the attach/config/view change above ran
    ' with ScreenUpdating off, so the workspace behind this form would otherwise be black).
    Application.ScreenUpdating = True
    Application.ScreenRefresh

    ' Get the braille translation type (UEB or EBAE) - forces a choice
    Dx_UEB_EBAE_String = ""
    Do While Dx_UEB_EBAE_String = ""
        Dx_Choose_Translation_Form.Show
    Loop

    Application.ScreenUpdating = False

    Dx_Attached_BANA_Template = Dx_BANA_Template_Name

    'converts Word Lang Tags into DBT foreign language tags - BANA template must be attached for this to work
    Application.Run MacroName:="Dx_Add_Color_To_Foreign_Language_Words"
    Application.Run MacroName:="Sh_Remove_Txt_Bxs_And_Frames"
    Application.ScreenUpdating = su_Prev ' Turn screen updating on
    Application.ScreenRefresh
    MsgBox (ActiveDocument.AttachedTemplate) + " template has been attached!", , "Braille Macros"

    If Dx_GP_String_1 <> "Doc_Is_Already_Brl" Then
        Dx_GP_String_1 = ""
        If ActiveDocument.Characters.count > 10 Then
            If MsgBox("Do you want to fix common file errors?", vbYesNo, "Braille Macros") = vbYes Then
                Application.Run MacroName:="Dx_Fix_Common_File_Errors"
            End If
            If MsgBox("Do you want to remove multiple consecutive paragraph marks? ", vbYesNo, "Braille Macros") = vbYes Then
                Application.Run MacroName:="Sh_Replace_Multiple_Para_Marks_No_Warning"
            End If
        End If
    End If

    ' LAST, after the optional cleanups above rather than beside the attach itself. Both of them
    ' rewrite text - Dx_Fix_Common_File_Errors alone runs some thirty passes over the document -
    ' so a font set any earlier is not the font the transcriber ends up looking at. (Until
    ' 8/18/2026 this note named Word's AutoFormat command as the reason; that call is gone, but
    ' the passes that replaced it rewrite text just the same, so this still belongs last.)
    Application.Run MacroName:="Dx_Set_Whole_Document_To_Times_New_Roman_14"

End Sub   '***** end of Dx_Attach_BANA_Template macro *****

Sub Dx_Set_Whole_Document_To_Times_New_Roman_14()
'
' Version: 2.0  Date: 8/18/2026 - Times New Roman 14 pt, was Courier New 12 pt (Jerry)
' Version: 1.0  Date: 8/5/2026
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Puts every character of the document into Times New Roman at 14 points, so the transcriber
' reads the whole file in one face and size however mixed the source was.
'
' WHY A PROPORTIONAL FACE, ON BRAILLE SOURCE. It reads like a mistake and is not, so do not
' change it back to a monospace one. Duxbury ships its own Word add-in, SWIFT, and when SWIFT
' attaches a template it sets the document to Times New Roman 14. A transcriber may use both,
' or move a file between them, and a file that changes face depending on which add-in touched
' it last is a file she has to think about. Matching SWIFT is the point. Jerry, 8/18/2026.
' It was Courier New 12 from 8/5/2026 until then, chosen for exactly the fixed-width reason
' you would expect - that reasoning was sound and this one outranks it.

    Sh_Set_Whole_Document_Font ActiveDocument, "Times New Roman", 14

End Sub   '***** end of Dx_Set_Whole_Document_To_Times_New_Roman_14 macro *****

Sub Sh_Set_Whole_Document_Font(ByVal targetDoc As Document, ByVal fontName As String, ByVal fontSize As Single)
'
' Version: 1.0  Date: 8/5/2026
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Puts every character of a document into one face and size.
'
' This is DIRECT formatting laid over the text, not a style change, so it overrides whatever the
' attached template's styles ask for and it survives them. It covers the main text story, tables
' included, and leaves headers and footers alone.
'
' It sets the face and the size ONLY. Color is untouched, so the red $pg tags and the red
' prodnotes a conversion produces come through it unchanged - which is the whole reason it is
' safe to run at the end of one.
'
' Called directly rather than through Application.Run: that marshals every argument as a Variant
' and cannot bind one to a typed "As Document" parameter.

    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    With targetDoc.Content.Font
        .Name = fontName
        .Size = fontSize
    End With

    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub   '***** end of Sh_Set_Whole_Document_Font macro *****

Function Sh_Is_Font_Installed(ByVal fontName As String) As Boolean
'
' Version: 1.0  Date: 8/8/2026
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' True when Word can see the named font on this machine.
'
' Lifted out of Lp_Attach_Lp_Template, where it was written inline for Tahoma alone. It answers
' the question for ANY face, and it is worth keeping for the reason it was written: a missing
' font is the ONE failure here that says nothing at all. Word substitutes silently, the
' substitute has different metrics, and a book that reads 18 point on screen prints at some
' other size.
'
' Its caller changed on 8/20/2026. The attach dialog used to ask about the bundled typeface, and
' that typeface is gone; Sh_Font_Status_Text below now asks on behalf of Sh_Doc_Info, about
' whatever face the document in front of the transcriber is actually set in.
'
' Adapted from: https://code.adonline.id.au/test-font-installed-microsoft-word/
'
' NOTE: Application.FontNames is a snapshot taken when Word started. A font installed while
' Word is open will not appear here until Word is restarted. The installer refuses to run with
' Word open, so a normal install is fine; a hand-installed font is the case to watch.

    Dim InstalledFontName As Variant

    Sh_Is_Font_Installed = False
    For Each InstalledFontName In Application.FontNames
        If UCase(InstalledFontName) = UCase(fontName) Then
            Sh_Is_Font_Installed = True
            Exit Function
        End If
    Next InstalledFontName

End Function   '***** end of Sh_Is_Font_Installed macro *****

Function Sh_Font_Status_Text(ByVal fontName As String, Optional ByVal targetDoc As Document) As String
'
' Version: 1.1  Date: 8/20/2026 - a THIRD answer, for a face that is not installed but is carried
'                                inside the document. Optional targetDoc; without one it behaves
'                                exactly as version 1.0 did
' Version: 1.0  Date: 8/8/2026
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' One line for Sh_Doc_Info saying whether this document's typeface is actually on this machine.
'
' It is here rather than inline in the message because a missing font is invisible everywhere
' else: Word substitutes without a word, and the substitute sets at a different size. "My
' document went small" is the support call, and this is the line that answers it.
'
' THE THIRD ANSWER, and why it had to be added on 8/20/2026. Dropping the bundled typeface means
' the installer takes it off the machine, while every book already set in it keeps that face -
' and those books are safe precisely because they carry their own embedded, unsubsetted copy.
' Ask Sh_Is_Font_Installed alone and every one of Jerry's and the testers' large print books
' answers "NO - Word is substituting, sizes will be wrong". That is FALSE for those books, and it
' is an emergency message: the natural reaction to it is to re-attach or reset the book to
' Tahoma, which moves every page break in a book that may already be printed. The change written
' to protect those books would have destroyed them through a different door.
'
' EmbedTrueTypeFonts is a document property, so this is what the document itself says it does. It
' does not name WHICH faces are embedded - Word offers no such list - so the answer is worded as
' "carries its own copy" rather than as a guarantee about this one face. That is honest, and it
' is enough to stop the transcriber acting on a false alarm.

    If Sh_Is_Font_Installed(fontName) Then
        Sh_Font_Status_Text = "Yes"
    ElseIf Not targetDoc Is Nothing Then
        If targetDoc.EmbedTrueTypeFonts Then
            Sh_Font_Status_Text = "No - but this document carries its own copy, so it sets correctly"
        Else
            Sh_Font_Status_Text = "NO - Word is substituting, sizes will be wrong"
        End If
    Else
        Sh_Font_Status_Text = "NO - Word is substituting, sizes will be wrong"
    End If

End Function   '***** end of Sh_Font_Status_Text macro *****

Sub Lp_Apply_Base_Font_To_Styles(ByVal targetDoc As Document, ByVal fontName As String)
'
' Version: 1.0  Date: 8/8/2026
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Puts the chosen typeface onto the document's STYLES. Sh_Set_Whole_Document_Font above lays the
' same face over the text as direct formatting; this is the other half, and both are needed.
' Direct formatting covers the main story only, so headers, footers and anything typed later
' follow the styles - and reading the face back later (Lp_Get_Doc_Setup_Params) only works
' because it is recorded on Normal.
'
' Why Normal plus a hand-written list, rather than a loop over every style:
'
'   Word resolves a run's font as docDefaults -> the paragraph style's basedOn chain -> the
'   character style -> direct formatting. LargePrintTemplate.dotx leaves docDefaults naming
'   Tahoma, and that is deliberate - it keeps the template the Tahoma baseline, so a book made
'   before this feature existed still reads back as Tahoma, correctly, without touching a thing.
'
'   Setting Normal therefore covers everything that reaches Normal, and everything that names no
'   font of its own reaches the same answer by simply not overriding it. What is left is the
'   short list below: styles that name a face outright and so ignore Normal completely.
'
' "No Spacing" is in that list and is the easy one to miss - it has no basedOn AT ALL, so it
' never touches Normal. A paragraph in it would sit in whatever face the style names, in the
' middle of a book set in another one.
'
' The other twelve are the colored character styles the transcriber applies by keyboard
' shortcut. Leave them out and coloring a phrase snaps it back to whatever those styles name -
' visible, baffling, and it looks like the typeface was never applied.
'
' Still needed after the bundled typeface was dropped on 8/20/2026: these 14 styles name a face
' outright, so nothing puts TAHOMA on them either unless this sub does.
'
' Complex-script (w:cs) fonts are left alone on purpose. Font.Name does not touch Font.NameBi,
' and pointing Arabic or Hebrew at a 352-glyph Latin face would be worse than leaving it.

    Dim sty As Style
    Dim i As Long
    Dim named As Variant

    If Len(Trim(fontName)) = 0 Then Exit Sub

    On Error Resume Next

    targetDoc.Styles(wdStyleNormal).Font.Name = fontName

    named = Array( _
        "Words Aqua", "Words Tan", "Words Blue", "Words Pink", "Words Yellow", _
        "Words Green", "Words Black", _
        "Text Blue", "Text Red", "Text Green", "Text Orange", "Text Violet", _
        "No Spacing")

    For i = LBound(named) To UBound(named)
        Set sty = Nothing
        Set sty = targetDoc.Styles(named(i))
        If Not sty Is Nothing Then sty.Font.Name = fontName
    Next i

    On Error GoTo 0

End Sub   '***** end of Lp_Apply_Base_Font_To_Styles macro *****

Function Lp_Indent_Factor_For_Font(ByVal fontName As String) As Double
'
' Version: 1.0  Date: 8/8/2026
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' How much wider this face is than Tahoma, for the hanging indents on bulleted lists.
'
' The indent table in Lp_Normalize_Styles was tuned by hand against Tahoma and is the authority
' for it - it is not derivable from font metrics, because most of the measurement is Word's tab
' behavior rather than the bullet. So the table stays, and a face that is not Tahoma scales off
' it instead of getting a second table of guesses.
'
' 1.054 is the bullet glyph's advance width in VistaTypeLP Legible divided by Tahoma's, measured
' from the two font files on 8/8/2026 (479 against 455 units per 1000-unit em). It moves an
' 18 point hang from -0.34" to -0.358".
'
' That case is KEPT although the typeface was dropped on 8/20/2026, and it is not dead code.
' A book already set in the face keeps it - that is the whole point of LP_FONT_LEGACY_LEGIBLE -
' so Lp_Normalize_Styles can still be run on one, and it should still get the hang that face
' needs rather than Tahoma's.
'
' 0.862 is VistaTypeLP Sans, measured the same way on 8/22/2026: 376 units on a 960-unit em
' against Tahoma's 931 on 2048, which is 391.7 against 454.6 per 1000. Note it goes the OTHER
' WAY from Legible - Noto's bullet is NARROWER than Tahoma's, where Atkinson's was wider - so
' this face SHRINKS the hang rather than growing it, an 18 point one from -0.34" to -0.293".
' That is worth an eye on a real bulleted list, because it is the first face here to move the
' number down.
'
' It is deliberately not derived from how wide the face sets overall. VistaTypeLP Sans runs
' about 8% wider than Tahoma in running text, so a general width ratio would say 1.08 and be
' wrong by a quarter of an inch on every bullet. The bullet's own advance is what the hang has
' to clear, and that is what both numbers here measure.
'
' A face this function does not know returns 1, i.e. Tahoma's numbers unchanged - the safe
' answer, since those are the ones that have been in the field for years.

    Select Case UCase(Trim(fontName))
        Case UCase(LP_FONT_LEGACY_LEGIBLE)
            Lp_Indent_Factor_For_Font = 1.054
        Case UCase(LP_FONT_SANS)
            Lp_Indent_Factor_For_Font = 0.862
        Case Else
            Lp_Indent_Factor_For_Font = 1#
    End Select

End Function   '***** end of Lp_Indent_Factor_For_Font macro *****

Sub Dx_Fix_Foreign_Languages()
    '
    ' Version: 1.0  Date: 2/28/2022
    '
    ' Attaching a template to a document can (and frequently does) remove all language tabs.
    ' This macro (when run before attaching a template) will preserve the tags within the document.
    '
    ' *************************
    '  Spanish
    '**************************
    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = wdSpanishModernSort  'Spanish (Spain)
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.LanguageID = wdSpanishModernSort  'Spanish (Spain)
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = wdMexicanSpanish
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.LanguageID = wdSpanishModernSort  'Spanish (Spain)
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = 22538 'Spanish (Latin America)
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.LanguageID = wdSpanishModernSort  'Spanish (Spain)
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = wdSpanish '(Spain, Traditional Sort)
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.LanguageID = wdSpanishModernSort  'Spanish (Spain)
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = 21514 'Spanish (United States)
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.LanguageID = wdSpanishModernSort  'Spanish (Spain)
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
    ' *******************
    ' French
    '********************
    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = wdFrench 'French (France)
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.LanguageID = wdFrench 'French (France)
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
    '***************
    ' Italian
    '***************
    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = wdItalian 'Italian (Italy)
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.LanguageID = wdItalian 'Italian (Italy)
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' **********************
    ' German
    '**********************
    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = wdGerman 'German (Germany)
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.LanguageID = wdGerman 'German (Germany)
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' ****************
    ' Latin
    ' ****************
    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = wdLatin
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.LanguageID = wdLatin
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub  '*** end of Dx_Fix_Foreign_Languages macro ***

Sub Dx_Fix_Para_Space_Errors()
'
' Dx_Fix_Para_Space_Errors Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version: 2.0 Date: 8/12/2026 - no temporary document; every pass is scoped to a range. Also
'                               gone: the Selection.MoveUp that stretched the selection up by a
'                               whole paragraph before copying out, and the run of
'                               delete-one-character calls that tidied up after the paste - one
'                               of which deleted the last character of the document on the
'                               no-selection route.
'
'                               NOT merged with the large-print version - the two do different
'                               work. This one handles middle dots and tabs after paragraph
'                               marks; the large-print one protects the "1 point" style.
' Version: 1.8 Date: 2/26/2024 - added Replace middle dot with space
' Version: 1.7 Date: 3/29/2017
'
    Dim rng As Range
    Dim su_Prev As Boolean

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    Set rng = Sh_Para_Fix_Range()

    ' rogue paragraph marks
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^013"
        .Replacement.Text = "^p"
        .Forward = True
        ' wdFindStop, not wdFindContinue: continue would run past the end of the range and
        ' quietly turn "the selection" into "the whole document".
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    ' middle dot (Unicode 00B7) becomes a space
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ChrW(183) & "{1,}"
        .Replacement.Text = " "
        .Forward = True
        ' wdFindStop, not wdFindContinue: continue would run past the end of the range and
        ' quietly turn "the selection" into "the whole document".
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    ' spaces before a paragraph mark
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^032{1,}^013"
        .Replacement.Text = "^p"
        .Forward = True
        ' wdFindStop, not wdFindContinue: continue would run past the end of the range and
        ' quietly turn "the selection" into "the whole document".
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    ' tabs after a paragraph mark
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^013^009{1,}"
        .Replacement.Text = "^p"
        .Forward = True
        ' wdFindStop, not wdFindContinue: continue would run past the end of the range and
        ' quietly turn "the selection" into "the whole document".
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    ' spaces after a paragraph mark
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^013^032{1,}"
        .Replacement.Text = "^p"
        .Forward = True
        ' wdFindStop, not wdFindContinue: continue would run past the end of the range and
        ' quietly turn "the selection" into "the whole document".
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub '***** End of Dx_Fix_Para_Space_Errors ********
' Dx_Remove_Multi_Spaces was here until 8/12/2026. Taking the temporary document out of the
' pair left the two byte-for-byte the same job - everything that differed between them existed
' only to serve the round trip - so they are now one macro, Sh_Remove_Multi_Spaces.
'
' THE RULE, Jerry 8/12/2026: whenever converting a pair leaves them identical, merge them.
' See docs/Temp-Doc-Conversion-Checklist.md.


' Dx_Replace_NonBreaking_Spaces was here until 8/12/2026. Three macros did this job and only
' Sh_ReplaceNonBreakingSpacesWithNormalSpace protected a "Print Pg Num" bar, so everything
' runs through that one now.


Sub Dx_Format_Tagged_Page_Numbers()
'
' Dx_Format_Tagged_Page_Numbers Macro
' Finds and replaces $pg paragraphs with Reference Page Number Style and removes the $pg
' and any spaces within the style.
'
' Version: 1.7  Date: 3/4/2024 - added "MS_Set_Word_Config_For_Braille"
' Version: 1.6  Date: 5/1/2023 - added call to Sh_Remove_Empty_Para_Before_Tables
' Version: 1.5  Date: 2/27/2022 - moved check for "ActiveDocument.Variables("BrailleType") " to the "Dx_Is_BANA_Template_Attached" macro
' Version: 1.4  Date: 1/17/2019
' Version: 1.3  Date: 3/19/2017
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
    ' is the BANA Template Attached... if not terminate macro
    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"

    ' move cursor to delete any selection
    Selection.HomeKey Unit:=wdLine
    
    If ActiveDocument.Bookmarks.Exists("TempPgNoFormat") = True Then
        ActiveDocument.Bookmarks("TempPgNoFormat").Delete
    End If
    
    ' create bookmark at cursor
    ActiveDocument.Bookmarks.Add Name:="TempPgNoFormat"
    
    If ActiveDocument.Variables("BrailleType") = "EBAT" Or ActiveDocument.Variables("BrailleType") = "UEBT" Then
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "RefPageNumber") Then
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles("RefPageNumber")
        With Selection.Find
            .Text = "$pg"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
        
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "RefPageNumber") Then
        Selection.Find.ClearFormatting
        Selection.Find.Style = ActiveDocument.Styles("RefPageNumber")
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^032{1,}"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
        
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "RefPageNumber") Then
        Selection.Find.ClearFormatting
        Selection.Find.Style = ActiveDocument.Styles("RefPageNumber")
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "$pg"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
        
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "RefPageNumber") Then
        Selection.Find.ClearFormatting
        Selection.Find.Style = ActiveDocument.Styles("RefPageNumber")
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find.Replacement.Font
            .Bold = False
            .Italic = False
            .Underline = wdUnderlineNone
        End With
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
    End If

    If ActiveDocument.Variables("BrailleType") = "EBAN" Or ActiveDocument.Variables("BrailleType") = "UEBN" Then 'using nemeth code
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "RefPageNemeth") Then
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles("RefPageNemeth")
        With Selection.Find
            .Text = "$pg"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
        
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "RefPageNemeth") Then
        Selection.Find.ClearFormatting
        Selection.Find.Style = ActiveDocument.Styles("RefPageNemeth")
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^032{1,}"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
        
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "RefPageNemeth") Then
        Selection.Find.ClearFormatting
        Selection.Find.Style = ActiveDocument.Styles("RefPageNemeth")
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "$pg"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
        
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "RefPageNemeth") Then
        Selection.Find.ClearFormatting
        Selection.Find.Style = ActiveDocument.Styles("RefPageNemeth")
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find.Replacement.Font
            .Bold = False
            .Italic = False
            .Underline = wdUnderlineNone
        End With
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
    End If

    ' Change all Textbook ref pg no styles to Nemeth ref pg no styles
        If ActiveDocument.Variables("BrailleType") = "EBAN" Or ActiveDocument.Variables("BrailleType") = "UEBN" Then ' Check for Nemeth
            ' Test for existance of styles
            Dx_GP_String_1 = "RefPageNumber"
            Dx_GP_String_2 = "RefPageNemeth"
            Application.Run MacroName:="Dx_Is_Style_Here"
            If Dx_GP_String_1 = "Here" Then
                ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
                ' this pass can do nothing without it - no style means nothing is formatted with it.
                If Sh_Style_Exists(ActiveDocument, "RefPageNumber") And Sh_Style_Exists(ActiveDocument, "RefPageNemeth") Then
                Selection.Find.ClearFormatting
                Selection.Find.Style = ActiveDocument.Styles("RefPageNumber")
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("RefPageNemeth")
                With Selection.Find
                    .Text = ""
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                End If
            End If
    
            ' Test for existance of styles
            Dx_GP_String_1 = "RefPageNumberEmbed"
            Dx_GP_String_2 = "RefPageNemethEmbed"
            Application.Run MacroName:="Dx_Is_Style_Here"
            If Dx_GP_String_1 = "Here" Then
            ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
            ' this pass can do nothing without it - no style means nothing is formatted with it.
            If Sh_Style_Exists(ActiveDocument, "RefPageNumberEmbed") Then
            Selection.Find.ClearFormatting
            Selection.Find.Style = ActiveDocument.Styles("RefPageNumberEmbed")
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Style = ActiveDocument.Styles( _
                "RefPageNemethEmbed")
            With Selection.Find
                .Text = ""
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            End If
            End If
        End If

    ' Change all Nemeth ref pg no styles to textbook ref pg no styles
    If ActiveDocument.Variables("BrailleType") = "EBAT" Or ActiveDocument.Variables("BrailleType") = "UEBT" Then ' is Textbook
    
        ' Test for existance of styles
        Dx_GP_String_1 = "RefPageNemeth"
        Dx_GP_String_2 = "RefPageNumber"
        Application.Run MacroName:="Dx_Is_Style_Here"
        
            If Dx_GP_String_1 = "Here" Then
                ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
                ' this pass can do nothing without it - no style means nothing is formatted with it.
                If Sh_Style_Exists(ActiveDocument, "RefPageNemeth") And Sh_Style_Exists(ActiveDocument, "RefPageNumber") Then
                Selection.Find.ClearFormatting
                Selection.Find.Style = ActiveDocument.Styles("RefPageNemeth")  ' error here
                Selection.Find.Replacement.ClearFormatting
                Selection.Find.Replacement.Style = ActiveDocument.Styles("RefPageNumber")
                With Selection.Find
                    .Text = ""
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
                End If
            End If
        
            ' Test for existance of styles
            Dx_GP_String_1 = "RefPageNemethEmbed"
            Dx_GP_String_2 = "RefPageNumberEmbed"
            Application.Run MacroName:="Dx_Is_Style_Here"
            If Dx_GP_String_1 = "Here" Then
        
            ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
            ' this pass can do nothing without it - no style means nothing is formatted with it.
            If Sh_Style_Exists(ActiveDocument, "RefPageNemethEmbed") And Sh_Style_Exists(ActiveDocument, "RefPageNumberEmbed") Then
            Selection.Find.ClearFormatting
            Selection.Find.Style = ActiveDocument.Styles("RefPageNemethEmbed")
            Selection.Find.Replacement.ClearFormatting
            Selection.Find.Replacement.Style = ActiveDocument.Styles("RefPageNumberEmbed")
            With Selection.Find
                .Text = ""
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindContinue
                .Format = True
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
            End If
        End If
    End If
    
    Application.Run MacroName:="Sh_Remove_Empty_Para_Before_Tables"
    
    If ActiveDocument.Bookmarks.Exists("TempPgNoFormat") = True Then
        ActiveDocument.Bookmarks("TempPgNoFormat").Select
        ActiveDocument.Bookmarks("TempPgNoFormat").Delete
    End If
        
    ActiveDocument.UndoClear

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    ActiveDocument.UndoClear

End Sub   '****** end of Dx_Format_Tagged_Page_Numbers macro *****



Sub Dx_Embed_Ref_Pg_No()
'
' Convert_Reference_Page_Number_to_Embedded_Reference_Page_Number
' Place cursor in paragraph of reference page number before execution
'
' Version 1.5: Date: 3/14/2024 - supressed instruction msg to when the style reqested style is the same as the current style
'                                   and added message when non-reference page number selected
' Version 1.4: Date: 12/14/2018 - added error check for bad cursor location
' Version 1.4: Date: 9/21/2018 - fixed Nemeth code error
' Version 1.3: Date: 3/19/2017
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
    ' is the BANA Template Attached... if not terminate macro
    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"

    If Selection.Style = "RefPageNumberEmbed" Or Selection.Style = "RefPageNemethEmbed" Then
        End
    End If

    If Selection.Style <> "RefPageNumber" And Selection.Style <> "RefPageNemeth" And _
        Selection.Style <> "RefPageNumberEmbed" And Selection.Style <> "RefPageNemethEmbed" Then
        MsgBox "Select the Reference Page Number or place cursor to the right of or within the number.", , "Braille Macros"
        End
    End If
    
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating off
    
    Selection.HomeKey Unit:=wdLine
    Selection.TypeBackspace
    Selection.TypeText Text:=" "
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    
    If (ActiveDocument.Variables("BrailleType") = "EBAT" Or ActiveDocument.Variables("BrailleType") = "UEBT") And Sh_Style_Exists(ActiveDocument, "RefPageNumberEmbed") Then ' is textbook
            Selection.Style = ActiveDocument.Styles("RefPageNumberEmbed")
    End If
    
    If (ActiveDocument.Variables("BrailleType") = "UEBN" Or ActiveDocument.Variables("BrailleType") = "EBAN") And Sh_Style_Exists(ActiveDocument, "RefPageNemethEmbed") Then  'is Nemeth
            Selection.Style = ActiveDocument.Styles("RefPageNemethEmbed")
    End If

    Selection.EndKey Unit:=wdLine
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.TypeText Text:=" "
    ' move cursor to next para
    Selection.MoveDown Unit:=wdParagraph, count:=1
    Application.Run MacroName:="Dx_Set_DBT_Codes_Color_and_Style"
    
    Application.ScreenUpdating = su_Prev ' Turn screen updating on
    
End Sub '*** end of Dx_Embed_Ref_Pg_No *****
Sub Dx_UnEmbed_Ref_Pg_No()
'
' Embedded_Reference_Page_NumberConvert_Reference_Page_Number_to_
' Double Click the embedded page number before execution
'
' Version: 1.9 Date: 3/14/2024 - supressed instruction msg to when the style reqested style is the same as the current style
'                                   and added message when non-reference page number selected
' Version: 1.8: Date: 9/21/2018 - fixed Nemeth code error
' Version: 1.7 Date: 9/9/2018 - Added BANA Macros label to msgbox
' Version: 1.6 Date: 3/19/2017
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
    ' is the BANA Template Attached... if not terminate macro
    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating off
    
    
    ' check to see if selection is a reference page number of any kind
    If Selection.Style <> "RefPageNumber" And Selection.Style <> "RefPageNemeth" And _
        Selection.Style <> "RefPageNumberEmbed" And Selection.Style <> "RefPageNemethEmbed" Then
        MsgBox "Select the Reference Page Number or place cursor to the right of or within the number.", , "Braille Macros"
        End
    End If
       
    ' selection is to unembed then no changes
    If Selection.Style <> "RefPageNumberEmbed" And Selection.Style <> "RefPageNemethEmbed" Then
        End
    End If

    ' Move cursor to the left until the embeded style is no longer valid
    Do While Selection.Style = "RefPageNumberEmbed" Or Selection.Style = "RefPageNemethEmbed"
        Selection.MoveLeft Unit:=wdCharacter, count:=1
    Loop
    
    Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdExtend
    
    Do While Selection.Style = "RefPageNumberEmbed" Or Selection.Style = "RefPageNemethEmbed"
        Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdExtend
    Loop
    
    Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
    Selection.Copy

    ' ****** end of Find and select the embedded reference page number ***************
    Selection.Cut
    Selection.TypeBackspace
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.MoveLeft Unit:=wdCharacter, count:=1
    Selection.Paste
    Selection.EndKey
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    Selection.ClearFormatting

    If (ActiveDocument.Variables("BrailleType") = "EBAT" Or ActiveDocument.Variables("BrailleType") = "UEBT") And Sh_Style_Exists(ActiveDocument, "RefPageNumber") Then ' is textbook
            Selection.Style = ActiveDocument.Styles("RefPageNumber")
    End If

    If (ActiveDocument.Variables("BrailleType") = "UEBN" Or ActiveDocument.Variables("BrailleType") = "EBAN") And Sh_Style_Exists(ActiveDocument, "RefPageNemeth") Then  'is Nemeth
            Selection.Style = ActiveDocument.Styles("RefPageNemeth")
    End If
    
    Application.ScreenUpdating = False
    Selection.HomeKey Unit:=wdLine
    Application.Run MacroName:="Dx_Set_DBT_Codes_Color_and_Style"
    Application.ScreenUpdating = su_Prev ' Turn screen updating on
    
End Sub   '*** end of Dx_UnEmbed_Ref_Pg_No macro ***

Sub Dx_Is_BANA_Template_Attached()
'
' Version: 1.5  Date: 2/27/2022 - Added check for "ActiveDocument.Variables("BrailleType")"
' Version: 1.4  Date: 10/27/2021 - fixed logic bug in determining if BANA template attached
' Version: 1.3  Date: 10/26/2021 - removed code to read SWIFT values
' Version: 1.2  Date: 1/8/2016
' Version: 1.3  Date: 12/6/2018 - added check and fix for blocked files
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Description:  Call: Application.Run MacroName:="Dx_Is_BANA_Template_Attached"
'                   Checks to see if a BANA Braille Template is attached
'                   and that that template is 2014 or later

    Application.Run MacroName:="Sh_Is_Doc_Open"
    If InStr(ActiveDocument.AttachedTemplate, "BANA Braille") = 0 Then
        Application.Run MacroName:="Dx_Attach_BANA_Template"   'template was NOT attached - attach it
    End If
    
    ' BANA Braille template is attached but is it a version which is too old?
    If InStr(ActiveDocument.AttachedTemplate, "BANA Braille") <> 0 And Val(Mid(ActiveDocument.AttachedTemplate, 14, 4)) < 2014 Then
        Dim lngQuery As Long
        lngQuery = MsgBox("The BANA template attached to this document is " + ActiveDocument.AttachedTemplate + "." & vbCr _
                       & vbCr & "These macros will not work with BANA templates prior to BANA Braille 2014" & vbCr _
                       & vbCr & "Attach a newer BANA template to this document before proceeding.", "Braille Macros")
        End
    End If
    
    On Error GoTo GetTempTypeFromUser  ' will crash if these are not in the file - need to get it from user
     If ActiveDocument.Variables("BrailleType") = "EBAT" Or ActiveDocument.Variables("BrailleType") = "UEBT" Or _
                ActiveDocument.Variables("BrailleType") = "EBAN" Or ActiveDocument.Variables("BrailleType") = "UEBN" Then
        GoTo GetTempTypeFromUserExit
    Else
        GoTo GetTempTypeFromUser
    End If
GetTempTypeFromUser:
        Dx_Choose_Translation_Form.Show
GetTempTypeFromUserExit:
    
End Sub   '*** end of Dx_Is_BANA_Template_Attached macro ***

Sub Dx_Convert_Auto_List_To_Text(Optional ByVal target As Range)
    '
    ' Original title "AutoListOff2"
    ' From: Computer Tools for Editors(and Proofreaders)by Paul Beverley, LCGI
    '        http://www.archivepub.co.uk/book.html
    ' Initial version by Paul Beverley; modified by Jerry Whittaker - jerry@thewhittakers.org
    '
    ' Changes auto-bulleted, auto-numbered and auto-outline listing to real bullets and numbers,
    ' and takes the tab out from behind the marker.
    '
    ' target  the range to work on. Omit it and the selection is used, or the whole document when
    '         there is no selection.
    '
    ' Version: 2.0  Date: 8/13/2026 - no temporary document; takes a range.
    '
    '                                ActiveDocument.ConvertNumbersToText is a whole-document
    '                                operation, and every Find was built on ActiveDocument.Range
    '                                with wdFindContinue - the comments even say "replaces all in
    '                                the document". The scratch document was the only thing
    '                                confining any of it to the transcriber's selection, so this
    '                                is another one where the round trip was doing real work.
    '
    '                                Converted now because Dx_Remove_Bullets calls it: a macro
    '                                that has been scoped to a range cannot call one that goes
    '                                through a temporary document, because the paste back
    '                                invalidates the range it was holding.
    '
    '                                Gone with the round trip: the four-character delete run and
    '                                "delete one character at the end of the document".
    ' Version: 1.3  Date: 11/16/2016
    '
    Dim rng As Range
    Dim NewCharacter As String
    Dim NormalFont As String
    Dim su_Prev As Boolean

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    If target Is Nothing Then
        If Selection.Type = wdSelectionNormal Then
            Set rng = Selection.Range
        Else
            Set rng = ActiveDocument.Content
        End If
    Else
        Set rng = target
    End If
    NewCharacter = ChrW(8226): ' a bullet
    NormalFont = ActiveDocument.Styles(wdStyleNormal).Font.Name

    rng.ListFormat.ConvertNumbersToText   ' this gets rid of the protected nature

    ' One common type of bullet uses Symbol font
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .Text = ChrW(&HF0B7) & "^t"
        .Forward = True
        .Font.Name = "Symbol"
        .Replacement.Text = NewCharacter & " "
        .Replacement.Font.Name = NormalFont
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With

    ' The other type of bullet uses Wingding font
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .Text = ChrW(&HF0FC) & "^t"
        .Font.Name = "Wingding"
        .Replacement.Text = NewCharacter & " "
        .Replacement.Font.Name = NormalFont
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With

    ' The tab behind the marker becomes a space, for each shape of list marker in turn
    Sh_Marker_Tab_To_Space rng, "([0-9]{1,}^046)^009", NormalFont      ' 1.
    Sh_Marker_Tab_To_Space rng, "(^0149)^009", NormalFont              ' bullet
    Sh_Marker_Tab_To_Space rng, "(^013[A-Za-z]{1,}^046)^009", NormalFont   ' a.
    Sh_Marker_Tab_To_Space rng, "([A-Za-z]{1,}\))^009", NormalFont     ' a)
    Sh_Marker_Tab_To_Space rng, "([0-9]{1,}\))^009", NormalFont        ' 1)

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart
    Application.ScreenUpdating = su_Prev

End Sub  '**** end of Dx_Convert_Auto_List_To_Text Macro ***********

Public Sub Sh_Marker_Tab_To_Space(ByVal rng As Range, ByVal pattern As String, ByVal fontName As String)
'
' Version: 1.0  Date: 8/13/2026
'
' One "list marker followed by a tab becomes list marker followed by a space" pass. Both copies
' of Convert_Auto_List_To_Text had five or six of these written out longhand and identical apart
' from the pattern.
'
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = True
        .Text = pattern
        .Replacement.Text = "\1^032"
        .Replacement.Font.Name = fontName
        .Forward = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With

End Sub   '*** end of Sh_Marker_Tab_To_Space ***

Sub Dx_Replace_Tabs_With_Single_Space()
'
' Dx_Replace_Tabs_With_Single_Space Macro
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Runs of tabs become one space - but UNDERLINED tabs become underscores first, because a
' scanner turns a ruled fill-in line into an underlined tab and a braille transcriber needs the
' underscores, not a space.
'
' Version: 2.0  Date: 8/12/2026 - no temporary document, the same conversion as
'                                Sh_Remove_Multi_Spaces.
'
'                                NOT merged with Lp_Replace_Tabs_With_Single_Space, and it is
'                                the exception that shows what the merge rule actually means:
'                                the two differ in WORK, not in plumbing. The underlined-tab
'                                passes below are braille's alone, so taking the round trip out
'                                leaves them still doing different jobs.
'
'                                One behavior change worth knowing: the first underlined-tab
'                                pass used to run before the temporary document was made, with
'                                Wrap:=wdFindContinue, so it could reach PAST a selection and
'                                change the rest of the document. Scoped to a range it now stays
'                                inside what the transcriber selected.
' Version: 1.4  Date: 3/8/2023
' Version: 1.4  Date: 3/6/2003 - added replacement of underlined tabs with single underscore
' Version: 1.3  Date: 11/16/2016
'
    Dim rng As Range
    Dim su_Prev As Boolean

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    If Selection.Type = wdSelectionNormal Then
        Set rng = Selection.Range
    Else
        Set rng = ActiveDocument.Content
    End If

    ' a single underlined tab becomes one underscore (often left by Abbyy)
    With rng.Find
        .ClearFormatting
        .Font.Underline = wdUnderlineSingle
        .Replacement.ClearFormatting
        .Replacement.Font.Underline = wdUnderlineNone
        .Text = "^t"
        .Replacement.Text = "_"
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' a RUN of underlined tabs becomes a spaced underscore
    With rng.Find
        .ClearFormatting
        .Font.Underline = wdUnderlineSingle
        .Replacement.ClearFormatting
        .Replacement.Font.Underline = wdUnderlineNone
        .Text = "^t{1,}"
        .Replacement.Text = " _ "
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    ' whatever tabs are left become a space
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^t{1,}"
        .Replacement.Text = "^032"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub  '*** end of Dx_Replace_Tabs_With_Single_Space Macro ***

Sub Dx_Fix_Common_File_Errors()

' Dx_Fix_Common_File_Errors Macro
'
' Version: 2.12  Date: 8/18/2026 - Word's own AutoFormat command is no longer run over the document. It ate empty
'                                  paragraphs (measured on the LP side, 8/13/2026), and everything it did that
'                                  braille wants is done properly by other steps in this same sequence. Jerry,
'                                  8/18/2026. The redundant configuration call that fed it went with it.
' Version: 2.11  Date: 8/5/2026 - colors the $pg tags red as its last content step, so they stay red when this macro is run on its own (Jerry)
' Version: 2.10  Date: 7/26/2026 - returns the user to where the cursor was when the macro started
' Version: 2.8 Date: 3/5/2024 - added "If ActiveDocument.Variables("BrailleType") = "EBAN" Or ActiveDocument.Variables("BrailleType") = "UEBN" then"
' Version: 2.7 Date: 2/6/2024 - moved Application.Run MacroName:="Dx_Fix_Para_Space_Errors" to last routine run
'                             - added Application.Run MacroName:="Dx_Fix_Equals_Before_Para_Mark"
' Version: 2.6 Date: 10/20/2021 - added Dx_Replace_Fraction_Text_With_Compact_Fractions
' Version: 2.5 Date: 4/20/2019 - added Dx_Fix_Abbyy_FineReader_Text_and_Headers
' Version: 2.4 Date: 12/6/2018 -  Removed attached template check - not needed - is called in top menu
' Version: 2.3 Date: 10/23/2018 - added white text to automatic
' Version: 2.2 Date: 9/14/2018  - Fix Body Text Styles positioned after conver autonumbers
' Version: 2.1 Date: 3/18/2017
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Calls a series or routines for global file cleanup
'
    '-------------------------------------------------------------
    ' Remember where the user is, so they can be put back at the end
    Sh_Save_User_Position
    '------------------------------------------------------

    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating off

    '------------- start cleanup ------------------

    ' Word's own AutoFormat command was run over the whole document here, first thing, from
    ' 3/18/2017 until 8/18/2026. It is gone, and nothing replaces it. Jerry, 8/18/2026.
    '
    ' The large print side removed the same call on 8/13/2026 and the reasons carry over intact:
    ' AutoFormat REMOVES EMPTY PARAGRAPHS, and turning off every AutoFormat option except the one
    ' wanted does not stop it - that was tried and measured. It was one of the three faults behind
    ' the empty paragraph marks that survived Full File Cleanup. It also restyles headings and
    ' rewrites quotes and fractions behind the transcriber's back.
    '
    ' Nothing here needed it. What it did for braille was already done properly elsewhere in this
    ' same sequence: quotes by Dx_Replace_Straight_Quotes_With_Smart_Quotes, fractions by
    ' Dx_Replace_Fraction_Text_With_Compact_Fractions, dashes by Dx_Fix_En_Dash_Errors and
    ' Dx_Fix_Em_Dash_Space_Errors. It also made plain URLs into live hyperlinks, which is pointless
    ' here - Dx_Convert_Hyperliks_To_Text turns them straight back into text a few steps later,
    ' because a DBT source file has no use for a clickable link. Deliberately NOT replaced with
    ' Sh_Linkify_Range for that reason; large print keeps its links, braille does not.
    '
    ' What is genuinely gone: AutoFormat's symbol replacement ((c) to the copyright sign) and its
    ' ordinals (1st to a superscript st). Neither is wanted in a braille source file.
    '
    ' This also settles a second thing. That call was the ONLY place in the whole project that ran
    ' Word's AutoFormat command, which is what made the on-demand AutoFormat settings load-bearing
    ' rather than cosmetic. With it gone they are the transcriber's, like the rest of the fifteen
    ' removed from the three configuration subs on 8/18/2026.
    Application.Run MacroName:="MS_Set_Word_Config_For_Braille"

    If ActiveDocument.Variables("BrailleType") = "EBAT" Or ActiveDocument.Variables("BrailleType") = "EBAN" Then 'are cleaned up only for EBAE Textbook and EBAE Nemeth
        Application.Run MacroName:="Dx_Fix_Em_Dash_Space_Errors"
    End If
    
    Application.Run MacroName:="Dx_Delete_Square_Bullet"
    
    Application.Run MacroName:="Dx_Fix_Abbyy_FineReader_Text_and_Headers"

    Application.Run MacroName:="Dx_Delete_Images"

    Application.Run MacroName:="Sh_Replace_White_Text_With_Automatic"

    Application.Run MacroName:="Dx_Fix_Primes"

    Application.Run MacroName:="Sh_Remove_Spaces_Before_Punctuation"

    Application.Run MacroName:="Dx_Fix_En_Dash_Errors"

    Application.Run MacroName:="Sh_Remove_Txt_Bxs_And_Frames"
    
    Application.Run MacroName:="Dx_Convert_Auto_List_To_Text"

    Application.Run MacroName:="Dx_Fix_Body_Text_Styles"
    
    Application.Run MacroName:="Dx_Add_Color_To_Foreign_Language_Words"
      
    Application.Run MacroName:="Dx_Replace_Straight_Quotes_With_Smart_Quotes"
    
    Application.Run MacroName:="Sh_Replace_Small_Caps_With_All_Caps"

    Application.Run MacroName:="Dx_Remove_Optional_Hyphens"
    
    Application.Run MacroName:="Dx_Remove_Breaks"
    
    Application.Run MacroName:="Dx_Remove_Column_Breaks"

    Application.Run MacroName:="Sh_RemoveHeadAndFoot"

    Application.Run MacroName:="Sh_Para_Before_Dollar" 'Fixes DAISY Page Problems
    
    Application.Run MacroName:="Dx_Remove_Keep_With_Next"

    Application.Run MacroName:="Dx_Convert_Hyperliks_To_Text"   'convert hidden hyperlink to actual address
    
    Application.Run MacroName:="Dx_Replace_Word_NonBreaking_Hyphen_With_Unicode_Non_Breaking_Hypen"
    
    Application.Run MacroName:="Dx_Replace_Tabs_With_Single_Space" ' also converts underlined tabs BANA template only to single Space
    
    Application.Run MacroName:="Dx_Remove_Breaks" 'page breaks
    
    ' "Para" supplied, so this does NOT stop to ask - Full File Cleanup must run straight
    ' through. See the note on Sh_Replace_Manual_Line_Break.
    ' A direct call, not Application.Run. Once MacroName:= is written as a NAMED argument every
    ' argument after it must be named too, so "Application.Run MacroName:=..., ""Para""" does not
    ' compile. This is a Public Sub in this same module, so it can simply be called.
    Sh_Replace_Manual_Line_Break "Para"
    
    Application.Run MacroName:="Dx_Replace_Underscore_With_Single_Underscore"
    
    Application.Run MacroName:="Sh_ReplaceNonBreakingSpacesWithNormalSpace"

    Application.Run MacroName:="Dx_Replace_Spaces_Before_Punctuation"

    Application.Run MacroName:="Sh_Remove_Multi_Spaces"
    
    If ActiveDocument.Variables("BrailleType") = "EBAN" Or ActiveDocument.Variables("BrailleType") = "UEBN" Then
        Application.Run MacroName:="Dx_Replace_Function_Application_With_Space" ' code used in math - U+2061 or chrW8289"
        Application.Run MacroName:="Dx_Fix_Equals_Before_Para_Mark"
    End If
    
    Application.Run MacroName:="Dx_Fix_Para_Space_Errors"
       
    Application.Run MacroName:="Dx_Add_Qmark_To_Incomplete_Equations"

    ' LAST of the content steps, and before the cursor goes home below - this one uses
    ' Selection.Find, which moves it. The cleanups above lose the red on the $pg tags, and this
    ' macro can be run on its own from the braille ribbon with no Dx_AutoTag_Page_Numbers
    ' afterwards to put it back.
    Application.Run MacroName:="Dx_Color_Dollar_PG_Red"

    ActiveDocument.UndoClear
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Application.ScreenUpdating = su_Prev ' Turn screen updating on

    Sh_Return_User_To_Start_Position

    MsgBox "End of Fix Common File Errors", , "Braille Macros"
    
End Sub '***** end of Dx_Fix_Common_File_Errors Macro *****

Sub Dx_Color_Dollar_PG_Red()
'
' Version: 1.0  Date: 8/5/2026
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Colors every $pg tag red, and nothing else.
'
' NOT Sh_Color_Dollar_PG_Red, which the large print side uses for this. That one also sets the
' replacement STYLE to Normal, and Word applies a paragraph style to the whole paragraph rather
' than to the three characters matched - in a braille file it would strip the BANA style off
' every paragraph holding a tag. This is the closing pass of Dx_AutoTag_Page_Numbers instead:
' color only, then put the DBT code colors back, since a tag can be followed by a code such
' as [[*ii*]] on the same paragraph.

    Application.Run MacroName:="Sh_Is_Doc_Open"

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Color = wdColorRed
    End With
    With Selection.Find
        .Text = "$pg"
        .Replacement.Text = ""    ' empty + Format:=True means "keep the text, take the formatting"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Put the DBT code colors back. Guarded: Dx_Set_DBT_Codes_Color_and_Style asks for the
    ' "DBT Code" style BY NAME, and only the BANA template carries it. On a document without it
    ' that call raises 5941 with no handler, which in Word means a modal dialog and a wedged
    ' session. No such style means no DBT codes to restore, so skipping is the right answer.
    If Sh_Style_Exists(ActiveDocument, "DBT Code") Then
        Application.Run MacroName:="Dx_Set_DBT_Codes_Color_and_Style"
    End If

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"

End Sub  '*** end of Dx_Color_Dollar_PG_Red Macro ***

Private Function Sh_Style_In_Use(ByVal targetDoc As Document, ByVal styleName As String) As Boolean
    '
    ' Version: 1.0  Date: 8/5/2026
    '
    ' .InUse, but False instead of run-time error 5941 when the document has no such style.

    If Not Sh_Style_Exists(targetDoc, styleName) Then Exit Function

    On Error Resume Next
    Sh_Style_In_Use = targetDoc.Styles(styleName).InUse
    Err.Clear
    On Error GoTo 0

End Function   '*** end of Sh_Style_In_Use function ***

Private Function Sh_Style_Exists(ByVal targetDoc As Document, ByVal styleName As String) As Boolean
    '
    ' Version: 1.0  Date: 8/5/2026
    '
    ' Does the document carry that style at all? Styles(...) raises 5941 when it does not, which
    ' is the answer rather than a fault.

    Dim probe As Style
    On Error GoTo NoStyle
    Set probe = targetDoc.Styles(styleName)
    Sh_Style_Exists = True
    Exit Function

NoStyle:
    Sh_Style_Exists = False

End Function   '*** end of Sh_Style_Exists function ***



Sub Dx_Selected_File_CleanUp()
'
' Version: 1.2  Date: 4/7/2023 - show form before text selected error
' Version: 1.1  Date: 10/18/2018 - added Is Text Selected to this startup

    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"
    Dx_Selected_Cleanup_Form.Show

End Sub  '*** end of Dx_Selected_File_CleanUp macro ***

Sub Dx_Fix_Body_Text_Styles()
'
' Dx_Fix_Body_Text_Styles Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version Date: 8/24/2015
'
' Call: Application.Run MacroName:="Dx_Fix_Body_Text_Styles"
'
' Replace styles "Normal (Web)", "Normal Indent" and "Normal", with style "Body Text"
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Body Text") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Normal")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Body Text")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Normal (Web)") And Sh_Style_Exists(ActiveDocument, "Body Text") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Normal (Web)")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Body Text")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Normal Indent") And Sh_Style_Exists(ActiveDocument, "Body Text") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Normal Indent")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Body Text")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
End Sub '***** end of Dx_Fix_Body_Text_Styles Macros *****

Sub Dx_Manual_Tag_with_Dollar_pg()
'
' Dx_Manual_Tag_with_Dollar_pg Macro
'
' Version: 1.5  Date: 9/25/2018 - incorporated proper tagging of lower case roman in EBAE documents'
' Version: 1.4  Date: 5/4/2017
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"
    Application.Run MacroName:="Sh_Remove_DollarPG_For_Retag"

    Selection.HomeKey Unit:=wdLine
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Copy ' put the possible roman numeral in the clipboard
     
    'copy the clipboard into a variable - From: http://www.vbaexpress.com/forum/showthread.php?27996-Solved-Copy-string-to-clipboard-in-VBA
    Dim Clipboard_Data As New DataObject
    Clipboard_Data.GetFromClipboard
    Dim Possible_LC_Roman As Variant
    Possible_LC_Roman = Clipboard_Data.GetText
    
    'validate the selection
    If Not Sh_IsValidRomanNumeral(UCase(Possible_LC_Roman)) Then  'see if it is a roman numeral - validates only in Ucase
       Selection.HomeKey Unit:=wdLine
       MsgBox "This is not a valid roman numeral.", , "Braille Macros"
       End
    End If
    
    ' if this is an EBAE document - then it might be a lower case roman numeral needing a [[*ii*]] code
    If ActiveDocument.Variables("BrailleType") = "EBAT" Or ActiveDocument.Variables("BrailleType") = "EBAN" Then
        If UCase(Possible_LC_Roman) <> Possible_LC_Roman Then  'the roman numeral is lower case
            Selection.HomeKey Unit:=wdLine
            Selection.TypeText Text:="[[*ii*]]"
            Selection.HomeKey Unit:=wdLine
        End If
    End If
          
    ' in either case (UEB or EBAE) place the $pg tag
    Selection.EndKey Unit:=wdLine
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    
    Selection.HomeKey Unit:=wdLine

    With Selection.Font
        .Color = wdColorRed
    End With
    
    Selection.TypeText Text:="$pg"
    Selection.HomeKey Unit:=wdLine
    
    Application.Run MacroName:="Dx_Set_DBT_Codes_Color_and_Style"
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    
End Sub  '***** end of Dx_Manual_Tag_with_Dollar_pg Macro *****

Sub Dx_Fix_En_Dash_Errors()
'
' Dx_Fix_En_Dash_Errors Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version 1.1
' Date: 12/11/2015

' Call: Application.Run MacroName:="Dx_Fix_En_Dash_Errors"
'
'----------------------------------------------------------------------------------------------
'  replace one or more en dashes with single hyphen
'----------------------------------------------------------------------------------------------
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^0150{1,}"
        .Replacement.Text = "^045"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub  '*** end of Dx_Fix_En_Dash_Errors macro ***

Sub Dx_Add_Color_To_Foreign_Language_Words()
'
' Dx_Add_Color_To_Foreign_Language_Words Macro
'
' Version 1.5
' Date 10/17/2016
'
' Revisons:
'       1.4 - now handles Spanish (Spain Modern Sort) and Spanish (Spain Traditional Sort)
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
' Description: Examines the language setting of each word and changes
'              them to the color styles used the BANA template
'
' Note: Works on the entire document
'
    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"
    
    '************ German *************
        
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "German") Then
    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = wdGerman
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("German")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    '*********** French **************
        
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "French") Then
    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = wdFrench
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("French")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    '*********** Italian ***************
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Italian") Then
    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = wdItalian
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Italian")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    '********* Latin *******************
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Latin") Then
    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = wdLatin
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Latin")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' ********** Spanish ************
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Spanish") Then
    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = wdSpanish
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Spanish")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Spanish") Then
    Selection.Find.ClearFormatting
    Selection.Find.LanguageID = wdSpanishModernSort
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Spanish")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    ActiveDocument.UndoClear
       
End Sub  '*** end of Dx_Add_Color_To_Foreign_Language_Words macro *****

Sub Dx_About()
'
' Dx_About Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'  Version: 1.3  Date 12/10/2019 - added code for alternative short URL in the user form code
'  Version: 1.2  Date: 3/17/2017
'
    'With Dx_About_Title_And_Agreement
        '.Version.Caption = "Version Beta .9.8  February 1, 2018"
        '.Version.ControlTipText = "This is BANA Macros Version Beta .9.8 February 1, 2018"
        '.Version.TextAlign = fmTextAlignCenter
    'End With
    Dx_About_Title_And_Agreement.Show
    Unload Dx_About_Title_And_Agreement
    
End Sub  '***** end of Dx_About Macro *****
Sub Dx_Video_Links()
'
' Version: 1.0  Date: 3/3/2021
'
    Dx_Video_Download_Link_Page.Show
    Unload Dx_Video_Download_Link_Page
    
End Sub  '***** end of Lp_Video_LinksMacro ****

Sub Dx_Close_with_no_Save()
    ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
End Sub  '*** end of Dx_Close_with_no_Save macro ***

' Dx_Horz_List_To_Vertical was here until 8/11/2026. There were two of these, one per side,
' and the braille copy was the 2018 version - it never got the bullet handling, the trapped
' crash on sorting a non-sortable selection, the 8/2/2026 change that stopped it seizing the
' Styles pane, or the sequence recognizer. Both ribbon buttons now run
' Lp_Horz_List_To_Vertical, which is the one that has been maintained. Its BANA-template guard
' went with it: the macro does not care what template is attached, and the large-print side had
' already commented its own template check out (Jerry, 8/11/2026).
'
' The round trip through the temporary document still differs by document type, and has to -
' see Sh_Copy_To_Temp_Doc.


Sub Dx_Fix_Em_Dash_Space_Errors()
'
' Dx_Fix_Em_Dash_Space_Errors Macro
'
' Removes leading and trailing spaces from em dashes for EBAE only
'
' Author: Jerry Whittaker  jerry@thewhittakers.org
'
' version 1.3
' date: 3/18/2017
'

        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "^+^032{1,}"
            .Replacement.Text = "^+"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        
        Selection.Find.Execute Replace:=wdReplaceAll
        With Selection.Find
            .Text = "^032{1,}^+"
            .Replacement.Text = "^+"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub '***** End of Dx_Fix_Em_Dash_Space_Errors Macro *****

Sub Dx_Spelling_List()
'
' Dx_Spelling_List Macro
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
' jerry@thewhittakers.org
'
' Version 1.4:  Date: 6/18/2018
' Version 1.5:  Date: 11/16/2018
'
' Duplicates the spelling list words...
' First word is contracted
' Second word is uncontacted
' Spelling word list is set to style "List1"
'
' Directions: Highlight spelling list and last para mark before executing macro
'

    Application.Run MacroName:="Sh_Is_Doc_Open"
    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"
      
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "Select the spelling list first!", , "Braille Macros"
        End
    End If

    Dx_Spelling_List_Options_Form.Show
    Unload Dx_Spelling_List_Options_Form

End Sub   '***** End of Dx_Spelling_List Macro *****

Sub Dx_AutoTag_Page_Numbers()
'
' Version: 2.8 Date: 8/5/2026 - consolidates back-to-back reference page numbers into one range, between the strip and the tagging
' Version: 2.7 Date: 8/5/2026 - the three passes that strip an EARLIER tagging no longer die with run-time error 5941 when the document has no RefPageNumber or RefPageNemeth style; same guard on the DBT code restore at the end
' Version: 2.6 Date: 8/2/2026 - the ten "^013(...)^013" passes now repeat until nothing is left to replace (Sh_Replace_All_Until_Done); a single Execute tagged only alternate numbers when two page numbers sat in consecutive paragraphs
' Version: 2.5 Date: 8/30/2025 - added new validation of roman numerals
' Version: 2.4 Date: 2/16/2024 - added code to automate validation
' Version: 2.3 Date: 1/18/2023 - fixed leaf continue from [[*lea*]] to [[*lec*]][[*i*]]
' Version: 2.2 Date 4/30/2023 - nearly full re-write
' Version: 2.1 Date: 9/21/2018 -  fixed problem where ref pg no is not found when it is the first para in the document
' Version: 2.0 Date: 2/8/2017
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Locates potential page numbers and tags with $pg
' Also locates continuation pages ##-## and enters the [[*lec*]][[*i*]] code
'

    ' FIRST, before anything reads or moves the cursor. Validation can leave every $pg
    ' paragraph selected at once, and Word then refuses most Selection work with error 4605.
    Sh_Clear_Multi_Selection

    Dim strLength As Integer

    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"
    Application.Run MacroName:="Dx_Fix_Para_Space_Errors"
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating off
    Sh_Save_User_Position
    
    ' place para mark at top of file
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.HomeKey Unit:=wdStory
    
'********************** clear all previous tags and [[*lec*]][[*i*]] codes ***************************
   
    ' place para mark before graphics (otherwise ref pg no before graphic will not be found
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^g"
        .Replacement.Text = "*~^p^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' place para mark after graphics
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^g"
        .Replacement.Text = "^&*~^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' remove $pg
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    Selection.Find.Replacement.Font.Color = wdColorAutomatic
    With Selection.Find
        .Text = "^013$pg"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' Guarded. This pass exists to STRIP an earlier tagging, so it only runs on a document
    ' that has been tagged before - which is exactly when Jerry hit run-time error 5941,
    ' 8/5/2026. Styles(name) raises 5941 when the document does not carry that style, and a
    ' braille file need not: RefPageNemeth belongs to the Nemeth math templates. No style
    ' means nothing is formatted with it, so there is nothing here to strip and skipping is
    ' correct. Unguarded it killed the whole macro before a single tag was placed.
    If Sh_Style_Exists(ActiveDocument, "RefPageNumber") Then
        ' remove RefPageNumber style from roman numerls
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "RefPageNumber") Then
        Selection.Find.ClearFormatting
        Selection.Find.Style = ActiveDocument.Styles("RefPageNumber")
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
    End If
    
    ' Guarded. This pass exists to STRIP an earlier tagging, so it only runs on a document
    ' that has been tagged before - which is exactly when Jerry hit run-time error 5941,
    ' 8/5/2026. Styles(name) raises 5941 when the document does not carry that style, and a
    ' braille file need not: RefPageNemeth belongs to the Nemeth math templates. No style
    ' means nothing is formatted with it, so there is nothing here to strip and skipping is
    ' correct. Unguarded it killed the whole macro before a single tag was placed.
    If Sh_Style_Exists(ActiveDocument, "RefPageNemeth") Then
        ' remove RefPageNumber style from roman numerls
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "RefPageNemeth") Then
        Selection.Find.ClearFormatting
        Selection.Find.Style = ActiveDocument.Styles("RefPageNemeth")
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
    End If
    
    ' Guarded. This pass exists to STRIP an earlier tagging, so it only runs on a document
    ' that has been tagged before - which is exactly when Jerry hit run-time error 5941,
    ' 8/5/2026. Styles(name) raises 5941 when the document does not carry that style, and a
    ' braille file need not: RefPageNemeth belongs to the Nemeth math templates. No style
    ' means nothing is formatted with it, so there is nothing here to strip and skipping is
    ' correct. Unguarded it killed the whole macro before a single tag was placed.
    If Sh_Style_Exists(ActiveDocument, "RefPageNemeth") Then
        ' replace nemeth RefPageNemeth with normal style
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "RefPageNemeth") Then
        Selection.Find.ClearFormatting
        Selection.Find.Style = ActiveDocument.Styles("RefPageNemeth")
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
    End If
    
    ' lec + letter + number
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = "\[\[*lec*\]\][\[*i*\]\][A-Za-z]{1,}[0-9]{1,}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' lea + number + letter
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = "\[\[*lec*\]\][\[*i*\]\][0-9]{1,}[A-Za-z]{1,}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'lea + number
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    Selection.Find.Replacement.Font.Color = wdColorAutomatic
    With Selection.Find
        .Text = "\[\[*lec*\]\]\[\[*i*\]\][0-9]{1,}^013"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'lea + + letter + number
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    Selection.Find.Replacement.Font.Color = wdColorAutomatic
    With Selection.Find
        .Text = "\[\[*lec*\]\][\[*i*\]\][A-Za-z]{1,}[0-9]{1,}^013"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'lea + number + letter
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    Selection.Find.Replacement.Font.Color = wdColorAutomatic
    With Selection.Find
        .Text = "\[\[*lec*\]\]\[*i*\]\][0-9]{1,}[A-Za-z]{1,}^013"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'lea + letter + number + letter
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    Selection.Find.Replacement.Font.Color = wdColorAutomatic
    With Selection.Find
        .Text = "\[\[*lec*\]\]\[\[*i*\]\][A-Za-z]{1,}[0-9]{1,}[A-Za-z]{1,}^013"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' Back-to-back reference page numbers become ONE range, HERE at the seam between the two
    ' halves of this macro. Everything above has stripped the old tags, DBT codes and styles, so
    ' the numbers are bare - 12, 13, 14 - and joining them into 12-14 lets the "number hyphen
    ' number" pass below produce $pg12-14[[*lec*]][[*i*]]14, continuation code and all. Do it
    ' after the tagging instead and that code would have to be built by hand. (Jerry, 8/5/2026)
    Dx_Merge_Adjacent_Pg_Numbers ActiveDocument

'******************* begin tagging *********************************
  
    Application.Run MacroName:="Sh_Fix_Ref_Pages_Before_and_After_Tables"
  
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "$pn"
        .Replacement.Text = "$pg"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "$ppn"
        .Replacement.Text = "$pg"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ppn"
        .Replacement.Text = "$pg"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = wdColorAutomatic
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013pn([0-9]{1,})"
        .Replacement.Text = "^013$pg\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "$pg^032{1,}"
        .Replacement.Text = "$pg"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' letter any length numb any length
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([A-Za-z]{1,})([0-9]{1,})^013"
        .Replacement.Text = "^p$pg\1\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done
    
    ' numb any length hyphen numb any length
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([0-9]{1,}^045)([0-9]{1,})^013"
        .Replacement.Text = "^p$pg\1\2[[*lec*]][[*i*]]\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done

    'letter , Number, Hyphen, letter, Number
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([A-Za-z]{1,}[0-9]{1,}^045)([A-Za-z]{1,}[0-9]{1,})^013"
        .Replacement.Text = "^p$pg\1\2[[*lec*]][[*i*]]\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done

    'letter ,Hyphen, letter
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([A-Za-z]{1,}^045)([A-Za-z]{1,})^013"
        .Replacement.Text = "^p$pg\1\2[[*lec*]][[*i*]]\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done

    'letter ,Number, letter
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([A-Za-z]{1,}[0-9]{1,}[A-Za-z]{1,})^013"
        .Replacement.Text = "^p$pg\1^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done

    'Number only
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([0-9]{1,})^013"
        .Replacement.Text = "^p$pg\1^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done
    
    'Number, Letter
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([0-9]{1,}[A-Za-z]{1,})^013"
        .Replacement.Text = "^p$pg\1^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done
    
    'Letter,Number
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([A-Za-z]{1,}[0-9]{1,}[A-Za-z]{1,}[0-9]{1,})^013"
        .Replacement.Text = "^p$pg\1^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done

    'Number, letter, Hyphen, Number, letter,
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([0-9]{1,}[A-Za-z]{1,}^045)([0-9]{1,}[A-Za-z]{1,})^013"
        .Replacement.Text = "^p$pg\1\2[[*lec*]][[*i*]]\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done

    'letter, Number, letter, Hyphen, letter, Number, letter,
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([A-Za-z]{1,}[0-9]{1,}[A-Za-z]{1,}^045)([A-Za-z]{1,}[0-9]{1,}[A-Za-z]{1,})^013"
        .Replacement.Text = "^p$pg\1\2[[*lec*]][[*i*]]\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done
     
    ' remove para mark before graphics (placed at top of macro)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "*~^013"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' remove para mark after graphics (placed at top of macro)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "*~^013"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' remove para mark before graphics
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "*~"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
   
'------------------------------------------------------------------------------------
' replace the plum color paragraph marks
'-------------------------------------------------------------------------------------

    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = wdColorPlum
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorAutomatic
    With Selection.Find
        .Text = "^013"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '********** begin tag roman numerals ***************
    Dim para As Paragraph
    Dim tempStr As String
    Dim txt As String
    Dim Lp_Base_Font_Size As String
    Dim ActualStr As String
     
    ' look at each paragraph in the document
    For Each para In ActiveDocument.Paragraphs
            
        txt = para.Range.Text
        
        ' the longest roman numeral is 10 characters plus 1 for the para mark = 11
        ' when length is greater than 11, then the paragraph is too long to be a pg numb
        ' when > 11 go to next paragraph
        If Len(txt) > 11 Then
          GoTo LoopEnd
        End If
        
        'bypass finding the roman numeral - already tagged - go to next paragraph
        If Left(txt, 3) = "$pg" Then
            GoTo LoopEnd
        End If
        
        ActualStr = Trim(txt) ' actual string can be upper or lower case
        tempStr = Trim(Left(UCase(txt), 11)) ' change to Upper case and take up to 11 characters
        strLength = Len(tempStr) - 1 'set length not including the para mark
        tempStr = Left(tempStr, strLength) ' set comparison string
        'is the string in the lower roman array
        
        If Len(tempStr) > 0 Then
            On Error Resume Next 'prevents crash in a table
            If Sh_IsValidRomanNumeral(tempStr) Then  'is it a valid roman numeral
               If ActiveDocument.Variables("BrailleType") = "UEBT" Or ActiveDocument.Variables("BrailleType") = "UEBN" Then ' this is for UEB
                        para.Range.InsertBefore ("$pg") ' put $pg at front of paragraph
                Else 'this is for EBAE
                    If UCase(ActualStr) = ActualStr Then  'the roman numeral is upper case
                        para.Range.InsertBefore ("$pg") ' put $pg at front of upper roman ref pg
                    Else ' the roaman numeral is lower case
                        para.Range.InsertBefore ("$pg[[*ii*]]") ' put $pg at front of lower roman ref pg
                    End If
                End If
            End If
        End If
LoopEnd:
    Next
    '********** end tag roman numerals ***************
'------------------------------------------------------------------------------------
' make $pg red
'-------------------------------------------------------------------------------------
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Color = wdColorRed
    End With
    With Selection.Find
        .Text = "$pg"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' Replace color and style of DBT Codes. Guarded for the same reason as the passes above:
    ' that macro asks for the "DBT Code" style by name and only the BANA template carries it.
    If Sh_Style_Exists(ActiveDocument, "DBT Code") Then
        Application.Run MacroName:="Dx_Set_DBT_Codes_Color_and_Style"
    End If
    
    Application.ScreenUpdating = su_Prev ' Turn screen updating on
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    ActiveDocument.UndoClear
    
    If Not ActiveWindow.ActivePane.View.ShowAll Then
       ActiveWindow.ActivePane.View.ShowAll = Not ActiveWindow.ActivePane.View.ShowAll
    End If
    
    ' remove para mark placed at top of file
    Selection.HomeKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1
    
    Sh_Return_User_To_Start_Position
    
    'Count the tags
    Dim TagCounter As Integer
    TagCounter = 0
    Dim par As Word.Paragraph

    For Each par In ActiveDocument.Paragraphs
        txt = par.Range.Text
        If InStr(txt, "$pg") > 0 Then
            TagCounter = TagCounter + 1
        End If
    Next
        
    If TagCounter = 0 Then
        MsgBox "There are no tagged page numbers in this document.", , "Braille Macros"
    Else
        If MsgBox("There are " + Trim(Str(TagCounter)) + " page numbers in the document." + vbCr + vbCr _
                + "Do you want to validate the tagged page numbers?", vbYesNo, "Braille Macros") = vbYes Then
            Sh_Validation_Choices_Form.Show
        End If
    End If
    
End Sub   '***** end of Dx_AutoTag_Page_Numbers Macro *************

Function Dx_Merge_Adjacent_Pg_Numbers(ByVal targetDoc As Document) As Long
'
' Combines a RUN of back-to-back reference page numbers into ONE range.
'
'   12 / 13         ->  12-14 is what three give; two give 12-13
'   12 / 13 / 14    ->  12-14
'   12-13 / 14      ->  12-14      (not 12-13-14)
'   B13 / B14       ->  B13-B14
'
' Real print books contain blank pages, and every page still has to carry a number, so DAISY and
' NIMAS coders emit two back to back - almost always at a chapter change. Left alone each gets
' its own reference-page line and they stack up. The large print side has done this since
' 8/2/2026; this is the braille answer, and Jerry asked for it on 8/5/2026.
'
' WHERE THIS RUNS MATTERS, and it is why this works on bare numbers rather than on $pg tags.
' Dx_AutoTag_Page_Numbers has two halves: the first strips every earlier tag, DBT code and
' reference-page style, and the second tags what is left. Called at the seam between them, the
' page numbers are PLAIN - no $pg, no codes, Normal style - so joining 12/13/14 into 12-14 lets
' the existing "number hyphen number" pass produce $pg12-14[[*lec*]][[*i*]]14, DBT continuation
' code and all. Merging after tagging, which is how this was first written, would have meant
' building that code by hand.
'
' ANY two adjacent numbers merge, in sequence or not, so 12 next to 99 gives 12-99. Jerry's call
' on the large print side, 8/2/2026, and it holds here: the "number" is often not a number.
'
' Silent. Screen updating belongs to the caller, which has already turned it off. Returns the
' number of runs merged; nothing reads it, it is there so the work can be checked from the
' Immediate window.
'
' Version: 1.0  Date: 8/5/2026
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
    Dim runs As Collection          ' each item is itself a Collection of paragraph Ranges
    Dim thisRun As Collection
    Dim para As Paragraph
    Dim prevRange As Range
    Dim r As Long

    ' --- 1. Group the page-number paragraphs into runs of genuinely adjacent ones ---
    ' A run is only stored once it reaches two: a lone number is the normal case.
    Set runs = New Collection
    Set thisRun = Nothing

    For Each para In targetDoc.Paragraphs
        If Dx_Is_Bare_Pg_Number_Paragraph(para.Range) Then
            If thisRun Is Nothing Then
                Set thisRun = New Collection
            ElseIf Not Dx_Numbers_Can_Merge(prevRange, para.Range) Then
                Set thisRun = New Collection        ' a section boundary starts a new run
            End If
            thisRun.Add para.Range
            If thisRun.count = 2 Then runs.Add thisRun   ' stored once, then it grows in place
            Set prevRange = para.Range
        Else
            Set thisRun = Nothing                   ' any other paragraph ends the run
        End If
    Next para

    If runs.count = 0 Then Exit Function            ' the ordinary case -- nothing doubled up

    ' --- 2. Merge, working BACK TO FRONT ---
    ' Later runs first, so shortening the document can never move a run still waiting to be
    ' looked at.
    For r = runs.count To 1 Step -1
        Set thisRun = runs(r)
        If Dx_Merge_One_Pg_Number_Run(targetDoc, thisRun) Then
            Dx_Merge_Adjacent_Pg_Numbers = Dx_Merge_Adjacent_Pg_Numbers + 1
        End If
    Next r

End Function   '*** end of Dx_Merge_Adjacent_Pg_Numbers function ***

Private Function Dx_Is_Bare_Pg_Number_Paragraph(ByVal r As Range) As Boolean
'
' True when this paragraph holds nothing but a reference page number, at the point in
' Dx_AutoTag_Page_Numbers where the tags have been stripped and none have been put back.
'
' The test is deliberately close to what the tagging passes below will actually recognize:
' short, and made only of letters, digits and hyphens. It must also carry at least one DIGIT.
'
' That last rule is what keeps a PURE ROMAN pair - v / vi - out of this. A LONE roman numeral is
' not matched by any of the Find passes; it is tagged by the loop near the end of the macro,
' which tests each paragraph with Sh_IsValidRomanNumeral and, for EBAE, puts the lower-case
' marker [[*ii*]] in front.
'
' Merged, the pair would still be tagged - the "letter hyphen letter" pass catches it - but by
' the wrong route. MEASURED, 8/5/2026: v / vi in EBAE gives $pg[[*ii*]]v and $pg[[*ii*]]vi
' today, while a merged v-vi gives $pgv-vi[[*lec*]][[*i*]]vi. The range is right; the lower-case
' marker is gone, because only the roman loop applies it and "v-vi" is not a valid roman
' numeral. In UEB there is no marker either way, so nothing would be lost there.
'
' Leaving roman pairs alone costs nothing - they are still tagged one by one, exactly as before.
' Merging them would mean deciding what a marked-up roman RANGE should look like to Duxbury,
' which is Jerry's call, not something to slip in as a detail.
'
' Requiring a digit also keeps ordinary short paragraphs out: two lines reading "Yes" and "No"
' are not a page range.
'
' Version: 1.0  Date: 8/5/2026
'
    Dim t As String
    Dim s As String
    Dim ch As String
    Dim i As Long
    Dim hasDigit As Boolean
    Dim inTable As Boolean

    ' Ignore anything in a table, exactly as the large print side does. Adjacent in
    ' Document.Paragraphs is NOT adjacent on the page - that collection walks every cell in
    ' document order, so the paragraph after a cell's last one is the first one in the NEXT
    ' cell. A cell must also always keep one paragraph, so deleting into its end-of-cell marker
    ' raises error 4605 and would abort the macro half way through.
    ' If Word cannot tell us, assume it IS in a table - the safe answer is "do not merge".
    inTable = True
    On Error Resume Next
    inTable = r.Information(wdWithInTable)
    On Error GoTo 0
    If inTable Then Exit Function

    t = Sh_Para_Visible_Text(r)
    s = Trim$(Replace(t, ChrW(160), " "))

    ' 12 characters: the roman-numeral loop below uses 11 plus the paragraph mark, and no real
    ' reference page number is longer.
    If Len(s) = 0 Or Len(s) > 12 Then Exit Function

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            hasDigit = True
        ElseIf UCase$(ch) >= "A" And UCase$(ch) <= "Z" Then
            ' a letter is allowed: B13, 14B, A15B
        ElseIf Lp_Is_Hyphen_Char(ch) Then
            ' a hyphen is allowed: 16-18 is already a range and can still grow
        Else
            ' anything else - a space, a tab, a break, the *~ graphic marker this macro
            ' plants earlier - means this is not a bare page number.
            Exit Function
        End If
    Next i

    If Not hasDigit Then Exit Function

    Dx_Is_Bare_Pg_Number_Paragraph = True

End Function   '*** end of Dx_Is_Bare_Pg_Number_Paragraph function ***

Private Function Dx_Numbers_Can_Merge(ByVal a As Range, ByVal b As Range) As Boolean
'
' True when two page-number paragraphs really are stacked one above the other in the same text
' flow. Tables are already out - Dx_Is_Bare_Pg_Number_Paragraph rejects them - so two tests:
'
'   1. They are touching: a ends exactly where b starts.
'   2. They are in the same SECTION. A section break lives on a paragraph mark, so a run that
'      spanned one would lose the break and that section's page setup with it. Two page numbers
'      back to back at a chapter change is exactly where a break turns up.
'
' Version: 1.0  Date: 8/5/2026
'
    On Error GoTo NoMerge

    If a.End <> b.start Then Exit Function
    If a.Sections(1).Index <> b.Sections(1).Index Then Exit Function

    Dx_Numbers_Can_Merge = True
    Exit Function

NoMerge:                                            ' could not tell, so do not merge
    Dx_Numbers_Can_Merge = False

End Function   '*** end of Dx_Numbers_Can_Merge function ***

Private Function Dx_Merge_One_Pg_Number_Run(ByVal targetDoc As Document, _
                                            ByVal numRun As Collection) As Boolean
'
' Turns one run of page-number paragraphs into a single range: writes the merged number into the
' LAST paragraph of the run, then deletes the earlier ones WHOLE, back to front.
'
' Both choices come from the large print side and both were learned the hard way on 8/2/2026:
' whole-paragraph deletes, because a single span across the run left a stray paragraph mark
' behind whenever the run was three or more; and the LAST paragraph survives, because Word
' refuses to delete a document's final paragraph mark.
'
' Version: 1.0  Date: 8/5/2026
'
    Dim firstRng As Range, lastRng As Range
    Dim writeRng As Range
    Dim firstText As String, lastText As String
    Dim mergedNum As String
    Dim i As Long

    If numRun.count < 2 Then Exit Function

    Set firstRng = numRun(1)
    Set lastRng = numRun(numRun.count)

    firstText = Sh_Para_Visible_Text(firstRng)
    lastText = Sh_Para_Visible_Text(lastRng)

    mergedNum = Dx_Merged_Pg_Number(Trim$(firstText), Trim$(lastText))
    If Len(mergedNum) = 0 Then Exit Function

    ' --- 1. Write the range into the LAST paragraph, text only, not the paragraph mark ---
    Set writeRng = targetDoc.Range(lastRng.start, lastRng.start + Len(lastText))

    ' Character positions and Range offsets only line up while the paragraph is plain text. If a
    ' field or an inline shape has crept in they will not, so check before writing.
    If StrComp(writeRng.Text, lastText, vbBinaryCompare) <> 0 Then Exit Function
    If StrComp(lastText, mergedNum, vbBinaryCompare) <> 0 Then writeRng.Text = mergedNum

    ' --- 2. Delete the earlier paragraphs of the run, whole, back to front ---
    For i = numRun.count - 1 To 1 Step -1
        numRun(i).Delete
    Next i

    Dx_Merge_One_Pg_Number_Run = True

End Function   '*** end of Dx_Merge_One_Pg_Number_Run function ***

Private Function Dx_Merged_Pg_Number(ByVal firstNum As String, ByVal lastNum As String) As String
'
' The merged range: the FIRST part of the first number and the LAST part of the last number,
' joined with one hyphen. That is what stops an already-hyphenated number from growing a second
' hyphen -- "12-13" next to "14" gives "12-14", not "12-13-14".
'
' A hyphen at either extreme is a stray, not a range boundary, so "-13" keeps its whole text.
' Two identical numbers give a single one rather than "12-12".
'
' Lp_Is_Hyphen_Char is reused rather than copied: it is a pure character test with nothing large
' print about it, and two copies would be two things to keep in step.
'
' Version: 1.0  Date: 8/5/2026
'
    Dim head As String, tail As String
    Dim i As Long

    ' head: everything before the FIRST hyphen of the first number
    For i = 1 To Len(firstNum)
        If Lp_Is_Hyphen_Char(Mid$(firstNum, i, 1)) Then Exit For
    Next i
    If i > 1 And i <= Len(firstNum) Then head = Trim$(Left$(firstNum, i - 1))
    If Len(head) = 0 Then head = Trim$(firstNum)

    ' tail: everything after the LAST hyphen of the last number
    For i = Len(lastNum) To 1 Step -1
        If Lp_Is_Hyphen_Char(Mid$(lastNum, i, 1)) Then Exit For
    Next i
    If i >= 1 And i < Len(lastNum) Then tail = Trim$(Mid$(lastNum, i + 1))
    If Len(tail) = 0 Then tail = Trim$(lastNum)

    If Len(head) = 0 Or Len(tail) = 0 Then Exit Function

    If StrComp(head, tail, vbTextCompare) = 0 Then
        Dx_Merged_Pg_Number = head
    Else
        Dx_Merged_Pg_Number = head & "-" & tail
    End If

End Function   '*** end of Dx_Merged_Pg_Number function ***



Sub Dx_Replace_Straight_Quotes_With_Smart_Quotes()
'
' Dx_Replace_Straight_Quotes_With_Smart_Quotes Macro
'
' Version 1.4
' Date 12/11/2015
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Call: Application.Run MacroName:="Dx_Replace_Straight_Quotes_With_Smart_Quotes"
'
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating off

    '*********************************************************************
    ' Find replace single straight quotes with curly quotes
    '*********************************************************************
    With Selection.Find
        .Text = "'"
        .Replacement.Text = "'"
        .Forward = True

        If Selection.Type <> wdSelectionNormal Then   'text is NOT selected"
            .Wrap = wdFindContinue  'replaces all in the document
        Else
            .Wrap = wdFindStop  'replaces only the selected text
        End If
        
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    '*********************************************************************
    ' Find replace double straight quotes with curly quotes
    '*********************************************************************
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = """"
            .Replacement.Text = """"
            .Forward = True

            If Selection.Type <> wdSelectionNormal Then   'text is NOT selected"
                .Wrap = wdFindContinue  'replaces all in the document
            Else
                .Wrap = wdFindStop  'replaces only the selected text
            End If
        
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        Application.ScreenUpdating = su_Prev ' Turn screen updating on
        
End Sub  '***** end of Dx_Replace_Straight_Quotes_With_Smart_Quotes Macro ******

Sub Dx_Format_Exercise_Lv_1_and_Lv_2()

' Dx_Format_Exercise_Lv_1_and_Lv_2 Macro
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Version: 1.8  Date: 7/26/2026 - returns the user to where the cursor was when the macro started
' Version: 1.6 Date: 8/21/2018 - Modifed to work with Nemeth
'
    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"
    
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "Select the exercise list first!", , "Braille Macros"
        End
    End If

    Sh_Save_User_Position   ' record the spot HERE, in the user's document, before Dx_Copy_To_Temp_Doc makes the temp file active

    Dx_UEB_EBAE_Fill_In_YN_Form.Show  'ask user if fill-in indicators are wanted for answers
    Unload Dx_UEB_EBAE_Fill_In_YN_Form

    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    'Application.ScreenUpdating = False ' Turn screen updating off
    Application.Run MacroName:="Dx_Copy_To_Temp_Doc"
    Application.Run MacroName:="Sh_Is_End_Paragraph_Mark_Included"
    Application.ScreenUpdating = False ' Turn screen updating off
    Application.Run MacroName:="Dx_Convert_Auto_List_To_Text"
    Application.Run MacroName:="Dx_Fix_Para_Space_Errors"
    Application.Run MacroName:="Sh_Remove_Multi_Spaces"
    Application.ScreenUpdating = False ' Turn screen updating off
    Selection.WholeStory
    If Sh_Style_Exists(ActiveDocument, "Body Text") Then
    Selection.Style = ActiveDocument.Styles("Body Text")
    End If
    Selection.HomeKey Unit:=wdStory


'----------------- oops section begin -------------------------
' --this section removes previous formatting so that-----------
'---the macro can be run again with different fill-in types----

    '*****************************************************
    'Remove kps if item previously marked
    '*****************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[[*kps*]]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '*****************************************************
    'Remove kpe if item previously marked
    '*****************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "[[*kpe*]]"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '*****************************************************
    'Remove UEB and EBAE tab fill-ins if item previously marked
    '*****************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032----^032"
        .Replacement.Text = "^032^t^032"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' replace EBAE fill-ins with tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032----^044"
        .Replacement.Text = "^032^t^044"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' replace EBAE fill-ins with tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032----^046"
        .Replacement.Text = "^032^t^046"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
'------------
    ' replace UEB fill-ins with tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032_^044"
        .Replacement.Text = "^032^t^044"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' replace UEB fill-ins with tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032_^046"
        .Replacement.Text = "^032^t^046"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

   ' replace space before underscore with tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032_^032"
        .Replacement.Text = "^032^t^032"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' replace multiple underscores with single tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^095{1,}"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' replace space underscore followed with comma with tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032_^044"
        .Replacement.Text = "^032^t^044"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' replace underscore followed by a period with a tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032_^046"
        .Replacement.Text = "^032^t^046"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
    '*****************************************************
    'Remove EBAE exercise level 2 fill-ins if item previously marked
    ' para mark folloed by EBAE hyphens followed by a space replace with para mark
    '*****************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013----^032"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    '*****************************************************
    'Remove UEB exercise level 2 fill-ins if item previously marked
    ' para mark folloed by UEB Underscore followed by a space with para mark
    '*****************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013_^032"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
           
    ' replace tabs followed by space at beging of level 2 items
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013^009^032"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' ---------- end of oops section ------------------------------

    '*****************************************************
    ' remove mulitiple para marks
    '*****************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

     With Selection.Find
        .Text = "^013{1,}"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' **************************************************************
    ' force the F&R to place [[*kpe*]] at the end of the last real
    ' entry in the list by putting in fake number - removed later
    ' **************************************************************
    
    Selection.EndKey Unit:=wdStory
    Selection.TypeText Text:="99. "
    
    ' **************************************************************
    ' force the F&R to place [[*kps*]] at the start of the first real
    ' entry in the list by putting in fake number - removed later
    ' **************************************************************
    
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.HomeKey Unit:=wdStory
    Selection.TypeText Text:="99. "
    'Selection.TypeParagraph

    ' **************************************************************
    ' Set all para to Exercise 2
    ' **************************************************************
    If Dx_GP_String_1 = "UEBN" Or Dx_GP_String_1 = "EBAN" Then ' this is Nemeth
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "Ex2Nemeth2") Then
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles("Ex2Nemeth2")
        With Selection.Find
            .Text = "^013"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
    Else ' this is text
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "Exercise2") Then
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        Selection.Find.Replacement.Style = ActiveDocument.Styles("Exercise2")
        With Selection.Find
            .Text = "^013"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
    End If

    ' **************************************************************
    ' find para marks followed numbers followd by period and space
    ' **************************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Exercise1") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Exercise1")
     With Selection.Find
        .Text = "^013([0-9]{1,}^046^032)"
        .Replacement.Text = "[[*kpe*]]^p[[*kps*]]\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' **************************************************************
    ' find para marks followed numbers followd by a space only
    ' **************************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Exercise1") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Exercise1")
     With Selection.Find
        .Text = "^013([0-9]{1,}^032)"
        .Replacement.Text = "[[*kpe*]]^p[[*kps*]]\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If

    ' **************************************************************
    ' find para marks followed numbers followd by period
    ' **************************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Exercise1") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Exercise1")
     With Selection.Find
        .Text = "^013([0-9]{1,}^046)"
        .Replacement.Text = "[[*kpe*]]^p[[*kps*]]\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' ****************************************************************************
    ' find para marks followed numbers followed a closed paren followed by a period
    ' ****************************************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Exercise1") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Exercise1")
     With Selection.Find
        .Text = "^013([0-9]{1,}\)^046)"
        .Replacement.Text = "[[*kpe*]]^p[[*kps*]]\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' ****************************************************************************
    ' find para marks followed numbers followed a closed paren not followed by a period
    ' ****************************************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Exercise1") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Exercise1")
     With Selection.Find
        .Text = "^013([0-9]{1,}\))"
        .Replacement.Text = "[[*kpe*]]^p[[*kps*]]\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' ****************************************************************************
    ' find para marks followed numbers enclosed in parens followed by a period
    ' ****************************************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Exercise1") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Exercise1")
     With Selection.Find
        .Text = "^013(\([0-9]{1,}\)^046)"
        .Replacement.Text = "[[*kpe*]]^p[[*kps*]]\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' ****************************************************************************
    ' find para marks followed numbers enclosed in parens not followed by a period
    ' ****************************************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Exercise1") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Exercise1")
     With Selection.Find
        .Text = "^013(\([0-9]{1,}\))"
        .Replacement.Text = "[[*kpe*]]^p[[*kps*]]\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If

    ' **************************************************************
    ' change color of kpe
    ' **************************************************************
        If Dx_GP_String_1 = "UEBN" Or Dx_GP_String_1 = "EBAN" Then ' this is Nemeth
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "Ex2Nemeth2") Then
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find.Replacement.Font
        Selection.Find.Replacement.Style = ActiveDocument.Styles("Ex2Nemeth2")
            .Hidden = True
            .Color = wdColorPlum
        End With
         With Selection.Find
            .Text = "(\[\[\*kpe\*\]\])"
            .Replacement.Text = "[[*kpe*]]"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
    Else ' this is text
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "Exercise2") Then
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find.Replacement.Font
        Selection.Find.Replacement.Style = ActiveDocument.Styles("Exercise2")
            .Hidden = True
            .Color = wdColorPlum
        End With
         With Selection.Find
            .Text = "(\[\[\*kpe\*\]\])"
            .Replacement.Text = "[[*kpe*]]"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
    End If

    ' **************************************************************
    ' change color of kps
    ' **************************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Exercise1") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Exercise1")
        .Hidden = True
        .Color = wdColorPlum
    End With
     With Selection.Find
        .Text = "(\[\[\*kps\*\]\])"
        .Replacement.Text = "[[*kps*]]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If

    ' **************************************************************
    ' make plum color style DBT Code
    ' **************************************************************
'    Selection.Find.ClearFormatting
'    Selection.Find.Font.Color = wdColorPlum
'    Selection.Find.Replacement.ClearFormatting
'    Selection.Find.Replacement.Style = ActiveDocument.Styles("DBT Code")
'    With Selection.Find
'        .Text = ""
'        .Replacement.Text = ""
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = True
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
'    End With
'    Selection.Find.Execute Replace:=wdReplaceAll

    '*****************************************************
    ' Make sure that all [[*kps*]] are level 1
    '*****************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Exercise1") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Exercise1")
    With Selection.Find.Replacement.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
    End With
    With Selection.Find
        .Text = "[[*kps*]]"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If

    '*****************************************************
    ' add exercise level 2 fill-in indicators
    '*****************************************************
    If Dx_UEB_EBAE_Boolean = True Then 'Fill-in indicators are wanted
        ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
        ' this pass can do nothing without it - no style means nothing is formatted with it.
        If Sh_Style_Exists(ActiveDocument, "Exercise2") Then
        Selection.Find.ClearFormatting
        Selection.Find.Style = ActiveDocument.Styles("Exercise2")
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "(*{1,}^013)"
            
            If Dx_GP_String_1 = "EBAT" Or Dx_GP_String_1 = "EBAN" Then 'EBAE fill-in indicators
                .Replacement.Text = "---- \1"
            End If
            
            If Dx_GP_String_1 = "UEBT" Or Dx_GP_String_1 = "UEBN" Then 'UEB fill-in indicators
                .Replacement.Text = "_ \1"
            End If

            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        End If
    End If
    
    ' Convert Tabs to fill Ins
    Application.Run MacroName:="Dx_Tabs_To_Fill_Ins"
    
    '*****************************************************
    ' remove spaces preceeding square left brace
    '*****************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "DBT Code") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("DBT Code")
    With Selection.Find.Replacement.Font
        .Hidden = True
        .Color = wdColorPlum
    End With
    With Selection.Find
        .Text = " ["
        .Replacement.Text = "["
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    '*****************************************************
    ' remove the bogus items from top
    '*****************************************************
    Selection.HomeKey Unit:=wdStory
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.Delete Unit:=wdCharacter, count:=1
    
    '*****************************************************
    ' remove the bogus items from bottom
    '*****************************************************
    Selection.EndKey Unit:=wdStory
    Selection.HomeKey Unit:=wdLine
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.EndKey Unit:=wdLine
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.Delete Unit:=wdCharacter, count:=1


    Selection.HomeKey Unit:=wdStory

    Application.Run MacroName:="Sh_Remove_Multi_Spaces"
    Application.Run MacroName:="Dx_Copy_From_Temp_Doc"

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    'ActiveDocument.UndoClear ' No undo
    Application.ScreenUpdating = su_Prev ' Turn screen updating on

    Sh_Return_User_To_Start_Position

End Sub  '***** end of Dx_Format_Exercise_Lv_1_and_Lv_2 Macro *****

Sub Dx_Fix_Ellipsis_Errors()
'
' Dx_Fix_Ellipsis_Errors() Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Date: 12/2/2015
' Version: 1.0
'
' *************************************************************
' Find ellipsis.. place space before and after
' *************************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^0133"
        .Replacement.Text = "^032^0133^032"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
' *************************************************************
' Find one or more spaces before ellipsis... make it a single space before
' *************************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}^0133"
        .Replacement.Text = "^032^0133"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' *************************************************************
'  Find one or more spaces after ellipsis... make it a single space after
' *************************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^0133^032{1,}"
        .Replacement.Text = "^0133^032"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' *************************************************************
' Find ellipsis followed by space followed by para mark... delete space
' *************************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^0133^032^013"
        .Replacement.Text = "^0133^013"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' *************************************************************
' Find ellipsis followed by space followed by question mark... delete space
' *************************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^0133^032^063"
        .Replacement.Text = "^0133^063"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' *************************************************************
' Find ellipsis followed by space followed by exclmation point... delete space
' *************************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^0133^032^033"
        .Replacement.Text = "^0133^033"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' *************************************************************
' Find ellipsis followed by space followed by period... delete space
' *************************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^0133^032^046{1,}"
        .Replacement.Text = "^0133^046"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' *************************************************************
' replace ellipsis with three periods
' *************************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^0133"
        .Replacement.Text = "^046^046^046"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

' *************************************************************
' Find ellipsis followed by alpha with no space... add space
' *************************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^0133([A-Za-z]{1,})"
        .Replacement.Text = "^0133^032\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
' *************************************************************
' replace all ellipsis with three periods
' *************************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^0133"
        .Replacement.Text = "^046^046^046"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub  ' ***** end of Dx_Fix_Ellipsis_Errors() Macro *****

Sub Dx_Is_Text_Selected()
'
' Dx_Is_Text_Selected Macro

' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version 1.0
' Date: 2/13/2015
'
'--------------------------------------------------------------------------------
' Checks if text is selected
'--------------------------------------------------------------------------------

    If Selection.Type <> wdSelectionNormal Then
        MsgBox "Text must be selected first!", , "Braille Macros"
        End
    End If
    
End Sub '***** End of Dx_Is_Text_Selected *************

Sub Dx_Copy_To_Temp_Doc()
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version: 1.3  Date: 3/19/2021 - added Dx_Attach_Same_BANA_Template to end of procedure
' Version: 1.2  Date: 9/9/2018
' Dx_Attached_BANA_Template is a public variable

    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"
    Dx_Attached_BANA_Template = ActiveDocument.AttachedTemplate 'get the name of the orig doc template
    Dx_GP_String_1 = ActiveDocument.Variables("BrailleType") ' put the braille xlation type in public var
    Selection.Copy 'copy the selected text in the main document
    Application.ScreenUpdating = False ' Turn screen updating off
    Documents.Add 'create a new temp wd doc
    Sh_SleepForSeconds 3
    Selection.Paste 'AndFormat (wdFormatOriginalFormatting)
    Application.Run MacroName:="Dx_Attach_Same_BANA_Template"

End Sub '***** end of Dx_Copy_To_Temp_Doc Macro *****

Sub Dx_Copy_From_Temp_Doc()
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
' Date: 9/9/2018
' Version: 1.5
'
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating off
    Selection.WholeStory
    Selection.Copy 'copy the selected text to the clipboard
    ActiveDocument.Close SaveChanges:=False 'close the temp doc without saving
    Application.ScreenUpdating = su_Prev ' Turn screen updating on
    Selection.Paste 'paste the clipboard back into the original document
    Application.Run MacroName:="MS_Set_Word_Config_For_Braille"

End Sub '***** end of Dx_Copy_From_Temp_Doc Macro *****

Sub Dx_Attach_Same_BANA_Template()
'
' Attaches the same BANA Template as the original DOC to the copied selection
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version: 1.6  Date: 7/9/2021 - added existance check and message for attached BANA template
' Version: 1.5  Date: 10/17/2020 - improved attachment process
' Version: 1.4  Date: 3/16/2017
'

     Dim TemplatePathandName As String
    TemplatePathandName = Options.DefaultFilePath(wdUserTemplatesPath) + "\" + Dx_Attached_BANA_Template
    
    With ActiveDocument
        If Sh_FileExists(TemplatePathandName) Then
            .UpdateStylesOnOpen = True
            .AttachedTemplate = TemplatePathandName
            .UpdateStylesOnOpen = False  ' supresses any further style updates
        Else
            MsgBox " Cannot continue!" + vbCr + vbCr + "The template file: " + TemplatePathandName + " does not exist." + vbCr + vbCr + "Install the file and try again.", , "Braille Macros"
            End
        End If
    End With
    
End Sub  '***** End of Dx_Attach_Same_BANA_Template Macro *****

' Dx_Kill_The_Hyperlinks was here until 8/12/2026. With the round trip gone the two sides were
' the same macro written twice, so they are now Sh_Kill_The_Hyperlinks.


Sub Dx_Remove_Bullets()
'
' Removes bullets and automatic numbering from a list and, if the transcriber asked for it on
' Dx_Bullet_Removal_Form, restyles the paragraphs and strips the hyperlinks.
'
' Version: 2.0  Date: 8/13/2026 - no temporary document; everything is scoped to a range.
'
'                                Three things here reached past a selection on their own:
'                                Sh_Remove_Hyperlinks, whose default is the whole document;
'                                Dx_Convert_Auto_List_To_Text, which was whole-document
'                                throughout; and every Find below, built on Selection.Find with
'                                wdFindContinue. The scratch document was holding all three
'                                back, so this is real work the round trip was doing.
'
'                                The bullet-removal form is still only shown when text is
'                                selected. That is not plumbing either - with nothing selected
'                                this runs over the whole document with the defaults, which is
'                                how the cleanup sequences call it.
'
'                                Gone with the round trip: the four-character delete run.
' Version: 1.1  Date: 11/16/2016
'
    Dim rng As Range
    Dim su_Prev As Boolean
    Dim Limited_Selection As Boolean

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    If Selection.Type <> wdSelectionNormal Then   'text is NOT selected"
        Limited_Selection = False
    Else
        Limited_Selection = True
        Dx_Bullet_Removal_Form.Show
    End If

    ' AFTER the form - it is modal, and the transcriber can change the selection behind it.
    If Limited_Selection = True Then
        Set rng = Selection.Range
    Else
        Set rng = ActiveDocument.Content
    End If

    If Dx_GP_String_1 = "Remove_Hyper" Then
        Sh_Remove_Hyperlinks rng
    End If

    Dx_Convert_Auto_List_To_Text rng

    Sh_Delete_Text_In_Range rng, ChrW(61623)   ' the Symbol-font bullet
    Sh_Delete_Text_In_Range rng, "^0149"       ' the ordinary bullet
    Sh_Delete_Text_In_Range rng, "_9"
    Sh_Delete_Text_In_Range rng, "_4"

    If Dx_GP_String_2 <> "U" Then
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting

            If Dx_GP_String_2 = "T1" And Sh_Style_Exists(ActiveDocument, "TOC 1") Then
                .Replacement.Style = ActiveDocument.Styles("TOC 1")
            End If
            If Dx_GP_String_2 = "T2" And Sh_Style_Exists(ActiveDocument, "TOC 2") Then
                .Replacement.Style = ActiveDocument.Styles("TOC 2")
            End If
            If Dx_GP_String_2 = "L1" And Sh_Style_Exists(ActiveDocument, "List1") Then
                .Replacement.Style = ActiveDocument.Styles("List1")
            End If
            If Dx_GP_String_2 = "L2" And Sh_Style_Exists(ActiveDocument, "List2") Then
                .Replacement.Style = ActiveDocument.Styles("List2")
            End If
            If Dx_GP_String_2 = "B" And Sh_Style_Exists(ActiveDocument, "Body Text") Then
                .Replacement.Style = ActiveDocument.Styles("Body Text")
            End If

            .Text = "^013"
            .Replacement.Text = "^p"
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    End If

    If Limited_Selection = True Then
        ' Dx_Fix_Para_Space_Errors scopes itself to the selection, which is still the
        ' transcriber's own - nothing here took it away.
        Application.Run MacroName:="Dx_Fix_Para_Space_Errors"
    End If

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub '****** end of Dx_Remove_Bullets Macro *****

Public Sub Sh_Delete_Text_In_Range(ByVal rng As Range, ByVal what As String)
'
' Version: 1.0  Date: 8/13/2026
'
' Deletes every occurrence of a literal string from a range. Dx_Remove_Bullets had four of these
' written out longhand, forty lines of Find options differing only in one line.
'
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = what
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

End Sub   '*** end of Sh_Delete_Text_In_Range ***

Sub Dx_Remove_Optional_Hyphens()
'
' Dx_Remove_Optional_Hyphens Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
' Version 1.0
' Date: 1/15/2017
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^031"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub '***** end of Dx_Remove_Optional_Hyphens macro *****
Sub Dx_Remove_Page_Breaks()
'
' Dx_Remove_Page_Breaks Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
' Version: 1.0 Date: 1/15/2017
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^m^013{1,}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub '****** end of Dx_Remove_Page_Breaks Macro ******
Sub Dx_Remove_Section_Breaks()
'
' Dx_Remove_Section_Breaks Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
' Version: 1.0  Date: 1/15/2017
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^b"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub '***** end of Dx_Remove_Section_Breaks Macro *****

Sub Dx_Remove_Column_Breaks()
'
' Dx_Remove_Column_Breaks Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
' Version: 1.0 Date: 1/15/2017
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^n"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub '***** end of Dx_Remove_Column_Breaks Macro ******

Sub Dx_Tabs_To_Fill_Ins()
'
' Dx_Tabs_To_Fill_Ins macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Date: 1/9/2017 Version: 1.4
'

    'remove underlines from tabs
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    'put space on each side of tabs
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^009"
        .Replacement.Text = "^032^09^032"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'remove extra spaces
    'Application.Run MacroName:="Sh_Remove_Multi_Spaces"

    'remove spaces before periods
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032^046"
        .Replacement.Text = "^046"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    'replace tab folowed by space followed by para-mark
    'and insert a period in place of the space
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^09^032^013"
        .Replacement.Text = "^09^046^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'replace each tab with user choice for fill-in indicator
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t"


        If Dx_GP_String_1 = "EBAT" Or Dx_GP_String_1 = "EBAN" Then 'EBAE fill-in indicator
            .Replacement.Text = "----"
        End If
        
        If Dx_GP_String_1 = "UEBT" Or Dx_GP_String_1 = "UEBN" Then  'UEB fill-in indicator
            .Replacement.Text = "_"
        End If

        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' remove spaces to left of commas
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032^044"
        .Replacement.Text = "^044"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub  '***** end of Dx_Tabs_To_Fill_Ins macro *****

Sub Dx_Convert_Hyper_To_Addresses()
'
' Replaces hyperlinks with the address of the hyperlink. The address is left as plain text -
' a Duxbury source file has no use for a clickable link, which is the one thing that separates
' this from the large-print copy.
'
' from: http://stackoverflow.com/questions/16493791/
'     extract-hyperlink-address-from-hyperlink-field-code
'
' Version: 2.0 Date: 8/13/2026 - no temporary document; everything is scoped to a range. See the
'                               note in Lp_Convert_Hyper_To_Addresses - Hyperlinks is a
'                               whole-document collection, so the scratch document was doing
'                               real work. Gone with it: the four-character delete run and
'                               "delete one character at the end of the document".
' Version: 1.2 Date: 11/16/2016
'
    Dim rng As Range
    Dim r As Range
    Dim strLinkText As String
    Dim i As Long
    Dim su_Prev As Boolean

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    If Selection.Type = wdSelectionNormal Then
        Set rng = Selection.Range
    Else
        Set rng = ActiveDocument.Content
    End If

    ' Every hyperlink in scope becomes its own address, as plain text. Backwards by index:
    ' writing r.Text takes the hyperlink out of the collection and changes the length of the
    ' text, so anything counted from the front has moved by the next turn of the loop.
    On Error Resume Next
    For i = rng.Hyperlinks.count To 1 Step -1
        With rng.Hyperlinks(i)
            strLinkText = .Address
            ' optional, should be OK for HTML links
            If .SubAddress <> "" Then
                strLinkText = strLinkText & "#" & .SubAddress
            End If
            Set r = .Range
        End With
        r.Text = strLinkText
        Set r = Nothing
    Next i
    On Error GoTo 0

    ' "mailto:" is how Word writes an e-mail address and is not part of the address a reader
    ' wants to see
    Sh_Delete_Text_In_Range rng, "mailto:"
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub  '*** end of Dx_Convert_Hyper_To_Addresses ***

' Dx_Replace_Manual_Line_Break was here until 8/12/2026 - see Sh_Replace_Manual_Line_Break.


Sub Dx_Delete_Images()
'
'   Dx_Delete_Images Macro
'
'   Jerry Whittaker - jerry@thewhittakers.org
'
'   Version: 1.6 Date: 2/20/2024 - added caution about deleteing images created by MathType
'   Version: 1.5 Date: 9/26/2023 - turned screen off to prevent scrolling text on screen
'   Version: 1.4 Date: 9/4/2020 - Commened out the section to put red border around images.
'   Version: 1.3 Date: 9/23/2018 - added converstion of all images to inline and added new para mark following each image
'   Version: 1.2 Date: 9/9/2018 - added count and delete question
'   Version: 1.1 Date: 5/12/2016
'
'  inline image count routine from: https://answers.microsoft.com/en-us/office/forum/office_2010-word/is-there-a-way-to-count-how-many-images-are-in-a/28061f1a-403b-42a1-829f-fd65c9a66395?db=5
    
    Application.ScreenUpdating = False    ' Turn screen updating off
    
    ' Convert Images and Shapes to Inline
    ' From: https://stackoverflow.com/questions/15919815/vba-macro-make-all-shapes-and-pictures-to-inline-with-text?rq=1
    Dim oShp As Shape
    For Each oShp In ActiveDocument.Shapes
       oShp.Select
       Selection.ShapeRange.ConvertToInlineShape
    Next oShp
    
    ' place a para mark following each image - following text is sometime part of the image paragraph
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^g"
        .Replacement.Text = "^&^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' count inline images and if found ask user if they wish to delete
    Dim inlines As Long

    With ActiveDocument
        inlines = .InlineShapes.count
    End With
    
    If inlines > 0 Then
        Dim Answer As String
           Answer = MsgBox("Found " & Str(inlines) & " images in the document. Delete all?" _
            & vbCrLf & vbCrLf & "************************************************" _
            & vbCrLf & "CAUTION!  CAUTION!   CAUTION!   CAUTION!" _
            & vbCrLf & "************************************************" _
            & vbCrLf & "A 'Yes' answer will also delete equations created" _
            & vbCrLf & "by MathType but will leave OMML* equations" _
            & vbCrLf & "intact." _
            & vbCrLf & vbCrLf & "*Office Math Mark-up Language equations" _
            & vbCrLf & "created by the Word Equation Editor", _
                       VBA.VbMsgBoxStyle.vbInformation + VBA.VbMsgBoxStyle.vbYesNo + VBA.VbMsgBoxStyle.vbDefaultButton2, _
                       "Braille Macros")
    
            If Answer = vbYes Then
                Selection.Find.ClearFormatting
                Selection.Find.Replacement.ClearFormatting
                With Selection.Find
                    .Text = "^g"
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Selection.Find.Execute Replace:=wdReplaceAll
            Else
                ' From: http://www.vbaexpress.com/forum/showthread.php?48376-VBA-Macro-To-Insert-Border-For-All-Images
                ' Author: gmcnultnult
                'Dim oInlineShp As InlineShape
                '    For Each oInlineShp In ActiveDocument.InlineShapes
                '    With oInlineShp
                '        With .Borders(wdBorderLeft)
                '            .LineStyle = wdLineStyleSingle
                '            .LineWidth = wdLineWidth600pt
                '            .Color = wdColorRed
                '        End With
                '            With .Borders(wdBorderRight)
                '            .LineStyle = wdLineStyleSingle
                '            .LineWidth = wdLineWidth600pt
                '            .Color = wdColorRed
                '        End With
                '            With .Borders(wdBorderTop)
                '            .LineStyle = wdLineStyleSingle
                '            .LineWidth = wdLineWidth600pt
                '            .Color = wdColorRed
                '        End With
                '        With .Borders(wdBorderBottom)
                '            .LineStyle = wdLineStyleSingle
                '            .LineWidth = wdLineWidth600pt
                '            .Color = wdColorRed
                '        End With
                '    End With
                'Next
                'MsgBox "Images will have a heavy red border when Word is in Print View", , "Braille Macros"
            End If
    End If
    
End Sub  '*** end of Dx_Delete_Images Macro ***
Sub Dx_Type_Dashes()
'
' Dx_Type_Dashes Macro
'
' Description: Types long, em, en, and minus dashes
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Version: 1.1  Date: 10/5/2018 - added check for attached template
' Version: 1.0  Date: 10/9/2016
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"
    Dx_Type_Dashes_Form.Show
    
End Sub  '*** end of Dx_Type_Dashes macro ***

Sub Dx_Set_DBT_Codes_Color_and_Style()
'
' Dx_Set_DBT_Codes_Color_and_Style Macro
'
' Version: 1.1 Date: 3/14/2017
' Author: Jerry Whittaker - jerry@thewhittakers.org

    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "DBT Code") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("DBT Code")
    With Selection.Find.Replacement.Font
        .Hidden = True
        .Color = wdColorPlum
    End With
    With Selection.Find
        .Text = "\[\[\*?{1,}\*\]\]"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"

End Sub   '*** end of Dx_Set_DBT_Codes_Color_and_Style Macro ***
Sub Dx_Fix_Primes()

    ' Dx_Fix_Primes Macro
    '
    ' converts smart (curley) single and double quotes and
    '   single and double stright quotes to proper primes
    '
    ' Author: Jerry Whittaker -  jerry@thewhittakers.org
    '
    ' Version: 1.0Date: 11/13/2016
    '
    'replace single close smart (curly) quote with single prime
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "([0-9]{1,})^0146"
        .Replacement.Text = "\1" & ChrW(8242)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'replace double close smart (curly) quote with double prime
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "([0-9]{1,})^0148"
        .Replacement.Text = "\1" & ChrW(8243)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'replace single stright quote with single prime
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "([0-9]{1,})^039"
        .Replacement.Text = "\1" & ChrW(8242)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'replace double stright quote with double prime
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "([0-9]{1,})^034"
        .Replacement.Text = "\1" & ChrW(8243)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub   '*** end of Dx_Fix_Primes Macro ***
Sub Dx_Set_Doc_Braille_Type_Variable()
'
' Dx_Set_Doc_Braille_Type_Variable macro
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
' Version 1.1 Date: 3/17/2017
'
' Sets document variable "BrlType" which holds whether translation is UEB or EBAE / texbook or nemeth
'

    Dim BrlType As String
    
    'if the BrailleType stored in the document variables is non existant then
    ' create an undefined BrailleType
    BrlType = "Undefined"
    On Error Resume Next 'error occurs when doc variable does not exist
    BrlType = ActiveDocument.Variables("BrailleType") ' if Doc Var does not exist then BrlType = ""
    'then as ask the user if this document is for UEB or EBAE ' EBAE and UEB are legacy types from previous versions of macros
    If BrlType = "Undefined" Or BrlType = "" Or BrlType = "UEB" Or BrlType = "EBAE" Then 'get the braille type the user wants
        Dx_UEB_EBAE_String = ""
        Do While Dx_UEB_EBAE_String = ""
            Dx_Choose_Translation_Form.Show ' sets the value of the doc variable and sets Dx_UEB_EBAE_String
        Loop
    End If
    
    Unload Dx_Choose_Translation_Form
    
End Sub   '*** end of Dx_Set_Doc_Braille_Type_Variable macro ***

Sub Dx_Remove_Keep_With_Next()
'
' Dx_Remove_Keep_With_Next Macro
'
' removes the keep with next (paragraph) paramater from all paragraphs
' Author: Jerry Whittaker - jerry@thewhittakers.org
' Version: 1.0  Date: 3/13/2018
'
    Selection.Find.ClearFormatting
    With Selection.Find.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .KeepWithNext = True
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .KeepWithNext = False
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub  '***end of Dx_Remove_Keep_With_Next macro ***

Sub Dx_Fix_Abbyy_FineReader_Text_and_Headers()
'
' Macro Dx_Fix_Abbyy_FineReader_Text_and_Headers()
'
' Word can have up to nine heading levels.  In Braille there are only three Heading levels.
'
' This macro searches for up to nine Abbyy FineReader 14 headers styles (i.e. "Heading #1" through "Heading #9")
'   and converts them to standard heading styles (without the # sign). Abbyy Headings #4 through
'   Headings #9 are set to braille heading level 3.
'
' This macro also changes all all of Abbyy's "Normal" and "Body text (1)" through
'    "Body text (9)" styles to body text style.
'
' This macro ONLY works with Abbyy 14 documents saved as "Formatted Text". Running this
' macro on earlier version of Abbyy may produce undesirable results.
' The macro should have no effect on any other documents.
'
' Author: Jerry Whittaker   jerry@thewhittakers.org
' Version: 1.5  Date: 7/9/2021 - modified attach code to check existance for currenly attached template and messag
' Version: 1.4  Date: 9/17/2019 - removed code that deleted color text
' Version: 1.3  Date: 9/5/2019 - removed end of macro message
' Version: 1.2  Date: 4/28/2018
' Version: 1.1  Date: 4/24/2018
' Version: 1.0  Date: 4/7/2018
'
    Dim doc As Document
    Dim para As Paragraph
    Dim StyleCntr As Integer
    Dim styleName As String
    Dim NewStyleName As String
    Dim ReloadTemplateSwitch As Boolean
    Set doc = ActiveDocument
    
    ' Convert Abbyy 14 heading styles (e.g. "Heading #1" thru "Heading #9" to normal heading styles
    For Each para In doc.Paragraphs
    
        For StyleCntr = 1 To 9  'up to nine types of Heading styles
            styleName = "Heading #" + LTrim(Str(StyleCntr))
            NewStyleName = "Heading " + LTrim(Str(StyleCntr))
            If StyleCntr > 3 Then
               NewStyleName = "Heading 3"
            End If
            If para.Style = styleName Then
                para.Range.Style = ActiveDocument.Styles(NewStyleName)
                ReloadTemplateSwitch = True
                StyleCntr = 10
            End If
        Next StyleCntr
    
        'Change all Abbyy 14 Body Text styles (e.g. "Body text (1)" through "Body text (9)to Body Text
        'For Each Para In Doc.Paragraphs
        For StyleCntr = 1 To 9  'up to nine types of body text styles
            styleName = "Body text (" + LTrim(Str(StyleCntr)) + ")"
            If para.Style = styleName And Sh_Style_Exists(ActiveDocument, "Body Text") Then
                para.Range.Style = ActiveDocument.Styles("Body Text")
                ReloadTemplateSwitch = True
                StyleCntr = 10
            End If
        Next StyleCntr
   
        ' chang Abbyy's "normal" and "Other" style to Body Text
        If para.Style = "Other" And Sh_Style_Exists(ActiveDocument, "Body Text") Then
            para.Range.Style = ActiveDocument.Styles("Body Text")
            ReloadTemplateSwitch = True
        End If
        
    Next para

    ' Delete all unused (but not built-in)Styles from document when an Abbyy 14 style change was detected
    ' Template is reloaded with the same template that was attached but without the Abbyy 14 styles
    '    showing in the Styles pane.
    ' From: https://word.tips.net/T001337_Removing_Unused_Styles.html - modified to include re-attachment
    ' of the BANA Template
    If ReloadTemplateSwitch = True Then 'An Abbyy 14 style was found and changed
        Dim oStyle As Style
        
        For Each oStyle In ActiveDocument.Styles
            'Only check out non-built-in styles
            If oStyle.BuiltIn = False Then
                With ActiveDocument.Content.Find
                    .ClearFormatting
                    .Style = oStyle.NameLocal
                    .Execute findText:="", Format:=True
                    If .found = False Then oStyle.Delete
                End With
            End If
        Next oStyle
        
        Dim CurrentTemplate As String
        CurrentTemplate = ActiveDocument.AttachedTemplate
        
        Dim TemplatePathandName As String
        TemplatePathandName = Options.DefaultFilePath(wdUserTemplatesPath) + "\" + CurrentTemplate
        
        With ActiveDocument
            If Sh_FileExists(TemplatePathandName) Then
                .UpdateStylesOnOpen = True
                .AttachedTemplate = TemplatePathandName
                .UpdateStylesOnOpen = False  ' supresses any further style updates
            Else
                MsgBox " Cannot continue!" + vbCr + vbCr + "The template file: " + TemplatePathandName + " does not exist." + vbCr + vbCr + "Install the file and try again.", , "Braille Macros"
                End
            End If
        End With
        
    End If
    
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Application.ScreenUpdating = True    ' Turn screen updating on
    Application.ScreenRefresh
    
End Sub   '*** end of Dx_Fix_Abbyy_FineReader_Text_and_Headers ***

Sub Dx_Ref_Pg_Number_Sequence_Menu()
'
' presents menu to locate, tag, validate reference page numbers
'
' Author:   Jerry Whittaker     jerry@thewhittakers.org
'
' Version: 1.0  Date: 10/19/2018

'
    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"
    Sh_Validation_Choices_Form.Show
    Unload Sh_Validation_Choices_Form
    Application.ScreenUpdating = True    ' Turn screen updating on
    Application.ScreenRefresh
    
End Sub   '*** Dx_Ref_Pg_Number_Sequence_Menu macro ***

Sub Dx_File_Fix_Sequence()
'
' Dx_File_Fix_Sequence macro
'
' presents menu to format all or selecte table(s)
'
' Author:   Jerry Whittaker
'           jerry@thewhittakers.org

' Version: 1.0  Date: 6/13/2018
'
    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"
    Unload Dx_File_Cleanup_Sub_Menu_Form
    Dx_File_Cleanup_Sub_Menu_Form.Show
    Unload Dx_File_Cleanup_Sub_Menu_Form
    Application.ScreenUpdating = True    ' Turn screen updating on
    Application.ScreenRefresh
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    
End Sub  '*** end of Dx_File_Fix_Sequence Macro ***

Sub Dx_Red_Border_Images()
    ' From: http://www.vbaexpress.com/forum/showthread.php?48376-VBA-Macro-To-Insert-Border-For-All-Images
    ' Author: gmcnultnult
    '
    ' Version: 1.0  Date: 10/12/2018
    '
    Dim oInlineShp As inlineShape
        For Each oInlineShp In ActiveDocument.InlineShapes
        With oInlineShp
            With .Borders(wdBorderLeft)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth600pt
                .Color = wdColorRed
            End With
                With .Borders(wdBorderRight)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth600pt
                .Color = wdColorRed
            End With
                With .Borders(wdBorderTop)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth600pt
                .Color = wdColorRed
            End With
            With .Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth600pt
                .Color = wdColorRed
            End With
        End With
    Next
End Sub   '*** end of Dx_Red_Border_Images macro ***

Sub Dx_Is_Style_Here()
'
' Checks to see both styles are attached
' style name are sent in Dx_GP_String_1 and Dx_GP_String_2
' status is returned in Dx_GP_String_1
'
' Version: 1.0  Date: 1/17/2019
'

Dim StrSty1
Dim StrSty2

StrSty1 = Dx_GP_String_1
StrSty2 = Dx_GP_String_2

On Error GoTo ErrHandler

    If Not ActiveDocument.Styles(StrSty1) Then
        Exit Sub
    End If
    
    If Not ActiveDocument.Styles(StrSty2) Then
        Exit Sub
    End If
    
ErrHandler:
    If Err.Number = 5941 Then
        Dx_GP_String_1 = "Not Here"
    Else
       Dx_GP_String_1 = "Here"
    End If
    
End Sub   '*** end of Dx_Is_Style_Here ***
Sub Dx_Compress_Linear_Math()
'
'  Macro suggestion by Katherine Thomison and Shelley Mack
'
'  Purpose: Removes all spaces from selected text and places spaces before and after signs of comparison
        '  = equal
        '  approximately equal (double tilda)
        '  <> not equal
        '  Not equal (slashed equal sign)
        '  > greater than
        '  < less than
        '  greater than or equal (underscrored greater than)
        '  less than or equal (underscored less than)
'
'  Author: Jerry Whittaker   jerry@thewhittakers.org
'
'  Version: 1.1  Date: 4/7/2023 - added check for BANA template
'  Version: 1.0  Date: 9/13/2019 - Modification from LP Version
'
    Application.Run MacroName:="Dx_Is_BANA_Template_Attached"
    
    If Selection.Type <> wdSelectionNormal Then
        Selection.Paragraphs(1).Range.Select
    End If
    
    If MsgBox("Compress this math expression?", vbYesNo, "Braille Macros") = vbYes Then
        GoTo CompressThis:
    Else
        Selection.Collapse
        MsgBox "Select only the math to be compressed and run this macro again.", , "Braille Macros"
        End
    End If
    
CompressThis:

    'convert normal "x" (multiply) with math x
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "x"
        .Replacement.Text = "×" 'unicode 00D7
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' set the find and replace parameters
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}"    ' the text to be found - one or more spaces (^032 is code for a space)
                                        ' the {1,} means find one or more - more efficient code
        .Replacement.Text = ""  ' the replacement text is nothing
        .Forward = True
        .Wrap = wdFindStop ' will not ask if you want to search the rest of document
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True  ' is a wildcard search because of the {1,} in the find text
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll  ' do the replacement
    
    Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "=" ' this finds an equal sign
            .Replacement.Text = " ^& "
            .Wrap = wdFindStop
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
    Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "\>" ' this finds the greater than sign
            .Replacement.Text = " ^& "
            .Wrap = wdFindStop
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
    Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "\<" ' this finds the less than sign
            .Replacement.Text = " ^& "
            .Wrap = wdFindStop
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = "\< \>" ' this finds the not-equal sign
            .Replacement.Text = " ^& "
            .Wrap = wdFindStop
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
    Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = ChrW(8805) 'unicode 2265 for underscored greater than sign (greater than or equal to)
            .Replacement.Text = " ^& "
            .Wrap = wdFindStop
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
    Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = ChrW(8804) 'unicode 2264 for underscore less than sign (less than or equal to)
            .Replacement.Text = " ^& "
            .Wrap = wdFindStop
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
    Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = ChrW(8800) 'unicode 2260 for slashed equal sign (not equal to)
            .Replacement.Text = " ^& "
            .Wrap = wdFindStop
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
    Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = ChrW(8776) ' double tilda - approximately  equal to
            .Replacement.Text = " ^& "
            .Wrap = wdFindStop
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        
    Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = " \<  \> "    ' double spaces between <> (fix for problems created above)
            .Replacement.Text = " <> "
            .Wrap = wdFindStop
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    
End Sub   '*** End of Dx_Compress_Linear_Math macro ***

Sub Dx_Convert_Hyperliks_To_Text()
    '
    ' Version: 1.0  Date: 9/27/2021 - complete rewrite of Sh_Show_Hidden_HLink macro
    '
    ' Converts the web-link or email address when they indicated by a link word like "here" or "my email address"
    ' also removes the "mailto:" header in email addresses
    ' Converts internal document hyperlinks to the text they point to
    
    Dim i As Long, rng As Range
    Dim LinkString As String
    
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False    ' Turn screen updating off
    Application.ScreenRefresh
    
    ' convert internal hyperlinks
    With ActiveDocument
        For i = .Hyperlinks.count To 1 Step -1
            LinkString = .Hyperlinks(i).SubAddress
            If LinkString <> "" Then ' it is an internal hyperlink
                .Range.Fields(i).Unlink
            End If
            Next i
    End With
    
    ' convert web links and "here" type email addresses
    With ActiveDocument
        For i = .Hyperlinks.count To 1 Step -1
                .Hyperlinks(i).Range.Text = .Hyperlinks(i).Address
         Next i
    End With
    
    ' remove "mailto:" from email addresses
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "mailto:"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Application.ScreenUpdating = su_Prev    ' Turn screen updating on
    Application.ScreenRefresh

End Sub   '*** end of Dx_Convert_Hyperliks_To_Text macro ***

Sub Dx_Replace_Fraction_Text_With_Compact_Fractions()
'
' Replaces typed fractions with compact fractions. e.g 1/2 to ½ (limited to the 18 compact fractions supported by Word)
'
' Version: 1.2  Date: 3/5/2023 - removed encode fractions with DBT fraction codes
' Version: 1.1  Date: 11/14/2021 - minor bug fix
' Version: 1.0  Date: 10/1/2021

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = "½"
        .Text = "1/2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8531)
        .Text = "1/3"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8532)
        .Text = "2/3"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = "¼"
        .Text = "1/4"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
   
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = "¾"
        .Text = "3/4"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8533)
        .Text = "1/5"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8534)
        .Text = "2/5"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8535)
        .Text = "3/5"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8536)
        .Text = "4/5"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8537)
        .Text = "1/6"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8538)
        .Text = "5/6"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8528)
        .Text = "1/7"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8539)
        .Text = "1/8"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8540)
        .Text = "3/8"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8541)
        .Text = " 5/8"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8542)
        .Text = "7/8"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8529)
        .Text = "1/9"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Replacement.Text = ChrW(8530)
        .Text = "1/10"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' find any DBT code - set to hidden wdColorPlum
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Hidden = False
        .Color = wdColorAutomatic
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Hidden = True
        .Color = wdColorPlum
    End With
    With Selection.Find
        .Text = "\[\[\*(?{1,})\*\]\]"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' replace hidden plum para marks with auto color - not hidden
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Hidden = True
        .Color = wdColorPlum
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Hidden = False
        .Color = wdColorAutomatic
    End With
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    
End Sub   '*** end of Dx_Replace_Fraction_Text_With_Compact_Fractions macro ***

Sub Dx_Replace_Word_NonBreaking_Hyphen_With_Unicode_Non_Breaking_Hypen()
'
' Version: 1.1  Date: 2/13/2024 - find code changed
' Version: 1.0  Date: 10/27/2021
'
Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8209) 'non-breaking hyphen
        .Replacement.Text = "-"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub  '*** end of Dx_Replace_Word_NonBreaking_Hyphen_With_Unicode_Non_Breaking_Hypen ***

' Dx_Replace_NonBreaking_Space_With_Space was here until 8/12/2026. Three macros did this job and only
' Sh_ReplaceNonBreakingSpacesWithNormalSpace protected a "Print Pg Num" bar, so everything
' runs through that one now.


Sub Dx_Replace_Underscore_With_Single_Underscore()
'
' Version: 1.0  Date: 3/4/2023
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "_{1,}"
        .Replacement.Text = "_"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub   '*** End of Dx_Replace_Underscore_With_Single_Underscore macro ***

Sub Dx_Remove_Breaks()
'
' Remove Section Break (next page)
' Remove Secion Break (continuous)
' Remove Section Break (even page)
' Remove Secion Break (odd page)
'
' Version: 1.0  Date: 3/4/2023
'

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^b"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub   '*** End of Dx_Remove_Breaks macro ***

Sub Dx_Replace_Spaces_Before_Punctuation()
'
' Version: 1.0  Date:3/8/2023
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " {1,}."
        .Replacement.Text = "."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " {1,}(\!)"
        .Replacement.Text = "\1"
        .Forward = False
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub   '*** end of Dx_Replace_Spaces_Before_Punctuation macro ***

Sub Dx_Replace_Function_Application_With_Space()
'
' Converts Function Application code (U+2061) into regular space - U+2061 is used in TeX.
'   TeX is a popular means of typesetting complex mathematical formulae
'
' needed for conversion of math into Nemeth code - math created by MathPix or MathKicker.ai
'
' Version: 1.0  Date: 2/9/2024
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8289) 'U+2061 - AKA
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub   '***End of Dx_Replace_Function_Application_With_Space macro ***

Function Dx_Is_The_Attached_Template_BANA_Braille()
'
' Version: 1.1  Date: 9/29/2021 - changed lookup style to "RefPageNumberEmbed"
' Version: 1.0  Date: 4/15/2021
'
' Returns True if the style "RefPageNumberEmbed" is in the attached template
' When true, the attached template is a large print template
'
    Dim oStyle As Style
    Dim styleName As String

    styleName = "RefPageNumberEmbed"
    Set oStyle = Nothing
    
    On Error Resume Next
    Set oStyle = ActiveDocument.Styles(styleName)
    
    If Not oStyle Is Nothing Then ' the style was found
        Dx_Is_The_Attached_Template_BANA_Braille = True
    End If

End Function '*** end of Dx_Is_The_Attached_Template_BANA_Braille Function ***

Sub Dx_Replace_Multiple_Para_Marks_With_Warning()
'
' Version: 1.0   Date:  2/17/2024
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
       
    If Selection.Type <> wdSelectionNormal Then  'there is no selected text

        Dim lngQuery As Long
        lngQuery = MsgBox("CAUTION!" & vbCr _
                      & vbCr & " Because no text is selected, this macro will replace ALL" & vbCr _
                                  & "multiple paragraph marks in the document with a single" & vbCr _
                                  & "paragraph mark." & vbCr _
                       & vbCr & "If you have purposely added extra blank lines to your" & vbCr _
                                  & "document this macro will remove those lines!" & vbCr _
                       & vbCr & "To limit the replacement scope, select only the text" & vbCr _
                                  & "to be changed." & vbCr _
                       & vbCr & "Do you wish to continue?", vbYesNo + vbCritical + vbDefaultButton2, "Braille Macros")
           
        If lngQuery = vbNo Then
            End
        End If
        
    End If
    
    Application.Run MacroName:="Sh_Replace_Multiple_Para_Marks_No_Warning"
    
End Sub    '***   end of  Dx_Replace_Multiple_Para_Marks_With_Warning macro ***

' Dx_Replace_Multiple_Para_Marks_No_Warning was here until 8/12/2026. It was still the original
' "^013{2,} -> ^p" find-and-replace through a temporary document, which cannot see a paragraph
' that LOOKS empty but holds a space - the run of marks is broken by the space, so the pattern
' never matches and the blank line survives (Jerry, 8/12/2026). The large-print side was rewritten
' in July to walk the paragraphs and test each one with Trim$, which does see them. Both sides now
' use that: Sh_Replace_Multiple_Para_Marks_No_Warning.


Sub Dx_Fix_Equals_Before_Para_Mark()
'
' for math problems... cannot end with just an equals character - adds question mark
'
' Version: 1.0  Date: 2/26/2024
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "=^p"
        .Replacement.Text = "=?^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub '**** end of Dx_Fix_Equals_Before_Para_Mark macro ***

Sub Dx_Delete_Square_Bullet()
'
' Version 1.0:  Date: 11/25/2025
'

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(61623)
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub   '*** end of Dx_Delete_Square_Bullet macro ***

Sub Dx_Add_Qmark_To_Incomplete_Equations()
'
' Fixes problem for Math Problems to be used with MathType
'
' Version: 1.0  Date: 4/1/2026
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "=^p"
        .Replacement.Text = "=?^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub   '*** end of Dx_Add_Qmark_To_Incomplete_Equations macro ***


'------------------------------------------------------------------------------------
'/ / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
' ** Top of LP Project macros Version:  **
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
'------------------------------------------------------------------------------------

Sub Lp_Attach_Lp_Template()
    '
    ' Description: Attaches large print template to document
    '
    ' Version: 2.4  Date: 8/12/2026 - records "was this document already large print?" in its own
    '                               Lp_Doc_Was_Already_LP, on EVERY path through this macro.
    '
    '                               It used to be recorded by writing "Doc_Is_Already_LP" into
    '                               Lp_GP_String_1 - and only on the re-attach branch. Nothing
    '                               ever cleared it. So once a transcriber re-attached the
    '                               template to a document that was already large print, that
    '                               string stayed set for the rest of the Word session, and the
    '                               NEXT attach - to a raw scanned file - silently skipped both
    '                               blocks it guards in Lp_Attach_The_Template: the whole of
    '                               Lp_Fix_Common_File_Errors, the table color pass, and the
    '                               empty-paragraph removal. That is exactly what Jerry saw:
    '                               empty paragraph marks surviving the attach, then coming out
    '                               the moment Full File Cleanup was run by hand.
    '
    '                               Lp_GP_String_1 is also written by Lp_Columns_Wanted_Form and
    '                               Lp_TOC_Format_And_Color_Form, so the flag could equally flip
    '                               back the other way. It is not a place to keep a decision.
    ' Version: 2.3  Date: 8/8/2026 - removed the inline Tahoma-installed check. The transcriber
    '                               now chooses the typeface on the attach dialog, which grays out
    '                               one that is not installed - a better answer than a warning, and
    '                               it takes an End statement out of this path that wiped every
    '                               Public on the way past
    ' Version: 2.2  Date: 8/30/2021 - added check for Tahoma Font
    ' Version: 2.1  Date: 3/3/2020 - added check and fix foreign file
    ' Version: 2.0  Date: 1/31/2020 - added warning message when LP template is already attached
    ' Version: 1.9  Date: 6/12/2019 - added run of 'Lp_Check_Compatibility' to check for doc extention or compatibility mode
    ' Version: 1.8  Date: 2/1/2018
    '
    '------------------------------------------------------------------------------------
    ' If no document is active then create a new blank document
    '------------------------------------------------------------------------------------
    If Documents.count = 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument ' will trigger AutoNew
    End If
    
    Application.Run MacroName:="Lp_Check_Compatibility"  'check to see if doc is .docx or .doc
    
    ' The check for a missing Tahoma used to sit here, and there is now NOTHING in its place.
    ' That is deliberate. It moved to the attach dialog on 8/8/2026, when the transcriber briefly
    ' chose the typeface and the dialog grayed out a face that was not installed; the typeface was
    ' dropped on 8/20/2026 and that check went with the choice.
    '
    ' Nothing replaced it because there is nothing left to warn about. Every Windows machine has
    ' Tahoma - a missing one means a broken Windows installation, not a VistaType problem - and
    ' Sh_Doc_Info reports the document's face and whether it is present, on demand, which is the
    ' right place for an answer nobody needs before they have asked.
    '
    ' Its "End" statement went with it, and good riddance: End wipes every Public, so a
    ' transcriber who answered No there would have cleared Lp_Base_Font_Size and every page
    ' setting on the way out.

    '------------------------------------------------------------------------------------
    ' Housekeeping
    '------------------------------------------------------------------------------------
    Application.ScreenUpdating = False ' Turn screen updating off
    Selection.HomeKey Unit:=wdStory  'move cursor to top of document
    
    '------------------------------------------------------------------------------------
    ' Set the style area view width
    '------------------------------------------------------------------------------------
    With ActiveDocument.ActiveWindow
        .View.Type = wdPrintView
        .StyleAreaWidth = InchesToPoints(1)
    End With
    '------------------------------------------------------------------------------------
    ' Give warning if already and LP Template
    '------------------------------------------------------------------------------------
    ' Answer this fresh every time. Both blocks in Lp_Attach_The_Template that are meant to run
    ' only on a document which is NOT already large print read this, and a stale True there is
    ' silent - the attach just does less work than it says it does.
    ' The LOOSER test, deliberately - 8/20/2026. What this decides is whether she is warned that
    ' re-attaching will re-flow her text, and that is true of a book on an OBSOLETE large print
    ' template just as much as one on the current template. Asking Lp_Is_The_Attached_Template_LP
    ' here would send exactly the book that most needs the warning through without one.
    Lp_Doc_Was_Already_LP = (Lp_Was_Made_As_An_Lp_Book() = True And _
                             InStr(UCase(ActiveDocument.AttachedTemplate.Name), UCase("Normal.do")) = 0)

    If Lp_Doc_Was_Already_LP Then
        Lp_ReAttachWarning_Form.Show
        GoTo eom  ' by-pass the "Lp_Attach_An_Lp_Template_Form.show" - which will be called from "Lp_ReAttachWarning_Form.show"
    End If

    '------------------------------------------------------------------------------------
    ' Invoke the form called "Lp_Attach_Lp_Template_Form"
    '------------------------------------------------------------------------------------

    Lp_Attach_An_Lp_Template_Form.Show
    
eom:
        
End Sub      '   *********** end of Lp_Attach_Lp_Template Macro ******************************

Sub Lp_Remove_Box_Bullets_Bullets_and_Numbers()
'
' Lp_Remove_Box_Bullets_Bullets_and_Numbers Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Turns automatic numbering and bullets into nothing, removes the box bullet and the tab behind
' it, drops hyperlinks, underlining and colored text.
'
' Version: 2.0  Date: 8/13/2026 - no temporary document; everything is scoped to a range.
'
'                                The scratch document was doing real work in three places.
'                                ActiveDocument.ListParagraphs is a whole-document collection;
'                                Selection.WholeStory before RemoveNumbers means what it says;
'                                and Sh_Remove_Hyperlinks with no argument strips every link in
'                                the document. All three are held to the range now.
'
'                                Gone with the round trip: the "delete one character at the end
'                                of the document" and the TypeParagraph/TypeBackspace pair that
'                                tidied up after the paste.
' Version: 1.3  Date: 11/18/2021 - fixed bug which removed first char of selection
' Version: 1.2  Date: 4/21/2021 - added new remove underline and blue color of links
' Version: 1.1  Date: 1/11/2019
' Version: 1.0  Date: 2/13/2015
'
    Dim doc As Document
    Dim rng As Range
    Dim i As Long
    Dim su_Prev As Boolean

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    Set doc = ActiveDocument
    If Selection.Type = wdSelectionNormal Then
        Set rng = Selection.Range
    Else
        Set rng = doc.Content
    End If

    ' Backwards: converting a list paragraph to text takes it out of ListParagraphs, and a
    ' forward For Each over a collection that is shrinking under it misses every other one.
    For i = rng.ListParagraphs.count To 1 Step -1
        rng.ListParagraphs(i).Range.ListFormat.ConvertNumbersToText
    Next i

    Sh_Delete_Text_In_Range rng, "^t"

    Sh_Remove_Hyperlinks rng

    rng.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph   ' bullets too

    Sh_Delete_Text_In_Range rng, ChrW(61623) & " "
    Sh_Delete_Text_In_Range rng, ChrW(61623)

    ' underlining off
    With rng.Find
        .ClearFormatting
        .Font.Underline = wdUnderlineSingle
        .Replacement.ClearFormatting
        .Replacement.Font.Underline = wdUnderlineNone
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' every character back to automatic color
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Replacement.Font.Color = wdColorAutomatic
        .Text = "^?"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' Lp_Fix_Para_Space_Errors scopes itself to the selection, which is still the transcriber's
    ' own - nothing above took it away.
    Application.Run MacroName:="Lp_Fix_Para_Space_Errors"

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub '****** End of Lp_Remove_Box_Bullets_Bullets_and_Numbers Macro *****

Sub Lp_Fix_Common_File_Errors()
' Version: 2.1  Date: 8/13/2026 - now removes empty paragraphs too, as its last repair.
'                               Jerry, 8/13/2026. Note for the File Cleanup form: checking
'                               "fix common file errors" alone now takes the blank lines out,
'                               without the caution the "remove para marks" box shows first.
'
' Lp_Fix_Common_File_Errors
'
' Version: 3.12  Date: 8/3/2026 - its 28 DoEvents are now Sh_Spin_DoEvents, so the please-wait spinner turns with each one; OnTime alone fires only once a second and the spinner looked stuck
' Version: 3.11  Date: 8/2/2026 - runs Sh_Color_Dollar_PG_Red as the LAST step, so the $pg tags are still red when File Cleanup is run on its own rather than as part of the attach sequence; the cleanups above lose the color and no Lp_Normalize_Styles follows to restore it
' Version: 3.10  Date: 7/26/2026 - returns the user to where the cursor was when the macro started
' Version: 3.8  Date: 3/2/2026 - added Sh_ReplaceNonBreakingSpacesWithNormalSpace
' Version: 3.7  Date: 11/25/2025 - added Lp_Fix_Para_Space_Errors
' Version: 3.6  Date: 10/16/2025 - commented out time-consuming routines
' Version: 3.4  Date: 6/18/2025 - added delete zero width spaces
' Version: 3.3  Date: 4/21/2025 - removed stats bar updates
' Version: 3.2  Date: 4/21/2025 - added compleded message bypass
' Version: 3.0  Date: 3/17/2024 - added "Lp_Convert_Ordinal_Numbers"
' Version: 2.9  Date: 10/22/2023 - added status bar updates and fixed Drop Caps for speed
' Version: 2.8  Date: 10/9/2023 - added Selection.Collapse to fix bug where char following cursor was deleted when text was selected
' Version: 2.7  Date: 10/28/2021 - added "Lp_Replace_Strong_With_Bold"
' Version: 2.6  Date: 10/19/2021 - Added "Lp_Replace_Compact_Fractions_With_Fraction_Text"
' Version: 2.5  Date: 5/31/2019 - added "ActiveDocument.SetCompatibilityMode (wdWord2013)" ' turns off compatability mode - makes table boarders stay within margin
' Version: 2.4  Date: 5/14/2019 - Added Lp_Add_Para_After_Image
' Version: 2.3  Date: 1/29/2019  - removed code that killed image placement
' Version: 2.2  Date: 1/8/2019
' Version: 2.1  Date: 12/27/2018
' Version: 2.0  Date: 10/12/2018
' Version: 1.9  Date: 2/14/2018

    ' Description:  Fixes common errors in entire  file
    '
    '   Converts Abbyy FineReader styles to Word styles
    '   Removes spaces before punctuation
    '   Removes drop caps
    '   Removes all spaces before and after paragraph marks
    '   Replaces multiple spaces with a single space
    '   Replaces multiple hyphens with a single hyphen
    '   Removes spaces before and after hyphens
    '   Replaces multiple "en" dashes with a single "en" dash
    '   Removes spaces before and after "en" dashes
    '   Replaces multiple "em" dashes with a single "em" dash
    '   Removes spaces before and after "em" dashes
    '   Removes optional hyphens
    '   Changes italics to dashed underline
    '   Replaces variations of normal styles with with Word normal style
    '   Places a paragraph mark before each $pg and turns them red
    '   Remove tabs before and after para marks
    '   makes sure all images are NOT followed immediatly by a line of text
    '   Converts superscript ordinals to normal style size
    '   Delete zero width spaces - often place by AI
    '   Resize pictures in tables to comfortably fit with the cell
    
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating off

    ' Save the spot BEFORE anything runs. The old CleanupBookmark was added further down,
    ' after Selection.Collapse and Sh_Color_Dollar_PG_Red had already moved the cursor.
    Sh_Save_User_Position

    Selection.Collapse 'clear selection
    Application.Run MacroName:="Sh_Color_Dollar_PG_Red"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Replace_Underline_Tab_With_Underlined_Underscore"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Delete_Square_Bullet"  ' run before Lp_Fix_Para_Space_Errors
Sh_Spin_DoEvents
    Application.Run MacroName:="Sh_ReplaceNonBreakingSpacesWithNormalSpace"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Add_Para_After_Image"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Fix_Abbyy_Text_and_Headers"
Sh_Spin_DoEvents
    Application.Run MacroName:="Sh_Remove_Spaces_Before_Punctuation"
Sh_Spin_DoEvents
    Application.Run MacroName:="Sh_Remove_Txt_Bxs_And_Frames"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Remove_Tab_Plus_Space_Combos"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Fix_Para_Space_Errors"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Italics_To_Dashed_Underline"
Sh_Spin_DoEvents
    Application.Run MacroName:="Sh_Replace_White_Text_With_Automatic"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Fix_EnDash_Errors"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Fix_Em_Dash_Space_Errors"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Fix_Normal_Styles"  'no longer ruins picture placement (left, right, center)
           'takes too long on large docs
Sh_Spin_DoEvents
    Application.Run MacroName:="Sh_Remove_Multi_Spaces"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Fix_Hyphen_Errors"
Sh_Spin_DoEvents
    Application.Run MacroName:="Sh_Para_Before_Dollar" 'Fixes DAISY Page Problems
Sh_Spin_DoEvents
    Application.Run MacroName:="Sh_Replace_Small_Caps_With_All_Caps"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Remove_Tabs_Before_and_After_Para_Marks"
Sh_Spin_DoEvents
    'Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Application.Run MacroName:="Lp_Convert_Hyperliks_To_Text"  'convert all (including hidden links) links to text (except internal links)
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Convert_Hyper_To_Addresses" 'convert all text links to active links
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Replace_Compact_Fractions_With_Fraction_Text"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Replace_Strong_With_Bold"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_RemoveHeadAndFoot"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Convert_Ordinal_Numbers"
Sh_Spin_DoEvents
    Application.Run MacroName:="Lp_Delete_Zero_Width_Spaces"
Sh_Spin_DoEvents

    ' Empty paragraphs go LAST of the repairs, and after Lp_Delete_Zero_Width_Spaces in
    ' particular: a paragraph holding nothing but a zero-width space is not empty until that
    ' pass has run. Nothing above this line creates a blank paragraph that is meant to stay --
    ' Lp_Add_Para_After_Image splits an image away from the text stuck to it, and
    ' Sh_Para_Before_Dollar already reduces its own runs to a single mark. Added 8/13/2026 at
    ' Jerry's request; until then the only callers were the attach and the cleanup form, each
    ' running it as a step of their own. See the note in that macro about what "multiple
    ' paragraph marks replaced by a single paragraph mark" means -- every blank line goes.
    Application.Run MacroName:="Sh_Replace_Multiple_Para_Marks_No_Warning"
Sh_Spin_DoEvents

    ' Color the $pg tags red LAST, so they are still red when this macro is run ON ITS OWN.
    ' File Cleanup is not only a step inside the attach sequence -- a transcriber can run it by
    ' itself, with no Lp_Normalize_Styles afterwards to restore the color. The cleanups above
    ' lose the original red, and nothing else would put it back. Jerry, 8/2/2026.
    Application.Run MacroName:="Sh_Color_Dollar_PG_Red"
Sh_Spin_DoEvents

    ActiveDocument.UndoClear

    Application.ScreenUpdating = su_Prev    ' Turn screen updating on - was commented out, so the
                                            ' screen stayed frozen and never followed the cursor back

    Sh_Return_User_To_Start_Position

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"

End Sub '*** end of Lp_Fix_Common_File_Errors macro ***

Sub Lp_Convert_Hyperliks_To_Text()
    '
    ' Version: 1.1  Date: 1/18/2003 - set On Errors for bug fix
    ' Version: 1.0  Date: 9/27/2021 - complete rewrite of Sh_Show_Hidden_HLink macro
    '
    ' Converts the web-link or email address when they indicated by a link word like "here" or "my email address"
    ' also removes the "mailto:" header in email addresses
    ' Converts internal document hyperlinks to the text they point to
    
    Dim i As Long, rng As Range
    Dim LinkString As String
    
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False    ' Turn screen updating off
    ' convert internal hyperlinks - not used for large print
    On Error Resume Next
    With ActiveDocument
        For i = .Hyperlinks.count To 1 Step -1
            LinkString = .Hyperlinks(i).SubAddress
            If LinkString <> "" Then ' it is an internal hyperlink
                .Range.Fields(i).Unlink
           End If
           Next i
    End With
    
    ' convert web links and "here" type email addresses
    With ActiveDocument
        For i = .Hyperlinks.count To 1 Step -1
                .Hyperlinks(i).Range.Text = .Hyperlinks(i).Address
         Next i
    End With
    
    ' remove "mailto:" from email addresses
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "mailto:"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Application.ScreenUpdating = su_Prev    ' Turn screen updating on
    Application.ScreenRefresh

End Sub   '*** end of Lp_Convert_Hyperliks_To_Text macro ***

Sub Lp_Italics_To_Dashed_Underline()
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Version: 1.2  Date: 2/12/2024 - removed dashed underlines from spaces and punctuation marks -
'                                 removed ".MatchWholeWord = True" replaces with ".MatchWholeWord = False"
' Version: 1.1  Date: 1/30/2024 - converted all forms (Bold, Underlined) of italics with dashed underlines - words only
' Version: 1.0  Date: 5/23/2016
'
' Description:  changes italics to dashed underline
'               Works on whole file - converts Italics with bold and/or underline


    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Bold = True
        .Italic = True
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = True
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Italic = True
        .Underline = wdUnderlineSingle
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Italic = True
        .Underline = wdUnderlineNone
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Font.Italic = True
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Italic = False
        .Underline = wdUnderlineDash
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Italic = True
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Italic = False
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
'*** remove dash inderlines from punctuation and spaces

    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineDash
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .Text = " "
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineDash
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .Text = "."
        .Replacement.Text = "."
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineDash
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .Text = "!"
        .Replacement.Text = "!"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineDash
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .Text = "?"
        .Replacement.Text = "?"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineDash
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .Text = ","
        .Replacement.Text = ","
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineDash
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .Text = ":"
        .Replacement.Text = ":"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineDash
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .Text = ";"
        .Replacement.Text = ";"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineDash
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .Text = "¿"
        .Replacement.Text = "¿"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineDash
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .Text = "¡"
        .Replacement.Text = "¡"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub '  ***** End of Lp_Italics_To_Dashed_Underline Macro *****************

Sub Lp_Fix_Hyphen_Errors()
'
' Lp_Fix_Hyphen_Errors Macro
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Date: 5/31/2016
'
'------------------------------------------------------------------------------------
'  remove optional hyphens
'------------------------------------------------------------------------------------

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^-{1,}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

'------------------------------------------------------------------------------------
'  replace multiple hyphens with single hyphen
'------------------------------------------------------------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^045{1,}"
        .Replacement.Text = "^045"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

'------------------------------------------------------------------------------------
'  remove spaces before hyphens
'------------------------------------------------------------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}^045"
        .Replacement.Text = "^045"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

'------------------------------------------------------------------------------------
'  remove spaces after hyphens
'------------------------------------------------------------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^045^032{1,}"
        .Replacement.Text = "^045"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub  '***** End of Lp_Fix_Hyphen_Errors Macro **************************

Sub Lp_Fix_EnDash_Errors()
'
' Lp_Fix_EnDash_Errors Macro - for large print file clean-up
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Date: 5/31/2016

'------------------------------------------------------------------------------------
'  replace multiple en dashes with single en dash
'------------------------------------------------------------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^0150{1,}"
        .Replacement.Text = "^0150"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

'------------------------------------------------------------------------------------
'  remove spaces before en dashes
'------------------------------------------------------------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}^0150"
        .Replacement.Text = "^0150"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

'------------------------------------------------------------------------------------
'  remove spaces after en dashes
'------------------------------------------------------------------------------------
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^0150^032{1,}"
        .Replacement.Text = "^0150"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub  '******* End of Lp_Fix_EnDash_Errors Macro  ****************

Sub Lp_Fix_Normal_Styles()

' Lp_Fix_Normal_Styles Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version: 2.1  Date: 2/12/2024 - added many styles from "Normal" template
' Version: 2.0  Date: 8/27/2021 - added "Normal Indent" and "Normal Indent 2" to Normal
' version: 1.9  Date:  6/4/2021 - added remove left indent from normal style para
' Version: 1.8  Date:  2/15/2021 - added replace normal with normal
' Version: 1.7  Date: 11/24/2020 - removed Lp_Base_Font_Size and replaces with Normal_Style_Font_Size
' Version: 1.6  Date: 2/20/2020 - added F&R to fix foreing "Normal" text font sized to Lp "Normal" font size
' Version: 1.5  Date: 3/26/2019
' Version: 1.4  Date: 1/29/2019
' Version: 1.3  Date: 1/13/2019
' Version: 1.2  Date: 4/27/2018
' Version: 1.1  Date: 2/22/2016
'
'
'------------------------------------------------------------------------------------
' Replace styles "Body Text", "Normal (Web)", "Normal Indent",
'                "HTML Address", "Plain Text" with style "Normal"
'------------------------------------------------------------------------------------
'


    Dim doc As Document
    Dim para As Paragraph
    Dim Normal_Style_Font_Size As Integer
    Set doc = ActiveDocument
 
    ' remove left para indent in normal styles - leaves auto lists indents intact
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Normal")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .FirstLineIndent = InchesToPoints(0)
        .CharacterUnitFirstLineIndent = 0
        .MirrorIndents = False
        .CollapsedByDefault = False
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Body Text") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Body Text")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If

    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Body Text 2") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Body Text 2")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Body Text 3") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Body Text 3")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Body Text First Indent") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Body Text First Indent")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Body Text First Indent 2") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Body Text First Indent 2")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Body Text Indent") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Body Text Indent")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Body Text Indent 2") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Body Text Indent 2")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Body Text Indent 3") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Body Text Indent 3")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Closing") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Closing")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Comment Reference") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Comment Reference")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Comment Subject") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Comment Subject")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Comment Text") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Comment Text")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Date") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Date")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Document Map") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Document Map")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "E-mail Signature") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("E-mail Signature")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Endnote Reference") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Endnote Reference")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find.Replacement.Font
        .Superscript = False
    End With
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Endnote Text") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Endnote Text")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Normal (Web)") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Normal (Web)")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Normal Indent") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Normal Indent")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Note Heading") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Note Heading")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' get size of base font
    On Error GoTo Next2
    'Application.Run MacroName:="Lp_Get_Doc_Setup_Params"  ' Places size of normal style font in public variable "Lp_Base_Font_Size"
    Normal_Style_Font_Size = ActiveDocument.Styles(wdStyleNormal).Font.Size
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Normal")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = Normal_Style_Font_Size
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
Next2:
On Error GoTo 0

Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"

End Sub  '******* Lp_Fix_Normal_Styles Macro ************************

Sub Lp_Is_Text_Selected()
'
' Lp_Is_Text_Selected Macro

' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version 1.0
' Date: 2/13/2015
'
'--------------------------------------------------------------------------------
' Checks if text is selected
'--------------------------------------------------------------------------------

    If Selection.Type <> wdSelectionNormal Then
        MsgBox "Text must be selected first!", , "VistaType LP (125)"
        End
    End If
    
End Sub '***** End of Lp_Is_Text_Selected *************

Sub Lp_Set_Page_To_Black()
'
' Lp_Set_Page_To_Black Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version: 1.2  Date - 12/3/2020 - removed old code used to determine black/white status
' Version: 1.1  Date: 9/26/2015
'
    
    ActiveDocument.Background.Fill.ForeColor.ObjectThemeColor = wdThemeColorText1
    'ActiveDocument.Background.Fill.ForeColor.RGB = RGB(0, 0, 0)  'alternate form for next code line
    ActiveDocument.Background.Fill.ForeColor.TintAndShade = 0#
    ActiveDocument.Background.Fill.Visible = msoTrue
    ActiveDocument.Background.Fill.Solid
    ActiveDocument.ActiveWindow.View.DisplayBackgrounds = True
    
 '-----------------------------------------------------------------------------------
 ' change any black boxes to white
 '-----------------------------------------------------------------------------------
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Box Black") And Sh_Style_Exists(ActiveDocument, "Box White") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Box Black")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Box White")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
 '-----------------------------------------------------------------------------------
 ' change any tables to yellow on black paper
 '-----------------------------------------------------------------------------------
    
    Dim t As Table
        For Each t In ActiveDocument.Tables
         t.Style = "Yellow on Black Screen Table"
    Next

 '-----------------------------------------------------------------------------------
 ' change Black Para to white para on black paper
 '-----------------------------------------------------------------------------------
 
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Para Black") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Para Black")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles _
        ("Para Black Inverted")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
 '-----------------------------------------------------------------------------------
 ' change Black words to white words on black paper
 '-----------------------------------------------------------------------------------

    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Words Black") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Words Black")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles _
        ("Words Black Inverted")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If

 '-----------------------------------------------------------------------------------
 ' cleanup
 '-----------------------------------------------------------------------------------

    Selection.Collapse 'clear selection
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    ActiveDocument.UndoClear
    
End Sub  '***** Lp_Set_Page_To_Black Macro *********

Sub Lp_Set_Page_To_White()
'
' Lp_Set_Page_To_White Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version: 1.3  Date: 1/23/2024 - added changes to black words and paragraphs
' Version: 1.2  Date: 12/3/2020 - removed old code to check on current black/white setting
' Version: 1.1  Date: 9/26/2015
'
'-----------------------------------------------------------------------------------
' set page color to white
'-----------------------------------------------------------------------------------

    ActiveDocument.Background.Fill.ForeColor.ObjectThemeColor = wdThemeColorBackground1
    ActiveDocument.Background.Fill.ForeColor.TintAndShade = 0#
    ActiveDocument.Background.Fill.Visible = msoTrue
    ActiveDocument.Background.Fill.Solid
   
'-----------------------------------------------------------------------------------
' change white boxes to black boxes
'-----------------------------------------------------------------------------------
   
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Box White") And Sh_Style_Exists(ActiveDocument, "Box Black") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Box White")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Box Black")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
'-----------------------------------------------------------------------------------
' change white tables to yellow for white pages
'-----------------------------------------------------------------------------------

    Dim t As Table
        For Each t In ActiveDocument.Tables
         t.Style = "Yellow on White Paper Table"
    Next
    
'-----------------------------------------------------------------------------------
' change previously black bkgrnd para to white
'-----------------------------------------------------------------------------------
    
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Para Black Inverted") And Sh_Style_Exists(ActiveDocument, "Para Black") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Para Black Inverted")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Para Black")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If

'-----------------------------------------------------------------------------------
' change previously black words to white
'-----------------------------------------------------------------------------------

    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Words Black Inverted") And Sh_Style_Exists(ActiveDocument, "Words Black") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Words Black Inverted")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Words Black")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If

'-----------------------------------------------------------------------------------
' cleanup
'-----------------------------------------------------------------------------------

    Selection.Collapse 'clear selection
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    ActiveDocument.UndoClear
    
End Sub  ' ***** End of Lp_Set_Page_To_White Macro *********

Sub Lp_Convert_Auto_List_To_Text(Optional ByVal target As Range)
    '
    ' Original title "AutoListOff2"
    ' From: Computer Tools for Editors(and Proofreaders)by Paul Beverley, LCGI
    '        http://www.archivepub.co.uk/book.html
    ' Modified by Jerry Whittaker - jerry@thewhittakers.org
    '
    ' Changes auto-bulleted, auto-numbered and auto-outline listing to plain text, takes the tab
    ' out from behind the marker, closes up the space below list items, moves the list to the
    ' left margin and squeezes multiple spaces.
    '
    ' target  the range to work on. Omit it and the selection is used, or the whole document when
    '         there is no selection.
    '
    ' Version: 2.0  Date: 8/13/2026 - no temporary document; takes a range. See the note in the
    '                                braille copy - ConvertNumbersToText and every Find here were
    '                                whole-document, so the scratch document was the scope.
    '
    '                                NOT merged with the braille copy: this one also closes the
    '                                space below list items, outdents three times and squeezes
    '                                multiple spaces, none of which a braille transcriber asked
    '                                for. Different work, not different plumbing.
    '
    '                                Gone with the round trip: "delete one character at the end of
    '                                the document", which ran whether or not there had been a
    '                                temporary document to tidy up after.
    ' Version: 1.5  Date: 9/30/2023 - added convert automatic numbers, bullets, and multi-level
    '                                lists to plain text - convert tabs to spaces - remove space
    '                                lines below list items - Move list to left margin - remove
    '                                multiple spaces from list items
    ' Version: 1.4  Date: 1/15/2019
    ' Version: 1.3  Date: 1/8/2019
    ' Version: 1.2  Date: 1/8/2016
    '
    Dim rng As Range
    Dim NewCharacter As String
    Dim NormalFont As String
    Dim su_Prev As Boolean

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    If target Is Nothing Then
        If Selection.Type = wdSelectionNormal Then
            Set rng = Selection.Range
        Else
            Set rng = ActiveDocument.Content
        End If
    Else
        Set rng = target
    End If
    NewCharacter = ChrW(8226): ' a bullet
    NormalFont = ActiveDocument.Styles(wdStyleNormal).Font.Name

    rng.ListFormat.ConvertNumbersToText  ' automatic numbers, bullets and multi-level lists

    ' One common type of bullet uses Symbol font
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .Text = ChrW(&HF0B7) & "^t"
        .Forward = True
        .Font.Name = "Symbol"
        .Replacement.Text = NewCharacter & " "
        .Replacement.Font.Name = NormalFont
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With

    ' The other type of bullet uses Wingding font
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .Text = ChrW(&HF0FC) & "^t"
        .Font.Name = "Wingding"
        .Replacement.Text = NewCharacter & " "
        .Replacement.Font.Name = NormalFont
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With

    Sh_Marker_Tab_To_Space rng, "(^0149)^009", NormalFont              ' bullet
    Sh_Marker_Tab_To_Space rng, "(^013[A-Za-z]{1,}^046)^009", NormalFont   ' a.
    Sh_Marker_Tab_To_Space rng, "([A-Za-z]{1,}\))^009", NormalFont     ' a)
    Sh_Marker_Tab_To_Space rng, "([0-9]{1,}\))^009", NormalFont        ' 1)
    Sh_Marker_Tab_To_Space rng, "([0-9]{1,}^046)^009", NormalFont      ' 1.

    ' Close the space below the list items and bring them to the left margin.
    ' Lp_Toggle_Space_After_Current_Para reads the selection, so the range is selected for it.
    ' That costs nothing on screen - the flashing came from ACTIVATING another document, not
    ' from moving the selection inside this one.
    rng.Select
    Application.Run MacroName:="Lp_Toggle_Space_After_Current_Para"
    rng.Paragraphs.Outdent
    rng.Paragraphs.Outdent
    rng.Paragraphs.Outdent

    ' Squeeze multiple spaces
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = True
        .Text = "^032{1,}"
        .Replacement.Text = "^032"
        .Replacement.Font.Name = NormalFont
        .Forward = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub  '**** end of Lp_Convert_Auto_List_To_Text Macro ***********

Sub Lp_About()
'
' Lp_About Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version: 1.3  Date 12/10/2019 - added code for alternative short URL in the user form code
' Version: 1.2  Date: 3/16/2016
'
' These Strings are changed with find and replace
    'With Lp_About_Title_And_Agreement
        '.Version.Caption = "This is Large Print Version Beta .9.8  February 1, 2018"
        '.Version.ControlTipText = "This is Large Print Version Beta .9.8  February 1, 2018"
        '.Version.TextAlign = fmTextAlignCenter
    'End With
    Lp_About_Title_And_Agreement.Show
    Unload Lp_About_Title_And_Agreement
    
End Sub  '***** end of Lp_About Macro ****
Sub Lp_Video_Links()
'
' Version: 1.0  Date: 3/3/2021
'
    Lp_Video_Download_Link_Page.Show
    Unload Lp_Video_Download_Link_Page
    
End Sub  '***** end of Lp_Video_LinksMacro ****

Sub Lp_Format_Page_Numbers()
'
' Formats tagged page numbers
'
' Version: 2.3  Date: 8/2/2026 - merges back-to-back $pg tags into one hyphenated tag before the bar is built (Lp_Merge_Adjacent_Pg_Tags)
' Version: 2.2  Date: 3/5/2026 - forced all to base font size
' Version: 2.1  Date: 3/27/2024 - added home key before starting to make sure cursor is a the begining of a line for a newly type $pg
' version: 2.0  Date: 11/7/2023 - added tag count
' Version: 1.9  Date: 5/1/2023 - added call to Sh_Remove_Empty_Para_Before_Tables
' Version: 1.8  Date: 2/3/2018
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
    Application.Run MacroName:="Lp_Is_Lp_Template_Attached"

    Sh_Save_User_Position

    'Count the tags
    Dim TagCounter As Integer
    TagCounter = 0
    Dim par As Word.Paragraph
    Dim txt As String

    For Each par In ActiveDocument.Paragraphs
        txt = par.Range.Text
        If InStr(txt, "$pg") > 0 Then
            TagCounter = TagCounter + 1
        End If
    Next
    
    If TagCounter = 0 Then
        MsgBox "There are no tagged page numbers in this document", , "VistaType LP (126)"
        End
    End If
    
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating off
    
    If ActiveDocument.Bookmarks.Exists("TempPgNoFormat") = True Then
        ActiveDocument.Bookmarks("TempPgNoFormat").Delete
    End If
    ' create bookmark at cursor
    ActiveDocument.Bookmarks.Add Name:="TempPgNoFormat"

    Lp_Get_Doc_Setup_Params 'need the base font size

    Selection.HomeKey Unit:=wdLine
    
    'remove tag when followed by only spaces followed by para mark
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "$pg^032{1,}^013"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' remove tag when followed only by para mark
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "$pg^013"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Two reference page numbers back to back become ONE hyphenated tag, so a blank print page
    ' does not stack two pink bars on top of each other. Position matters in both directions:
    ' AFTER the two passes above, because deleting an empty "$pg" takes its paragraph mark with
    ' it and is often what brings two real tags together, and BEFORE the wildcard below,
    ' because once the bar is built there is no tag left to merge.
    Lp_Merge_Adjacent_Pg_Tags ActiveDocument

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    With Selection.Find.Replacement.Font
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
        .Color = wdColorAutomatic
        .Size = Val(Lp_Base_Font_Size)
    End With
    With Selection.Find
        .Text = "($pg)(*)(^013)"
        ' ^s, a NON-BREAKING SPACE, at each end. Briefly an en space, ChrW(8194), on 8/4/2026,
        ' to keep Sh_ReplaceNonBreakingSpacesWithNormalSpace from flattening the bar's ends.
        ' Reverted the next day: the en space worked - a test document confirmed 8194 at both
        ' ends - but it did not fix what prompted it. That macro now skips any paragraph
        ' styled Print Pg Num instead, which protects the bar without changing what it is
        ' made of. If the character here ever changes, change the ^s the color pass below
        ' searches for to match, or the ends silently stay red.
        ' Separately: the trailing space falls outside the pink shading, and always did. That
        ' is the single right-aligned tab stop driving pn<number> to the margin, so anything
        ' after it overruns the stop. It has nothing to do with which space character is used.
        .Replacement.Text = "\1^s\2^009pn\2^s\3"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    'put page numbers in styled bar
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Print Pg Num") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Print Pg Num")
    With Selection.Find
        .Text = "$pg"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    'Remove $pg
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "$pg"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' Change the non-breaking spaces at each end of the bar to automatic color - they inherit
    ' the red from the $pg tag they replaced. This pass MUST always look for the same character
    ' the bar is built from above, or it silently finds nothing and the ends stay red.
    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = wdColorRed
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorAutomatic
    With Selection.Find
        .Text = "^s"
        .Replacement.Text = "^s"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'make pn pink
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Print Pg Num") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Print Pg Num")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = 13935604
    With Selection.Find
        .Text = "pn"
        .Replacement.Text = "pn"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    Application.Run MacroName:="Sh_Remove_Empty_Para_Before_Tables"
    
    If ActiveDocument.Bookmarks.Exists("TempPgNoFormat") = True Then
        ActiveDocument.Bookmarks("TempPgNoFormat").Select
        ActiveDocument.Bookmarks("TempPgNoFormat").Delete
    End If
        
    ActiveDocument.UndoClear

    Application.ScreenUpdating = su_Prev ' Turn screen updating on
    
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    
    ActiveWindow.DocumentMap = False

    Sh_Return_User_To_Start_Position

    MsgBox "Reference page formatting complete", , "VistaType LP (127)"

End Sub   '****end of Lp_Format_Page_Numbers Macro ***********

Function Lp_Merge_Adjacent_Pg_Tags(ByVal targetDoc As Document) As Long
'
' Combines a RUN of back-to-back "$pg" tag paragraphs into ONE tag, hyphenated first to last.
'
'   $pg12 / $pg13            ->  $pg12-13
'   $pg12 / $pg13 / $pg14    ->  $pg12-14
'
' Real print books contain blank pages, and every page still has to carry a page number, so
' DAISY and NIMAS coders emit two numbers back to back -- almost always at a chapter change.
' Left alone, Lp_Format_Page_Numbers builds a separate pink bar for each and the two stack up
' into a visual mess. The $pg validation never showed this to the user: it is an entirely
' visual review with no adjacency or sequence checking anywhere in it.
'
' LARGE PRINT ONLY. The braille twin Dx_Format_Tagged_Page_Numbers builds its bar from the
' RefPageNumber / RefPageNemeth styles and is deliberately untouched.
'
' ANY two adjacent tags merge, whether or not the numbers run in sequence, so $pg12 next to
' $pg99 gives $pg12-99. Jerry's call, 8/2/2026, and the right one: the "number" is often not a
' number. Roman numerals, "12a" and already-hyphenated "12-13" all occur, and "the next one"
' means nothing for any of them.
'
' Called from Lp_Format_Page_Numbers only, between the passes that delete an empty "$pg" and
' the wildcard that builds the bar. Silent -- the caller reports its own completion. Screen
' updating belongs to the caller, which has already turned it off. Returns the number of runs
' merged; nothing reads it, it is there so the work can be checked from the Immediate window.
'
' Version: 1.0  Date: 8/2/2026
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
    Dim runs As Collection          ' each item is itself a Collection of paragraph Ranges
    Dim thisRun As Collection
    Dim para As Paragraph
    Dim prevRange As Range
    Dim r As Long

    ' --- 1. Group the tag paragraphs into runs of genuinely adjacent ones ---
    ' A run is only stored once it reaches two: a lone tag is the normal case and there is
    ' nothing to do with it.
    Set runs = New Collection
    Set thisRun = Nothing

    For Each para In targetDoc.Paragraphs
        If Lp_Is_Pg_Tag_Paragraph(para.Range) Then
            If thisRun Is Nothing Then
                Set thisRun = New Collection
            ElseIf Not Lp_Tags_Can_Merge(prevRange, para.Range) Then
                Set thisRun = New Collection        ' a section boundary starts a new run
            End If
            thisRun.Add para.Range
            If thisRun.count = 2 Then runs.Add thisRun   ' stored once, then it grows in place
            Set prevRange = para.Range
        Else
            Set thisRun = Nothing                   ' any other paragraph ends the run
        End If
    Next para

    If runs.count = 0 Then Exit Function            ' the ordinary case -- no doubled-up tags

    ' --- 2. Merge, working BACK TO FRONT ---
    ' Later runs first, so shortening the document can never move a run that is still waiting
    ' to be looked at. Same reason Sh_Strip_Prodnote_Enclosing_Quotes works backwards.
    For r = runs.count To 1 Step -1
        Set thisRun = runs(r)
        If Lp_Merge_One_Pg_Tag_Run(targetDoc, thisRun) Then
            Lp_Merge_Adjacent_Pg_Tags = Lp_Merge_Adjacent_Pg_Tags + 1
        End If
    Next r

End Function   '*** end of Lp_Merge_Adjacent_Pg_Tags function ***

Private Function Lp_Is_Pg_Tag_Paragraph(ByVal r As Range) As Boolean
'
' True only when this paragraph IS a reference page tag: "$pg" at the START of it, with a page
' number after it. A paragraph that merely CONTAINS "$pg" somewhere in the middle, such as
' "see $pg12 above", is not one and must never be pulled into a merge.
'
' Leading spaces, tabs and non-breaking spaces before the "$" are ignored. The paragraph mark
' is not text -- Sh_Para_Visible_Text strips it.
'
' The test is CASE-INSENSITIVE on purpose. The macros always write lower-case "$pg", but every
' Find in Lp_Format_Page_Numbers runs MatchCase = False, so "$PG12" becomes a pink bar whether
' this notices it or not; being stricter here than the code that follows would only miss it.
' (This module has Option Explicit but no Option Compare, so StrComp is binary unless told.)
'
' Version: 1.0  Date: 8/2/2026
'
    Dim t As String
    Dim p As Long
    Dim inTable As Boolean

    ' Jerry, 8/2/2026: ignore any $pg in a table. Adjacent in Document.Paragraphs is NOT
    ' adjacent on the page -- that collection walks every cell in document order, so the
    ' paragraph after a cell's last one is the first one in the NEXT cell. A cell must also
    ' always keep at least one paragraph, so deleting into its end-of-cell marker raises error
    ' 4605 and would abort the whole macro half way through, screen still frozen. Tags in
    ' tables are still FORMATTED into bars exactly as before; only merging skips them.
    ' If Word cannot tell us, assume it IS in a table -- the safe answer is "do not merge".
    inTable = True
    On Error Resume Next
    inTable = r.Information(wdWithInTable)
    On Error GoTo 0
    If inTable Then Exit Function

    t = Sh_Para_Visible_Text(r)
    p = Sh_First_Visible_Char(t)
    If p = 0 Then Exit Function                     ' empty, or nothing but spaces

    If StrComp(Mid$(t, p, 3), "$pg", vbTextCompare) <> 0 Then Exit Function

    ' An EMPTY tag is not one. The two passes above have already deleted those, but a tag with
    ' nothing to hyphenate would give "$pg-13", so let it end the run instead.
    If Len(Lp_Pg_Tag_Number(t)) = 0 Then Exit Function

    ' The braille lower-roman prefix, "$pg[[*ii*]]". It cannot reach a large print document.
    If InStr(1, t, "$pg[[", vbTextCompare) > 0 Then Exit Function

    ' A break or a tab means this is not a plain tag sitting on its own. The page break matters
    ' most: it is the one case where two numbers back to back must NOT be joined.
    If InStr(t, vbTab) > 0 Then Exit Function       ' Chr(9)
    If InStr(t, Chr$(11)) > 0 Then Exit Function    ' line break
    If InStr(t, Chr$(12)) > 0 Then Exit Function    ' page or section break
    If InStr(t, Chr$(14)) > 0 Then Exit Function    ' column break

    Lp_Is_Pg_Tag_Paragraph = True

End Function   '*** end of Lp_Is_Pg_Tag_Paragraph function ***

Private Function Lp_Tags_Can_Merge(ByVal a As Range, ByVal b As Range) As Boolean
'
' True when two tag paragraphs really are stacked one above the other in the same text flow.
' Tables are already out -- Lp_Is_Pg_Tag_Paragraph rejects them -- so only two tests remain:
'
'   1. They are touching: a ends exactly where b starts.
'   2. They are in the same SECTION. A section break lives on a paragraph mark, so a run that
'      spanned one would lose the break and that section's page setup with it. Two page
'      numbers back to back at a chapter change is exactly where a break turns up.
'
' Version: 1.0  Date: 8/2/2026
'
    On Error GoTo NoMerge

    If a.End <> b.Start Then Exit Function
    If a.Sections(1).Index <> b.Sections(1).Index Then Exit Function

    Lp_Tags_Can_Merge = True
    Exit Function

NoMerge:                                            ' could not tell, so do not merge
    Lp_Tags_Can_Merge = False

End Function   '*** end of Lp_Tags_Can_Merge function ***

Private Function Lp_Merge_One_Pg_Tag_Run(ByVal targetDoc As Document, _
                                         ByVal tagRun As Collection) As Boolean
'
' Turns one run of tag paragraphs into a single tag: writes the merged number into the LAST
' paragraph of the run, then deletes the earlier ones WHOLE, back to front. True when it
' changed something.
'
' Two deliberate choices, both of them learned the hard way on 8/2/2026:
'
' WHOLE-PARAGRAPH deletes, not one span across the run. Version 1.0 deleted a single range
' running from the first paragraph's mark to the end of the last paragraph's text. Measured on
' a real document the offsets were provably right -- first paragraph 12..21, last 27..33,
' delete 20..32, exactly its mark plus the middle paragraph plus the last one's text -- but
' Word did not honor them: one paragraph mark survived every time the run was three or more,
' leaving a blank paragraph after the merged bar. Deleting whole paragraphs is the pattern
' Sh_Delete_Prodnote_Paragraphs already uses, and it behaves.
'
' The LAST paragraph is the survivor, not the first. Word refuses to delete a document's final
' paragraph mark, so a run that ends at the end of a document would strand an empty paragraph
' if the survivor were the first. Nothing is lost by keeping the last one: the paragraphs are
' adjacent, so the bar lands in the same place on the page either way, and the "Print Pg Num"
' pass restyles the whole paragraph a few lines later regardless.
'
' Version: 1.1  Date: 8/2/2026 - whole-paragraph deletes, and the LAST paragraph of the run is
'                               the one kept. Was a single span delete, which left a blank
'                               paragraph behind whenever the run was three or more
' Version: 1.0  Date: 8/2/2026
'
    Dim firstRng As Range, lastRng As Range
    Dim numRng As Range
    Dim firstText As String, lastText As String
    Dim oldNum As String, mergedNum As String
    Dim tagPos As Long
    Dim i As Long

    If tagRun.count < 2 Then Exit Function

    Set firstRng = tagRun(1)
    Set lastRng = tagRun(tagRun.count)

    firstText = Sh_Para_Visible_Text(firstRng)
    lastText = Sh_Para_Visible_Text(lastRng)

    mergedNum = Lp_Merged_Pg_Number(Lp_Pg_Tag_Number(firstText), Lp_Pg_Tag_Number(lastText))
    If Len(mergedNum) = 0 Then Exit Function

    ' --- 1. Write the merged number into the LAST tag, and nothing else ---
    ' The range covers the characters after "$pg" up to the end of the visible text: not the
    ' red "$pg" itself, and not the paragraph mark. Word gives the inserted text the formatting
    ' of what it replaced, which is the old number.
    tagPos = InStr(1, lastText, "$pg", vbTextCompare)
    If tagPos = 0 Then Exit Function
    oldNum = Mid$(lastText, tagPos + 3)
    Set numRng = targetDoc.Range(lastRng.start + tagPos + 2, lastRng.start + Len(lastText))

    ' Character positions and Range offsets only line up while the paragraph is plain text. If
    ' a field or an inline shape has crept in they will not, so check before writing.
    If StrComp(numRng.Text, oldNum, vbBinaryCompare) <> 0 Then Exit Function
    If StrComp(oldNum, mergedNum, vbBinaryCompare) <> 0 Then numRng.Text = mergedNum

    ' --- 2. Delete the earlier paragraphs of the run, whole, back to front ---
    ' Back to front so that removing one cannot move another that is still waiting to go.
    For i = tagRun.count - 1 To 1 Step -1
        tagRun(i).Delete
    Next i

    Lp_Merge_One_Pg_Tag_Run = True

End Function   '*** end of Lp_Merge_One_Pg_Tag_Run function ***

Private Function Lp_Pg_Tag_Number(ByVal visibleText As String) As String
'
' The page number out of a tag paragraph's visible text: everything after the "$pg", trimmed.
' "  $pg 12 " gives "12". Empty when the tag carries no number at all.
'
' Version: 1.0  Date: 8/2/2026
'
    Dim p As Long

    p = InStr(1, visibleText, "$pg", vbTextCompare)
    If p = 0 Then Exit Function

    Lp_Pg_Tag_Number = Trim$(Replace(Mid$(visibleText, p + 3), ChrW(160), " "))

End Function   '*** end of Lp_Pg_Tag_Number function ***

Private Function Lp_Merged_Pg_Number(ByVal firstNum As String, ByVal lastNum As String) As String
'
' The number for a merged tag: the FIRST part of the first tag's number and the LAST part of
' the last tag's number, joined with one hyphen. That is what stops an already-hyphenated tag
' from growing a second hyphen -- "12-13" next to "14" gives "12-14", not "12-13-14".
'
' A hyphen at either extreme is a stray, not a range boundary, so "-13" keeps its whole text
' rather than contributing an empty first part. Two identical tags give a single number rather
' than "12-12".
'
' Version: 1.0  Date: 8/2/2026
'
    Dim head As String, tail As String
    Dim i As Long

    ' head: everything before the FIRST hyphen of the first number
    For i = 1 To Len(firstNum)
        If Lp_Is_Hyphen_Char(Mid$(firstNum, i, 1)) Then Exit For
    Next i
    If i > 1 And i <= Len(firstNum) Then head = Trim$(Left$(firstNum, i - 1))
    If Len(head) = 0 Then head = Trim$(firstNum)

    ' tail: everything after the LAST hyphen of the last number
    For i = Len(lastNum) To 1 Step -1
        If Lp_Is_Hyphen_Char(Mid$(lastNum, i, 1)) Then Exit For
    Next i
    If i >= 1 And i < Len(lastNum) Then tail = Trim$(Mid$(lastNum, i + 1))
    If Len(tail) = 0 Then tail = Trim$(lastNum)

    If Len(head) = 0 Or Len(tail) = 0 Then Exit Function

    If StrComp(head, tail, vbTextCompare) = 0 Then
        Lp_Merged_Pg_Number = head
    Else
        Lp_Merged_Pg_Number = head & "-" & tail
    End If

End Function   '*** end of Lp_Merged_Pg_Number function ***

Private Function Lp_Is_Hyphen_Char(ByVal c As String) As Boolean
'
' A plain hyphen, Word's non-breaking hyphen, an en dash or an em dash. A converted book can
' arrive carrying any of them. The dashes are written as ChrW codes so they survive every
' export and re-import of this module; what gets written back is always a plain hyphen, which
' is what the auto-tagger's own patterns match (^045).
'
' Version: 1.0  Date: 8/2/2026
'
    Select Case c
        Case "-", Chr$(30), ChrW(8211), ChrW(8212)
            Lp_Is_Hyphen_Char = True
    End Select

End Function   '*** end of Lp_Is_Hyphen_Char function ***

Sub Lp_AutoTag_Page_Numbers()
'
' Version: 2.5  Date: 8/2/2026 - the twelve "^013(...)^013" passes now repeat until nothing is left to replace (Sh_Replace_All_Until_Done); a single Execute tagged only alternate numbers when two page numbers sat in consecutive paragraphs
' Version: 2.4  Date: 7/26/2026 - returns the user to where the cursor was when the macro started
' Version: 2.2  Date: 4/30/2023 - complete rewrite to eliminate false tagging
'
'  Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Locates potential reference page numbers and tags with $pg
'

    ' FIRST, before anything reads or moves the cursor. Validation can leave every $pg
    ' paragraph selected at once, and Word then refuses most Selection work with error 4605.
    Sh_Clear_Multi_Selection

    Dim strLength As Integer

    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating off

    Sh_Save_User_Position

    Application.Run MacroName:="Sh_Is_Doc_Open"
    Application.Run MacroName:="MS_Set_Word_Config_For_Large_Print"
    Application.Run MacroName:="Lp_Fix_Para_Space_Errors"
    Application.Run MacroName:="Lp_Fix_Hyphen_Errors"

    ' add para mark at begining of doc to assure that ref pg no on first line is detected (removed at close of macro)
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.HomeKey Unit:=wdStory

    ' ***************** fix publisher file page codeing ***********
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "$pn"
        .Replacement.Text = "$pg"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "$ppn"
        .Replacement.Text = "$pg"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "ppn"
        .Replacement.Text = "$pg"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = wdColorAutomatic
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013pn([0-9]{1,})"
        .Replacement.Text = "^013$pg\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    '*************** begin tagging Tagging Roman Numerals ***********************
    Dim para As Paragraph
    Dim tempStr As String
    Dim txt As String
     
    ' look at each paragraph in the document
    For Each para In ActiveDocument.Paragraphs
            
        txt = para.Range.Text
        
        ' the longest roman numeral is 10 characters plus 1 for the para mark = 11
        ' when length is greater than 11, then the paragraph is too long to be a Roman numeral pg numb
        ' when > 11 go to next paragraph
        If Len(txt) > 11 Then
          GoTo LoopEnd
        End If
        
        'bypass finding the roman numeral - already tagged - go to next paragraph
        If Left(txt, 3) = "$pg" Then
            GoTo LoopEnd
        End If
        
        tempStr = Trim(Left(UCase(txt), 11)) ' change to upper case and take up to 11 characters
        strLength = Len(tempStr) - 1 'set length not including the para mark
        tempStr = Left(tempStr, strLength)  ' set comparison string
        
        'is the trimmed upper case string a valid roman numeral
        If Len(tempStr) > 0 Then
            On Error Resume Next 'prevents crash in a table
            If Sh_IsValidRomanNumeral(tempStr) Then
                para.Range.InsertBefore ("$pg") ' put $pg at front of paragraph
            End If
        End If
LoopEnd:
    Next
    '*************** end tagging Tagging Roman Numerals ***********************

    '*************** begin tagging arabic page number variations ***********************
    
    Application.Run MacroName:="Sh_Fix_Ref_Pages_Before_and_After_Tables"
    
    ' place para mark before graphics (otherwise ref pg no before graphic will not be found
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^g"
        .Replacement.Text = "*~^p^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' place para mark after graphics
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^g"
        .Replacement.Text = "^&*~^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll


    ' any length numb hyphen numb any length
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([0-9]{1,})([0-9]{1,})^013"
        .Replacement.Text = "^p$pg\1\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done
    
    ' any length number, hyphen, number
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013([0-9]{1,}^045)([0-9]{1,})^013"
        .Replacement.Text = "^p$pg\1\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done

    'letter , any lengthNumber, Hyphen, letter, Number
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([A-Za-z]{1,}[0-9]{1,})([A-Za-z]{1,}[0-9]{1,})^013"
        .Replacement.Text = "^p$pg\1\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done

    'letter ,Hyphen, letter
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([A-Za-z]{1,})([A-Za-z]{1,})^013"
        .Replacement.Text = "^p$pg\1\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done

    'letter , any length Number, letter
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([A-Za-z]{1,}[0-9]{1,}[A-Za-z]{1,})^013"
        .Replacement.Text = "^p$pg\1^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done

    ' any length Number
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([0-9]{1,})^013"
        .Replacement.Text = "^p$pg\1^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done

    ' any length Number, Letter
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([0-9]{1,}[A-Za-z]{1,})^013"
        .Replacement.Text = "^p$pg\1^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done
    
    ' Letter,  any length Number
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([A-Za-z]{1,}[0-9]{1,})^013"
        .Replacement.Text = "^p$pg\1^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done
    
    'Letter, any length Number, hypen, letter, number
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([A-Za-z]{1,}[0-9]{1,}^045)([A-Za-z]{1,}[0-9]{1,})^013"
        .Replacement.Text = "^p$pg\1\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done
    
    ' any length Number, letter, Hyphen, Number, letter,
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([0-9]{1,}[A-Za-z]{1,}^045)([0-9]{1,}[A-Za-z]{1,})^013"
        .Replacement.Text = "^p$pg\1\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done

    ' any length Number, letter, Hyphen, Number, letter,
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([A-Za-z]{1,}[0-9]{1,}[A-Za-z]{1,})([A-Za-z]{1,}[0-9]{1,}[A-Za-z]{1,})^013"
        .Replacement.Text = "^p$pg\1\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done
    
    'letter, any length number, letter, Hyphen, letter, number, letter
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
       .Text = "^013([A-Za-z]{1,}[0-9]{1,}[A-Za-z]{1,}^045)([A-Za-z]{1,}[0-9]{1,}[A-Za-z]{1,})^013"
        .Replacement.Text = "^p$pg\1\2^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Sh_Replace_All_Until_Done
    
    ' remove para mark before graphics (placed at top of macro)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "*~^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' remove para mark after graphics (placed at top of macro)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "*~^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' remove para mark before graphics
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "*~"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
  
    Application.Run MacroName:="Sh_Color_Dollar_PG_Red"
    
    ActiveDocument.UndoClear
    
    'remove top para mark
    Selection.HomeKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1

    Application.ScreenUpdating = su_Prev ' Turn screen updating on
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    
    'Count the tags
    Dim TagCounter As Integer
    TagCounter = 0
    Dim par As Word.Paragraph

    For Each par In ActiveDocument.Paragraphs
        txt = par.Range.Text
        If InStr(txt, "$pg") > 0 Then
            TagCounter = TagCounter + 1
        End If
    Next
        
    ' Put the user back BEFORE the validate prompt: if they say yes, Lp_Validate_Dollar_pg
    ' is supposed to leave them sitting at the tag it is complaining about.
    Sh_Return_User_To_Start_Position

    If TagCounter = 0 Then
        MsgBox "There are no tagged page numbers in this document.", , "VistaType LP (128)"
    Else
        If MsgBox("There are " + Trim(Str(TagCounter)) + " page numbers in the document." + vbCr + vbCr _
        + "Do you want to validate the tagged page numbers?", vbYesNo, "VistaType LP (182)") = vbYes Then
            Application.Run MacroName:="Lp_Validate_Dollar_pg"
        End If
    End If
    
End Sub   '****end of Lp_AutoTag_Page_Numbers Macro ***********

Sub Lp_Manual_Tag_with_Dollar_pg()
'
' Lp_Manual_Tag_with_Dollar_pg Macro
'
' Version: 1.5  Date: 3/16/2017
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
    Application.Run MacroName:="Sh_Is_Doc_Open"

    Application.Run MacroName:="Sh_Remove_DollarPG_For_Retag"

    Selection.EndKey Unit:=wdLine
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend

    With Selection.Font
        .Color = wdColorAutomatic
    End With
    
    Selection.HomeKey Unit:=wdLine

    With Selection.Font
        .Color = wdColorRed
    End With
    
    Selection.TypeText Text:="$pg"
    Selection.HomeKey Unit:=wdLine
    
End Sub  '***** end of Lp_Manual_Tag_with_Dollar_pg Macro *****

Sub Sh_Kill_The_Hyperlinks()
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Strips hyperlinks from the selection if there is one, and from the whole document if not.
' Shared by both sides.
'
' Version: 2.0  Date: 8/12/2026 - no temporary document, and ONE macro. The whole body of both
'                                copies was one call to Sh_Remove_Hyperlinks wrapped in
'                                round-trip plumbing - but the scratch document was doing real
'                                work here, because Sh_Remove_Hyperlinks deleted every hyperlink
'                                in the active document. It takes a range now, so the scope is
'                                explicit. With that, the two sides were identical and merged.
'                                Gone with the round trip: the braille copy's Selection.MoveUp
'                                and its delete-one-character run, and the large-print copy's
'                                "delete one character at the end of the document", which ran on
'                                every route out and ate the last character when nothing was
'                                selected.
' Version: 1.3  Date: 10/6/2023
' Version: 1.2  Date: 1/8/2019
' Version: 1.1  Date: 1/9/2016
'
    Dim rng As Range
    Dim su_Prev As Boolean

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    If Selection.Type = wdSelectionNormal Then
        Set rng = Selection.Range
    Else
        Set rng = ActiveDocument.Content
    End If

    Sh_Remove_Hyperlinks rng

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub '***** End of Sh_Kill_The_Hyperlinks Macro ******************

' The range the paragraph-and-space passes should work on, with two adjustments that the
' temporary document used to make by hand. Both were found by testing the conversion of
' Lp_/Dx_Fix_Para_Space_Errors on 8/12/2026, and both are easy to lose:
'
' 1. NEVER include the document's own final paragraph mark. These passes rewrite ^013, and
'    rewriting the final mark ADDS a paragraph. The old code cleaned that up afterwards with
'    "go to the end, delete one character" - which on the no-selection route also deleted the
'    last character of the transcriber's text. Not touching the mark is better than tidying
'    up after it.
'
' 2. When there IS a selection, reach back one character to take in the paragraph mark in
'    FRONT of it. "Spaces after a paragraph mark" cannot match the first selected paragraph
'    otherwise, because its opening mark sits outside the selection - so leading spaces on
'    the first line survived. The temporary document got this right by accident: it began by
'    typing a paragraph mark at the top so the pattern had something to match. Including the
'    real one is the same trick without the scratch document. Nothing before the selection is
'    harmed: every one of these patterns puts the paragraph mark back.
'
' Version: 1.0  Date: 8/12/2026
Public Function Sh_Para_Fix_Range() As Range
    Dim doc As Document
    Dim a As Long, b As Long

    Set doc = ActiveDocument

    If Selection.Type = wdSelectionNormal Then
        a = Selection.Range.Start
        b = Selection.Range.End
        If a > 0 Then
            If doc.Range(a - 1, a).Text = vbCr Then a = a - 1
        End If
    Else
        a = 0
        b = doc.Content.End
    End If

    If b > doc.Content.End - 1 Then b = doc.Content.End - 1
    If b < a Then b = a

    Set Sh_Para_Fix_Range = doc.Range(a, b)
End Function

Sub Lp_Fix_Para_Space_Errors()
'
' Lp_Fix_Para_Space_Errors Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version: 2.0 Date: 8/12/2026 - no temporary document; every pass is scoped to a range.
'
'                               NOT merged with the braille version: the two do different work.
'                               This one protects the "1 point" style; the braille one deals with
'                               middle dots and with tabs after paragraph marks. Merging would
'                               mean each side gaining passes nobody asked for.
'
'                               A REAL BUG went with the round trip. The old tail did
'                               Selection.EndKey then Delete one character on EVERY route out,
'                               not just the temp-document one - so running this with nothing
'                               selected deleted the last character of the transcriber's
'                               document. It existed to tidy up after the paste.
'
'                               Also gone: ActiveDocument.UndoClear, which threw away the WHOLE
'                               undo history rather than this macro's part of it.
' Version: 1.3 Date: 10/17/2023 - fixed bug adding para mark and deleted last char in selection
' Version: 1.2 Date:  2/8/2020 - added protection for all "1 point" styles
' Version: 1.1 Date: 1/8/2019
' Version: 1.0 Date: 2/13/2015
'
    Dim rng As Range
    Dim su_Prev As Boolean

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    Set rng = Sh_Para_Fix_Range()

    ' rogue paragraph marks
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^013"
        .Replacement.Text = "^p"
        .Forward = True
        ' wdFindStop, not wdFindContinue: continue would run past the end of the range and
        ' quietly turn "the selection" into "the whole document".
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    ' Protect the "1 point" style from the rest of the macros: every "1 point" paragraph is left
    ' holding a single non-breaking space and nothing else. Guarded - Styles(name) raises 5941
    ' when the document does not carry the style, and with no style there is nothing to protect.
    If Sh_Style_Exists(ActiveDocument, "1 point") Then
        With rng.Find
            .ClearFormatting
            .Style = ActiveDocument.Styles("1 point")
            .Replacement.ClearFormatting
            .Text = "?{1,}^013"
            .Replacement.Text = "^s^013"
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
    End If

    ' spaces before a paragraph mark
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^032{1,}^013"
        .Replacement.Text = "^p"
        .Forward = True
        ' wdFindStop, not wdFindContinue: continue would run past the end of the range and
        ' quietly turn "the selection" into "the whole document".
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    ' spaces after a paragraph mark
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^013^032{1,}"
        .Replacement.Text = "^p"
        .Forward = True
        ' wdFindStop, not wdFindContinue: continue would run past the end of the range and
        ' quietly turn "the selection" into "the whole document".
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub '***** End of Lp_Fix_Para_Space_Errors ********

Sub Lp_Fix_Em_Dash_Space_Errors()
'
' Lp_Fix_Em_Dash_Space_Errors Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Date: 5/31/2016
' Version: 1.1
'

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^+{1,}"
        .Replacement.Text = "^+"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^+^032{1,}"
        .Replacement.Text = "^+"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "^032{1,}^+"
        .Replacement.Text = "^+"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub '***** End of Lp_Fix_Em_Dash_Space_Errors Macro *****

' target  a range to work on. Omit it and the macro decides for itself, as it does when a
'         transcriber runs it: the selection if there is one, the whole document if not. A
'         converted caller that is already working on a range passes that range, which also
'         skips the cursor save-and-restore - the caller owns that.
'
Sub Sh_Remove_Multi_Spaces(Optional ByVal target As Range)
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Collapses runs of spaces to one, over the selection if there is one and the whole document if
' there is not. Shared by both sides.
'
' Version: 2.1  Date: 8/12/2026 - ONE macro. Lp_ and Dx_Remove_Multi_Spaces were separate, and
'                                after 2.0 below the two were the same job written twice -
'                                everything that had differed existed to serve the round trip.
'                                The rule from here: converting a pair that leaves them identical
'                                means merging them (Jerry).
' Version: 2.0  Date: 8/12/2026 - no temporary document. It was only ever there to keep the
'                                replace inside the selection, and a RANGE does that on its own -
'                                so the scratch window is never created, shown or activated, and
'                                the flashing goes with it. Two things went with the round trip:
'                                the large-print copy's ActiveDocument.UndoClear, which threw away
'                                the WHOLE undo history and not just this macro's part of it, and
'                                the braille copy's Selection.MoveUp, which quietly stretched the
'                                work up by a paragraph beyond what was selected.
' Version: 1.3  Date: 7/22/2025 - added code for table pasting
' Version: 1.2  Date: 10/17/2023 - bug fix for deleting last char and adding para mark
' Version: 1.1  Date: 1/8/2019
' Version: 1.0  Date: 2/13/2015
'
    Dim rng As Range
    Dim su_Prev As Boolean
    Dim onOwn As Boolean

    onOwn = (target Is Nothing)

    If onOwn Then
        su_Prev = Application.ScreenUpdating
        Application.ScreenUpdating = False
        Sh_Save_User_Position
        ' Text selected means "just this bit"; nothing selected means the whole document. That
        ' was the ONLY thing the round trip decided, and Find on a range answers it directly.
        If Selection.Type = wdSelectionNormal Then
            Set rng = Selection.Range
        Else
            Set rng = ActiveDocument.Content
        End If
    Else
        Set rng = target
    End If

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^032{2,}"
        .Replacement.Text = "^032"
        .Forward = True
        ' wdFindStop, not wdFindContinue: continue would carry on past the end of the range and
        ' quietly treat "the selection" as "the whole document".
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    If onOwn Then
        Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
        Sh_Return_User_To_Start_Position
        Selection.Collapse Direction:=wdCollapseStart
        Application.ScreenUpdating = su_Prev
        Application.ScreenRefresh
    End If

End Sub  '***** End of Sh_Remove_Multi_Spaces ********

Sub Lp_Copy_To_Temp_Doc()
    '
    ' Version: 2.4 Date: 8/2/2026 - no longer hides the Styles pane: that setting is Word-wide, so it closed the pane in the user's own book, and nothing ever put it back
    ' Version: 2.3 Date: 3/5/2026 - full rewrite of previous versions
    '
    Dim origDoc As Document
    Dim tempDoc As Document
    Dim sourceRng As Range
    Dim destRng As Range
    Dim wasInTable As Boolean
    Dim strTemplatePath As String
    
    ' 1. Set the Template Path and Verify
    strTemplatePath = Options.DefaultFilePath(wdUserTemplatesPath) & "\" & LP_TEMPLATE_FILE
    
    If Dir(strTemplatePath) = "" Then
        MsgBox "Template not found at: " & strTemplatePath, vbCritical, "Template Error"
        Exit Sub
    End If

    ' 2. Capture selection from original document
    If Selection.Information(wdWithInTable) Then
        wasInTable = True
        Set sourceRng = Selection.Tables(1).Range
    Else
        Set sourceRng = Selection.Range
    End If
    
    Set origDoc = ActiveDocument
    
    ' START GLOBAL FREEZE
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' 3. Create the document
    Set tempDoc = Documents.Add(Template:=strTemplatePath, Visible:=False)
    
    ' 4. Transfer content
    Set destRng = tempDoc.Range
    If wasInTable Then
        destRng.InsertParagraphBefore
        Set destRng = tempDoc.Paragraphs(2).Range
        destRng.Collapse wdCollapseStart
    End If
    
    destRng.FormattedText = sourceRng.FormattedText
    
    ' 5. UI Cleanup
    ' 8/2/2026 - removed "Application.TaskPanes(wdTaskPaneFormatting).Visible = False". The
    ' intent was to keep the scratch window uncluttered, but pane visibility is a WORD-WIDE
    ' setting, not a per-document one, so there is no way to tidy the temp document without
    ' closing the pane in the user's book as well - and nothing ever put it back. With
    ' eighteen callers (Table Tools, Fill-In Line, Horiz-to-Vert list, image color) the pane
    ' died on the first one used and stayed dead for the rest of the session. Save-and-restore
    ' is not the answer either: several callers never reach Lp_Copy_From_Temp_Doc, so the
    ' saved state would strand.

    ' 6. THE REVEAL - Fix for Error 5941
    ' If no window exists for this hidden doc, we create one now
    If tempDoc.Windows.count = 0 Then
        tempDoc.Windows.Add
    End If

    ' Now that we are sure Windows(1) exists, we configure it
    With tempDoc.Windows(1)
        .Visible = True
        .WindowState = wdWindowStateMaximize
    End With

    ' 7. Final handoff
    tempDoc.Activate
    Application.ScreenUpdating = su_Prev

End Sub '*** end of Lp_Copy_To_Temp_Doc Macro ***

Sub Lp_Copy_From_Temp_Doc()
    '
    ' copies changes from a temp file back into the original file
    '
    ' Version: 1.6  Date: 2/8/2026 - full rewrite
    ' Version: 1.5  Date: 1/23/2021
    '
    Dim masterDoc As Document
    Dim tempDoc As Document
    Dim targetRange As Range
    Dim tempName As String

    ' 1. Identify the Temp Doc (currently active)
    Set tempDoc = ActiveDocument
    tempName = tempDoc.Name

    ' 2. Identify the Master Doc
    If Documents.count > 1 Then
        If Documents(1).Name <> tempName Then
            Set masterDoc = Documents(1)
        Else
            Set masterDoc = Documents(2)
        End If
    Else
        MsgBox "Master document not detected.", vbCritical
        Exit Sub
    End If

    ' 3. Set the target to the current selection in the Master Doc
    ' We use the Master's active window to find where your cursor is
    Set targetRange = masterDoc.ActiveWindow.Selection.Range

    ' 4. Transfer the content directly (Formatting included)
    ' This replaces the targetRange with the tempDoc's content
    targetRange.FormattedText = tempDoc.Content.FormattedText
    
    ' 5. Close the Temp file
    tempDoc.Close SaveChanges:=wdDoNotSaveChanges
    
    ' Optional: Bring Master to front so you can see the result
    masterDoc.Activate
End Sub

Sub Lp_Toggle_Page_Color()
'
' Lp_Toggle_Page_Color macro
'
' Version: 1.3  Date: 10/18/2020 - added Lp_Set_Doc_Background_Form
' Version: 1.0   Date: 5/18/2016
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    Application.Run MacroName:="Lp_Is_Lp_Template_Attached"
    
    Lp_Set_Doc_Background_Form.Show
    
End Sub  '*** end of Lp_Toggle_Page_Color macro ***

Sub Lp_Selected_File_CleanUp()
'
' Lp_Selected_File_CleanUp
'
' Version: 1.4  Date: 1/28/2020 - moved all "is text selected" to internals in the form
' version: 1.2  Date:  2/13/2019 - added check for selected text
' Version: 1.1  Date: 10/19/2018 - check for selected text before showing menu
' Version: 1.0  Date: 5/18/2016
'
' Author: Jerry Whittaker - jerry@thewhittakers.org

    Application.Run MacroName:="Sh_Is_Doc_Open"
    Lp_Selected_Cleanup_Form.Show
    Unload Lp_Selected_Cleanup_Form

End Sub  '*** end of Lp_Selected_File_CleanUp macro ***

Public Sub Sh_Linkify_Range(ByVal target As Range)
'
' Version: 1.0  Date: 8/13/2026
'
' Makes a live hyperlink out of every plain URL or e-mail address in the range. This is the
' replacement for Selection.Range.AutoFormat, which did the same job as a side effect of
' reformatting the entire document - and took empty paragraphs with it. See the note in
' Lp_Convert_Hyper_To_Addresses.
'
' Text that is already part of a hyperlink is left alone, so running this twice is harmless.
'
    Dim doc As Document
    Dim p As Paragraph
    Dim raw As String, flat As String, low As String
    Dim parts() As String
    Dim k As Long, pos As Long, lead As Long
    Dim offs() As Long
    Dim tok As String, addr As String
    Dim r As Range

    If target Is Nothing Then Exit Sub
    Set doc = target.Document

    On Error Resume Next

    For Each p In target.Paragraphs
        raw = p.Range.Text

        ' Everything that separates one token from the next becomes a space, one character for
        ' one character, so an offset into flat is still an offset into the paragraph.
        flat = Replace(raw, vbTab, " ")
        flat = Replace(flat, Chr(160), " ")      ' non-breaking space
        flat = Replace(flat, Chr(11), " ")       ' manual line break
        flat = Replace(flat, vbCr, " ")
        flat = Replace(flat, Chr(7), " ")        ' end-of-cell marker

        ' Cheap gate: most paragraphs are prose and can be skipped without splitting them.
        low = LCase$(flat)
        If InStr(low, "://") > 0 Or InStr(low, "www.") > 0 Or InStr(low, "@") > 0 Then

            parts = Split(flat, " ")

            ' Offsets first, then work BACK TO FRONT. Hyperlinks.Add turns the anchor into a
            ' field, and the field takes up room, so every offset later in the paragraph moves
            ' the moment the first link is made. Going backwards means nothing we have yet to
            ' handle has moved. (Front to back linked the first address in a paragraph and
            ' quietly missed the rest - measured, on a line holding an e-mail and a URL.)
            ReDim offs(UBound(parts))
            pos = 0
            For k = 0 To UBound(parts)
                offs(k) = pos
                pos = pos + Len(parts(k)) + 1    ' + 1 for the space Split ate
            Next k

            For k = UBound(parts) To 0 Step -1
                pos = offs(k)
                tok = parts(k)
                If Len(tok) > 0 Then
                    ' Trim what a sentence puts around an address. Count what comes off the
                    ' FRONT - the anchor has to start after it.
                    lead = 0
                    Do While Len(tok) > 0
                        If InStr("([{<" & Chr(34) & "'", Left$(tok, 1)) = 0 Then Exit Do
                        tok = Mid$(tok, 2)
                        lead = lead + 1
                    Loop
                    Do While Len(tok) > 0
                        If InStr(".,;:!?)]}>" & Chr(34) & "'", Right$(tok, 1)) = 0 Then Exit Do
                        tok = Left$(tok, Len(tok) - 1)
                    Loop

                    addr = Sh_Link_Address(tok)
                    If Len(addr) > 0 Then
                        Set r = doc.Range(p.Range.Start + pos + lead, _
                                          p.Range.Start + pos + lead + Len(tok))
                        If r.Hyperlinks.count = 0 Then
                            doc.Hyperlinks.Add Anchor:=r, Address:=addr
                        End If
                        Set r = Nothing
                    End If
                End If
            Next k

        End If
    Next p

    Err.Clear

End Sub   '*** end of Sh_Linkify_Range ***

Private Function Sh_Link_Address(ByVal tok As String) As String
'
' The address a token should point at, or "" when it is not an address at all.
'
    Dim low As String
    Dim at As Long

    If Len(tok) < 5 Then Exit Function
    low = LCase$(tok)

    If Left$(low, 7) = "http://" Or Left$(low, 8) = "https://" Or Left$(low, 6) = "ftp://" Then
        Sh_Link_Address = tok

    ElseIf Left$(low, 4) = "www." Then
        ' www.example.com - needs a dot after the "www." or it is not a host name
        If InStr(5, low, ".") > 0 Then Sh_Link_Address = "http://" & tok

    Else
        ' name@example.com - an @ with something either side and a dot in the domain
        at = InStr(low, "@")
        If at > 1 Then
            If InStr(at + 2, low, ".") > 0 And InStr(at + 1, low, "@") = 0 Then
                Sh_Link_Address = "mailto:" & tok
            End If
        End If
    End If

End Function   '*** end of Sh_Link_Address ***

Sub Lp_Convert_Hyper_To_Addresses()
'
' Replaces hyperlinks with the address of the hyperlink, then turns the plain addresses back
' into live links.
'
' from: http://stackoverflow.com/questions/16493791/
'     extract-hyperlink-address-from-hyperlink-field-code
'
' Version: 2.0   Date: 8/13/2026 - no temporary document; everything is scoped to a range.
'
'                                 ActiveDocument.Hyperlinks is a whole-document collection and
'                                 the "mailto:" pass was a Selection.Find with wdFindContinue, so
'                                 the scratch document was the only thing holding either of them
'                                 to the transcriber's selection. Sh_Linkify_Range was being
'                                 handed ActiveDocument.Content outright.
'
'                                 Gone with the round trip: "go to the end of the document and
'                                 delete one character", which ran on EVERY route out and ate the
'                                 last character of the book when nothing was selected, and
'                                 ActiveDocument.UndoClear, which threw away the whole undo
'                                 history rather than this macro's part of it.
'
'                                 NOT merged with the braille copy, which stops after turning
'                                 the links into text and does not make them live again. A DBT
'                                 source file has no use for a clickable link.
' Version: 1.3   Date: 8/13/2026 - AutoFormat gone; see the note on Sh_Linkify_Range
' Version: 1.2   Date: 1/8/2019
' Version: 1.1   Date: 6/9/2016
'
    Dim rng As Range
    Dim r As Range
    Dim strLinkText As String
    Dim i As Long
    Dim su_Prev As Boolean

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    If Selection.Type = wdSelectionNormal Then
        Set rng = Selection.Range
    Else
        Set rng = ActiveDocument.Content
    End If

    ' Every hyperlink in scope becomes its own address, as plain text. Backwards by index:
    ' writing r.Text takes the hyperlink out of the collection and changes the length of the
    ' text, so anything counted from the front has moved by the next turn of the loop.
    On Error Resume Next
    For i = rng.Hyperlinks.count To 1 Step -1
        With rng.Hyperlinks(i)
            strLinkText = .Address
            ' optional, should be OK for HTML links
            If .SubAddress <> "" Then
                strLinkText = strLinkText & "#" & .SubAddress
            End If
            Set r = .Range
        End With
        r.Text = strLinkText
        Set r = Nothing
    Next i
    On Error GoTo 0

    ' "mailto:" is how Word writes an e-mail address and is not part of the address a reader
    ' wants to see
    Sh_Delete_Text_In_Range rng, "mailto:"
    Sh_Linkify_Range rng

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub  '*** end of Lp_Convert_Hyper_To_Addresses ***

Sub Sh_Remove_Txt_Bxs_And_Frames()
    '
    ' Original Macro Name "TextBoxFrameCut" by author below
    '
    ' From: Computer Tools for Editors(and Proofreaders)by Paul Beverley, LCGI
    '        downloadable at no cost from http://www.archivepub.co.uk/book.html
    ' Also from: https://wordribbon.tips.net/T009169_Removing_All_Text_Boxes_In_a_Document.html
    ' Initial version by richardwalshe@prufrock.co.uk
    '
    ' Takes the text out of every text box and frame in scope, leaves it where the box was
    ' anchored, marked top and bottom so the transcriber can check that it landed somewhere
    ' sensible, and deletes the box.
    '
    ' Version: 4.0 Date: 8/13/2026 - ONE macro. At Jerry's word the large-print copy took on the
    '                               braille behavior - ungroup first, take text from ANY shape
    '                               that has some rather than text boxes only, and keep the text's
    '                               formatting - and with that the two were the same job written
    '                               twice, so they merged. The old names are gone and all five
    '                               callers point here.
    '
    '                               Two things the large-print copy loses in the bargain, both
    '                               worth having lost: it pulled the text out as a plain STRING,
    '                               so bold, italics and everything else in a text box arrived
    '                               flat; and its closing marker had no paragraph mark after it,
    '                               so it welded itself onto the front of the anchor paragraph.
    '
    ' Version: 3.0 Date: 8/13/2026 - no temporary document.
    '
    '                               The scratch document was REAL WORK here, not plumbing.
    '                               ActiveDocument.Shapes and .Frames are collections of the whole
    '                               document - there is no such thing as "the shapes in this
    '                               selection" - so copying the selection out was the only thing
    '                               keeping this off the rest of the book. Scope is explicit now:
    '                               a shape belongs to the range when its ANCHOR does
    '                               (Sh_Shape_In_Range), which is the same test Word itself uses
    '                               when it decides what travels with copied text.
    '
    '                               Gone with the round trip: the large-print "go to the end of
    '                               the document and delete one character", which ran on EVERY
    '                               route out and ate the last character of the book when nothing
    '                               was selected; its ActiveDocument.UndoClear, which threw away
    '                               the whole undo history rather than this macro's part of it;
    '                               and the braille four-character delete run.
    '
    '                               The clipboard is gone too. The braille copy kept formatting by
    '                               copying and pasting; a FormattedText assignment does the same
    '                               thing without taking the transcriber's clipboard away.
    '
    '                               The shape loop runs backwards by index. For Each over Shapes
    '                               while deleting from it SKIPS the next shape, which is what the
    '                               old large-print "do the whole pass four times" loop was
    '                               working around. Backwards, one pass is enough.
    ' Version: 2.0 Date: 1/17/2023 - corrected spacing and removed added para mark and char deletion
    ' Version: 1.1 Date: 11/16/2016
    '
    Dim doc As Document
    Dim rng As Range
    Dim shp As Shape
    Dim fr As Frame
    Dim dest As Range
    Dim i As Long
    Dim ipos As Long
    Dim hasText As Boolean
    Dim landed As Boolean
    Dim su_Prev As Boolean

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    Set doc = ActiveDocument
    Set rng = doc.Content
    If Selection.Type = wdSelectionNormal Then
        ' A selection INSIDE a text box is Selection.Type = wdSelectionNormal too, but its story
        ' is the text box's own, not the document's - and the shapes we are looking for are
        ' anchored in the document. Scoping to it matches nothing and the macro appears to do
        ' nothing at all, which is what Jerry hit on 8/13/2026 after clicking into a box. A
        ' selection that is not document text is treated as no selection.
        If Selection.Range.StoryType = wdMainTextStory Then Set rng = Selection.Range
    End If

    ' Ungroup first - a text box inside a group is not reachable as a shape of its own.
    For i = doc.Shapes.count To 1 Step -1
        Set shp = doc.Shapes(i)
        If shp.Type = msoGroup Then
            If Sh_Shape_In_Range(shp, rng) Then shp.Ungroup
        End If
    Next i

    ' Backwards by index: deleting from Shapes inside a For Each skips the next one.
    For i = doc.Shapes.count To 1 Step -1
        Set shp = doc.Shapes(i)
        If Sh_Shape_In_Range(shp, rng) Then

            ' NOT TextFrame.HasText. Word answers True for a text box holding nothing but its
            ' own paragraph mark, so an empty box came out as a pair of markers with nothing
            ' between them - Jerry's "phantom converted text box", 8/13/2026. An empty box has
            ' nothing to rescue, so it is deleted without ceremony.
            hasText = Sh_Shape_Has_Real_Text(shp)

            If Not hasText Then
                ' Only an actual EMPTY TEXT BOX goes. Every other shape without text is a
                ' picture, a line or a drawing, and this macro has always left images alone -
                ' "Leaves images intact", as the braille copy's own comment put it.
                If shp.Type = msoTextBox Then
                    On Error Resume Next
                    shp.Delete
                    On Error GoTo 0
                End If
            Else
                ' Build at the anchor paragraph, forwards: marker, the text with its formatting
                ' intact, marker. Assigning FormattedText to a COLLAPSED range inserts.
                '
                ' The shape is deleted ONLY if the text actually arrived. This whole block used
                ' to sit under one On Error Resume Next, so a failed FormattedText assignment
                ' was swallowed and the macro carried on to delete the box - leaving a pair of
                ' markers with nothing between them and the text gone for good. Losing a text
                ' box's contents is far worse than leaving the box where it is for the
                ' transcriber to deal with. Jerry, 8/13/2026.
                ipos = shp.Anchor.Paragraphs(1).Range.start
                Set dest = doc.Range(ipos, ipos)
                dest.InsertAfter SH_TBX_BELOW & vbCr
                dest.Collapse Direction:=wdCollapseEnd

                On Error Resume Next
                Err.Clear
                dest.FormattedText = shp.TextFrame.TextRange.FormattedText
                landed = (Err.Number = 0)
                Err.Clear
                On Error GoTo 0

                If landed Then
                    dest.Collapse Direction:=wdCollapseEnd
                    dest.InsertAfter vbCr & SH_TBX_ABOVE & vbCr
                    shp.Delete
                Else
                    ' take the opening marker back out and leave the box alone
                    doc.Range(ipos, ipos + Len(SH_TBX_BELOW) + 1).Delete
                End If
            End If
        End If
    Next i

    For i = doc.Frames.count To 1 Step -1
        Set fr = doc.Frames(i)
        If Sh_Frame_In_Range(fr, rng) Then
            ' An empty frame gets no markers either - same reason as the empty text box above.
            If Len(Trim$(Replace(Replace(fr.Range.Text, vbCr, ""), Chr(11), ""))) > 0 Then
                fr.Range.InsertBefore SH_TBX_BELOW & vbCr
                fr.Range.InsertAfter vbCr & SH_TBX_ABOVE & vbCr
            End If
            fr.Delete
        End If
    Next i

    Sh_Text_Frame_Warning_To_Red rng

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub '***** end of Sh_Remove_Txt_Bxs_And_Frames macro *****

Public Function Sh_Shape_In_Range(ByVal shp As Shape, ByVal rng As Range) As Boolean
'
' Version: 1.0  Date: 8/13/2026
'
' Is this floating shape part of the range? A shape has no position in the text of its own - it
' hangs off an ANCHOR, and the anchor is what moves with the text when it is copied. So the
' anchor is what decides.
'
' The story has to match before the offsets are worth comparing: a shape anchored in a header or
' a footnote has an anchor whose .Start counts from the beginning of THAT story, and it would
' otherwise land inside the range's numbers by coincidence.
'
    Dim a As Range

    On Error Resume Next
    Set a = shp.Anchor
    If a Is Nothing Then Exit Function
    If a.StoryType <> rng.StoryType Then Exit Function
    ' STRICTLY less than rng.End. A selection of one paragraph ends at the START of the
    ' next one, so "<=" pulled in a shape anchored to the paragraph just past the
    ' selection - selecting one paragraph converted a shape belonging to the one below it.
    Sh_Shape_In_Range = (a.start >= rng.start And a.start < rng.End)
    Err.Clear

End Function   '*** end of Sh_Shape_In_Range ***

Public Function Sh_Shape_Has_Real_Text(ByVal shp As Shape) As Boolean
'
' Version: 1.0  Date: 8/13/2026
'
' Has this shape any text worth rescuing?
'
' TextFrame.HasText cannot answer that. It returns TRUE for a text box that holds nothing but
' its own paragraph mark - which is what an empty text box IS - so a document with a stray
' empty box came out of Remove_Txt_Bxs_And_Frames with a pair of markers and nothing between
' them. Jerry found it on 8/13/2026 and called it a phantom converted text box, which is
' exactly what it looked like.
'
' A shape that is not a text container at all (a picture, a line, a group) raises an error on
' TextFrame and answers False, which is also what we want.
'
    Dim t As String

    On Error Resume Next
    If shp.TextFrame.HasText Then t = shp.TextFrame.TextRange.Text
    Err.Clear
    On Error GoTo 0

    t = Replace(t, vbCr, "")
    t = Replace(t, Chr(11), "")      ' manual line break
    t = Replace(t, Chr(160), " ")    ' non-breaking space
    t = Replace(t, vbTab, " ")

    Sh_Shape_Has_Real_Text = (Len(Trim$(t)) > 0)

End Function   '*** end of Sh_Shape_Has_Real_Text ***

Public Function Sh_Frame_In_Range(ByVal fr As Frame, ByVal rng As Range) As Boolean
'
' Version: 1.0  Date: 8/13/2026
'
' A frame does sit in the text, so its own range answers the question. Same story guard as
' Sh_Shape_In_Range.
'
    Dim a As Range

    On Error Resume Next
    Set a = fr.Range
    If a Is Nothing Then Exit Function
    If a.StoryType <> rng.StoryType Then Exit Function
    ' STRICTLY less than rng.End. A selection of one paragraph ends at the START of the
    ' next one, so "<=" pulled in a frame anchored to the paragraph just past the
    ' selection - selecting one paragraph converted a frame belonging to the one below it.
    Sh_Frame_In_Range = (a.start >= rng.start And a.start < rng.End)
    Err.Clear

End Function   '*** end of Sh_Frame_In_Range ***

Sub Sh_Replace_Manual_Line_Break(Optional ByVal answer As String)
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Turns manual line breaks into paragraph marks or spaces, over the selection if there is one and
' the whole document if not. Shared by both sides.
'
' answer  "Para" or "Space" from a caller that already knows, and then NOTHING IS ASKED. Omit it
'         and the transcriber is asked.
'
'         That distinction is the point. The braille copy used to decide by whether text was
'         selected - ask if it was, silently use "Para" if it was not - and that quietly did the
'         right thing for the wrong reason: Full File Cleanup runs it with nothing selected, and
'         must not stop to ask. Making it always ask broke Full File Cleanup on the braille side
'         (Jerry, 8/12/2026). What actually separates the two cases is WHO IS CALLING: a sequence
'         knows the answer, a person is asked. So the sequence passes it.
'
' Version: 2.2  Date: 8/12/2026 - one macro again, with the answer as a parameter
' Version: 2.1  Date: 8/12/2026 - merged, then reverted when it interrupted Full File Cleanup
' Version: 2.0  Date: 8/12/2026 - no temporary document; the passes are scoped to a range. Gone
'                                with it: ActiveDocument.UndoClear, which threw away the WHOLE
'                                undo history rather than this macro's part of it, and on the
'                                braille side the Selection.MoveUp that stretched the selection up
'                                by a paragraph and the delete-one-character calls that tidied up
'                                after the paste. The Dx_Fix_Para_Space_Errors call is dropped
'                                from the large-print side - a braille macro called from there
'                                looked like a slip (Jerry).
' Version: 1.5  Date: 10/17/2023 - bug fix - no more leaving temp docs
' Version: 1.4  Date: 9/7/2021 - removed accidental delete of doc bug
' Version: 1.3  Date: 3/15/2021
' Version: 1.2  Date: 1/8/2019
' Version: 1.1  Date: 4/19/2016
'
    Dim rng As Range
    Dim su_Prev As Boolean

    If answer = "Para" Or answer = "Space" Then
        Sh_GP_String_1 = answer
    Else
        Sh_Space_Or_Para_Form.Show
        ' End, not Exit Sub, and deliberately: this is called from the cleanup sequences, and
        ' cancelling the question is meant to stop the whole sequence rather than let it carry on.
        If Sh_GP_String_1 <> "Para" And Sh_GP_String_1 <> "Space" Then
            End
        End If
    End If

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    If Selection.Type = wdSelectionNormal Then
        Set rng = Selection.Range
    Else
        Set rng = ActiveDocument.Content
    End If

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^l"
        If Sh_GP_String_1 = "Space" Then
            .Replacement.Text = " "
        Else
            .Replacement.Text = "^p"
        End If
        .Forward = True
        ' wdFindStop, not wdFindContinue: continue would run past the end of the range and
        ' quietly turn "the selection" into "the whole document".
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' Only for "Space": a line break turned into a space can land beside spaces already there.
    ' A paragraph mark cannot, so there is nothing to tidy in that case (Jerry).
    If Sh_GP_String_1 = "Space" Then Sh_Remove_Multi_Spaces rng

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub  '*** end of Sh_Replace_Manual_Line_Break ***

Sub Lp_Replace_Tabs_With_Single_Space()
'
' Lp_Replace_Tabs_With_Single_Space Macro
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Runs of tabs become one space, over the selection if there is one and the whole document if
' there is not.
'
' Version: 2.0  Date: 8/12/2026 - no temporary document, the same conversion as
'                                Sh_Remove_Multi_Spaces. Gone with it: ActiveDocument.UndoClear,
'                                which threw away the WHOLE undo history and not just this
'                                macro's part of it, and the trailing "delete one character at the
'                                end of the document", which existed to tidy up after the paste.
'                                NOT merged with the braille version - see the note there.
' Version: 1.3  Date: 10/17/2023 - minor bug fix
' Version: 1.2  Date: 1/8/2019
' Version: 1.1  Date: 1/7/2016
'
    Dim rng As Range
    Dim su_Prev As Boolean

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    If Selection.Type = wdSelectionNormal Then
        Set rng = Selection.Range
    Else
        Set rng = ActiveDocument.Content
    End If

    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^t{1,}"
        .Replacement.Text = "^032"
        .Forward = True
        ' wdFindStop, not wdFindContinue: continue would run past the end of the range and
        ' quietly turn "the selection" into "the whole document".
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    ' Tabs turned into spaces can sit next to spaces that were already there. Hand the range
    ' over rather than letting it work out its own scope again.
    Sh_Remove_Multi_Spaces rng

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub  '*** end of Lp_Replace_Tabs_With_Single_Space Macro ***

Sub Sh_Replace_Small_Caps_With_All_Caps()
'
' Sh_Replace_Small_Caps_With_All_Caps Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Small caps are a look, not letters - the text underneath is still lower case, and a screen
' reader or a braille translator reads it as lower case. This turns the look into real capitals.
'
' Version: 2.0 Date: 8/13/2026 - no temporary document, and ONE macro.
'
'                               The whole body of both copies was a single formatted
'                               find-and-replace. Everything else was round-trip plumbing, so
'                               once that came out the two were identical and merged.
'
'                               Gone with the round trip: the large-print "go to the end of the
'                               document and delete one character", which ran on EVERY route out
'                               and ate the last character of the book when nothing was selected;
'                               the braille "delete one character" before the paste back; and
'                               ActiveDocument.UndoClear on both sides, which threw away the
'                               whole undo history rather than this macro's part of it.
'
'                               The "turn small caps and all caps off" at the end now runs on a
'                               COLLAPSED selection, which is what it has always amounted to on
'                               the large-print side: it is there so the transcriber's next
'                               keystrokes are not in small caps, not to strip the capitals this
'                               macro has just applied.
' Version: 1.2 Date: 10/19/2023 - fixed removal of characters in selected text
' Version: 1.1 Date: 1/8/2019
' Version: 1.0 Date: 1/8/2016
'
    Dim rng As Range
    Dim su_Prev As Boolean

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    If Selection.Type = wdSelectionNormal Then
        Set rng = Selection.Range
    Else
        Set rng = ActiveDocument.Content
    End If

    With rng.Find
        .ClearFormatting
        .Font.SmallCaps = True
        .Font.AllCaps = False
        .Replacement.ClearFormatting
        .Replacement.Font.SmallCaps = False
        .Replacement.Font.AllCaps = True
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        ' wdFindStop, not wdFindContinue: continue would run past the end of the range and
        ' quietly turn "the selection" into "the whole document".
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position
    Selection.Collapse Direction:=wdCollapseStart

    ' so the next thing typed is not in small caps
    With Selection.Font
        .SmallCaps = False
        .AllCaps = False
    End With

    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

End Sub '***** end of Sh_Replace_Small_Caps_With_All_Caps Macro *****
Sub Lp_Replace_Section_Break_With_Page_Break()
'
' Lp_Replace_Section_Break_With_Page_Break Macro
'
' Version 1.3  Date: 1/8/2019
' Version 1.2  Date: 11/15/2018
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    Lp_Section_Brk_Caution.Show
    MsgBox "End of Macro", , "VistaType LP (129)"
    
End Sub  '*** end of Lp_Replace_Section_Break_With_Page_Break Macro ***
Sub Lp_Replace_Multiple_Para_Marks_With_Warning()
'
' Version: 2.0  Date: 8/3/2026 - shows the please-wait box while the replacement runs; placed AFTER the caution prompt, not around it
' Version: 1.9  Date: 5/7/2025 - altered message content
' Version: 1.8  Date:  1/8/2019
' Version: 1.7  Date: 12/27/2018
' Version: 1.6  Date: 2/14/2018

' Author: Jerry Whittaker - jerry@thewhittakers.org
'

    Application.Run MacroName:="Sh_Is_Doc_Open"
       
    If Selection.Type <> wdSelectionNormal Then  'there is no selected text

        Dim lngQuery As Long
        lngQuery = MsgBox("CAUTION!" & vbCr _
                      & vbCr & "Because no text range is selected, this macro will replace ALL" & vbCr _
                                  & "multiple paragraph marks in the document with a single" & vbCr _
                                  & "paragraph mark." & vbCr _
                       & vbCr & "If you have purposely added extra blank lines to your" & vbCr _
                                  & "document this macro will remove those lines!" & vbCr _
                       & vbCr & "To limit the replacement, select a text range" & vbCr _
                                  & "where paragraph marks are to be removed." & vbCr _
                       & vbCr & "Do you wish to continue?", vbYesNo + vbCritical + vbDefaultButton2, "VistaType LP (183)")
           
        If lngQuery = vbNo Then
            End
        End If
        
    End If
    
    ' The box goes HERE, not in the cleanup menu around this whole macro: the caution
    ' above asks "Do you wish to continue?", and a "Working - Please Wait" window sitting
    ' over a question the user has not answered yet would be nonsense. Jerry, 8/3/2026.
    Sh_Show_Please_Wait "Removing consecutive empty paragraph marks"
    Application.Run MacroName:="Sh_Replace_Multiple_Para_Marks_No_Warning"
    Sh_Hide_Please_Wait
    
End Sub    '***   end of  Lp_Replace_Multiple_Para_Marks_With_Warning macro ***
     
Sub Sh_Replace_Multiple_Para_Marks_No_Warning()
'
'  Version: 2.2  Date: 8/13/2026 - it removes EVERY empty paragraph in range, leaving none.
'
'                                 2.0 and 2.1 left one behind on the reading that "collapse a
'                                 run" meant "down to one blank line". It does not. The caution
'                                 in Sh_Replace_Multiple_Para_Marks_With_Warning has said what
'                                 this macro does since 2018 - replace multiple paragraph marks
'                                 with A SINGLE PARAGRAPH MARK, and "if you have purposely added
'                                 extra blank lines this macro will remove those lines". One
'                                 paragraph mark between two paragraphs is no blank line at all.
'                                 A lone blank line is two consecutive marks, so it goes too.
'  Version: 2.1  Date: 8/13/2026 - the blank test now also counts a paragraph holding only tabs
'                                 or only non-breaking spaces - see Sh_IsBlankParaMark 1.1
'  Version: 2.0  Date: 8/12/2026 - ONE macro, and it respects the selection.
'
'                                 The braille side was still doing "^013{2,} -> ^p" through a
'                                 temporary document. That pattern needs the paragraph marks to
'                                 be TOUCHING, so a line holding a single space breaks the run
'                                 and survives - which is what Jerry was seeing. The walk below
'                                 tests each paragraph with Trim$ instead, so a spaces-only line
'                                 counts as blank. Both sides use this now.
'
'                                 It also now works on the SELECTION when there is one. The walk
'                                 added in July went over the whole document however much was
'                                 selected - while the caution in
'                                 Sh_Replace_Multiple_Para_Marks_With_Warning was telling the
'                                 transcriber that selecting text is how you limit it. The
'                                 warning was right and the code was not.
'  Version: 1.8  Date: 8/3/2026 - the throttled DoEvents is now Sh_Spin_DoEvents, so it turns whichever progress box is showing
'  Version: 1.7  Date: 7/18/2026 - throttle DoEvents to every 200 paragraphs (was every one)
'  Version: 1.6  Date: 7/18/2026 - walk paragraphs via .Previous (linked) instead of
'                                  indexed paras(i); ~O(n) vs ~O(n^2) on large files.
'  Version: 1.5  Date: 7/2/2026 - complete rewrite
'
    Dim doc As Document
    Dim rng As Range
    Dim p As Paragraph
    Dim prevP As Paragraph
    Dim stopBefore As Long
    Dim deCount As Long

    Set doc = ActiveDocument

    If Selection.Type = wdSelectionNormal Then
        Set rng = Selection.Range
    Else
        Set rng = doc.Content
    End If
    stopBefore = rng.Start

    Set p = rng.Paragraphs.Last

    ' Walk backward via the linked .Previous so we never random-index the (slow) Paragraphs
    ' collection. Capture prevP BEFORE any delete: deleting p invalidates p, not prevP.
    ' Every empty paragraph goes, however many there are in a row and whether or not it has a
    ' blank neighbor. DoEvents only every 200 paragraphs.
    Do While Not (p Is Nothing)
        Set prevP = p.Previous          ' Nothing at the first paragraph of the document

        ' The document's own final paragraph mark cannot be removed - Word always keeps one -
        ' so leave it alone rather than ask and be refused.
        If p.Range.End < doc.Content.End Then
            If Sh_IsBlankParaMark(p) Then p.Range.Delete
        End If

        If Not (prevP Is Nothing) Then
            ' .Previous walks out of the selection quite happily, so stop at its edge.
            If prevP.Range.Start < stopBefore Then Exit Do
        End If

        Set p = prevP
        deCount = deCount + 1
        If deCount Mod 200 = 0 Then Sh_Spin_DoEvents
    Loop

End Sub   '***** Sh_Replace_Multiple_Para_Marks_No_Warning ********

Private Function Sh_IsBlankParaMark(p As Paragraph) As Boolean
    '
    ' Version: 1.1  Date: 8/13/2026 - a tab-only or non-breaking-space-only paragraph now counts
    '                                as empty. This is Jerry's own Lp_IsEmptyPara test, written
    '                                for exactly this and never wired to anything; that function
    '                                is gone and this is now the one blank test in the project.
    '                                The narrower Trim$-only version that used to be here was
    '                                preserving the behavior of the old indexed loop, and a line
    '                                holding one tab looks every bit as empty on the page as a
    '                                line holding one space.
    '
    ' NOT stripped, deliberately: the end-of-cell marker, Chr(7). The last paragraph of a table
    ' cell reads as vbCr & Chr(7), so leaving Chr(7) in place is what stops this returning True
    ' for an empty cell - and deleting that paragraph would damage the table.
    '
    Dim txt As String

    txt = p.Range.Text
    txt = Replace(txt, vbCr, "")      ' paragraph mark
    txt = Replace(txt, Chr(160), "")  ' non-breaking space
    txt = Replace(txt, Chr(9), "")    ' tab
    txt = Trim$(txt)                  ' ordinary spaces

    Sh_IsBlankParaMark = (Len(txt) = 0)
End Function   '***** Sh_IsBlankParaMark ********

Function Sh_Close_And_Reopen(ByVal targetDoc As Document) As Document
'
' Version: 1.1  Date: 8/20/2026 - never returns a document other than the one asked for. 1.0 fell
'                               back to ActiveDocument when the reopen raised, which on a machine
'                               with a second document open handed back THAT one - and the caller
'                               went on to activate it and announce the save as finished, leaving
'                               the transcriber looking at the wrong book. Also captures
'                               ConfirmConversions before arming the error handler, closes with
'                               screen updating off, and says so when the file comes back read-only
' Version: 1.0  Date: 8/20/2026
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Saves a document, closes it, and opens it again from disk. Returns the reopened Document, or
' NOTHING if it could not be reopened - the caller's old object variable is dead the moment the
' close happens, so use what this hands back and check it.
'
' WHY THIS EXISTS: Word does not repaint the Quick Style gallery - the row of style thumbnails
' on the Home tab - when the attached template changes which styles belong in it. Closing the
' file and opening it again is the only thing that does, and that is Jerry's own finding, made
' on 8/20/2026 after he put qFormat on the five List Bullet styles.
'
' Two other answers were tried first and are recorded so nobody spends the afternoon again:
'
'   The document is ALREADY CORRECT, so nothing about the styles needs fixing. Measured on the
'   build box straight after an attach, before any save: List Bullet through List Bullet 5 each
'   read QuickStyle = True and Visibility = False, exactly as the template says. The fault is
'   the ribbon, not the style sheet.
'
'   Toggling each gallery style QuickStyle False and back True, with screen updating on, does
'   NOT refresh it. That was 3.0.201 and Jerry tested it - the gallery stayed stale. Writing the
'   property is not what Word watches.
'
'   Document.UpdateStyles DOES refresh the gallery and MUST NEVER BE USED HERE. It works by
'   re-copying every style from the attached template, which throws away the base font and size
'   Lp_Attach_The_Template has just applied - and drags a book set in the dropped VistaTypeLP
'   Legible back to Tahoma, moving every page break in it.
'
' WHAT MAKES THE REOPEN ITSELF SAFE, and it is one line elsewhere: Lp_Attach_The_Template sets
' UpdateStylesOnOpen = False immediately after attaching, so the saved file does not re-copy
' styles from the template when it is opened. Without that, this reopen would do the very damage
' the paragraph above forbids. Do not "tidy" that line away.
'
' The open handler is deliberately NOT suppressed. Sh_Skip_Open_Handler exists for documents the
' add-in opens that are not the transcriber's work, like the GPL text; this one IS her work, and
' it needs exactly what any opened large print document gets. Reopening makes a NEW window, so
' the per-window settings the attach applied - view type, rulers, style area - are gone with the
' old one, and Sh_HandleDocumentOpened putting them back is the point rather than a side effect.
' It cannot loop: that handler only re-opens the attach dialog for an OBSOLETE template, and the
' document reaching here has the current one attached.
'
' Screen updating: OFF for the close, ON for the open. The open is what has to repaint, so it
' must be on by then. The close is different - Sh_HandleDocumentActivated treats screen updating
' being off as "a macro is running" and stands down, and with a second document open the close
' activates it and would otherwise run a full configuration switch for that document's type,
' rewriting the transcriber's AutoCorrect list on the way out and again on the way back.

    Dim docPath As String
    Dim cc_Prev As Boolean
    Dim su_Prev As Boolean
    Dim reopened As Document

    Set Sh_Close_And_Reopen = targetDoc

    ' Never been saved, so there is nothing on disk to reopen. Leave it exactly as it is - a
    ' stale gallery is a great deal better than losing the transcriber's document.
    If targetDoc Is Nothing Then Exit Function
    If Len(targetDoc.Path) = 0 Then Exit Function

    docPath = targetDoc.FullName

    ' Both captured BEFORE the error handler is armed. Read them after, and a raise inside the
    ' close sends Bail off to "restore" them from uninitialized variables - which for
    ' ConfirmConversions means switching OFF, application-wide and for good, a setting a
    ' transcriber working from raw publisher .doc files may have turned on deliberately.
    cc_Prev = Application.Options.ConfirmConversions
    su_Prev = Application.ScreenUpdating

    On Error GoTo Bail

    ' Saved explicitly rather than closing with wdDoNotSaveChanges. The caller has just been
    ' through the Save As, so this is normally a no-op - but if anything is ever added between
    ' that save and this call, discarding it silently is how work disappears.
    If Not targetDoc.Saved Then targetDoc.Save

    Application.ScreenUpdating = False
    targetDoc.Close SaveChanges:=wdDoNotSaveChanges
    Application.ScreenUpdating = True

    ' Same guard Sh_Show_Full_License uses: without it Word can stop on a "Convert File"
    ' question and the macro waits for a dialog nobody is looking at.
    Application.Options.ConfirmConversions = False
    Set reopened = Documents.Open(FileName:=docPath, AddToRecentFiles:=True)
    Application.Options.ConfirmConversions = cc_Prev
    On Error GoTo 0

    Set Sh_Close_And_Reopen = reopened

    ' Word writes a hidden ~$ owner file beside an open document and does not always release it
    ' at once. Reopening the same path milliseconds after closing it - on OneDrive, SharePoint or
    ' a network share - is the setup for coming back READ-ONLY. Say so. The alternative is a
    ' transcriber being told her file is saved and then typing into a book that will not take it.
    If Not reopened Is Nothing Then
        If reopened.ReadOnly Then
            MsgBox "Your document was saved, but Word has reopened it as READ-ONLY." & vbCr & vbCr _
                 & docPath & vbCr & vbCr _
                 & "Close it and open it again before you carry on editing. This usually means a " _
                 & "sync program - OneDrive or SharePoint - had not finished with the file.", _
                 vbExclamation, "VistaType LP (232)"
        End If
    End If
    Exit Function

Bail:
    ' The close or the open failed. NEVER answer with whatever document happens to be active -
    ' with a second book open that hands back the wrong one, and the caller activates it and
    ' announces the save as finished. Returning Nothing is the honest answer and the caller
    ' checks for it.
    On Error Resume Next
    Application.Options.ConfirmConversions = cc_Prev
    Application.ScreenUpdating = su_Prev
    On Error GoTo 0
    Set Sh_Close_And_Reopen = Nothing

End Function   '***** end of Sh_Close_And_Reopen macro *****

' The navigation pane has TWO halves and needs both written, which is not obvious and cost a
' build to find out. ActiveWindow.DocumentMap is a WINDOW property; CommandBars("Navigation") is
' application-wide. Writing DocumentMap alone works when a document is being CONFIGURED - the
' window is there and stays - but does nothing at all during DocumentBeforeClose, because the
' window it names is on its way out and Word's own remembered pane state is the thing still
' standing. Jerry, testing 3.0.207: the Styles pane went and the navigation pane stayed.
'
' Dx_Attach_BANA_Template has written both since 2021, with the On Error added in 2023 (2.3)
' because the CommandBars line raises on some Word versions. That pair was the answer all along.
'
' Version: 1.0  Date: 8/20/2026
Sub Sh_Tidy_Panes_After_Close()
'
' Version: 1.1  Date: 8/20/2026 - the Styles pane goes through Sh_Set_Styles_Pane, because
'                               Application.TaskPanes does nothing when Word has no document open
' Version: 1.0  Date: 8/20/2026
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Scheduled by Sh_HandleDocumentClosing one second after a close was ATTEMPTED, and the whole
' point of it is that by the time it runs the attempt has been settled. Word raises
' DocumentBeforeClose before the "save your changes?" prompt and says nothing afterwards about
' which button was pressed - so this does not try to know. It waits and looks.
'
' Application.OnTime cannot run while a modal dialog is up, which is what makes one second
' enough however long she takes to answer: the tick simply cannot fire until the prompt is gone.
' One second is also the floor - OnTime will not schedule closer than that.
'
' THE TEST IS "IS WORD EMPTY", and that is deliberate rather than lazy:
'
'   She cancelled the close, or closed one of two documents - Documents.count is not zero, and
'   this does NOTHING. A document is still in front of her, its panes are the ones she has been
'   working with, and a close that happened to some other document is not a reason to take them.
'   This is the case versions 1.3 to 1.5 got wrong by hiding first and trying to undo it after.
'
'   She closed the last document - Word is empty, and the panes are put down so the next document
'   she opens does not inherit them. That is what Jerry asked for on 8/20/2026, and it is also
'   where the panes are most obviously wrong: an empty Word with a Styles pane in it.
'
' Both panes go through a setter that writes TWO handles each - see Sh_Set_Styles_Pane and
' Sh_Set_Navigation_Pane. With no document open, Application.TaskPanes has nothing to hang a pane
' on and quietly does nothing; the CommandBars name for the same pane is application-wide and
' still answers. That caught the navigation pane in 3.0.207 and the Styles pane again in 3.0.209.
'
' Nothing here can raise into Word: a tick that fails is a pane left showing, which is untidy and
' nothing worse.
'
    On Error Resume Next

    ' Still working. Leave her screen exactly as it is.
    If Documents.count > 0 Then
        Err.Clear
        Exit Sub
    End If

    Sh_Set_Styles_Pane False
    Sh_Set_Navigation_Pane False

    Err.Clear

End Sub  '*** end of Sh_Tidy_Panes_After_Close ***

' Put the Styles pane up or down, from a place where there may be NO DOCUMENT OPEN.
'
' Everywhere else in this module the pane is set with Application.TaskPanes(wdTaskPaneFormatting),
' which is right and should stay - but it needs a document window to hang the pane on, and quietly
' does nothing when Word is empty. Jerry, testing 3.0.209: he closed the last document and the
' navigation pane went while the Styles pane stayed sitting there. Exactly the same shape as the
' navigation pane fault one build earlier, for exactly the same reason.
'
' CommandBars("Styles") is the other handle on the same pane and is application-wide, so it
' answers and can be written with nothing open - measured on the build box 8/20/2026, where with
' no document open CommandBars("Navigation") and CommandBars("Styles") were both still readable.
' Both are written here, in that order, because the CommandBar is the one that works when Word is
' empty and TaskPanes is the one that has always worked when it is not.
'
' Sh_Tidy_Panes_After_Close is the ONLY caller and the only place that needs this. Every other
' site sets the pane while configuring a document that is open in front of it, where TaskPanes
' alone is correct - do not sweep them into this.
'
' Version: 1.0  Date: 8/20/2026
Private Sub Sh_Set_Styles_Pane(ByVal wanted As Boolean)
    On Error Resume Next
    Application.TaskPanes(wdTaskPaneFormatting).Visible = wanted
    CommandBars("Styles").Visible = wanted
    Err.Clear
    On Error GoTo 0
End Sub  '*** end of Sh_Set_Styles_Pane ***

Private Sub Sh_Set_Navigation_Pane(ByVal wanted As Boolean)
    On Error Resume Next
    ActiveWindow.DocumentMap = wanted
    CommandBars("Navigation").Visible = wanted
    Err.Clear
    On Error GoTo 0
End Sub  '*** end of Sh_Set_Navigation_Pane ***

' Is the navigation pane showing? Asks the application-wide half first: it is the half that
' survives a document closing, and therefore the half worth putting back.
'
' Version: 1.0  Date: 8/20/2026
Private Function Sh_Navigation_Pane_Showing() As Boolean
    On Error Resume Next

    Sh_Navigation_Pane_Showing = CommandBars("Navigation").Visible
    If Err.Number <> 0 Then
        Err.Clear
        Sh_Navigation_Pane_Showing = ActiveWindow.DocumentMap
    End If

    Err.Clear
    On Error GoTo 0
End Function  '*** end of Sh_Navigation_Pane_Showing ***

Sub Lp_Turn_on_Styles_Pane()
    '
    ' Version: 1.2 Date:  10/26/2021 - added "Application.RestrictLinkedStyles = True"
    ' Version: 1.1 Date: 10/4/2018 -  Show recommended finally fixed
        
    Application.TaskPanes(wdTaskPaneFormatting).Visible = True 'turn on styles pane
    ActiveDocument.FormattingShowNextLevel = False
    ActiveDocument.StyleSortMethod = wdStyleSortRecommended
    ActiveDocument.FormattingShowFilter = wdShowFilterFormattingRecommended
    Application.RestrictLinkedStyles = True

End Sub   '***end of Lp_Turn_on_Styles_Pane Macro ***

Sub Lp_Type_Fill_In_Line()
'
' Lp_Type_Fill_In_Line Macro
'
' Author: Jerry Whittaker jerry@thewhittakers.org
'
' Version: 1.5  Date: 7/11/2025 - bug fix for blank lines on page break
' Version: 1.4  Date: 11/7/2023 - added Application.Run MacroName:="Sh_Is_Doc_Open" and "Lp_Is_Lp_Template_Attached"
' Version: 1.3  Date: 9/26/2023 - macro ends if LP template is not attached (no message on non-lp documents)
' Version: 1.2  Date: 12/28/2016
'
' Shows Fill-in menu
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    Application.Run MacroName:="Lp_Is_Lp_Template_Attached"
    
    Dim oStyle As Style
    Dim styleName As String
    
    styleName = "Box Black"
    Set oStyle = Nothing
    
    On Error Resume Next
    Set oStyle = ActiveDocument.Styles(styleName)
    
    If oStyle Is Nothing Then ' the style was not found
        End
    End If
    
    'Application.Run MacroName:="Lp_Is_Lp_Template_Attached"
    Lp_Type_Fill_In_Line_Form.Show
    Unload Lp_Type_Fill_In_Line_Form

    
End Sub   '*** end of Lp_Type_Fill_In_Line macro ***

Sub Lp_Format_Exercise_Lv_1_and_Lv_2()

' Lp_Format_Exercise_Lv_1_and_Lv_2 Macro
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Version: 1.6  Date: 7/26/2026 - returns the user to where the cursor was when the macro started
' Version: 1.4  Date: 3/20/2021 - added remove multiple spaces
' Version: 1.4  Date: 1/24/2021 - ajusted formatting
' Version: 1.3  Date: 2/11/2020 - permit LoopCounter - Z to undo
' Version: 1.1  Date: 2/14/2017
'
    Application.Run MacroName:="Lp_Is_Lp_Template_Attached"
    
    If Selection.Type <> wdSelectionNormal Then
        MsgBox "Select the exercise list first!", , "VistaType LP (133)"
        End
    End If

    Sh_Save_User_Position   ' record the spot HERE, before Lp_Copy_To_Temp_Doc makes the temp file active

    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating off
    Application.Run MacroName:="Lp_Copy_To_Temp_Doc"
    Application.ScreenUpdating = False ' Turn screen updating off
    Application.Run MacroName:="Lp_Convert_Auto_List_To_Text"
    Application.Run MacroName:="Lp_Fix_Para_Space_Errors"
    Application.Run MacroName:="Sh_Remove_Multi_Spaces"
    Application.ScreenUpdating = False ' Turn screen updating off
    Selection.WholeStory
    Selection.Style = ActiveDocument.Styles("Normal")
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    
    ' remove previous fill-ins
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "_{2,}"
        .Replacement.Text = "_"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

   ' replace space before underscore with tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032_^032"
        .Replacement.Text = "^032^t^032"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    'replace two tabs with tab-comma-tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t{2,}"
        .Replacement.Text = "^t^044^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineNone
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^095{2,}"
        .Replacement.Text = "^t^044^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' replace multiple underscores with single tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^095{1,}"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' replace space underscore followed with comma with tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032_^044"
        .Replacement.Text = "^032^t^044"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' replace underscore followed by a period with a tab
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032_^046"
        .Replacement.Text = "^032^t^046"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
   
    '*****************************************************
    ' remove mulitiple para marks
    '*****************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
     With Selection.Find
        .Text = "^013{2,}"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '*****************************************************
    ' double para marks
    '*****************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

     With Selection.Find
        .Text = "^013"
        .Replacement.Text = "^p^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' **************************************************************
    ' Set all para to List 2
    ' **************************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "List 2") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("List 2")
    With Selection.Find
        .Text = "^013"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If

    ' **************************************************************
    ' find para marks followed numbers followd by period and space
    ' **************************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "List") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("List")
     With Selection.Find
        .Text = "^013([0-9]{1,}^046^032)"
        .Replacement.Text = "^p\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    ' **************************************************************
    ' find para marks followed numbers followed by a space only
    ' **************************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "List") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("List")
     With Selection.Find
        .Text = "^013([0-9]{1,}^032)"
        .Replacement.Text = "^p\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If

    ' **************************************************************
    ' find para marks followed numbers followed by period
    ' **************************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "List") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("List")
     With Selection.Find
        .Text = "^013([0-9]{1,}^046)"
        .Replacement.Text = "^p\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If

    ' **************************************************************
    ' find para marks followed numbers followed by period
    ' **************************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "List") Then
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("List")
     With Selection.Find
        .Text = "^013([0-9]{1,})"
        .Replacement.Text = "^p\1"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
    
    '*****************************************************
    ' remove mulitiple para marks
    '*****************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
     With Selection.Find
        .Text = "^013{2,}"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '*****************************************************
    ' Convert Tabs to fill Ins
    '*****************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = " ________ "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '*****************************************************
    ' remove underlines from spaces
    '*****************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineNone
    With Selection.Find
        .Text = "^032"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    '*****************************************************
    ' remove spaces before commas
    '*****************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}^044"
        .Replacement.Text = "^044"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '*****************************************************
    ' place para mark before each "List" Style
    '*****************************************************
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "List") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("List")
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = "^013^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If

    Application.Run MacroName:="Lp_Fix_Para_Space_Errors"
    Application.Run MacroName:="Sh_Remove_Multi_Spaces"
    Application.Run MacroName:="Sh_Remove_Spaces_Before_Punctuation"

    '*****************************************************
    ' clean-up
    '*****************************************************
    Selection.HomeKey Unit:=wdStory
    Selection.MoveDown Unit:=wdLine, count:=2, Extend:=wdExtend
    If Sh_Style_Exists(ActiveDocument, "List") Then
    Selection.Style = ActiveDocument.Styles("List")
    End If
    
    Selection.HomeKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1
    
    Application.Run MacroName:="Lp_Copy_From_Temp_Doc"

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    'ActiveDocument.UndoClear ' No undo
    Application.ScreenUpdating = su_Prev ' Turn screen updating on

    Sh_Return_User_To_Start_Position

End Sub  '***** end of Lp_Format_Exercise_Lv_1_and_Lv_2 Macro *****
Sub Lp_Remove_Tabs_Before_and_After_Para_Marks()
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Date: 1/23/2017
' Version: 1.0
'
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating off

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^009{1,}^013"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "^013^009{1,}"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
   
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Application.ScreenUpdating = su_Prev ' Turn screen updating on
    
End Sub  '***** End of Lp_Remove_Tabs_Before_and_After_Para_Marks macro ********
Sub Lp_Remove_Tab_Plus_Space_Combos()
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Date: 1/23/2017
' Version: 1.0
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}^009{1,}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^009{1,}^032{1,}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub   '*** end of Lp_Remove_Tab_Plus_Space_Combos macro ***
Sub Lp_Validate_Dollar_PG()
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version: 1.6  Date: 11/7/2023 - added count of $pg
' Version: 1.5  Date: 1/3/2019 - removed requirment to have LP template attached
' Version: 1.4  Date: 9/27/2019 - Added Calls to Sh_Validation_Choices_Form Form
' Version: 1.3  Date: 9/25/2018 - added Application.Run MacroName:="Sh_Color_Dollar_PG_Red" and changed sendkeys sequence
' Version: 1.2  Date: 3/22/2018
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    
    'Count the tags
    Dim TagCounter As Integer
    TagCounter = 0
    Dim par As Word.Paragraph
    Dim txt As String

    For Each par In ActiveDocument.Paragraphs
        txt = par.Range.Text
        If InStr(txt, "$pg") > 0 Then
            TagCounter = TagCounter + 1
        End If
    Next
    
    If TagCounter = 0 Then
        MsgBox "There are no tagged page numbers in this document", , "VistaType LP (134)"
        End
    End If

    Application.Run MacroName:="Sh_Color_Dollar_PG_Red"
    Sh_Validation_Choices_Form.Show
    Unload Sh_Validation_Choices_Form
    
End Sub

' --- Recognizing a horizontal ordered list -------------------------------------------------
'
' "a) 149 B.C. b) 323 B.C. c) 44 A.D. d) 70 A.D." used to come apart at every period, because
' the passes in Lp_Horz_To_Vert_List_Form match a marker by what it looks like on its own - a
' letter or number, a period or bracket - and "B." looks exactly like "b." does. No amount of
' local pattern-matching can tell them apart; "a.Alpha b.Beta" is a real list written with no
' spaces, and it is the same shape as "B.C." (Jerry, 8/10/2026).
'
' What DOES tell them apart is the sequence. A list's markers share a style - the same kind of
' character, the same bracket or period, parenthesized or not - and they count up by one. So
' these routines collect every candidate marker, group them by style, and look for the longest
' run that ascends a, b, c or 1, 2, 3 or i, ii, iii. In the example above the a/b/c/d run is
' four long; the B/A "run" never gets past one, because the values go 2, 2, 1, 1. The list wins
' and the dates are left alone.
'
' Lp_Split_Ordered_List_Sequence returns False when it finds nothing convincing, and the form
' then falls back to the passes that were always there. It only ever takes over when it is sure.

' The value of a marker body: a-z as 1-26, a number as itself, i/ii/iii/iv... as 1-10. Zero
' means "not a marker".
'
' Version: 1.0  Date: 8/10/2026
Public Function Lp_Marker_Value(ByVal body As String) As Long
    Dim low As String
    low = LCase$(Trim$(body))

    If Len(low) = 0 Then Exit Function

    If IsNumeric(low) Then
        If Val(low) >= 1 And Val(low) <= 999 Then Lp_Marker_Value = Val(low)
        Exit Function
    End If

    If Len(low) = 1 And low >= "a" And low <= "z" Then
        ' A single letter is BOTH a letter marker and a roman numeral when it is i, v or x.
        ' Treated as a letter here; the roman reading is tried separately by the caller.
        Lp_Marker_Value = Asc(low) - Asc("a") + 1
        Exit Function
    End If
End Function

' The roman reading of a marker body, or zero. Kept separate from Lp_Marker_Value because "i"
' and "v" are both letters and numerals, and which one a list means only shows in the run.
'
' Version: 1.0  Date: 8/10/2026
Public Function Lp_Roman_Value(ByVal body As String) As Long
    Select Case LCase$(Trim$(body))
        Case "i": Lp_Roman_Value = 1
        Case "ii": Lp_Roman_Value = 2
        Case "iii": Lp_Roman_Value = 3
        Case "iv": Lp_Roman_Value = 4
        Case "v": Lp_Roman_Value = 5
        Case "vi": Lp_Roman_Value = 6
        Case "vii": Lp_Roman_Value = 7
        Case "viii": Lp_Roman_Value = 8
        Case "ix": Lp_Roman_Value = 9
        Case "x": Lp_Roman_Value = 10
    End Select
End Function

' Splits the active document's text at the markers of the best ordered-list run it can find.
' True when it split something, False when it found nothing convincing and the caller should
' carry on with the old passes.
'
' Version: 1.0  Date: 8/10/2026
' target   the text to work on. Nothing means the whole document.
' leadMark  True to leave a paragraph mark at the very start, which the caller that ends by
'           deleting the first character needs. The hidden-document route does not.
'
' Version: 1.1  Date: 8/12/2026 - takes a range
Public Function Lp_Split_Ordered_List_Sequence(Optional ByVal target As Range, _
                                              Optional ByVal leadMark As Boolean = True) As Boolean
    Const MAX_MARKERS As Long = 400
    Dim txt As String
    Dim i As Long, n As Long, j As Long, k As Long
    Dim ch As String
    Dim prevCh As String
    Dim body As String
    Dim delim As String
    Dim opened As Boolean
    Dim pos() As Long, val_() As Long, roman() As Long, style() As String
    Dim keep() As Boolean, bestKeep() As Boolean
    Dim bestScore As Long, sc As Long
    Dim sIdx As Long, r As Integer, v As Long
    Dim runStart As Long, runCount As Long, lastVal As Long
    Dim seenBefore As Boolean
    Dim base As Long
    Dim doc As Document

    On Error GoTo eom
    Lp_ListSplitError = ""
    Set doc = ActiveDocument
    If Not target Is Nothing Then Set doc = target.Document
    If target Is Nothing Then Set target = doc.Content
    base = target.Start
    txt = target.Text
    If Len(txt) = 0 Then Exit Function

    ReDim pos(1 To MAX_MARKERS)
    ReDim val_(1 To MAX_MARKERS)
    ReDim roman(1 To MAX_MARKERS)
    ReDim style(1 To MAX_MARKERS)
    n = 0
    bestScore = 0

    ' --- collect the candidates -----------------------------------------------------------
    For i = 1 To Len(txt)
        ' A marker starts the text, or follows a space or a paragraph mark. That single rule is
        ' what keeps "C." inside "B.C." from ever being considered.
        '
        ' The character before, read WITHOUT asking for character zero. VBA does not
        ' short-circuit Or - it evaluates every operand - so "i = 1 Or Mid$(txt, i - 1, 1) = ..."
        ' calls Mid$ with a start of 0 on the first character and raises error 5 every time.
        If i = 1 Then
            prevCh = " "
        Else
            prevCh = Mid$(txt, i - 1, 1)
        End If

        If prevCh = " " Or prevCh = vbCr Then
            j = i
            opened = False
            If Mid$(txt, j, 1) = "(" Then
                opened = True
                j = j + 1
            End If
            body = ""
            Do While j <= Len(txt) And Len(body) < 4
                ch = Mid$(txt, j, 1)
                If (ch >= "0" And ch <= "9") Or (LCase$(ch) >= "a" And LCase$(ch) <= "z") Then
                    body = body & ch
                    j = j + 1
                Else
                    Exit Do
                End If
            Loop
            If Len(body) > 0 And j <= Len(txt) Then
                delim = Mid$(txt, j, 1)
                If (delim = ")" Or (delim = "." And Not opened) Or (delim = ":" And Not opened)) Then
                    If Not (opened And delim <> ")") Then
                        If n < MAX_MARKERS Then
                            n = n + 1
                            pos(n) = i
                            val_(n) = Lp_Marker_Value(body)
                            roman(n) = Lp_Roman_Value(body)
                            ' The style a marker belongs to: parenthesized or not, which
                            ' delimiter, and whether it is a number, lower case or capitals. Two
                            ' markers only count as part of the same list when all three agree.
                            '
                            ' The LENGTH of the marker is deliberately NOT part of this. It was
                            ' at first, and it quietly broke every roman-numeral list - i, ii,
                            ' iii and iv are one, two, three and four characters long, so no two
                            ' of them were ever the same "style" and the run never got past one.
                            ' It would have done the same to 9 followed by 10.
                            style(n) = CStr(Abs(opened)) & delim & _
                                       IIf(IsNumeric(body), "9", IIf(body = LCase$(body), "a", "A"))
                        End If
                    End If
                End If
            End If
        End If
    Next i

    If n < 2 Then Exit Function
    ReDim keep(1 To n)
    ReDim bestKeep(1 To n)

    ' --- score every style, in both readings ----------------------------------------------
    '
    ' Not "the single longest run". A scanned page often carries the list across two adjacent
    ' paragraphs, and the numbering may RESTART in the second one - a, b then a, b again. Taking
    ' only the longest run split the first pair and left the second alone (Jerry, 8/11/2026).
    '
    ' So: pick the best STYLE, then split at every marker in that style which belongs to any run
    ' of two or more that ascends by one. A restarted run counts. A marker standing on its own
    ' does not, which is what keeps "B.C." and "J. R. R." out - their values never ascend, so
    ' every run they could form is one long and is thrown away.
    For sIdx = 1 To n
        seenBefore = False
        For j = 1 To sIdx - 1
            If style(j) = style(sIdx) Then seenBefore = True
        Next j

        If Not seenBefore Then
            For r = 0 To 1        ' 0 = letters and numbers, 1 = roman numerals
                For j = 1 To n
                    keep(j) = False
                Next j
                runStart = 0
                runCount = 0
                lastVal = 0

                For j = 1 To n
                    If style(j) = style(sIdx) Then
                        If r = 1 Then
                            v = roman(j)
                        Else
                            v = val_(j)
                        End If
                        If v > 0 Then
                            If runCount > 0 And v = lastVal + 1 Then
                                runCount = runCount + 1
                                keep(j) = True
                            Else
                                If runCount = 1 Then keep(runStart) = False   ' a run of one is not a list
                                runStart = j
                                runCount = 1
                                keep(j) = True
                            End If
                            lastVal = v
                        End If
                    End If
                Next j
                If runCount = 1 Then keep(runStart) = False

                sc = 0
                For j = 1 To n
                    If keep(j) Then sc = sc + 1
                Next j

                If sc > bestScore Then
                    bestScore = sc
                    For j = 1 To n
                        bestKeep(j) = keep(j)
                    Next j
                End If
            Next r
        End If
    Next sIdx

    If bestScore < 2 Then Exit Function

    ' --- split, working backwards so the earlier positions stay valid ---------------------
    For j = n To 1 Step -1
        If bestKeep(j) Then
            i = pos(j)
            If i > 1 Then
                If Mid$(txt, i - 1, 1) = " " Then
                    ' the space becomes the line break rather than being left dangling
                    doc.Range(base + i - 2, base + i - 1).Text = vbCr
                ElseIf Mid$(txt, i - 1, 1) <> vbCr Then
                    doc.Range(base + i - 1, base + i - 1).InsertBefore vbCr
                End If
            End If
        End If
    Next j

    ' A paragraph mark at the very start, ALWAYS - even when the list begins part-way down and
    ' the first paragraph is a lead-in of ordinary prose. The passes this function stands in for
    ' always left one there, and the form's last two lines - go to the top, delete one character
    ' - exist to take it away again. Leave it out and that Delete eats the first real character
    ' instead: "A. Grant's Tomb" came back as ".Grant's Tomb", and a lead-in line lost its first
    ' letter (Jerry, 8/10 and 8/11/2026).
    If leadMark Then
        If Left$(doc.Range(base, base + 1).Text, 1) <> vbCr Then doc.Range(base, base).InsertBefore vbCr
    End If

    Lp_Split_Ordered_List_Sequence = True
    Exit Function

eom:
    Lp_ListSplitError = "err " & Err.Number & ": " & Err.Description & " (markers found: " & n & ")"
End Function

' Diagnostic breadcrumb: why Lp_Split_Ordered_List_Sequence gave up. Empty when it did not.
Public Function Lp_Get_List_Split_Error() As String
    Lp_Get_List_Split_Error = Lp_ListSplitError
End Function

' --- Horizontal list to vertical, in a temporary document that never appears -----------------
'
' Same architecture as the round trip this project has always used - copy the selection out,
' work on it, put it back - and the same single undo, because the text goes home in ONE
' assignment. What is different is that the temporary document is never shown.
'
' Lp_Copy_To_Temp_Doc creates it hidden and then deliberately gives it a window, makes the
' window visible, maximizes it and activates it. That is the flashing, and it is four
' statements, not an accident. It has to do that because its callers work through Selection,
' and Selection only reaches the document that is active.
'
' So the passes here work on a RANGE instead. The range is the temporary document's own
' content, which holds nothing but the selection, so there is no region to track and nothing
' of the transcriber's to damage - the two things that made the in-place attempt of 8/11/2026
' risky. No custom undo record either: one FormattedText assignment is already one undo step.
'
' The temporary document is built from the SOURCE document's own attached template, so styles
' resolve the same way on both sides. That is why this needs no separate braille route: a BANA
' document gets a BANA temporary document.
'
' Version: 1.0  Date: 8/12/2026
Public Sub Lp_Horz_To_Vert_Hidden(ByVal src As Range, ByVal listKind As String, _
                                  ByVal sortWanted As Boolean)
    Dim origDoc As Document
    Dim tempDoc As Document
    Dim tpl As String
    Dim su_Prev As Boolean
    Dim body As Range
    Dim work As Range
    Dim srcTrail As Long

    If src Is Nothing Then Exit Sub
    Set origDoc = src.Document

    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    On Error Resume Next
    tpl = origDoc.AttachedTemplate.FullName
    On Error GoTo 0

    ' How many paragraph marks the selection ends with. They have to travel WITH the text and
    ' come back in the same number, and both halves of that matter:
    '
    '   * a paragraph's formatting - its indentation among it - lives IN its paragraph mark.
    '     Leaving the last mark behind cost the last item its indentation (Jerry, 8/12/2026).
    '   * the cleanup passes collapse runs of marks, so the text can come back ending with
    '     fewer than it left with. That deleted the mark after the last item and welded it onto
    '     the paragraph below ("A. Jack Nicklaus ... D. Ben Hogan", Jerry, 8/12/2026).
    '
    ' So: send everything, then put the count back before handing it home.
    srcTrail = 0
    Do While Len(src.Text) > srcTrail
        If Mid$(src.Text, Len(src.Text) - srcTrail, 1) = vbCr Then
            srcTrail = srcTrail + 1
        Else
            Exit Do
        End If
    Loop
    Set work = src

    On Error GoTo eom
    If Len(tpl) > 0 Then
        Set tempDoc = Documents.Add(Template:=tpl, Visible:=False)
    Else
        Set tempDoc = Documents.Add(Visible:=False)
    End If

    ' NOTHING here shows, maximizes or activates the temporary document. That is the whole point.
    tempDoc.Content.FormattedText = work.FormattedText

    Lp_Hv_Passes tempDoc, listKind, sortWanted

    ' Put the trailing mark count back to what the selection had - see above.
    Lp_Hv_Set_Trailing_Marks tempDoc, srcTrail

    ' Home in one assignment, which Word records as a single undo step. The temporary
    ' document's own final paragraph mark is never part of the text and is left behind.
    Set body = tempDoc.Range(0, tempDoc.Content.End - 1)
    work.FormattedText = body.FormattedText

    tempDoc.Close SaveChanges:=wdDoNotSaveChanges
    Application.ScreenUpdating = su_Prev
    Exit Sub

eom:
    On Error Resume Next
    If Not tempDoc Is Nothing Then tempDoc.Close SaveChanges:=wdDoNotSaveChanges
    Application.ScreenUpdating = su_Prev
End Sub

' Every pass, against the temporary document's whole content. listKind is "ORDERED", "SPACED"
' or "TABBED".
'
' Version: 1.0  Date: 8/12/2026
Private Sub Lp_Hv_Passes(ByVal doc As Document, ByVal listKind As String, _
                         ByVal sortWanted As Boolean)
    Lp_Hv_Repl doc, "^l", "^p"                 ' manual line break becomes a paragraph mark
    Lp_Hv_Repl doc, "^0149^032*", "^p^&"       ' ordinary bullet followed by a space
    Lp_Hv_Repl doc, "^046{2,}", "^046"         ' multiple periods
    Lp_Hv_Repl doc, "^032{2,}", "^032"         ' multiple spaces
    Lp_Hv_Repl doc, "^032^046", "^046"         ' space before a period

    Select Case listKind
    Case "ORDERED"
        If Not Lp_Split_Ordered_List_Sequence(doc.Content, True) Then
            Lp_Hv_Repl doc, "[A-z0-9]{1,}^046^032", "^p^&"
            Lp_Hv_Repl doc, "\(([A-z0-9]{1,})\)", "~\1" & ChrW(338)
            Lp_Hv_Repl doc, "(~[A-z0-9]{1,}" & ChrW(338) & "^032)", "^p^&"
            Lp_Hv_Repl doc, "(~[A-z0-9]{1,}" & ChrW(338) & "^046)", "^p^&"
            Lp_Hv_Repl doc, "^0149^032*", "^p^&"
            Lp_Hv_Repl doc, "([A-z0-9]{1,}^046)", "^p\1^032"
            Lp_Hv_Repl doc, "([A-z0-9]{1,}\)^046)", "^p\1^032"
            Lp_Hv_Repl doc, "([A-z0-9]{1,}\))", "^p\1^032"
        End If

    Case "SPACED"
        Lp_Hv_Repl doc, "^032{1,}", "^p"

    Case Else   ' TABBED
        Lp_Hv_Repl doc, "^032{2,}", "^032"
        Lp_Hv_Repl doc, "^009^032{1,}", "^009"
        Lp_Hv_Repl doc, "^032{1,}^009", "^009"
        Lp_Hv_Repl doc, "^009{2,}", "^009"
        Lp_Hv_Repl doc, "^009{1,}", "^p"
    End Select

    Lp_Hv_Repl doc, "^032{1,}", "^032"
    Lp_Hv_Repl doc, "~", "("
    Lp_Hv_Repl doc, ChrW(338), ")"
    Lp_Hv_Repl doc, "^032{1,}^013", "^p"
    Lp_Hv_Repl doc, "^013^032{1,}", "^p"
    Lp_Hv_Repl doc, "^032{1,}^046", "^046"
    Lp_Hv_Repl doc, "^013{2,}", "^p"

    If sortWanted Then Lp_Hv_Sort doc

    ' The passes leave a spare paragraph mark at the front, which is what the old route removed
    ' by going to the top and deleting one character. AFTER the sort, not before: an empty
    ' paragraph sorts above everything, so trimming first only let it come back to the top.
    On Error Resume Next
    Do While Left$(doc.Content.Text, 1) = vbCr
        doc.Range(0, 1).Delete
    Loop
    On Error GoTo 0
End Sub

' Make the temporary document's text end with exactly `wanted` paragraph marks, not counting
' the document's own final one, which is never part of the text.
'
' Version: 1.0  Date: 8/12/2026
Private Sub Lp_Hv_Set_Trailing_Marks(ByVal doc As Document, ByVal wanted As Long)
    Dim have As Long
    Dim body As String
    Dim guard As Long

    On Error GoTo eom
    Do
        body = doc.Range(0, doc.Content.End - 1).Text
        have = 0
        Do While Len(body) > have
            If Mid$(body, Len(body) - have, 1) = vbCr Then
                have = have + 1
            Else
                Exit Do
            End If
        Loop

        If have = wanted Then Exit Do
        If have > wanted Then
            doc.Range(doc.Content.End - 2, doc.Content.End - 1).Delete
        Else
            doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertBefore vbCr
        End If

        guard = guard + 1
    Loop While guard < 20
eom:
End Sub

' Version: 1.0  Date: 8/12/2026
Private Sub Lp_Hv_Repl(ByVal doc As Document, ByVal findText As String, ByVal replText As String)
    On Error GoTo eom
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
eom:
End Sub

' A selection Word cannot sort used to close the temporary document and say so. Here the list is
' already converted and correct - only the sort was refused - so it is left as it is and the
' reason goes to the breadcrumb rather than into a dialog.
'
' Version: 1.0  Date: 8/12/2026
Private Sub Lp_Hv_Sort(ByVal doc As Document)
    On Error GoTo SortMsg
    doc.Content.Sort ExcludeHeader:=False, FieldNumber:="Paragraphs", _
        SortFieldType:=wdSortFieldAlphanumeric, SortOrder:=wdSortOrderAscending, _
        CaseSensitive:=False, LanguageID:=wdEnglishUS
    Exit Sub
SortMsg:
    Lp_ListSplitError = "sort refused: " & Err.Description
End Sub

' --- One horizontal-to-vertical macro, two kinds of document -------------------------------
'
' The two sides copy the selection into a temporary document in genuinely different ways, and
' the difference matters:
'
'   Lp_Copy_To_Temp_Doc   builds the temp document FROM LargePrintTemplate.dotx, and stops with
'                         "Template not found" when it is missing
'   Dx_Copy_To_Temp_Doc   plain new document, then re-attaches the SAME BANA template and keeps
'                         the document's BrailleType
'
' Send a braille file through the large-print route and three things go wrong: it fails outright
' on a machine with no large-print template, the text is round-tripped through a document
' carrying LP styles, and the BANA template and BrailleType are not put back. The last one would
' not show up until much later.
'
' So the route is chosen from the document being worked on. Sh_Doc_Config_Type answers that
' already. The choice is REMEMBERED, because by the time the text comes back the active document
' is the temporary one and asking again would give the wrong answer.
'
' Version: 1.0  Date: 8/11/2026
Public Sub Sh_Copy_To_Temp_Doc()
    If Sh_Doc_Config_Type() = "BRL" Then
        Sh_TempDocRoute = "BRL"
        Application.Run MacroName:="Dx_Copy_To_Temp_Doc"
    Else
        Sh_TempDocRoute = "LP"
        Application.Run MacroName:="Lp_Copy_To_Temp_Doc"
    End If
End Sub

' Version: 1.0  Date: 8/11/2026
Public Sub Sh_Copy_From_Temp_Doc()
    If Sh_TempDocRoute = "BRL" Then
        Application.Run MacroName:="Dx_Copy_From_Temp_Doc"
    Else
        Application.Run MacroName:="Lp_Copy_From_Temp_Doc"
    End If
    Sh_TempDocRoute = ""
End Sub

Sub Lp_Horz_List_To_Vertical()
    '
    ' Author: Jerry Whittaker -  jerry@thewhittakers.org
    '
    ' Version: 1.4: Date: 8/12/2026 - the temporary document is no longer shown, so no flashing
    ' Version: 1.3: Date: 8/11/2026 - now serves BOTH ribbon tabs; Dx_Horz_List_To_Vertical and its form are gone
    ' Version: 1.2: Date: 9/20/2018 - added optional manual selection of text (before execution) or automatic selection of current para
    ' Version: 1.1: Date: 9/10/2018 - added automatic paragraph selection
    ' Version: 1.0: Date: 3/13/2018
    '
    ' Shared by BOTH ribbon tabs from 8/11/2026 - see the note where Dx_Horz_List_To_Vertical
    ' used to be. No template check on either side: this macro does not care what is attached.
    Application.Run MacroName:="Sh_Is_Doc_Open"

    If Selection.Type <> wdSelectionNormal Then
        Selection.Paragraphs(1).Range.Select
    End If

    Application.Run MacroName:="Lp_Is_Text_Selected"

    Lp_Hv_Go = False
    Lp_Horz_To_Vert_List_Form.Show
    Unload Lp_Horz_To_Vert_List_Form
    If Not Lp_Hv_Go Then Exit Sub

    ' The document is touched HERE, with the dialog gone. No custom undo record: the text goes
    ' home in one FormattedText assignment, which Word already records as a single undo step.
    Lp_Horz_To_Vert_Hidden Selection.Range, Lp_Hv_Kind, Lp_Hv_Sort_Wanted
    Lp_Hv_Go = False

End Sub  '*** end of Lp_Horz_List_To_Vertical macro ***

Sub Lp_Table_Tools()
'
' Version: 1.7  Date: 10/17/2025 - changed selection msg content
' Version: 1.6  Date: 8/12/2025 - added Lp_GP_String_2 = ActiveDocument.FullName
' Version: 1.5  Date: 8/4/2025 - added  Application.Run MacroName:="Lp_ValidateTableIntegrityForListOrRotation" and
'                                       Application.Run MacroName:="Lp_DoesRangeHaveATOCStyle"
' Version: 1.4  Date: 6/24/2025 - added save index of slected table
' Version: 1.3  Date: 5/28/2025 - made menu choice made on what is selected in document (Table or TOC Range)
' Version: 1.2  Date: 11/7/2023 -  Added"Sh_Is_Doc_Open" and "Lp_Is_Lp_Template_Attached"
' Version: 1.1  Date: 9/26/2023 - macro ends if LP template is not attached (no message on non-lp documents)
' Version: 1.0  Date: 3/21/2018
'
    Dim i As Integer
    Dim tbl As Table

    Application.Run MacroName:="Sh_Is_Doc_Open"
    Application.Run MacroName:="Lp_Is_Lp_Template_Attached"

    If (ActiveDocument.Tables.count = 0 Or Not Selection.Information(wdWithInTable)) And Not Selection.Range.Paragraphs.count > 1 Then
        MsgBox "Select a TABLE (or place cursor in a table) or select a range containing a TOC. Selected TOC range may include embedded non-TOC Styles.", , "VistaType LP (172)"
        End
    End If

    If Selection.Information(wdWithInTable) Then
        Lp_GP_String_2 = ActiveDocument.fullName
        Application.Run MacroName:="Lp_ValidateTableIntegrityForListOrRotation"
        For i = 1 To ActiveDocument.Tables.count
            Set tbl = ActiveDocument.Tables(i)
            'save the index of the selected table in the original doc
            If Selection.Range.start >= tbl.Range.start And Selection.Range.End <= tbl.Range.End Then
                Lp_GP_Counter_1 = i
                Exit For
            End If
        Next i
        Lp_Table_Tools_Menu_Form.Show
        Unload Lp_Table_Tools_Menu_Form
    End If

    If Selection.Range.Paragraphs.count > 1 Then 'something is here
        Lp_TOC_Format_And_Color_Form.Show
    End If
    End
End Sub

Sub Lp_Fix_Abbyy_Text_and_Headers()
'
' Macro Lp_Fix_Abbyy_Text_and_Headers()
'
' Word can have up to nine heading levels.  In Large print there are only five Heading levels.
' This macro searches for up to nine Abbyy 14 headers styles (i.e. Heading #1 etc)
'   and converts them to standard heading styles (without the # sign). Headings 6 through
'   9 are set to heading level 5.
'
' This macro also changes all "Normal" and "Body Text" styles to the large print normal style.
'
' This macro ONLY works with Abbyy FineReader documents saved as "Formatted Text".
' The macro should have no effect on any other documents.
'
' Author: Jerry Whittaker   jerry@thewhittakers.org
'
' Version:  1.6  Date: 10/20/21 - removed destructive table style code
' Version:  1.5  Date: 7/9/2021 - added existance check and message for LargePrintTemplate.dotx
' Version:  1.4  Date: 11/23/2020 - changes to accomodate single template file
'
'
    Dim doc As Document
    Dim para As Paragraph
    Dim StyleCntr As Integer
    Dim styleName As String
    Dim NewStyleName As String
    Dim FontSize As Integer
    Set doc = ActiveDocument
    Dim ReloadTemplateSwitch As Boolean
    
    ' Convert Abbyy 14 header styles to LP header styles
    For Each para In doc.Paragraphs
        For StyleCntr = 1 To 9  'up to nine types of Heading styles
            styleName = "Heading #" + LTrim(Str(StyleCntr))
            NewStyleName = "Heading " + LTrim(Str(StyleCntr))
            If StyleCntr > 5 Then
               NewStyleName = "Heading 5"
            End If
            If para.Style = styleName Then
                ReloadTemplateSwitch = True 'An Abbyy 14 style was changed
                'para.Range.Font.Reset
                para.Range.Style = ActiveDocument.Styles(NewStyleName)
                StyleCntr = 10
            End If
        Next StyleCntr

        For StyleCntr = 1 To 9  'up to nine types of body text styles
            styleName = "Body text (" + LTrim(Str(StyleCntr)) + ")"
            If para.Style = styleName Then
                ReloadTemplateSwitch = True 'An Abbyy 14 style was changed
                para.Range.Style = ActiveDocument.Styles("Normal")
                StyleCntr = 10
            End If
        Next StyleCntr

        'change Abbyy's "Normal" and "Other" styles to LP Normal
        If para.Style = "Other" Then
            ReloadTemplateSwitch = True 'An Abbyy 14 style was changed
            para.Range.Style = ActiveDocument.Styles("Normal")
        End If
        
    Next para
    
End Sub   '*** end of Lp_Fix_Abbyy_Text_and_Headers ***

Sub Lp_File_Fix_Sequence()
'
' Lp_File_Fix_Sequence macro
'
' presents menu to format all or selecte table(s)
'
' Author:   Jerry Whittaker
'           jerry@thewhittakers.org
'
' Version: 1.0  Date: 5/22/2018
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    Unload Lp_File_Cleanup_Sub_Menu_Form
    Lp_File_Cleanup_Sub_Menu_Form.Show
    Unload Lp_File_Cleanup_Sub_Menu_Form
    Application.ScreenUpdating = True    ' Turn screen updating on
    Application.ScreenRefresh
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    
End Sub  '*** end of Lp_File_Fix_Sequence Macro ***

' The character immediately BEFORE the cursor, and the one immediately AFTER it. Both peek by
' moving the selection and then putting it back, which is the only way to read them - and which
' is why anything that types afterwards must set its own underlining: moving the cursor throws
' the pending character formatting away and picks it up afresh from whatever is beside it.
'
' "" means there is nothing there - the start or end of the document.
'
' Version: 1.0  Date: 8/10/2026
Public Function Lp_Char_Before_Cursor() As String
    On Error GoTo eom
    Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
    Lp_Char_Before_Cursor = Selection.Text
    Selection.MoveRight Unit:=wdCharacter, count:=1
eom:
End Function

' Version: 1.0  Date: 8/10/2026
Public Function Lp_Char_After_Cursor() As String
    On Error GoTo eom
    Selection.MoveRight Unit:=wdCharacter, count:=1, Extend:=wdExtend
    Lp_Char_After_Cursor = Selection.Text
    Selection.MoveLeft Unit:=wdCharacter, count:=1
eom:
End Function

' True when nothing PRECEDES the cursor on this line, so a space put in front of the fill-in line
' would be wrong: the start of a paragraph, the start of the document, just after a manual line
' break or a page break, the start of a table cell - or a space that is already there. A fill-in
' line at the start of a paragraph must sit hard against the left margin (Jerry, 8/10/2026).
'
' Version: 1.0  Date: 8/10/2026
Public Function Lp_Nothing_Precedes_On_Line() As Boolean
    Select Case Lp_Char_Before_Cursor()
        Case " ", "", vbCr, Chr(11), Chr(12), Chr(7)
            Lp_Nothing_Precedes_On_Line = True
    End Select
End Function

' The character at a position, or "" when that position is outside the document. Keeps the
' scanning below free of bounds arithmetic.
'
' Version: 1.0  Date: 8/10/2026
Private Function Lp_Char_At(doc As Document, ByVal pos As Long) As String
    If pos < 0 Then Exit Function
    If pos > doc.Content.End - 1 Then Exit Function
    Lp_Char_At = doc.Range(pos, pos + 1).Text
End Function

' True when the character at pos is part of a fill-in line: an underscore, or the manual line
' break that a fill to the right margin puts BETWEEN its lines. The break only counts when it has
' underscores on both sides - otherwise this would swallow a line break of the transcriber's that
' merely happens to sit next to a fill.
'
' Version: 1.0  Date: 8/10/2026
Private Function Lp_Is_Fill_Char(doc As Document, ByVal pos As Long) As Boolean
    Select Case Lp_Char_At(doc, pos)
        Case "_"
            Lp_Is_Fill_Char = True
        Case Chr(11)
            Lp_Is_Fill_Char = (Lp_Char_At(doc, pos - 1) = "_") And (Lp_Char_At(doc, pos + 1) = "_")
    End Select
End Function

' Deletes the fill-in line the cursor is touching, so that choosing a length REPLACES the line
' instead of adding to it. Without this a transcriber who made a line too short had to
' double-click it and delete, or backspace it away a character at a time, or copy and paste part
' of it - a guessing game, because how long a line LOOKS depends on the point size (Jerry,
' 8/10/2026).
'
' Scans out from the cursor in both directions for as long as it is still inside the fill, so it
' handles a fill to the right margin with extra lines too - that one carries a manual line break
' between each of its lines, and the whole thing goes.
'
' Minds the spaces: taking the underscores out of "Name: ____ is here." would leave two spaces
' where one belongs, so a doubled space at the join is collapsed. The spaces the new line needs
' are then put in by the macro that draws it, under its own rules, exactly as if the line were
' being typed for the first time.
'
' Does nothing at all unless an underscore is actually touching the cursor, so it can be called
' unconditionally.
'
' Version: 1.0  Date: 8/10/2026
Public Sub Lp_Remove_Fill_In_Line_At_Cursor()
    Dim doc As Document
    Dim startPos As Long
    Dim endPos As Long

    On Error GoTo eom
    Set doc = ActiveDocument

    If Selection.Start <> Selection.End Then Exit Sub   ' a real selection - not ours to touch
    startPos = Selection.Start
    endPos = startPos

    If Not (Lp_Is_Fill_Char(doc, startPos - 1) Or Lp_Is_Fill_Char(doc, startPos)) Then Exit Sub

    Do While Lp_Is_Fill_Char(doc, startPos - 1)
        startPos = startPos - 1
    Loop
    Do While Lp_Is_Fill_Char(doc, endPos)
        endPos = endPos + 1
    Loop

    doc.Range(startPos, endPos).Delete

    ' one space, not two, where the line used to be
    If (Lp_Char_At(doc, startPos - 1) = " ") And (Lp_Char_At(doc, startPos) = " ") Then
        doc.Range(startPos, startPos + 1).Delete
    End If

    Selection.SetRange startPos, startPos

eom:
End Sub

' True when the character AT the cursor sits on a different line from the one before it - which,
' called with the cursor between a fill-in line and the punctuation that follows it, means the
' punctuation has dropped to the next line and the fill needs to give a character back.
'
' Version: 1.0  Date: 8/10/2026
Private Function Lp_Next_Char_Wrapped() As Boolean
    Dim p As Long
    Dim doc As Document

    On Error GoTo eom
    Set doc = ActiveDocument
    p = Selection.Start
    If p < 1 Then Exit Function
    If p > doc.Content.End - 2 Then Exit Function
    If doc.Range(p - 1, p).Text <> "_" Then Exit Function   ' nothing of ours left to give back

    Lp_Next_Char_Wrapped = _
        (doc.Range(p, p + 1).Information(wdFirstCharacterLineNumber) <> _
         doc.Range(p - 1, p).Information(wdFirstCharacterLineNumber))
eom:
End Function

' True when a space belongs after a fill-in line. Answers No in three cases:
'
'   * nothing follows on the line - a paragraph mark, a manual line break, a page break, the end
'     of a table cell, the end of the document. A space there would show as nothing at all, and
'     on a line that already reaches the right margin it would push the wrap
'   * a space is already there
'   * the next character is punctuation that has to sit hard against what comes before it. A
'     fill-in line ending a sentence must read "____." and never "____ ." (Jerry, 8/10/2026)
'
' Quotation marks are deliberately NOT in the list: a closing one wants no space in front and an
' opening one does, and in straight-quote form the two are the same character.
'
' Version: 1.0  Date: 8/10/2026
Public Function Lp_Space_Wanted_After_Fill() As Boolean
    Select Case Lp_Char_After_Cursor()
        Case " ", "", vbCr, Chr(11), Chr(12), Chr(7)
            ' nothing follows on this line, or a space is there already
        Case Else
            Lp_Space_Wanted_After_Fill = Not Lp_Punctuation_Follows()
    End Select
End Function

' True when the very next character is one of LP_TIGHT_PUNCTUATION.
'
' Version: 1.0  Date: 8/10/2026
Public Function Lp_Punctuation_Follows() As Boolean
    Dim c As String
    c = Lp_Char_After_Cursor()
    If Len(c) = 1 Then Lp_Punctuation_Follows = (InStr(LP_TIGHT_PUNCTUATION, c) > 0)
End Function

' True when nothing follows the cursor on this line, so a space put there would be invisible:
' a paragraph mark, a manual line break, a page break, the end of a table cell, the end of the
' document - or a space that is already there.
'
' Version: 1.0  Date: 8/10/2026
Public Function Lp_Nothing_Follows_On_Line() As Boolean
    Select Case Lp_Char_After_Cursor()
        Case " ", "", vbCr, Chr(11), Chr(12), Chr(7)
            Lp_Nothing_Follows_On_Line = True
    End Select
End Function

Sub Lp_Type_Fill_In_Line_To_Margin()

    ' Called from: Lp_Type_Fill_In_Form
    '
    ' Version 2.2  Date: 8/10/2026 - proved WHY the last line has to be one underscore shorter when
    '                                punctuation follows: the two are one unbreakable word, so if they
    '                                do not fit Word moves the whole fill to the next line rather than
    '                                wrapping the period alone. Making the room conditional was tried
    '                                and stranded "Name:" on a line by itself. The comment now says so
    ' Version 2.1  Date: 8/10/2026 - the manual line break between the first and second lines is no
    '                                longer underlined. Word draws an underline for an underlined
    '                                break and runs it to the right margin, which made that line look
    '                                longer than the ones below it (Jerry)
    ' Version 2.0  Date: 8/10/2026 - the cursor sitting in a fill-in line now REPLACES it. See
    '                                Lp_Remove_Fill_In_Line_At_Cursor
    ' Version 1.9  Date: 8/10/2026 - never a space before a period, or before any of , ; : ! ? ) ] }
    '                                - a fill-in line that ends a sentence reads "____." (Jerry)
    ' Version 1.8  Date: 8/10/2026 - the fill now reaches the right margin when the cursor is in the
    '                                MIDDLE of a sentence. The loop stops when the cursor's line number
    '                                changes, and with text after the cursor that happened too early -
    '                                the following word wrapped, and the insertion point sitting in
    '                                front of it was reported on the new line. The shortfall was the
    '                                width of that word: 72 underscores with nothing after, 64 and 56
    '                                points of white space when "Whittaker." followed. The rest of the
    '                                sentence is now parked behind a paragraph mark for the duration
    '                                and put back straight afterwards (Jerry, 8/10/2026)
    ' Version 1.7  Date: 8/10/2026 - no second space when one is already in front of the cursor, and
    '                                a space after the fill when text follows it, so the underscores no
    '                                longer run straight into the next word. The trailing one matters
    '                                for more than looks: underscores and a word with no space between
    '                                them are ONE unbreakable run to Word, and an unbreakable run that
    '                                will not fit gets moved to the next line whole. Both mirror what
    '                                Lp_Type_Counted_Fill_In_Lines already did (Jerry, 8/10/2026)
    ' Version 1.6  Date: 5/30/2025 - fixed blank line error at top of new page - additional lines from fill to right margin now have no manual line feed
    ' Version 1.5: Date: 12/2/23 - removed version 1.4 fix
    ' Version 1.4: Date: 10/21/2021 - Removed manual line break at margin
    ' Version 1.3: Date: 8/1/2019  added code to assure that lines are not bold
    ' Version 1.2: Date: 5/30/2019  fixed fill to right margin problem when cursor is at right margin
    ' Version 1.1: Date: 1/10/2019
    ' Version 1.0: Date: 1/3/2019
    '
    ' Author: Jerry Whittaker  jerry@thewhittakers.org

    Dim CurrentLine As Integer
    Dim NextLine As Integer
    Dim StartLine As Integer
    Dim NoOfXtraLinesWanted As Integer
    Dim LineCounter As Integer
    Dim Loop_Cntr As Integer
    Dim CurrentColumn As Integer
    Dim CurrentPageNumber As Integer
    Dim Stubline As Boolean
    Dim tempSplit As Boolean
    Dim punctFollows As Boolean
    Dim splitPos As Long
    Dim endPos As Long
    
    NoOfXtraLinesWanted = Lp_GP_Counter_1
    LineCounter = 0
    
    CurrentPageNumber = Selection.Information(wdActiveEndPageNumber)

    If Selection.Type = wdSelectionNormal Then
        Selection.Delete Unit:=wdCharacter, count:=1
    End If

    ' Sitting in a fill-in line means REPLACE it, not add to it.
    Lp_Remove_Fill_In_Line_At_Cursor

    CurrentLine = Selection.Range.Information(wdFirstCharacterLineNumber) ' get the line number where the cursor is located
    StartLine = CurrentLine
    NextLine = CurrentLine + 1
    CurrentColumn = Selection.Range.Information(wdFirstCharacterColumnNumber)
    
    If Selection.Font.Bold = True Then
        Selection.Font.Bold = wdToggle  ' turn off bold
    End If

    ' Get the rest of the sentence out of the way while the line is being filled.
    '
    ' The loop below stops when the CURSOR's line number changes. With text after the cursor
    ' that happens too early: as the underscores grow, the following word stops fitting and
    ' wraps to the next line, and the insertion point - which sits immediately before that word -
    ' is then reported as being on the next line too. The loop reads that as "the underscores
    ' wrapped" and stops. Measured on the build box 8/10/2026: a fill with nothing after it took
    ' 72 underscores, the same fill followed by "Whittaker." took 64 and left 56 points of white
    ' space. The shortfall is the width of whatever word follows, which is why it looked like
    ' five or six characters (Jerry, 8/10/2026).
    '
    ' A paragraph mark parked after the cursor removes the question: nothing follows on the line,
    ' so nothing can wrap, and the fill behaves exactly as it already did at the end of a
    ' paragraph. It is taken out again below, before the trailing space is decided.
    punctFollows = Lp_Punctuation_Follows()   ' asked BEFORE the split, while the sentence is intact

    tempSplit = False
    If Not Lp_Nothing_Follows_On_Line() Then
        splitPos = Selection.Start
        ActiveDocument.Range(splitPos, splitPos).InsertAfter vbCr
        Selection.SetRange splitPos, splitPos
        tempSplit = True
    End If

    If NoOfXtraLinesWanted = 0 Then  ' no extra lines - just the fill to margin
        ' A space in front only when there is something for it to separate the fill from. Not at
        ' the start of a paragraph, and not when a space is already there - see
        ' Lp_Nothing_Precedes_On_Line.
        If Not Lp_Nothing_Precedes_On_Line() Then
            Selection.TypeText Text:=" "  ' non-underline space (if cursor is on last column then the following line (type x) will jump to next line
            Selection.TypeText Text:="x"  ' type the "x" to see if it is not
            If NextLine = Selection.Range.Information(wdFirstCharacterLineNumber) Then  ' has the cursor move to the next line
                CurrentLine = NextLine
                StartLine = CurrentLine  ' set the new line as the current line
                Selection.TypeBackspace  ' remove "x"
                Selection.TypeText Text:=Chr(11)  ' manual line break
            End If
                Selection.TypeBackspace  ' remove "x"
        End If
        
        With Selection.Font
            .Underline = wdUnderlineSingle  ' turn underline on
        End With
        
        Do While CurrentLine = StartLine
            Selection.TypeText Chr(95)   '  underscore
            CurrentLine = Selection.Range.Information(wdFirstCharacterLineNumber) ' get the new line number where the cursor is located
        Loop
        
        Selection.TypeBackspace
        
    Else  ' Extra lines wanted
    
        NoOfXtraLinesWanted = NoOfXtraLinesWanted + 1
        
        If CurrentColumn = Selection.Range.Information(wdFirstCharacterColumnNumber) Then
            Stubline = True
        End If
        
        CurrentColumn = Selection.Range.Information(wdFirstCharacterColumnNumber)
        
        If Not Lp_Nothing_Precedes_On_Line() Then
            With Selection.Font
                .Underline = wdUnderlineNone
            End With
            If Selection.Font.Bold = True Then
                Selection.Font.Bold = wdToggle  ' turn off bold
            End If
            Selection.TypeText Text:=" "  ' non-underline space
        End If

        Do While NoOfXtraLinesWanted > LineCounter
        
            If Selection.Information(wdActiveEndPageNumber) <> CurrentPageNumber Then
                CurrentPageNumber = Selection.Information(wdActiveEndPageNumber) ' reset the page number
                StartLine = 1
            Else
                CurrentLine = Selection.Range.Information(wdFirstCharacterLineNumber) ' get the line number where the cursor is located
                StartLine = CurrentLine
            End If
            
            With Selection.Font
                    .Underline = wdUnderlineSingle  ' turn underline on
            End With
            
            Do While CurrentLine = StartLine
                Selection.TypeText Chr(95)   '  underscore
                CurrentLine = Selection.Range.Information(wdFirstCharacterLineNumber) ' get the new line number where the cursor is located
            Loop

            LineCounter = LineCounter + 1
                             
            If NoOfXtraLinesWanted = LineCounter Then
                Selection.TypeBackspace
            Else
                Selection.TypeBackspace
                If Stubline Then
                    ' NOT underlined. Word draws an underline for an underlined line break and
                    ' runs it out to the right margin, so this line LOOKED longer than the ones
                    ' below it even though it carries the same number of underscores (Jerry,
                    ' 8/10/2026). The break is part of the layout, not part of the fill. The top
                    ' of this loop switches underlining back on for the next line.
                    With Selection.Font
                        .Underline = wdUnderlineNone
                    End With
                    Selection.TypeText Text:=Chr(11) ' manual line break
                    Stubline = False
                Else
                    Selection.TypeText Text:=Chr(95) ' an underscore
                End If
            End If
        Loop
    End If

    With Selection.Font
        .Underline = wdUnderlineNone  ' turn underline off
    End With

    ' One underscore back when punctuation follows, and it has to happen HERE, while the sentence
    ' is still parked - not after the rejoin.
    '
    ' The fill stops up to a character short of the margin already, so it is tempting to ask
    ' afterwards whether the period actually needs the room. That does not work: the period has no
    ' space in front of it, so the underscores and the period are ONE unbreakable word to Word. If
    ' the pair does not fit, Word does not wrap the period on its own - it picks the whole run up
    ' and moves it to the next line, leaving "Name:" stranded on a line by itself. Tried on
    ' 8/10/2026 and that is exactly what happened. Because they move together, "did the period
    ' wrap?" can never see the failure.
    '
    ' So the room is made in advance. The last line ends one underscore shorter than the lines
    ' above it and the period stands in that place. A period is narrower than an underscore, so
    ' the line does finish a fraction shorter than its neighbours - about a third of a character.
    ' That is the cost of keeping the period on the line, and it is the smaller of the two evils.
    If punctFollows Then Selection.TypeBackspace

    ' Put the sentence back together - see the note above. Only ever removes the mark this macro
    ' put there: it has to be the very next character, and it has to be a paragraph mark.
    If tempSplit Then
        endPos = Selection.End
        If ActiveDocument.Range(endPos, endPos + 1).Text = vbCr Then
            ActiveDocument.Range(endPos, endPos + 1).Delete
            Selection.SetRange endPos, endPos
        End If
    End If

    ' And a space after the fill when text follows it. Underlining is turned off HERE, after the
    ' peek inside Lp_Nothing_Follows_On_Line, not before it: moving the cursor to look at the next
    ' character throws the pending formatting away and takes it from what is beside it, which is
    ' an underscore, and underlined.
    If Lp_Space_Wanted_After_Fill() Then
        With Selection.Font
            .Underline = wdUnderlineNone
        End With
        If Selection.Font.Bold = True Then
            Selection.Font.Bold = wdToggle  ' turn off bold
        End If
        Selection.TypeText Text:=" "  ' non underlined space
    End If

    ' Underlining OFF as the very last thing, whatever route got here. Every peek at a
    ' neighbouring character moves the cursor, and moving it takes the pending formatting from
    ' what is beside it - an underscore, underlined. Turning it off any earlier means the
    ' transcriber's next keystroke comes out underlined (Jerry, 8/10/2026).
    Selection.Font.Underline = wdUnderlineNone

End Sub   '*** end of Lp_Type_Fill_In_Line_To_Margin macro ***


Sub Lp_Type_Counted_Fill_In_Lines()

    ' Called from: Lp_Type_Fill_In_Form
    '
    ' Version 1.6:  Date: 8/10/2026 the cursor sitting in a fill-in line now REPLACES it. See
    '                               Lp_Remove_Fill_In_Line_At_Cursor
    ' Version 1.5:  Date: 8/10/2026 never a space before a period, or before any of , ; : ! ? ) ] }
    '                               - a fill-in line that ends a sentence reads "____." (Jerry). Also
    '                               no space in FRONT when the fill starts a paragraph
    ' Version 1.4:  Date: 8/9/2026  that trailing space is no longer underlined. Peeking at the
    '                               next character moves the cursor, and moving it re-picks up the
    '                               formatting beside it - an underscore, underlined. Underlining
    '                               is now turned off AFTER the peek, immediately before the space
    '                               is typed, which is the order the leading space already used
    ' Version 1.3:  Date: 8/9/2026  a fill-in line dropped into the middle of a line of text now
    '                               ends with a space when text follows it, the mirror of the space
    '                               already put in front. Without it the underscores ran straight
    '                               into the next word (Jerry, 8/9/2026). Nothing is added when a
    '                               space is already there, or when the line ends after the
    '                               underscores - a paragraph mark, a manual line break, the end of
    '                               a table cell or the end of the document - because a trailing
    '                               space at the end of a line is invisible, and it would still
    '                               push a line that already reaches the margin onto the next one
    ' Version 1.2:  Date: 8/1/2019  added code to assure that lines are not bold
    ' Version 1.1:  Date: 1/10/2019
    ' Version 1.0:  Date: 12/13/2018
    '
    ' Author: Jerry Whittaker  jerry@thewhittakers.org
    
    Dim Chr_Cntr As Integer
    Dim Loop_Cntr As Integer

    Chr_Cntr = Lp_GP_Counter_1
    Loop_Cntr = 1

    If Selection.Type = wdSelectionNormal Then
        Selection.Delete Unit:=wdCharacter, count:=1
    End If

    ' Sitting in a fill-in line means REPLACE it, not add to it.
    Lp_Remove_Fill_In_Line_At_Cursor

    ' A space in front only when there is something for it to separate the fill from. The test
    ' used to be "is the character before a space", which said yes to a paragraph mark as well -
    ' so a fill-in line starting a paragraph was pushed off the left margin by a space nobody
    ' asked for (Jerry, 8/10/2026). Lp_Nothing_Precedes_On_Line is now the one rule, shared with
    ' Lp_Type_Fill_In_Line_To_Margin.
    If Not Lp_Nothing_Precedes_On_Line() Then
        With Selection.Font
            .Underline = wdUnderlineNone
        End With
        If Selection.Font.Bold = True Then
            Selection.Font.Bold = wdToggle  ' turn off bold
        End If
        Selection.TypeText Text:=" "  ' non underlined space
    End If
    
    With Selection.Font
        .Underline = wdUnderlineSingle  ' turn on underlineing
    End With

    Do
        Selection.TypeText Text:=Chr(95)    ' underscore
        Loop_Cntr = Loop_Cntr + 1
        If Loop_Cntr > Chr_Cntr Then Exit Do
    Loop
    
    With Selection.Font
        .Underline = wdUnderlineNone  ' turn off underlineing
    End With

    ' A space after the line as well, when the transcriber has put the fill-in line in the middle
    ' of a line of text. Same look as the space in front of it, and the same test read the other
    ' way round: peek at the character that FOLLOWS, then put the cursor back.
    ' Underlining OFF after the peek inside Lp_Nothing_Follows_On_Line, not before it. Moving
    ' the cursor to look at the next character throws away the pending character formatting and
    ' picks it up afresh from what is beside it - which is an underscore, and underlined. The
    ' space came out underlined for exactly that reason (Jerry, 8/9/2026).
    If Lp_Space_Wanted_After_Fill() Then
        With Selection.Font
            .Underline = wdUnderlineNone
        End With
        If Selection.Font.Bold = True Then
            Selection.Font.Bold = wdToggle  ' turn off bold
        End If
        Selection.TypeText Text:=" "  ' non underlined space
    End If

    ' Underlining OFF as the very last thing - see the same note in Lp_Type_Fill_In_Line_To_Margin.
    Selection.Font.Underline = wdUnderlineNone

    'Unload Lp_Type_Fill_In_Line_Form

End Sub   '*** end of Lp_Type_Counted_Fill_In_Lines macro ***

Sub Lp_Set_Display_For_Large_Print()

    ' turn on show/hide
    If Not ActiveWindow.ActivePane.View.ShowAll Then  ' show all is not active
       ActiveWindow.ActivePane.View.ShowAll = Not ActiveWindow.ActivePane.View.ShowAll
    End If
    
    ActiveDocument.FormattingShowFont = False
    ActiveDocument.FormattingShowParagraph = False
    ActiveDocument.FormattingShowNumbering = False
    ActiveWindow.ActivePane.View.Type = wdPrintView
    'Application.Options.ShowCropMarks = True
    ActiveWindow.View.ShowAll = True
    ActiveWindow.DisplayRulers = True
    ActiveWindow.DisplayVerticalRuler = True
    ' 8/2/2026 - the Styles pane is not opened or re-sorted HERE. This sub runs on every LP
    ' document open, and it was applying the Recommended sort and filter three times over (once
    ' itself, once through Lp_Turn_on_Styles_Pane, once more through the
    ' MS_Set_Word_Config_For_Large_Print call below), wiping out whatever the user had chosen.
    ' The sort and the filter are still hers. 8/20/2026 - the pane's VISIBILITY is a different
    ' question and it did come back: MS_Set_Word_Config_For_Large_Print, one line below, opens
    ' the pane for a large print document. Do not add a second copy of that here.
    ' RestrictLinkedStyles went with the removed Lp_Turn_on_Styles_Pane call but is still set,
    ' one line later, by MS_Set_Word_Config_For_Large_Print.
    ActiveDocument.FormattingShowNextLevel = False
    Application.Run MacroName:="MS_Set_Word_Config_For_Large_Print"
    
End Sub

Sub Lp_Is_Lp_Template_Attached()
'
' Version: 1.0  Date: 12/9/2018
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    If Lp_Is_The_Attached_Template_LP = False Then
        MsgBox "VistaType large print template is not attached", , "VistaType LP (135)"
        End
    End If
    
End Sub   '*** end of Lp_Is_Lp_Template_Attached macro ***

Function Lp_Is_The_Attached_Template_LP() As Boolean
'
' Version: 2.0  Date: 8/20/2026 - asks the ATTACHED TEMPLATE, which is what this function has
'                               always been named after. Was: does the style "Box Black" exist
' Version: 1.1  Date: 2/15/2026 - Changed from Print Pg Num to Box Black
' Version: 1.0  Date: 4/15/2021
'
' True when the document in front of us is a large print document - the template named
' LP_TEMPLATE_FILE is attached to it.
'
' READ THIS BEFORE CHANGING A CALLER. Until 8/20/2026 this tested for the presence of the style
' "Box Black" instead, and answered a subtly different question: not "is this a large print
' document" but "was this document ever MADE as a large print book". Jerry, 8/20/2026: the style
' test was written before he knew how to test for an attached template, and testing for the
' template is the better idea. The name was right all along; the implementation was the stopgap.
'
' The two questions are not the same, and the difference is a real document: a book made on an
' OBSOLETE large print template still has Box Black in it, but the template attached to it is
' not ours. That book must be recognized, warned about and re-attached - which is what
' Lp_Was_Made_As_An_Lp_Book, immediately below, now exists for. Three callers wanted that
' question and were moved to it; the other five wanted this one and were left alone.
'
' Compared on the NAME only. AttachedTemplate returns a Template object whose default property
' is its full path, so .Name is taken explicitly rather than letting VBA coerce the object -
' the folder differs between machines and between Office versions and must not be part of the
' test. Compared case-insensitively: Word reports the name as the file system has it.
'
    Dim attachedName As String

    ' No document open, or a document Word will not answer for. Not a large print document, and
    ' not a fault either - about a dozen callers ask this before knowing what is on screen.
    On Error GoTo eom

    attachedName = ActiveDocument.AttachedTemplate.Name
    Lp_Is_The_Attached_Template_LP = (StrComp(attachedName, LP_TEMPLATE_FILE, vbTextCompare) = 0)

eom:
    ' Whatever raised is this function's normal "no" answer, not something the caller caused.
    Err.Clear

End Function '*** end of Lp_Is_The_Attached_Template_LP Function ***

Function Lp_Was_Made_As_An_Lp_Book() As Boolean
'
' Version: 1.0  Date: 8/20/2026 - the body of Lp_Is_The_Attached_Template_LP 1.1, kept for the
'                               one question it is actually right for
'
' True when this document was ever made as a large print book, whatever template is attached to
' it NOW. The test is the presence of the style "Box Black", which comes from the large print
' template and is not something a transcriber would have by accident.
'
' This is deliberately the LOOSER of the two tests, and that is the whole point of it. Ask
' Lp_Is_The_Attached_Template_LP which configuration a document gets; ask this one whether a
' document needs the large print template re-attached, because the answer is yes precisely when
' this is True and that one is False - a book on an obsolete template, or on one Word reset to
' Normal when the file was mailed or the template was moved. Test it the other way round and
' such a book reads as an ordinary document, gets the ordinary configuration, and is never
' warned that it must be re-attached before she edits it.
'
' Three callers: Sh_HandleDocumentOpened (the obsolete-template warning), Lp_Attach_Lp_Template
' (whether to warn about re-attaching), and LP_Attach_An_Lp_Template_Form's Initialize (noting
' the face a legacy book is already set in). All three are about the book, not the template.
'
    Dim oStyle As Style

    Set oStyle = Nothing

    On Error Resume Next
    Set oStyle = ActiveDocument.Styles("Box Black")   ' was "Print Pg Num" before 2/15/2026

    If Not oStyle Is Nothing Then Lp_Was_Made_As_An_Lp_Book = True

    ' Styles(...) raises 5941 when the style is absent, which is this function's normal "no"
    ' answer, not a fault. Clear it so the caller does not inherit an error it never caused.
    Err.Clear

End Function '*** end of Lp_Was_Made_As_An_Lp_Book Function ***

Sub Lp_Para_To_Next_Page()
    Application.Run MacroName:="Sh_Is_Doc_Open"
    Selection.ParagraphFormat.PageBreakBefore = wdToggle
End Sub

Sub Lp_Compress_Linear_Math()
'
'  Version: 1.2  Date: 2/27/2003 - set spacing using "hare space"
'  Version: 1.1  Date: 2/22/2020 - no need to be in large print
'  Version: 1.0  Date: ???
'
'  Macro suggestion by Katherine Thomison and Shelley Mack
'
'  Purpose: Removes all spaces from selected text and places spaces before and after signs of comparison
        '  = equal
        '  approximately equal (double tilda)
        '  <> not equal
        '  Not equal (slashed equal sign)
        '  > greater than
        '  < less than
        '  greater than or equal (underscrored greater than)
        '  less than or equal (underscored less than)
        '  underscore less than sign (less than or equal to)
        '  slashed equal sign (not equal to)
        '  double tilda - approximately equal to
'
'  Author: Jerry Whittaker   jerry@thewhittakers.org
'
'  Version: 1.5  Date:  12/2/2019 - and 'End' after user respons to yes/no question - added length check for expression
'  Version: 1.4  Date:  10/7/2019 - Added code to bypass shortcut key when LP Template not attached.
'  Version: 1.2  Date:  1/13/2019 - added additional symbols
'  Version: 1.1  Date:  1/27/2019 - keep spaces before and after signs of comparison
'  Version: 1.0  Date:  1/26/2019
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    Application.Run MacroName:="Lp_Is_Lp_Template_Attached"

    Sh_Save_User_Position

    If Selection.Type <> wdSelectionNormal Then ' no text was select prior to running the macro
        Selection.Paragraphs(1).Range.Select
    End If

    If Len(Selection) > 50 Then
        If MsgBox("You may have selected more text than just a math expression." & vbCr _
                & vbCr & "This macro will remove ALL spaces from the selected text." & vbCr & vbCr _
                & "Select only the math expression, or if the math expression all that is in the paragraph (with no other text)," _
                & " simply place the cursor in the paragraph and run the macro again." & vbCr & vbCr _
                & "Do you wish to continue with the compression?", vbYesNo, "VistaType LP (184)") = vbYes Then
            GoTo CompressThis:
         Else
            Selection.Collapse 'Direction:=wdCollapseStart
            End
        End If
    End If
    
    If MsgBox("Compress this math expression?", vbYesNo, "VistaType LP (185)") = vbYes Then
        GoTo CompressThis:
    Else
        Selection.Collapse 'Direction:=wdCollapseStart
        End
    End If
    
CompressThis:
    Dim HS As String
    HS = ChrW(8202)  'Hair Space or unicode 8202

    ' set the find and replace parameters
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}"    ' the text to be found - one or more spaces (^032 is code for a space)
                                        ' the {1,} means find one or more - more efficient code
        .Replacement.Text = ""  ' the replacement text is nothing
        .Forward = True
        .Wrap = wdFindStop ' will not ask if you want to search the rest of document
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True  ' is a wildcard search because of the {1,} in the find text
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll  ' do the replacement
    
    'convert normal "x" (multiply) with math x
'    Selection.Find.ClearFormatting
'    Selection.Find.Replacement.ClearFormatting
'    With Selection.Find
'        .Text = "x" 'alphabet "x"
'        .Replacement.Text = HS + "×" + HS 'unicode 00D7 = math x
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .MatchCase = False
'        .MatchWholeWord = False
'        .MatchWildcards = False
'        .MatchSoundsLike = False
'        .MatchAllWordForms = False
'    End With
'    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "=" ' this finds an equal sign
        .Replacement.Text = HS + "^&" + HS
        .Wrap = wdFindStop
    End With
   Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\>" ' this finds the greater than sign
        .Replacement.Text = HS + "^&" + HS
        .Wrap = wdFindStop
    End With
   Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\<" ' this finds the less than sign
        .Replacement.Text = HS + "^&" + HS
        .Wrap = wdFindStop
    End With
   Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\< \>" ' this finds the not-equal sign
        .Replacement.Text = HS + "^&" + HS
        .Wrap = wdFindStop
    End With
   Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8805) 'unicode 2265 for underscored greater than sign (greater than or equal to)
        .Replacement.Text = HS + "^&" + HS
        .Wrap = wdFindStop
    End With
   Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8804) 'unicode 2264 for underscore less than sign (less than or equal to)
        .Replacement.Text = HS + "^&" + HS
        .Wrap = wdFindStop
    End With
   Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8800) 'unicode 2260 for slashed equal sign (not equal to)
        .Replacement.Text = HS + "^&" + HS
        .Wrap = wdFindStop
    End With
   Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8776) ' double tilda - approximately equal to
        .Replacement.Text = HS + "^&" + HS
        .Wrap = wdFindStop
    End With
   Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " \<  \> "    ' double spaces between <> (fix for '<  >' problems created above)
        .Replacement.Text = HS + "<>" + HS
        .Wrap = wdFindStop
    End With
   Selection.Find.Execute Replace:=wdReplaceAll
   
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"

    Sh_Return_User_To_Start_Position

End Sub   ' end of Lp_Compress_Linear_Math macro ***

Sub Lp_Add_Para_After_Image()

    ' place a para mark following each image - following text is sometime part of the image paragraph,
    '     fixes common problem with DAISY and NIMAS Files
    '
    '  Author: Jerry Whittaker   jerry@thewhittakers.org
    '  Version: 1.3  Date: 8/5/2026 - does nothing once the LP template is attached (Jerry)
    '  Version: 1.2  Date: 1/22/2026 - full rewrite
    '  Version: 1.1  Date: 4/9/2024 - complete rewrite - Much faster
    '  Version: 1.0  Date: 5/14/2019

    ' This repairs raw DAISY and NIMAS files, where text that belongs after an image is stuck in
    ' the image's own paragraph. That is a BEFORE-the-template job. Once the LP template is on,
    ' the images have been placed and sized and the paragraphs around them are styled, so adding
    ' more paragraph marks only damages a document that is already right.
    '
    ' The guard has to live HERE rather than at the caller. Lp_Fix_Common_File_Errors skips this
    ' whole macro during the attach sequence, but File Cleanup on the LP ribbon runs the same
    ' sequence on demand, and by then the template usually IS attached.
    If Lp_Is_The_Attached_Template_LP = True Then Exit Sub

    Dim ils As inlineShape
    Dim rng As Range

    ' Loop through inline images from last to first
    ' (backwards prevents range shifting issues)
    Dim i As Long
    For i = ActiveDocument.InlineShapes.count To 1 Step -1

        Set ils = ActiveDocument.InlineShapes(i)

        ' Create a range immediately after the image
        Set rng = ils.Range.Duplicate
        rng.Collapse Direction:=wdCollapseEnd

        ' If the next character is NOT a paragraph mark, insert one
        If rng.Characters.count > 0 Then
            If rng.Characters(1).Text <> vbCr Then
                rng.InsertAfter vbCr
            End If
        Else
            ' Image is at the very end of the document
            rng.InsertAfter vbCr
        End If

    Next i

End Sub   '*** end of Lp_Add_Para_After_Image macro ***

Sub Lp_Get_Doc_Setup_Params()

    ' get font and page settings for current document and place in public variables
    '
    ' Version: 1.6  Date: 8/8/2026 - also reads the typeface back off the Normal style, into
    '                               Lp_Base_Font_Name, the same way the size has always been read
    ' Version: 1.5  Date: 12/11/2020 - minor fixt to PPG
    ' Version: 1.4  Date: 12/3/2020 - fixed null DM - Set to "Unknown"
    ' Version: 1.3  Date: 12/1/2020 - added method to get gutter size from doc xml
    ' Version: 1.2  Date: 11/24/2020 - added orientation and Tab Settings
    ' Version: 1.1  Date: 10/15/2020 - new method for determining Lp_Base_Font_Size
    ' Version: 1.0  Date: 4/29/2020
    
    PTM = ActiveDocument.PageSetup.TopMargin / Application.InchesToPoints(1)
    PBM = ActiveDocument.PageSetup.BottomMargin / Application.InchesToPoints(1)
    PLM = ActiveDocument.PageSetup.LeftMargin / Application.InchesToPoints(1)
    PRM = ActiveDocument.PageSetup.RightMargin / Application.InchesToPoints(1)
    PPH = Round(PointsToInches(ActiveDocument.PageSetup.PageHeight), 2)
    PPW = Round(PointsToInches(ActiveDocument.PageSetup.PageWidth), 2)
    PMM = ActiveDocument.PageSetup.MirrorMargins  ' Zero = not mirrored
    Lp_Base_Font_Size = ActiveDocument.Styles(wdStyleNormal).Font.Size
    ' Same idea as the size on the line above: the document IS the setting, so there is nothing
    ' to store and nothing to go stale. A book made before the typeface could be chosen answers
    ' Tahoma here, correctly, because that is what its Normal style inherited from the template.
    Lp_Base_Font_Name = ActiveDocument.Styles(wdStyleNormal).Font.Name
    If Len(Trim(Lp_Base_Font_Name)) = 0 Then Lp_Base_Font_Name = LP_FONT_TAHOMA
    TOCTabSetting = Str(Val(PPW) - (Val(PLM) + Val(PRM)))
 
    If ActiveDocument.PageSetup.Orientation = 1 Then 'Landscape
        PPO = "L"
    Else
        PPO = "P"
    End If
    
    If PMM = 0 Then
        MirrorString = "No"
        PPG = "0"
    Else
        MirrorString = "Yes"
        ' get gutter size  - no way to get this from active document - is stored in the document xml
        Sh_GP_String_1 = ""
        Sh_Read_Document_Variables "GutterSize", "VarValue"
        PPG = Sh_GP_String_1
        Sh_GP_String_1 = ""
    End If
    
    ' get media type  - no way to get this from active document - is stored in the document xml
    Sh_Read_Document_Variables "Media", "VarValue" 'places "VarValue" into Sh_GP_String_1
    If Sh_GP_String_1 = "" Then '  this document has no media entry
        DM = "Unknown"
    Else
        DM = Sh_GP_String_1
    End If
    Sh_GP_String_1 = ""

   If PPO = "L" Then
        Sh_GP_String_1 = "Landscape"
    Else
        Sh_GP_String_1 = "Portrait"
    End If

    'Test Display - Leave here for template testing
    'MsgBox " Normal Style Font Size     = " + Trim(Lp_Base_Font_Size) & vbCr _
                & " Paper/Screen Height        = " + Trim(str(PPH)) & vbCr _
                & " Paper/ScreenWidth          = " + Trim(str(PPW)) & vbCr _
                & " Top Margin                       = " + Trim(str(PTM)) & vbCr _
                & " Bottom Margin                 = " + Trim(str(PBM)) & vbCr _
                & " Left Margin                       = " + Trim(str(PLM)) & vbCr _
                & " Right Margin                     = " + Trim(str(PRM)) & vbCr _
                & " Mirrored Margins              = " + MirrorString & vbCr _
                & " Gutter Size                         = " + PPG & vbCr _
                & " Binding Margin                 = " + Trim(str(Val(PLM) + Val(PPG))) & vbCr _
                & " Orientation                        = " + PPO & vbCr _
                & " Output Media Type           = " + DM & vbCr _
                & " Right Margin Tab Setting = " + Trim(TOCTabSetting), , "VistaType LP Document Settings"
                
      Sh_GP_String_1 = ""

End Sub   '*** end of Lp_Get_Doc_Setup_Params macro ***

Sub Lp_Check_Compatibility()

    ' Author: Jerry Whittaker - jerry@thewhittakers.org
    '
    ' Version: 1.3  Date: 2/22/23 - changed message text
    ' Version: 1.2  Date: 10/7/2021 - Bug fix - exiting even if compatabilty error found
    ' Version: 1.1  Date: 9/13/2021 - added verbage re other non-docx file types.
    ' Version: 1.0  Date: 6/12/2019 - assures that the file extention is .docx
    ' Version: 1.0  Date: 4/19/2019

    
    Dim DocPath As String
    DocPath = ActiveDocument.fullName

    If Right(DocPath, 5) = ".docx" Then Exit Sub  ' is a docx file
    
    If Left(DocPath, 8) = "Document" And InStr(DocPath, "\") = 0 Then Exit Sub    ' path does not contain"\" and is not a true path - This is a new unsaved file

    MsgBox "Cannot Continue!" & vbCr _
    & vbCr & "This document is running in 'Compatibility Mode', is a PDF file, " _
    & vbCr & "an HTML file, a text file, or other non-docx file type." & vbCr _
    & vbCr & "For large print documents, the file must be edited and saved with a .docx" _
    & " extension otherwise the document will have border formatting problems." & vbCr _
    & vbCr & "To correct this issue, save this file as a .docx file and continue." & vbCr _
    & vbCr & "To make .docx the default for Word, go to .File/Options/Save to set the default file type.", , "VistaType LP (136)"
    End

End Sub   '*** end of  Lp_Check_Compatibility macro ***


Sub Lp_Table_Convert_Table_Format_Error()
  '
  ' Version: 1.1  Date: 8/4/2025
  ' Version: 1.0  Date: 12/20/2019
  '
    MsgBox "Table conversion terminated." & vbCrLf _
    & vbCrLf & "The table contains merged or split cells, an embedded table, or is" _
    & vbCrLf & "not rectangular or square in shape." _
    & vbCrLf & vbCrLf & "All columns must have the same number of cells." _
    & vbCrLf & vbCrLf & "All rows must have the same number of cells." _
    & vbCrLf & vbCrLf & "If you are trying to convert the table to a list, fix the problems " _
    & vbCrLf & "in the original table and try again." _
    & vbCrLf & vbCrLf & "If you are trying to rotate the table, use Excel's Transposition feature.", , "VistaType LP (137)"
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    End
    
End Sub

Sub Lp_Table_Convert_R_Only_Table_To_List()
'
Dim su_Prev As Boolean
su_Prev = Application.ScreenUpdating
' Version: 1.4  Date: 7/26/2026 - returns the user to where the cursor was when the macro started
' Version: 1.3  Date: 8/14/2025 - removed 40% screen - added Application.ScreenUpdating = False
' Version: 1.2  Date: 7/22/2025 - Removed "Remove manual line breaks, tabs and extra spaces from table"
'                                 routines and places into Lp_Table_Cleanup_For_Roation_And_List()
' Version: 1.1   Date: 12/5/2019 - revisions compatible with Lp_Table_Convert_Options_Form Version 1.0  Date :12/5/2019
' Version: 1.0   Date: 11/27/2019
'

    Dim TempFileName As String
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    Lp_Table_Convert_Options_Form.Hide

    Application.Run MacroName:="Lp_Table_Style_InCell_Para_And_Image"
    
    If InStr(Lp_GP_String_3, "S") > 0 Then
        Application.Run MacroName:="Lp_Table_Mark_Keep_With_Next"
    End If

   ' convert table to text
    ActiveDocument.Tables(1).Select
    Selection.rows.ConvertToText Separator:=wdSeparateByParagraphs, NestedTables:=True

    ' Store full path + file name of the active (temp) document
    Dim TempDocName As String
    TempFileName = ActiveDocument.fullName
    

    'place transcriber note at top in temp file
     Selection.HomeKey Unit:=wdStory 'top of temp doc - move to the single para mark at top

     Application.Run MacroName:="Lp_Table_Insert_Transcriber_Note"

    'remove bottom para marks
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1
    'Selection.EndKey Unit:=wdStory
    'Selection.Delete Unit:=wdCharacter, Count:=1

     DoEvents
     Selection.WholeStory
     DoEvents
     Selection.Copy
     DoEvents

     'open original doc
     DoEvents
     Documents(Lp_GP_String_2).Activate
     DoEvents
     ActiveDocument.Tables(Lp_GP_Counter_1).Select
     DoEvents
     ActiveDocument.Tables(Lp_GP_Counter_1).Delete
     DoEvents
     'The table is gone, so the cursor now sits exactly where the converted block will land.
     'The transcriber note was put at the TOP of the temp document, so it is the first
     'paragraph of what we are about to paste - leave the user on it (Jerry, 7/26/2026).
     Sh_Set_Return_Position Selection.Range.start
     Selection.Paste
     DoEvents

     'Screen stays OFF through the temp-file cleanup below. Turning it on here and THEN
     'activating the temp document painted that document on screen - a splash of the table's
     'alternating row color - before it was closed again (Jerry, 7/26/2026).
     'delete temp file
     Documents(TempFileName).Activate
     ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
     DoEvents
     Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"

     Application.ScreenUpdating = su_Prev
     Sh_Return_User_To_Start_Position

End Sub  '*** end Lp_Table_Convert_R_Only_Table_To_List macro ****

Sub Lp_Convert_Table_To_Pseudo_Columns()
    '
    ' Retains table format but looks like columns
    '
    ' Version: 1.1  Date: 1/29/2026 - full rewrite
    ' Version: 1.0  Date: 1/14/2019
    '
    Dim mainDoc As Document
    Dim sourceTbl As Table, backupTbl As Table, newTbl As Table
    Dim targetRange As Range, backupRange As Range
    Dim i As Long, totalCells As Long
    Dim RequestedColumns As Integer, CalculatedRows As Integer
    Dim objUndo As UndoRecord
    
    ' 1. INITIALIZE UNDO
    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord "Reflow Table"
    
    On Error GoTo ErrorHandler
    
    ' 2. VALIDATION
    If Not Selection.Information(wdWithInTable) Then
        MsgBox "Please click inside the table first.", vbExclamation
        objUndo.EndCustomRecord
        Exit Sub
    End If
    
    Set mainDoc = ActiveDocument
    Set sourceTbl = Selection.Tables(1)
    Set targetRange = sourceTbl.Range ' Original location
    
    ' 3. CAPTURE DATA FOR THE USER FORM
    ' Store current table column count into the public variable
    Lp_GP_Counter_1 = sourceTbl.Columns.count
    
    ' 4. GET USER INPUT
    On Error Resume Next
    Lp_Table_Convert_Options_Form.Hide
    Lp_Columns_Wanted_Form.Show
    RequestedColumns = Val(Lp_GP_String_1)
    On Error GoTo ErrorHandler
    
    If RequestedColumns <= 0 Then
        objUndo.EndCustomRecord
        Exit Sub
    End If
    
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ' 5. CREATE BACKUP AT END OF DOCUMENT
    ' We put the table at the very end so it's out of the way
    totalCells = sourceTbl.Range.Cells.count
    Set backupRange = mainDoc.Content
    backupRange.Collapse Direction:=wdCollapseEnd
    backupRange.InsertBefore vbCr
    backupRange.Collapse Direction:=wdCollapseEnd
    
    backupRange.FormattedText = sourceTbl.Range.FormattedText
    Set backupTbl = backupRange.Tables(1)
    
    ' 6. DELETE ORIGINAL & CREATE NEW STRUCTURE
    sourceTbl.Delete
    
    CalculatedRows = -Int(-totalCells / RequestedColumns)
    Set newTbl = mainDoc.Tables.Add(Range:=targetRange, _
                                   NumRows:=CalculatedRows, _
                                   NumColumns:=RequestedColumns)
    
    ' 7. TRANSFER DATA FROM BACKUP
    For i = 1 To totalCells
        Dim sRng As Range, tRng As Range
        Set sRng = backupTbl.Range.Cells(i).Range
        sRng.MoveEnd Unit:=wdCharacter, count:=-1
        
        Set tRng = newTbl.Range.Cells(i).Range
        tRng.MoveEnd Unit:=wdCharacter, count:=-1
        
        tRng.FormattedText = sRng.FormattedText
    Next i

    ' 8. APPLY STYLES
    With newTbl
        On Error Resume Next
        .Style = "Table Grid"
        .ApplyStyleHeadingRows = True
        .ApplyStyleFirstColumn = True
        On Error GoTo 0
        
        ' REMOVE ALL BORDERS
        .Borders.Enable = False
        
        .TopPadding = InchesToPoints(0.15)
        .LeftPadding = InchesToPoints(0.18)
        .AutoFitBehavior (wdAutoFitWindow)
        .AllowAutoFit = False
        .Range.ParagraphFormat.SpaceBefore = 0
        .Range.ParagraphFormat.SpaceAfter = 0
    End With
    
    ' Only apply custom borders if "Y" is present in the parameter string
    If InStr(Lp_GP_String_3, "Y") > 0 Then
        Call ApplyTableBorders(newTbl, Lp_GP_String_3, Lp_Base_Font_Size)
    End If

    ' 9. DELETE THE BACKUP TABLE AND EXTRA SPACE
    backupTbl.Delete
    If mainDoc.Characters.Last.Previous.Text = vbCr Then
        mainDoc.Characters.Last.Previous.Delete
    End If

    ' 9. REFRESH AND FINISH
    Application.ScreenUpdating = su_Prev
    DoEvents               ' Yields execution so Word can catch up
    Application.ScreenRefresh ' Forces the visual update
    
    objUndo.EndCustomRecord
    
    MsgBox "Press Ctrl+Z to restore your original table.", , "VistaType LP (221)"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = su_Prev
    If Not objUndo Is Nothing Then objUndo.EndCustomRecord
    MsgBox "Error: " & Err.Description, vbCritical
End Sub


' --- BORDER HELPER (Essential) ---
Sub ApplyTableBorders(tbl As Table, params As String, FontSize As String)
    Dim cRed As Integer, cGreen As Integer, cBlue As Integer
    Dim bWidth As WdLineWidth
    If InStr(params, "2") > 0 Then: cRed = 255: cGreen = 0: cBlue = 0
    If InStr(params, "3") > 0 Then: cRed = 228: cGreen = 90: cBlue = 45
    If InStr(params, "4") > 0 Then: cRed = 74: cGreen = 93: cBlue = 255
    If InStr(params, "5") > 0 Then: cRed = 192: cGreen = 0: cBlue = 249
    If InStr(params, "6") > 0 Then: cRed = 0: cGreen = 128: cBlue = 0
    ' 7/30/2026: was its own Select Case, and it disagreed with every other border in the
    ' document on 14 different base sizes - 18 to 21 came out at 3 pt where the box and
    ' reference-page borders were 2 1/4, and anything above 42 fell to Case Else and got the
    ' THINNEST border of all, 2 1/4, exactly where it should have been heaviest.
    bWidth = Lp_Border_Weight_For_Base_Font(FontSize)
    With tbl.Borders
        .InsideLineStyle = wdLineStyleNone
        Dim side As Variant
        For Each side In Array(wdBorderLeft, wdBorderRight, wdBorderTop, wdBorderBottom)
            With .Item(side)
                .LineStyle = wdLineStyleSingle: .LineWidth = bWidth
                .Color = IIf(InStr(params, "1") > 0, wdColorAutomatic, RGB(cRed, cGreen, cBlue))
            End With
        Next side
    End With
End Sub
   
Sub Lp_Convert_Table_To_Real_Columns()
    '
    ' Converts table to multi-column list within two continuous page breaks
    '
    ' Version: 1.1  Date: 1/29/2026 - complete rewrite
    ' Version: 1.0  Date: 1/14/2019
    '
    Dim sourceTbl As Table
    Dim targetRange As Range
    Dim RequestedColumns As Integer
    Dim objUndo As UndoRecord
    
    ' 1. INITIALIZE & VALIDATE
    On Error Resume Next
    Set sourceTbl = Selection.Tables(1)
    On Error GoTo 0
    
    If sourceTbl Is Nothing Then
        MsgBox "Please place cursor inside the table first!", vbExclamation
        Exit Sub
    End If
    
    ' 2. CAPTURE DATA FOR THE USER FORM
    ' Store current table column count into the public variable
    Lp_GP_Counter_1 = sourceTbl.Columns.count
    
    ' 3. CONFIGURE USER FORM & GET INPUT
    Lp_Table_Tools_Menu_Form.Hide
    
    ' Gray out the border checkbox since this isn't a table-grid conversion
    Lp_Columns_Wanted_Form.BorderWanted.Enabled = False
    Lp_Columns_Wanted_Form.Show
    
    ' Value returned from the Form
    RequestedColumns = Val(Lp_GP_String_1)
    If RequestedColumns <= 0 Then Exit Sub
    
    ' 4. START UNDO
    Set objUndo = Application.UndoRecord
    objUndo.StartCustomRecord "Convert Table to Columns"
    
    ' 5. SCREEN CONTROL
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ' 6. ISOLATE SECTION WITH BREAKS
    Set targetRange = sourceTbl.Range
    targetRange.Collapse Direction:=wdCollapseEnd
    targetRange.InsertBreak Type:=wdSectionBreakContinuous
    
    Set targetRange = sourceTbl.Range
    targetRange.Collapse Direction:=wdCollapseStart
    targetRange.InsertBreak Type:=wdSectionBreakContinuous
    
    ' 7. CONVERT TO TEXT & CAPTURE RANGE
    ' We set the range here so we can re-select the text after the table object is gone
    Set targetRange = sourceTbl.Range
    sourceTbl.rows.ConvertToText Separator:=wdSeparateByParagraphs, NestedTables:=False
    
    ' 8. RE-SELECT ALL WORDS & APPLY ZERO SPACING
    targetRange.Select
    With Selection.ParagraphFormat
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
    End With
    
    ' 9. APPLY COLUMN LAYOUT
    With Selection.Sections(1).PageSetup.TextColumns
        .SetCount NumColumns:=RequestedColumns
        .EvenlySpaced = True
    End With
    
    ' 10. REFRESH VISUALS
    Application.ScreenUpdating = su_Prev
    DoEvents
    Application.ScreenRefresh
    
    ' 11. CLOSE UNDO & NOTIFY
    objUndo.EndCustomRecord
    
    MsgBox "Press Ctrl+Z one time to restore the original table.", , "VistaType LP (222)"
End Sub

Sub Lp_Toggle_Space_After_Current_Para()
    '
    ' Version: 1.1  Date: 2/20/2024 - added Application.Run MacroName:="Sh_Is_Doc_Open"
    ' Version: 1.0  Date: 1/26/2020
    '
    ' gets the font size of the current paragraph and either adds that
    ' much space following the paragraph or reduces the space to zero
    ' works on any style
    
    Application.Run MacroName:="Sh_Is_Doc_Open"

    Dim FontSizeOfStyle As Integer

    FontSizeOfStyle = ActiveDocument.Styles(ActiveDocument.Paragraphs(ActiveDocument.Range(0, Selection.End).Paragraphs.count).Style).Font.Size

    If Selection.ParagraphFormat.SpaceAfter = 0 Then
        Selection.ParagraphFormat.SpaceAfter = FontSizeOfStyle
    Else
        Selection.ParagraphFormat.SpaceAfter = 0
    End If
    
End Sub   '***end of Lp_No_Space_After_Para macro ***

Sub Lp_Write_Lp_Template_Path_Into_Document_Header()

    ' Writes full path path of the template attached to the document into the document header
    '  Typical entery would be  " C:\users\jerry\appdata\roaming\microsoft\templates\Large Print Templates\LP Templates for .50 inch margin\24 pt 8.5x11 paper .50 inch margin.dotx"
    '
    ' Version: 1.1  Date: 2/7/2020 - rewrite
    ' Version: 1.0  Date: 1/31/2020

    Dim templatePath As String
    Dim LpTemplatePath As String
    
    templatePath = Dialogs(wdDialogToolsTemplates).Template

    ' if the LpTemplatePath stored in the document variables is non existant then
    ' create an undefined LpTemplatePath
    
    LpTemplatePath = "Undefined"
    On Error Resume Next
    LpTemplatePath = ActiveDocument.Variables("Undefined")
    ActiveDocument.Variables("LpTemplatePath").Delete

    ActiveDocument.Variables.Add Name:="LpTemplatePath", Value:=templatePath
    
    'To view this setting in the file, rename the .docx to .zip, right click and select open
    ' select the folder "word" right click and select open
    ' select settings.xml and right click and select open

End Sub

Sub Lp_Attach_The_Template()

    ' Attaches the LP template with style changes
    '
    ' Version: 3.6  Date: 8/20/2026 - closes and reopens the document after the Save As, which is
    '                               the only thing that makes Word repaint the Quick Style gallery
    '                               after the attached template has changed which styles belong in
    '                               it - see Sh_Close_And_Reopen, and do NOT reach for
    '                               Document.UpdateStyles instead. Re-asserts the Styles pane
    '                               afterwards, because pane visibility is application state and
    '                               does not survive the last document closing
    ' Version: 3.5  Date: 8/20/2026 - the typeface is no longer a choice. Lp_Base_Font_Name arrives
    '                               as Tahoma for every new book, and as the book's own face for one
    '                               already set in the dropped VistaTypeLP Legible - see
    '                               LP_FONT_LEGACY_LEGIBLE and AttachOkay_Click. Nothing in this
    '                               macro changed; the value reaching it did
    ' Version: 3.4  Date: 8/12/2026 - the two "only on a document that is not already LP" blocks
    '                               now test Lp_Doc_Was_Already_LP instead of a leftover string in
    '                               the shared Lp_GP_String_1. See Lp_Attach_Lp_Template 2.4: the
    '                               old flag was never cleared, so after one re-attach in a Word
    '                               session every later attach skipped the file cleanup and the
    '                               empty-paragraph removal without saying so
    ' Version: 3.3  Date: 8/8/2026 - applies the typeface the transcriber chose (Lp_Base_Font_Name)
    '                               to the styles AND the text, instead of hard-coding Tahoma, and
    '                               embeds the font in the document when it is ours. Also records
    '                               the choice in the BaseFontName document variable
    ' Version: 3.2  Date: 8/2/2026 - now the ONE place that forces the Styles pane open at Recommended sort and Recommended filter, for both a new attach and a re-attach; placed before the Save As so it survives a canceled save
    ' Version: 3.1  Date: 7/23/2026 - calls Lp_Set_Prodnote_Style_Visibility after the attach so "Prodnote" shows in the Styles pane when the document uses it (the pre-attach hide-all loop hid it, and unhideWhenUsed does not fire for a style that was already in use, e.g. a converted DAISY/NIMAS document)
    ' Version: 3.0  Date: 7/23/2026 - reverted the 2.9 "Prodnote" exception: Style.Visibility = True sets <w:semiHidden/> (it HIDES), so the pre-attach loop hides every style as its comment says, and excluding Prodnote only stopped it being hidden
    ' Version: 2.9  Date: 7/22/2026 - (superseded) skipped "Prodnote" in the pre-attach visibility loop
    ' Version: 2.8  Date: 7/18/2026 - removed the "template has been attached" prompt; Save As now uses one dialog object so the file saves under the name the user types
    ' Version: 2.7  Date: 7/18/2026 - stabilize the document before saving; now saves only once (attach->stabilize->save)
    ' Version: 2.6  Date: 7/3/2026 - changed external app order
    ' Version: 2.5  Date: 3/5/2026 - added non-modal message
    ' Version: 2.4  Date: 10/13/2025 - Moved "Lp_ResizePicturesAndShapesToFitPageWidthAndPageHeight" from runnining only on new documents
    '                                  to running anytime template is added to new documents or added to existing LP doc
    ' Version: 2.3  Date: 9/1/2025 - added Application.Run MacroName:="Sh_Replace_Multiple_Para_Marks_No_Warning"
    ' Version: 2.2  Date: 5/13/2025 - added Application.Run MacroName:="MS_Set_Word_Config_For_Large_Print"
    ' Version: 2.1  Date: 5/5/2025 - added   Lp_ResizePicturesAndShapesToFitPageWidthAndPageHeight AND Lp_SetPicturesToInlineAndLockAspectRatio
    ' Version: 2.0  Date: 2/4/2024 - removed  Application.Run MacroName:="Lp_Fix_Normal_Styles"
    ' version: 1.9  Date: 1/30/2024 - removed call to Lp_Italics_To_Dashed_Underline
    ' Version: 1.8  Date: 10/24/2023 - added On Error Resume Next to Selection.PageSetup.Orientation = wdOrientPortrait
    ' Version: 1.7  Date: 10/23/2021 - added call to "Lp_Set_TOC_and_Print_Page_Num_Tab_Stops"
    ' Version: 1.6  Date: 10/20/2021 - will now update all styles in all files regargless of age
    ' Version: 1.5  Date: 7/9/2021 - added check and message for LargePrintTemplate.dotx
    ' Version: 1.4  Date: 3/20/2021 - turn document background to white
    ' Version: 1.3  Date: 12/9/2020 - now working as advertised
    ' Version: 1.2  Date: 12/1/2020 - Minor fix to gutter and added writing gutter size to doc xml file
    ' Version: 1.1  Date: 11/26/2020 - added code for gutter size
    ' Version: 1.1  Date: 11/12/2020 - if doc was previously LP (Lp_GP_String_1 = "Doc_Is_Already_LP") doc then bypass change table colors - added correction to tabs for Print Pg Num
    ' Version: 1.0  Date: 11/8/2020
    '
    
   'Lp_Attach_An_Lp_Template_Form.Hide ' hide the user form for fontsize and media type'+++++ new non-modal form code

    Unload Lp_Attach_An_Lp_Template_Form

    ' The form is gone by now - this macro is reached asynchronously through Sh_BridgeTargetMacro
    ' and Application.OnTime - so the Public set in AttachOkay_Click is the only carrier of the
    ' typeface choice. Default it if it is empty, because an "End" statement anywhere on the way
    ' in wipes every Public, and there is one at the "template does not exist" bail-out below.
    If Len(Trim(Lp_Base_Font_Name)) = 0 Then Lp_Base_Font_Name = LP_FONT_TAHOMA

    'save the name of the current document
    Dim currentdoc As Document
    Set currentdoc = ActiveDocument 'will work with blank, unsaved documents too
    
    If Not Lp_Doc_Was_Already_LP Then  ' this only needs to be done on docs which are not lp
        ' The message belongs INSIDE the If. It used to be set just above it as well, so a
        ' re-attach announced "Fixing common file errors" and then skipped the macro - which is
        ' correct behavior for a document that is already large print, but the message said
        ' otherwise. Jerry, 8/3/2026.
        '
        ' The spinner turns during the macro because its 28 DoEvents are now Sh_Spin_DoEvents,
        ' which advances whichever progress box is showing - this one here, or the please-wait
        ' box when File Cleanup is run from the ribbon instead.
Sh_NonModalMessageForm.SetActivityMessage "Fixing common file errors"
DoEvents

        Application.Run MacroName:="Lp_Fix_Common_File_Errors"
    End If
    
Sh_NonModalMessageForm.SetActivityMessage "Attaching the VistaType LP template"
DoEvents
 
    ' NOTE: do NOT disable Application.Options.Pagination here. It was tried on 7/26/2026 to
    ' stop screen flashing, did not stop it (the cause was the progress form's spinner
    ' repainting - see Sh_NonModalMessageForm.SpinTick), and it left every table black and
    ' white: banded row shading from "Yellow on White Paper Table" is applied by Word's
    ' LAYOUT engine, so switching background repagination off across the attach means the
    ' banding never gets computed.
    Application.ScreenUpdating = False ' Turn screen updating off
    ' hide all non-Word styles before attaching the LP template.
    On Error GoTo AvoidCrash
        'Adapted From: https://www.office-forums.com/threads/styles-styles-how-to-hide-unused-styles.1881281/
        Dim oSty As Style
            With ActiveDocument
            For Each oSty In .Styles
                ' Style.Visibility = True sets <w:semiHidden/>, i.e. it HIDES the style
                ' (verified against the saved XML), so this loop hides every existing style
                ' before the LP template is attached -- as the comment above says.
                .Styles(oSty.NameLocal).Visibility = True
            Next oSty
         End With

AvoidCrash:
    On Error GoTo 0

    'attach "LargePrintTemplate.dotx"
    With ActiveDocument
            Dim TemplatePathandName As String
            TemplatePathandName = Options.DefaultFilePath(wdUserTemplatesPath) + "\" + LP_TEMPLATE_FILE
            If Sh_FileExists(TemplatePathandName) Then  ' is the large print template available for this user?
                Lp_Fix_Normal_Styles
                .UpdateStylesOnOpen = True
                .AttachedTemplate = TemplatePathandName
                .UpdateStylesOnOpen = False  ' supresses any further style updates
                .Application.Run MacroName:="Lp_Remove_All_Styles_Except_Lp_Styles"
            Else
                Unload Sh_NonModalMessageForm
                Application.ScreenUpdating = True
                MsgBox " Cannot continue!" + vbCr + vbCr + "The template file: " + TemplatePathandName + " does not exist." + vbCr + vbCr + "Install the file and try again.", , "VistaType LP (141)"
                End
            End If
    End With

Sh_NonModalMessageForm.SetActivityMessage "Converting document font to " & Lp_Base_Font_Name
DoEvents

    ' Set the whole document to the chosen face and size - styles first, then the text.
    '
    ' This must stay AFTER Lp_Remove_All_Styles_Except_Lp_Styles above. The attach a few lines
    ' up runs with UpdateStylesOnOpen = True, which copies the template's styles in and puts
    ' Normal back to Tahoma; anything set before that point is thrown away.
    '
    ' Both calls are needed. The styles carry the choice - which is what makes it readable again
    ' after the document is closed and reopened - and the direct formatting is what makes a messy
    ' imported document actually conform.
    Lp_Apply_Base_Font_To_Styles ActiveDocument, Lp_Base_Font_Name
    Sh_Set_Whole_Document_Font ActiveDocument, Lp_Base_Font_Name, CSng(Val(Lp_Base_Font_Size))

    ' Keep this. The old code got here through Selection.WholeStory, which left the selection
    ' spanning the document; Sh_Set_Whole_Document_Font does not touch the selection at all, and
    ' the page-setup block below works on the section holding it.
    Selection.HomeKey Unit:=wdStory

    ' ----------------------------------------------------------------------------------------------------------------------------
    ' Embed the typeface, so a book still sets correctly on a machine that has not got it
    ' ----------------------------------------------------------------------------------------------------------------------------
    ' Set here, well before the Save As at the end of this macro - after it would produce a saved
    ' file with no font in it and nothing to notice. These are document properties, so they stick:
    ' every later Ctrl+S re-embeds.
    '
    ' Turned on only when the book is set in something other than Tahoma. There is nothing worth
    ' embedding in a Tahoma book - every Windows machine has Tahoma - and leaving it off keeps
    ' those files the size they have always been.
    '
    ' This line is what makes dropping the bundled typeface on 8/20/2026 safe. Every book attached
    ' in VistaTypeLP Legible between 8/8 and 8/20/2026 carries its own copy of the face, so it goes
    ' on setting and printing correctly after the installer takes that font off the machine. Do not
    ' reduce this to a constant False because nothing bundles a font any more - a legacy book keeps
    ' its face through a re-attach, and it must keep the embedded copy with it.
    '
    ' Be clear about what this switch is, though: it is per DOCUMENT, not per font. Word embeds
    ' every embeddable non-system face the document uses, so a running head left in some other
    ' font, or a style like Title that names Calibri Light outright, gets embedded too and the
    ' file grows accordingly. There is no "embed just this one" in Word. Worth measuring a real
    ' book on the build box; if it is fat, the cure is to stop the stray faces getting in.
    '
    ' DoNotEmbedSystemFonts is set for good order, not as the Tahoma guard - Word's "system
    ' fonts" list is the old Arial/Courier New/Times New Roman/Symbol/Wingdings set and does not
    ' include Tahoma. The line above is what keeps Tahoma books clean.
    '
    ' NOT subsetted. A subset holds only the characters the document already uses, so a teacher
    ' or a reader who edits the book and types a character it did not contain gets a substitute
    ' for that character. These files are handed on to be edited; take the extra size.
    ActiveDocument.DoNotEmbedSystemFonts = True
    ActiveDocument.EmbedTrueTypeFonts = (Lp_Base_Font_Name <> LP_FONT_TAHOMA)
    ActiveDocument.SaveSubsetFonts = False

    ' ----------------------------------------------------------------------------------------------------------------------------
    ' Setup Page Parameters
    ' ----------------------------------------------------------------------------------------------------------------------------
    
Sh_NonModalMessageForm.SetActivityMessage "Setting page/screen margins and gutter sizes"
DoEvents

    With ActiveDocument.PageSetup
         .TopMargin = InchesToPoints(PTM)
         .BottomMargin = InchesToPoints(PBM)
         .LeftMargin = InchesToPoints(PLM)
         .RightMargin = InchesToPoints(PRM)
         
        If PMM Then ' Page mirror margins
            .MirrorMargins = True
            .Gutter = InchesToPoints(PPG) ' PPG is gutter size
            .GutterPos = wdGutterPosLeft
            ' writes a variable name (GutterSize) and variable value into the document xml file
            Sh_Write_Document_Variables "GutterSize", PPG
        Else
             .GutterPos = wdGutterPosLeft
             .MirrorMargins = False
             .Gutter = InchesToPoints(0)
        End If
        
Sh_NonModalMessageForm.SetActivityMessage "Setting page orientation"
DoEvents

         ' writes a variable name (Media) and variable value (DM) into the document xml file
         Sh_Write_Document_Variables "Media", DM

         ' The typeface the transcriber ASKED for. The Normal style is the working answer and is
         ' what everything reads, but a style can be changed afterwards - "update style to match
         ' selection" on a Normal paragraph, or pasting a block of Tahoma - and then nothing
         ' records what this book was meant to be. Sh_Doc_Info shows both, so a disagreement
         ' between them is visible instead of being guessed at over the telephone.
         Sh_Write_Document_Variables "BaseFontName", Lp_Base_Font_Name

         ' *** set for all documents *****
         .HeaderDistance = InchesToPoints(0)
         .FooterDistance = InchesToPoints(0)
         
         ' *** Set Page Size and Orientation ***
         '  orientation must bet set before size
         If PPO = "L" Then
             Selection.PageSetup.Orientation = wdOrientLandscape
          Else
            On Error Resume Next
            Selection.PageSetup.Orientation = wdOrientPortrait 'will crash if first item in document is a drop cap
         End If

         .PageWidth = InchesToPoints(PPW)
         .PageHeight = InchesToPoints(PPH)
    End With

   ' ----------------------------------------------------------------------------------------------------------------------------
   ' Fix para styles
   ' ----------------------------------------------------------------------------------------------------------------------------

    If Not Lp_Doc_Was_Already_LP Then  ' this only needs to be done on docs which are not lp
    
Sh_NonModalMessageForm.SetActivityMessage "Setting table alternating color style"
DoEvents
        ' all tables to default - Convert existing tables to "Yellow on White Paper Table"
         Dim Tbl_Cnt As Integer
         Tbl_Cnt = 0
         Tbl_Cnt = ActiveDocument.Tables.count
         If Tbl_Cnt <> 0 Then
             Dim t As Table
             For Each t In ActiveDocument.Tables
                 t.Style = "Yellow on White Paper Table"
             Next
         End If

Sh_NonModalMessageForm.SetActivityMessage "Removing empty paragraphs"
DoEvents

        Application.Run MacroName:="Sh_Replace_Multiple_Para_Marks_No_Warning"
        
Sh_NonModalMessageForm.SetActivityMessage "Fixing Abbyy FineReader headings and normal styles"
DoEvents
        
        Application.Run MacroName:="Lp_Fix_Abbyy_Text_and_Headers"
        
    End If
    
Sh_NonModalMessageForm.SetActivityMessage "Adjusting oversize pictures to fit within margins"
DoEvents

    Application.Run MacroName:="Lp_SetPicturesToInlineAndLockAspectRatio"

    Application.Run MacroName:="Lp_ResizePicturesAndShapesToFitPageWidthAndPageHeight"
    
    Application.Run MacroName:="Lp_Normalize_Styles" 'this routine sets it's on non-modal messages
    
Sh_NonModalMessageForm.SetActivityMessage "Setting tabs for TOCs and reference page numbers."
DoEvents

    Application.Run MacroName:="Lp_Set_TOC_and_Print_Page_Num_Tab_Stops"

    ' Show "Prodnote" in the Styles pane only if this document actually uses it. Needed
    ' because the pre-attach loop above hides EVERY style, and the template's
    ' <w:unhideWhenUsed/> only fires when a style is newly APPLIED -- it does not
    ' retroactively un-hide a style that was already in use, as in a converted DAISY/NIMAS
    ' document whose prodnotes are already styled.
    Application.Run MacroName:="Sh_Set_Prodnote_Style_Visibility"

    ' Turn on print view
    Application.Run MacroName:="Lp_Set_Display_For_Large_Print"
    Application.Run MacroName:="MS_Set_Word_Config_For_Large_Print"

    '******************  cleanup  **************************
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    ActiveDocument.Background.Fill.Visible = msoFalse

    Application.ScreenUpdating = True ' Turn screen updating on
    Application.ScreenRefresh
    Sh_Set_Navigation_Pane False 'navigation pane - both halves, see Sh_Set_Navigation_Pane
    ActiveDocument.UndoClear ' clear the undo stack
    
DoEvents
    ' Make the document visible and active on screen
    currentdoc.Activate

    Sh_SetBarVisible "Styles", True

    ' The ONE place the Styles pane is forced to Recommended sort and Recommended filter (opening
    ' it is no longer exclusive to the attach - MS_Set_Word_Config_For_Large_Print does that for
    ' any large print document as of 8/20/2026; the sort and the filter are still only set here):
    ' attaching the LP template, whether it is new to this document or being re-attached (both
    ' arrive here - the re-attach path only adds the warning form first). Everywhere else the
    ' user's own pane settings are left alone, so they survive from session to session.
    ' Jerry, 8/2/2026.
    '
    ' Placed here, not after the Save As below: ScreenUpdating is back on and currentdoc is
    ' active, so the pane paints and the settings land in the transcriber's book rather than a
    ' leftover temp window - and this still runs when the user cancels the save, which is
    ' exactly when they are left working in a freshly attached document.
    Application.Run MacroName:="Lp_Turn_on_Styles_Pane"

    Dim doc As Document
    Dim userChoice As VbMsgBoxResult
    Set doc = ActiveDocument

    ' Stabilize the document FIRST, then save it exactly once (attach -> stabilize -> save).
    Sh_NonModalMessageForm.Show vbModeless
    Sh_NonModalMessageForm.SetActivityMessage "Repaginating the document"
    DoEvents
    Sh_PauseSeconds 3   'pause for nn seconds

    doc.Repaginate

    Sh_NonModalMessageForm.Show vbModeless
    Sh_NonModalMessageForm.SetActivityMessage "Updating document fields"
    DoEvents
    Sh_PauseSeconds 3   'pause for nn seconds

    doc.Fields.Update
    doc.UndoClear

    ' Hide the progress form momentarily so Windows can cleanly shift focus to the Save As dialog
    Sh_NonModalMessageForm.Hide

    ' Force the Word Application and your specific document to the front
    Application.Activate
    currentdoc.Activate
    DoEvents

    Dim dlgSaveAs As Dialog
SaveTheFile:
 'jw
    ' Capture ONE dialog object and use it for BOTH .Display and .Execute, so the file is
    ' saved under the name the user types. (A separate Dialogs(wdDialogFileSaveAs) reference
    ' for .Execute ignores the typed name and re-saves under the document's current/default name.)
    Set dlgSaveAs = Dialogs(wdDialogFileSaveAs)
    ' .Display ONLY opens the window to get the file name; it does NOT save yet
    If dlgSaveAs.Display <> -1 Then
        ' User canceled the dialog
        userChoice = MsgBox( _
            "You have canceled the Save." & vbCrLf & vbCrLf & _
            "Continuing without saving may result in an unstable Word document." & vbCrLf & vbCrLf & _
            "Do you want to reconsider saving this file?", _
            vbYesNo + vbExclamation, _
            "Save As Canceled")

        If userChoice = vbNo Then
            Unload Sh_NonModalMessageForm
            Application.ScreenUpdating = True
            Exit Sub
        Else
            GoTo SaveTheFile
        End If
    Else
        ' 1. Show the non-modal form BEFORE saving
        Sh_NonModalMessageForm.Show vbModeless

        'Sh_NonModalMessageForm.LblMessage ""
        Sh_NonModalMessageForm.SetActivityMessage "Saving the stabilized document. Activity spinner is idle."
        DoEvents

        ' 2. Execute the save on the SAME dialog object so the typed name is used
        dlgSaveAs.Execute

        ' 3. Keep the message up for a brief moment so they see it finish
        Sh_PauseSeconds 3
    End If

    'Unload the progress form completely
    Unload Sh_NonModalMessageForm
    DoEvents

    ' Close and reopen the saved file, which is the ONLY thing that makes Word repaint the Quick
    ' Style gallery after the attached template has changed which styles belong in it. Jerry's
    ' finding, 8/20/2026; see Sh_Close_And_Reopen for the two answers that were tried first and
    ' the one that must never be used.
    '
    ' AFTER the save and before the closing message, so the transcriber sees the window blink
    ' once and then "File has been stabilized and saved" - rather than a blink after being told
    ' the job was finished.
    '
    ' BOTH variables have to be re-pointed. The close kills whatever they referred to, and
    ' currentdoc is the one the earlier half of this macro used - leave it stale and the next
    ' person to add a line down here gets "object variable not set" on a document that is sitting
    ' open in front of them.
    Set doc = Sh_Close_And_Reopen(doc)
    Set currentdoc = doc

    ' Nothing means the reopen failed and the document is not on screen. Her work is on disk -
    ' the Save As above succeeded, which is the only way this line is reached - so say where it
    ' is and stop, rather than activating a document that is not there.
    If doc Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "Your document has been saved, but Word could not reopen it." & vbCr & vbCr _
             & "Open it again from where you saved it and carry on. Nothing has been lost.", _
             vbExclamation, "VistaType LP (233)"
        Exit Sub
    End If

    'Re-assert dominance for your saved document AFTER the form is entirely gone
    Application.Activate
    doc.Activate
    ActiveWindow.View.Type = wdPrintView
    DoEvents

    ' The Styles pane again, and it is not a duplicate of the call further up. Pane VISIBILITY is
    ' application state, not document state: on the common machine where this book is the only one
    ' open, closing it leaves Word with no documents and the pane goes with them. The sort and
    ' filter settings do survive - those are saved in the file - so this is here for the pane
    ' itself. This is the same attach, finishing the job after the reopen. Still needed after
    ' 8/20/2026 gave every large print document its pane on open: the reopen runs with the open
    ' handler live, but on a machine where this book was the only one open the pane went when the
    ' document did, and Word does not always restore it from that.
    Application.Run MacroName:="Lp_Turn_on_Styles_Pane"

    MsgBox "File has been stabilized and saved", vbInformation, "VistaType LP (201)"
    
    Unload Sh_NonModalMessageForm
    
'jw
    'MsgBox "The current time is: " & Time, vbInformation, "Current Time"
    
End Sub   '*** end of Lp_Attach_The_Template macro ***

Sub Lp_Make_All_Pictures_In_Selected_Table_Inline()

    'From: https://microsoft.public.word.vba.general.narkive.com/aHYD2M1F/convert-shape-in-table-to-inlineshape

    Dim oHeight As Single
    Dim oWidth As Single
    Dim s As Shape
    Dim oCell As cell
    
    On Error GoTo Bye ' if no table
    
    Set oCell = Selection.Tables(1).cell(1, 1)
    
    For Each s In ActiveDocument.Shapes
        If s.Anchor.InRange(oCell.Range) Then
            If s.Type = msoPicture Then
                s.ConvertToInlineShape
                
                oHeight = oCell.Range.InlineShapes(1).Height
                oWidth = oCell.Range.InlineShapes(1).Width
                
                oCell.Height = oHeight
                Selection.Tables(1).Columns(1).Width = oWidth
                Exit For
            End If
        End If
        Next s
Bye:

End Sub   '*** end of Lp_Convert_Shapes_to_Inline macro ***

Sub Lp_Keep_With_Next_Para()
'
' Version: 1.1  Date: 7/27/2026 - follow the paragraph if it reflows to the next page
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    With Selection.ParagraphFormat
        .KeepWithNext = wdToggle
    End With
    Sh_Keep_Cursor_In_View
End Sub

Sub Lp_Remove_Space_After_Para()
'
' Removes space following current para and glues it to the next paragraph (toggle)
' Version: 1.0  Date: 11/30/2020
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    With Selection.ParagraphFormat
        Application.Run MacroName:="Sh_Is_Doc_Open"
        If .SpaceAfter > 0 Then
            .SpaceAfter = 0
            .KeepWithNext = True
        Else
            .SpaceAfter = ActiveDocument.Styles(wdStyleNormal).ParagraphFormat.SpaceAfter
            .KeepWithNext = False
        End If
    End With
    
End Sub '** end of Lp_Remove_Space_After_Para **

Sub Lp_Picture_Tools_Menu_Starter()
'
'  Version: 1.0  Date: ??
    Application.Run MacroName:="Sh_Is_Doc_Open"
    Application.Run MacroName:="Lp_Is_Lp_Template_Attached"
    Lp_Bakgrnd_Picture_Menu_Form.Show
    Unload Lp_Bakgrnd_Picture_Menu_Form
End Sub
Sub Lp_Set_TOC_and_Print_Page_Num_Tab_Stops()
'
' Sets TOC and Print Page Number tab stops bases on the size of the paper/screen
'
' Version 1.0  Date: 4/2/2021

    With ActiveDocument.PageSetup
        
        If Sh_Style_Exists(ActiveDocument, "TOC 1") Then
        ActiveDocument.Styles("TOC 1").ParagraphFormat.TabStops.ClearAll
        ActiveDocument.Styles("TOC 1").ParagraphFormat.TabStops.Add Position:= _
            InchesToPoints(TOCTabSetting), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderDots
        End If
        
        If Sh_Style_Exists(ActiveDocument, "TOC 2") Then
        ActiveDocument.Styles("TOC 2").ParagraphFormat.TabStops.ClearAll
        ActiveDocument.Styles("TOC 2").ParagraphFormat.TabStops.Add Position:= _
            InchesToPoints(TOCTabSetting), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderDots
        End If
        
        If Sh_Style_Exists(ActiveDocument, "TOC 3") Then
        ActiveDocument.Styles("TOC 3").ParagraphFormat.TabStops.ClearAll
        ActiveDocument.Styles("TOC 3").ParagraphFormat.TabStops.Add Position:= _
            InchesToPoints(TOCTabSetting), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderDots
        End If
        
        If Sh_Style_Exists(ActiveDocument, "TOC 4") Then
        ActiveDocument.Styles("TOC 4").ParagraphFormat.TabStops.ClearAll
        ActiveDocument.Styles("TOC 4").ParagraphFormat.TabStops.Add Position:= _
            InchesToPoints(TOCTabSetting), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderDots
        End If
        
        If Sh_Style_Exists(ActiveDocument, "TOC 5") Then
        ActiveDocument.Styles("TOC 5").ParagraphFormat.TabStops.ClearAll
        ActiveDocument.Styles("TOC 5").ParagraphFormat.TabStops.Add Position:= _
            InchesToPoints(TOCTabSetting), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderDots
        End If
        
        If Sh_Style_Exists(ActiveDocument, "Print Pg Num") Then
        ActiveDocument.Styles("Print Pg Num").ParagraphFormat.TabStops.ClearAll
        ActiveDocument.Styles("Print Pg Num").ParagraphFormat.TabStops.Add Position:= _
            InchesToPoints(TOCTabSetting), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
        End If
                      
    End With
    
End Sub   '*** end of Lp_Set_TOC_and_Print_Page_Num_Tab_Stops macro ***

Function Lp_Border_Weight_For_Base_Font(Optional ByVal SizeOverride As String = "") As WdLineWidth
'
' The one place that decides how heavy a border should be for the current base font size.
'
' Version: 1.0  Date: 7/30/2026 - split out of Lp_Normalize_Styles so that macro and
'                                 Lp_Set_Table_Border_Weights cannot drift apart. They had:
'                                 Normalize Styles compared numerically, while
'                                 Lp_Set_Table_Border_Weights matched the size as TEXT
'                                 against fifteen literals ("14","16",..."42"), so any odd
'                                 size, and anything above 42, silently got no border at all.
'

    Dim fSize As Single

    If SizeOverride <> "" Then
        ' A caller that already holds the size (ApplyTableBorders is handed it).
        fSize = Val(SizeOverride)
    Else
        ' Same fallback (and same side effect of filling in Lp_Base_Font_Size) as before.
        If Lp_Base_Font_Size = "" Then
            Lp_Base_Font_Size = ActiveDocument.Styles(wdStyleNormal).Font.Size
        End If
        fSize = Val(Lp_Base_Font_Size)
    End If

    If fSize >= 38 Then
        Lp_Border_Weight_For_Base_Font = wdLineWidth600pt      ' border weight menu: 6
    ElseIf fSize >= 30 Then
        Lp_Border_Weight_For_Base_Font = wdLineWidth450pt      ' 4 1/2
    ElseIf fSize >= 22 Then
        Lp_Border_Weight_For_Base_Font = wdLineWidth300pt      ' 3
    ElseIf fSize >= 14 Then
        Lp_Border_Weight_For_Base_Font = wdLineWidth225pt      ' 2 1/4
    Else
        Lp_Border_Weight_For_Base_Font = wdLineWidth050pt      ' 1/2
    End If

End Function   '*** End of Lp_Border_Weight_For_Base_Font ***

Sub Lp_Set_Table_Border_Weights()

    ' sets the border weight of ALL tables relative to the Base Font Size (Normal style),
    ' clears any diagonal borders and table shadows, and sets Word's own default border
    ' width so a border the user draws afterwards matches.
    '
    ' Called by Lp_Normalize_Styles and by all four "color every table" buttons on
    ' Lp_Table_Tools_Menu_Form - applying a Word table style resets the borders, so those
    ' buttons call this to put the weights back.
    '
    ' Version: 2.0  Date: 7/30/2026 - weight now chosen by Lp_Border_Weight_For_Base_Font
    '                                 instead of matching the size as text against fifteen
    '                                 literals. Sizes 14-42 EVEN behave exactly as before;
    '                                 odd sizes and anything above 42 previously fell through
    '                                 every branch and left the borders untouched, so a
    '                                 Table Tools color button appeared to do nothing to
    '                                 them. Four near-identical 48-line blocks became one.
    ' Version: 1.2  Date: 1/9/2024 - added On Erro Resume if table has missing borders
    ' version: 1.1  Date: 10/16/2023 - set width 225pt range to include Lp_Base_Font_Size = "20"
    ' Version: 1.0  Date: 9/29/2022

    Dim CurrentTable As Table
    Dim targetWeight As WdLineWidth
    Dim edge As Variant

    targetWeight = Lp_Border_Weight_For_Base_Font
    Options.DefaultBorderLineWidth = targetWeight

    For Each CurrentTable In ActiveDocument.Tables
        With CurrentTable

            ' A table can be missing an edge (a single cell has no inside borders), and
            ' asking for one that is not there raises. Skip it and carry on, as before.
            On Error Resume Next

            For Each edge In Array(wdBorderLeft, wdBorderRight, wdBorderTop, _
                                   wdBorderBottom, wdBorderHorizontal, wdBorderVertical)
                With .Borders(edge)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = targetWeight
                    .Color = wdColorAutomatic
                End With
            Next edge

            .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
            .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
            .Borders.Shadow = False

            On Error GoTo 0

        End With
    Next CurrentTable

End Sub

Sub Lp_Replace_Compact_Fractions_With_Fraction_Text()
'
' Version: 1.0  Date: 10/19/2021
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "½"
        .Replacement.Text = "1/2"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8531)
        .Replacement.Text = "1/3"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8532)
        .Replacement.Text = "2/3"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "¼"
        .Replacement.Text = "1/4"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
   
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "¾"
        .Replacement.Text = "3/4"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8533)
        .Replacement.Text = "1/5"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8534)
        .Replacement.Text = "2/5"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8535)
        .Replacement.Text = "3/5"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8536)
        .Replacement.Text = "4/5"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8537)
        .Replacement.Text = "1/6"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8538)
        .Replacement.Text = "5/6"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8528)
        .Replacement.Text = "1/7"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8539)
        .Replacement.Text = "1/8"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8540)
        .Replacement.Text = "3/8"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8541)
        .Replacement.Text = "5/8"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8542)
        .Replacement.Text = "7/8"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8529)
        .Replacement.Text = "1/9"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8530)
        .Replacement.Text = "1/10"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub   '*** end of Lp_Replace_Compact_Fractions_With_Fraction_Text macro ***

Sub Lp_Replace_Strong_With_Bold()
    '
    ' Replace "Strong" style with Bold and delete "Strong"
    '
    ' Version 1.0  Date: 10/28/2021
    '
    ' Guarded: Styles(name) raises 5941 when the document does not carry that style, and
    ' this pass can do nothing without it - no style means nothing is formatted with it.
    If Sh_Style_Exists(ActiveDocument, "Strong") Then
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Strong")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End If
 
    Application.Run MacroName:="Lp_Remove_All_Styles_Except_Lp_Styles"

End Sub   '*** end of Lp_Replace_Strong_With_Bold macro ***

Sub Lp_Remove_All_Styles_Except_Lp_Styles()
'
' Removes all styles except LP styles from the Styles Pane. Non-LP custom
' PARAGRAPH/LINKED styles that are actually in use have their text reassigned to
' Normal first, then the now-unused style is deleted -- so foreign body styles
' carried in from OCR / imports / web paste are neutralized, not just left in the
' pane. Built-in Word styles (Heading 1-5, TOC 1-5, List Paragraph, List Bullet,
' Header, Footer, ...) are never touched: the BuiltIn = False guard protects them.
'
' Version: 1.4  Date: 7/21/2026 - added "Prodnote" to the LP keep-list so the DAISY/NIMAS
'                                 converter's Prodnote-styled paragraphs survive template attach
' Version: 1.3  Date: 7/20/2026 - in-use non-LP custom paragraph/linked styles are now
'                                 converted to Normal before deletion (were kept if in use);
'                                 whitelist test is exact (comma-delimited) instead of InStr substring;
'                                 collect-then-process so styles aren't deleted mid-enumeration
' Version: 1.2  Date: 1/23/2024 - added "Words Black Inverted" and "Para Black Inverted"
' Version: 1.1  Date: 2/19/2023 - added "Para Black" and "Words Black"
' Version: 1.0  Date: 11/2/2021
'
    Dim oStyle As Style
    Dim st As Style
    Dim LpStyles As String        ' LP styles to KEEP (comma-delimited, wrapped in commas)
    Dim styleName As String
    Dim doomed As Collection      ' non-LP custom style names, collected before any deletion
    Dim nm As Variant

    ' Wrapped in commas on both ends so membership is an EXACT match:
    '   InStr(LpStyles, "," & name & ",")  -- so "Para" no longer matches "Para Aqua", etc.
    LpStyles = ",1 point,Gray Scale Table,Yellow on Black Screen Table,Yellow on White Paper Table," _
    & "Normal,List,List 1,Text Blue,Text Green,Text Orange,Text Red,Text Violet," _
    & "Box Black,Box Blue,Box Orange,Box Red,Box Violet,Box White," _
    & "Words Aqua,Words Black,Words Blue,Words Green,Words Pink,Words Tan,Words Yellow," _
    & "Para Aqua,Para Black,Para Blue,Para Green,Para Tan,Para Yellow,Print Pg Num,Words Black Inverted," _
    & "Para Black Inverted,Prodnote,"

    ' Pass 1: collect the non-built-in, non-LP style names. We do NOT delete inside this
    ' For Each -- removing items from the Styles collection mid-enumeration skips styles.
    Set doomed = New Collection
    For Each oStyle In ActiveDocument.Styles
        If oStyle.BuiltIn = False Then
            styleName = oStyle.NameLocal
            If InStr(LpStyles, "," & styleName & ",") = 0 Then
                doomed.Add styleName
            End If
        End If
    Next oStyle

    ' Pass 2: neutralize + remove each collected style.
    For Each nm In doomed
        On Error Resume Next
        Set st = Nothing
        Set st = ActiveDocument.Styles(CStr(nm))
        If Not st Is Nothing Then
            Select Case st.Type
                Case wdStyleTypeParagraph, wdStyleTypeLinked
                    ' Reassign any text using this style to Normal, then drop the style.
                    With ActiveDocument.Content.Find
                        .ClearFormatting
                        .Style = CStr(nm)
                        .Replacement.ClearFormatting
                        .Replacement.Style = ActiveDocument.Styles(wdStyleNormal)
                        .Text = ""
                        .Replacement.Text = ""
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = True
                        .MatchWildcards = False
                        .Execute Replace:=wdReplaceAll
                    End With
                    ActiveDocument.Styles(CStr(nm)).Delete
                Case Else
                    ' Character / table / list styles can't become the paragraph style Normal --
                    ' keep the prior rule: delete only if the style is not in use.
                    With ActiveDocument.Content.Find
                        .ClearFormatting
                        .Style = CStr(nm)
                        .Execute findText:="", Format:=True
                        If .found = False Then ActiveDocument.Styles(CStr(nm)).Delete
                    End With
            End Select
        End If
        On Error GoTo 0
    Next nm
    
End Sub   '*** end of Lp_Remove_All_Styles_Except_Lp_Styles macro ***

Function Sh_Document_Has_Prodnotes(ByVal targetDoc As Document) As Boolean
'
' True when the document actually USES the Prodnote style -- in the body text or inside a
' table. The style merely EXISTING is not enough to go on: every LP document defines Prodnote
' whether anything is styled with it or not, so existence would nearly always be True.
'
' This is the one place the test lives. Sh_Set_Prodnote_Style_Visibility decides whether to
' keep or delete the style by it, Sh_Strip_Prodnote_Enclosing_Quotes decides whether there is
' anything to unwrap by it, and Sh_Convert_XML_File_To_Word_Document decides whether to show
' the prodnote note at the end of a conversion by it.
'
' WALKS THE PARAGRAPHS. Version 1.0 used a style-aware Find, which worked in isolation and
' failed at the end of a real conversion: Jerry saw the note not appear on a document that
' plainly had red prodnotes in it (8/3/2026). Word's Find settings are STICKY - the ones a
' caller does not set explicitly carry over from whatever used Find last, and by that point in
' a conversion Sh_Color_Dollar_PG_Red and others have been setting them all the way through.
' A walk depends on nothing but the document, and it is the same test
' Sh_Delete_Prodnote_Paragraphs has always used. It stops at the first hit, so on a document
' that has prodnotes it costs almost nothing.
'
' Version: 2.0  Date: 8/3/2026 - walks the paragraphs instead of using Find, which could
'                               silently answer False after other macros had left their own
'                               settings on the Find object
' Version: 1.0  Date: 8/1/2026
'
    Dim st As Style
    Dim para As Paragraph

    On Error Resume Next
    Set st = targetDoc.Styles("Prodnote")
    On Error GoTo 0
    If st Is Nothing Then Exit Function     ' no Prodnote style at all, so nothing can use it

    For Each para In targetDoc.Paragraphs
        If StrComp(para.Style.NameLocal, "Prodnote", vbTextCompare) = 0 Then
            Sh_Document_Has_Prodnotes = True
            Exit Function
        End If
    Next para

End Function   '*** end of Sh_Document_Has_Prodnotes function ***

Sub Sh_Set_Prodnote_Style_Visibility()
'
' Keeps "Prodnote" in the Styles pane ONLY when the document actually contains at least one
' paragraph styled Prodnote. When the document uses it, the style is made visible; when
' nothing uses it, the style is DELETED from the document. Shared by Large Print and Braille.
'
' Why delete rather than just hide (Style.Visibility = semiHidden): semiHidden only removes a
' style under the "Recommended" styles-pane filter. Large print and braille documents set that
' filter, but a document with the Normal template attached is configured by
' MS_Set_Word_Config_For_New_Install, which sets the pane to "All styles" (wdShowFilterStylesAll)
' -- and under "All styles" a semiHidden style still shows. So hiding cannot work there, and the
' behavior would depend on which template is attached. Deleting the unused style removes it from
' the pane under ANY filter and for ANY template.
'
' Deleting is safe: it happens only when the usage test finds no paragraph in the style, so no
' text reverts. The style is not lost -- it returns from the LP template on the next attach, and
' the DAISY/NIMAS converter re-creates it when it emits prodnotes.
'
' Usage is tested with a style Find (fast; this also runs on every document open) rather than a
' VBA paragraph loop.
'
' Version: 2.2  Date: 7/24/2026 - normalize a used style to Priority 1 (was 2)
' Version: 2.1  Date: 7/24/2026 - a used style is normalized to Priority 2 with semiHidden cleared, so files saved by earlier builds stop showing Prodnote at high priority / marked hide-until-used
' Version: 2.0  Date: 7/24/2026 - renamed Lp_ -> Sh_ (shared). Deletes the unused style instead of only setting semiHidden, so removal works for any attached template / pane filter, not just Recommended
' Version: 1.1  Date: 7/23/2026 - usage tested with Find instead of a paragraph loop (open-time speed)
' Version: 1.0  Date: 7/23/2026
'
    Dim st As Style
    Dim used As Boolean

    On Error Resume Next
    Set st = ActiveDocument.Styles("Prodnote")
    On Error GoTo 0
    If st Is Nothing Then Exit Sub          ' no Prodnote style in this document -- nothing to do

    used = Sh_Document_Has_Prodnotes(ActiveDocument)

    On Error Resume Next
    If used Then
        ' Normalize a used style so it shows as a plain Prodnote entry, matching the LP
        ' template, and fix up documents saved by earlier builds that hid it or gave it a
        ' high priority: Priority 1 (Word displays this; the .docx stores uiPriority 0) and
        ' cleared semiHidden. (VBA cannot clear <w:unhideWhenUsed/> on a pre-existing style;
        ' that legacy flag is corrected only by re-attaching the template, which now carries
        ' a clean Prodnote definition.)
        st.Priority = 1
        st.Visibility = False
    Else
        st.Delete                           ' unused -> remove it from the pane entirely
    End If
    Err.Clear
    On Error GoTo 0

End Sub   '*** end of Sh_Set_Prodnote_Style_Visibility macro ***

Sub Sh_Delete_Prodnote_Paragraphs()
'
' Deletes every paragraph styled "Prodnote" -- in the body text and inside tables.
' Prodnote paragraphs are produced by the DAISY/NIMAS converter (see
' Sh_Tag_Prodnotes_As_Prodnote_Style) and by the "Prodnote" style in the LP template.
'
' Requires an open document with at least one paragraph actually styled Prodnote. Otherwise
' the user is told and nothing happens. The deletion is confirmed Yes/No before anything is
' removed. Shared: usable from both the Large Print and Braille Macros tabs, so it does NOT
' require the LP template -- only that a document is open.
'
' Version: 1.4  Date: 7/24/2026 - message-box titles now follow the document type: "VistaType LP"
'                                 for a large print document, "Braille Macros" otherwise
' Version: 1.3  Date: 7/23/2026 - dropped the LP-template guard so it works on Braille
'                                 documents too; now only checks that a document is open
' Version: 1.2  Date: 7/23/2026 - after the deletion the Prodnote style is hidden again
'                                 (Visibility = True -> semiHidden), since Word drops
'                                 semiHidden from the local definition once a style is used
' Version: 1.1  Date: 7/22/2026 - now guards on Lp_Is_Lp_Template_Attached (open doc + LP
'                                 template) and tests Prodnote USAGE rather than the style
'                                 merely existing; "No prodnotes found" when none are used
' Version: 1.0  Date: 7/22/2026
'
    ' Guard: a document must be open. No LP-template requirement -- this is shared with the
    ' Braille tab, and Braille documents can carry prodnotes too.
    Application.Run MacroName:="Sh_Is_Doc_Open"

    ' This macro is shared. Title its message boxes for the document type: "VistaType LP" for
    ' a large print document, "Braille Macros" otherwise (a braille document, or a plain
    ' converted document that is neither) -- so a braille user does not see a "VistaType LP" title.
    Dim isLP As Boolean
    isLP = (Lp_Is_The_Attached_Template_LP = True)

    ' --- 1. Collect the paragraphs actually styled Prodnote ---
    ' This doubles as the usage test: the style existing in the template is not enough, and
    ' since every LP document now defines "Prodnote", existence would almost always be True.
    ' Comparing style names is safe even if no Prodnote style exists (nothing matches).
    ' Collecting up front also avoids deleting while walking the collection, which would
    ' skip paragraphs (For Each is O(n); indexed Paragraphs(i) would be O(n^2)).
    Dim para As Paragraph
    Dim marked As Collection
    Dim i As Long
    Dim deleted As Long
    Dim su_Prev As Boolean

    Set marked = New Collection
    For Each para In ActiveDocument.Paragraphs
        If StrComp(para.Style.NameLocal, "Prodnote", vbTextCompare) = 0 Then
            marked.Add para.Range
        End If
    Next para

    If marked.count = 0 Then
        MsgBox "No prodnotes found", vbInformation, IIf(isLP, "VistaType LP (228)", "Braille Macros")
        Exit Sub
    End If

    ' --- 2. Confirm before deleting anything ---
    If MsgBox("Do you want to delete all paragraphs styled as Prodnote?", _
              vbYesNo + vbQuestion, IIf(isLP, "VistaType LP (229)", "Braille Macros")) <> vbYes Then
        Exit Sub
    End If

    ' --- 3. Delete back-to-front so the remaining ranges stay valid ---
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim total As Long
    total = marked.count

    For i = total To 1 Step -1
        On Error Resume Next
        marked(i).Delete
        Err.Clear
        On Error GoTo 0
    Next i

    ' A Prodnote paragraph that is the ONLY content of a table cell (or the final paragraph
    ' of the document) keeps its paragraph mark -- Word will not delete an end-of-cell
    ' marker. Its text goes, but an EMPTY paragraph still styled Prodnote survives, which
    ' would keep the style in use and therefore visible in the Styles pane. Reset those to
    ' Normal so no Prodnote-styled paragraph is left behind.
    Dim residual As Long
    For Each para In ActiveDocument.Paragraphs
        If StrComp(para.Style.NameLocal, "Prodnote", vbTextCompare) = 0 Then
            On Error Resume Next
            para.Style = ActiveDocument.Styles(wdStyleNormal)
            Err.Clear
            On Error GoTo 0
            residual = residual + 1
        End If
    Next para

    deleted = total - residual

    ' Re-hide the Prodnote style now that nothing uses it (Word drops <w:semiHidden/> from
    ' the local definition once a style has been used, so it will not re-hide itself).
    Application.Run MacroName:="Sh_Set_Prodnote_Style_Visibility"

    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh

    Dim msg As String
    msg = deleted & " paragraph(s) styled as Prodnote were deleted."
    If residual > 0 Then
        msg = msg & vbCr & vbCr & residual & " more were the only content of a table cell, " & _
              "so the cell could not be removed. Their text was deleted and the empty " & _
              "paragraph reset to Normal."
    End If
    MsgBox msg, vbInformation, IIf(isLP, "VistaType LP (230)", "Braille Macros")

End Sub   '*** end of Sh_Delete_Prodnote_Paragraphs macro ***

Sub Dx_Change_Prodnotes_To_Transcriber_Notes()
'
' Braille: restyle every paragraph currently styled "Prodnote" -- in the body text and inside
' tables -- to the "TranscriberNote" style. Prodnote paragraphs come from the DAISY/NIMAS
' converter; in a braille document they should instead carry the braille TranscriberNote style.
'
' Requires an open document with a BANA Braille template attached (the attached-template name
' begins with "BANA Braille"; the ".dot"/".dotx" extension is part of the name and does not
' affect the begins-with test). The document must contain a "Prodnote" style and at least one
' paragraph that uses it; otherwise the user is told and nothing happens. The change is
' confirmed Yes/No before anything is restyled.
'
' Version: 1.0  Date: 7/24/2026
'
    Dim tmplName As String
    Dim stProd As Style
    Dim stTrans As Style
    Dim rng As Range
    Dim used As Boolean
    Dim su_Prev As Boolean

    ' 1. A document must be open.
    Application.Run MacroName:="Sh_Is_Doc_Open"

    ' 2. A BANA Braille template must be attached (name begins with "BANA Braille").
    On Error Resume Next
    tmplName = ActiveDocument.AttachedTemplate.Name
    On Error GoTo 0
    If StrComp(Left(tmplName, Len("BANA Braille")), "BANA Braille", vbTextCompare) <> 0 Then
        MsgBox "This macro needs a BANA Braille template attached to the document." & vbCr & vbCr & _
               "Attach a BANA Braille template and try again.", vbExclamation, "Braille Macros"
        Exit Sub
    End If

    Sh_Save_User_Position   ' after the guards, so an early Exit Sub above leaves nothing pending

    ' 3. The document must contain a "Prodnote" style.
    On Error Resume Next
    Set stProd = ActiveDocument.Styles("Prodnote")
    On Error GoTo 0
    If stProd Is Nothing Then
        MsgBox "This document does not contain a ""Prodnote"" style." & vbCr & vbCr & _
               "There is nothing to change.", vbInformation, "Braille Macros"
        Exit Sub
    End If

    ' 4. ...and at least one paragraph must actually use it.
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Style = stProd
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        .Execute
        used = .found
    End With
    If Not used Then
        MsgBox "No prodnotes found", vbInformation, "Braille Macros"
        Exit Sub
    End If

    ' 5. The target "TranscriberNote" style must exist (supplied by the BANA template).
    On Error Resume Next
    Set stTrans = ActiveDocument.Styles("TranscriberNote")
    On Error GoTo 0
    If stTrans Is Nothing Then
        MsgBox "This document does not contain a ""TranscriberNote"" style." & vbCr & vbCr & _
               "Attach a BANA Braille template that defines the TranscriberNote style and try again.", _
               vbExclamation, "Braille Macros"
        Exit Sub
    End If

    ' 6. Confirm before changing anything.
    If MsgBox("Change all paragraphs styled Prodnote into Transcriber Notes?", _
              vbYesNo + vbQuestion, "Braille Macros") <> vbYes Then
        Exit Sub
    End If

    ' 7. Restyle Prodnote -> TranscriberNote across the whole story (body text and tables) in a
    '    single style Find/Replace, which reassigns the paragraph style everywhere it is used.
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    With ActiveDocument.Content.Find
        .ClearFormatting
        .Style = stProd
        .Replacement.ClearFormatting
        .Replacement.Style = stTrans
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With

    ' Prodnote is now unused; apply the shared rule so it drops out of the Styles pane.
    Application.Run MacroName:="Sh_Set_Prodnote_Style_Visibility"

    Application.ScreenUpdating = su_Prev

    Sh_Return_User_To_Start_Position

    MsgBox "All paragraphs styled Prodnote have been changed to Transcriber Notes.", _
           vbInformation, "Braille Macros"

End Sub   '*** end of Dx_Change_Prodnotes_To_Transcriber_Notes macro ***

Sub Lp_RemoveHeadAndFoot()

' from: https://word.tips.net/T001777_Deleting_All_Headers_and_Footers.html

    Dim oSec As Section
    Dim oHead As HeaderFooter
    Dim oFoot As HeaderFooter

    For Each oSec In ActiveDocument.Sections
        For Each oHead In oSec.headers
            If oHead.Exists Then oHead.Range.Delete
        Next oHead

        For Each oFoot In oSec.Footers
            If oFoot.Exists Then oFoot.Range.Delete
        Next oFoot
    Next oSec
End Sub   '*** end of Lp_RemoveHeadAndFoot macro ***

Sub Lp_SetSelectedTableBorderWeight()
'
' called from Table Tools Menu - the same job as Lp_Set_Table_Border_Weights, but only for
' the table the cursor is in.
'
' Version: 2.0  Date: 7/30/2026 - weight now comes from Lp_Border_Weight_For_Base_Font, the
'                                 one place that decides it, so this matches the box styles,
'                                 the reference page number border and every other table.
'                                 It used to match the base size as TEXT against fifteen
'                                 literals ("14","16",..."42"), so an odd size, or anything
'                                 above 42, fell through every branch and the button quietly
'                                 did nothing. Four near-identical blocks became one loop.
' Version: 1.0  Date: 10/16/23 - extended 225pt range to inclue Lp_Base_Font_Size = "20"
' Version: 1.0  Date: 3/6/2023

    Dim CurrentTable As Table
    Dim targetWeight As WdLineWidth
    Dim edge As Variant

    If Selection.Tables.count = 0 Then Exit Sub
    Selection.Tables(1).Select

    targetWeight = Lp_Border_Weight_For_Base_Font

    For Each CurrentTable In Selection.Tables
        With CurrentTable

            ' A table can be missing an edge (a single cell has no inside borders), and
            ' asking for one that is not there raises. Skip it and carry on, as before.
            On Error Resume Next

            For Each edge In Array(wdBorderLeft, wdBorderRight, wdBorderTop, _
                                   wdBorderBottom, wdBorderHorizontal, wdBorderVertical)
                With .Borders(edge)
                    .LineStyle = wdLineStyleSingle
                    .LineWidth = targetWeight
                    .Color = wdColorAutomatic
                End With
            Next edge

            .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
            .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
            .Borders.Shadow = False

            On Error GoTo 0

        End With
    Next CurrentTable

End Sub
Sub Lp_Convert_Ordinal_Numbers()
'
' converts superscript "st", "nd", "rd", and "th" to normal text size
'
' Version: 1.0  Date: 3/17/2024
'
    Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Superscript = True
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Superscript = False
    End With
    With Selection.Find
        .Text = "st"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Superscript = True
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Superscript = False
    End With
    With Selection.Find
        .Text = "nd"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Superscript = True
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Superscript = False
    End With
    With Selection.Find
        .Text = "rd"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
        Selection.Find.ClearFormatting
    With Selection.Find.Font
        .Superscript = True
    End With
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find.Replacement.Font
        .Superscript = False
    End With
    With Selection.Find
        .Text = "th"
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub   '*** end of Lp_Convert_Ordinal_Numbers ***

Sub Lp_ResizePicturesAndShapesToFitPageWidthAndPageHeight()
'
' Loops through all shapes in the document.  Checks to see if they're too wide, and if they are, resizes them.
' Adapted From: https://superuser.com/questions/570121/fitting-images-to-documents-margins-in-a-docx-file
'
' Version: 1.0  Date: 4/5/2024
'
    Dim Shapes As Integer
    Dim inlines As Integer
    Dim WidthAvail As Single

    Shapes = ActiveDocument.Shapes.count
    inlines = ActiveDocument.InlineShapes.count
    
    Dim WidthPercent As String
    Dim ShapeLoop As Integer
    Dim InLineLoop As Integer
    
    'Sets the variables to loop through all shapes in the document, one for shapes and one for inline shapes.
    'Calculate usable width of page
    With ActiveDocument.PageSetup
        WidthAvail = .PageWidth - .LeftMargin - .RightMargin
    End With
    
    For ShapeLoop = 1 To Shapes
        'MsgBox Prompt:="Shape " & ShapeLoop & " width: " & ActiveDocument.Shapes(ShapeLoop).Width
        If ActiveDocument.Shapes(ShapeLoop).Width > WidthAvail Then
            ActiveDocument.Shapes(ShapeLoop).LockAspectRatio = msoTrue 'forces shape to fit within all margins
            ActiveDocument.Shapes(ShapeLoop).Width = WidthAvail
        End If
    Next ShapeLoop
    
    'Loops through all shapes in the document.  Checks to see if they're too wide, and if they are, resizes them.
    For InLineLoop = 1 To inlines
        If ActiveDocument.InlineShapes(InLineLoop).Width > WidthAvail Then
            ActiveDocument.InlineShapes(InLineLoop).Width = WidthAvail
            ' When Picture is be too tall to fit within the top and bottom margins
            If ActiveDocument.InlineShapes(InLineLoop).ScaleHeight > ActiveDocument.InlineShapes(InLineLoop).ScaleWidth Then
                WidthPercent = ActiveDocument.InlineShapes(InLineLoop).ScaleWidth
                ActiveDocument.InlineShapes(InLineLoop).ScaleHeight = WidthPercent
            End If
        End If
LoopAgain:
    Next InLineLoop

End Sub   '*** end of macro Lp_ResizePicturesToFitPageWidthAndPageHeight ***

Sub Lp_Resize_Images()
    '
    ' Lp_Resize_Images Macro
    '
    ' Author: Jerry Whittaker - jerry@thewhittakers.org
    '
    ' Version: 1.5  Date: 11/18/2020 - added unload of Lp_Resize_Images_Form
    ' Version: 1.3  Date: 2/11/2019 - complete rewirte - deletion of images from tables corrected
    ' Version: 1.2  Date: 1/10/2019
    ' Version: 1.1  Date: 12/9/2018
    ' Version: 1.0  Date: 1/22/2016

    Application.Run MacroName:="Sh_Is_Doc_Open"
    
    Dim i As Long
        
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating of
     
    ' A single image is selected - selection can be within text or in a table
    If Selection.Type = wdSelectionInlineShape Then
        With Selection
            For i = 1 To .InlineShapes.count
            With .InlineShapes(i)
                .ScaleHeight = Val(Lp_Pic_Percent)
                .ScaleWidth = Val(Lp_Pic_Percent)
            End With
            Next i
        End With
        Selection.Collapse 'clear selection
    ElseIf Lp_Pic_All_Selectd = "S" And Selection.Type = wdSelectionNormal Then  ' either a selected image or range (range may include a table)
            Application.Run MacroName:="Lp_Copy_To_Temp_Doc"
            With ActiveDocument
                For i = 1 To .InlineShapes.count
                With .InlineShapes(i)
                    .ScaleHeight = Val(Lp_Pic_Percent)
                    .ScaleWidth = Val(Lp_Pic_Percent)
                End With
                    Next i
            End With
            Selection.EndKey Unit:=wdStory  ' move to the bottom of the document
            Selection.Delete Unit:=wdCharacter, count:=1  ' delete ending para mark
            Application.Run MacroName:="Lp_Copy_From_Temp_Doc"
    ElseIf Selection.Information(wdWithInTable) Then   ' Table is selected
        Selection.Tables(1).Select 'Select the whole table
        Application.Run MacroName:="Lp_Copy_To_Temp_Doc"
            With ActiveDocument
                For i = 1 To .InlineShapes.count
                With .InlineShapes(i)
                    .ScaleHeight = Val(Lp_Pic_Percent)
                    .ScaleWidth = Val(Lp_Pic_Percent)
                End With
                    Next i
            End With
        Selection.WholeStory
        Selection.Copy 'copy the selected text to the clipboard
        ActiveDocument.Close SaveChanges:=False 'close the temp doc without saving
        Selection.rows.Delete
        Selection.TypeBackspace ' delete the table in the original document
        Selection.Delete Unit:=wdCharacter, count:=1
        Selection.Paste 'paste the clipboard back into the original document
    ElseIf Lp_Pic_All_Selectd = "A" Then  ' all images in the document
        Sh_Save_User_Position
        With ActiveDocument
            For i = 1 To .InlineShapes.count
            With .InlineShapes(i)
                .ScaleHeight = Val(Lp_Pic_Percent)
                .ScaleWidth = Val(Lp_Pic_Percent)
            End With
            Next i
        End With
        Sh_Return_User_To_Start_Position
    Else
        MsgBox "Select a specific image, a text range containing images (including tables with images) or select a table containing images.", , "VistaType LP (142)"
    End If
   
    Selection.Collapse 'clear selection
    Application.ScreenUpdating = su_Prev ' Turn screen updating on
    Application.ScreenRefresh
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    ActiveDocument.UndoClear
    
    Unload Lp_Resize_Images_Form
    
End Sub  '***** end of Lp_Resize_Images Macro *****

Sub Lp_SetPicturesToInlineAndLockAspectRatio()
'
' Version: 1.0  Date: 4/21/2025
'
    Dim shp As Shape
    Dim ilShp As inlineShape

    ' Convert floating shapes to inline and lock their aspect ratio
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            ' Convert the shape to an inline shape
            Set ilShp = shp.ConvertToInlineShape
            ' Lock the aspect ratio
            ilShp.LockAspectRatio = True
        End If
    
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            ' Convert the shape to an inline shape
            Set ilShp = shp.ConvertToInlineShape
            ' Lock the aspect ratio
            ilShp.LockAspectRatio = True
        End If
    Next shp
    
End Sub   '*** end of macro Lp_SetPicturesToInlineAndLockAspectRatio ***

Sub Lp_Delete_Zero_Width_Spaces()
'
' Version: 1.0 Date: 6/18/2025
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8203) 'zero width space (U+200B) common in AI generated text
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub   '*** end of Lp_Delete_Zero_Width_Spaces macro ***

Sub Lp_Table_Apply_Character_Case_To_Row_Headers()
'
' Version: 1.1 Date:9/19/2025 - removed lower case setting for row/column headers
' Version: 1.0 Date: 8/9/2025
'
    Dim tbl As Table
    Dim c   As cell
    Dim rng As Range
    Dim i   As Long

    ' Point to the first table
    Set tbl = ActiveDocument.Tables(1)
    
    ' Walk each cell in row 1
    For Each c In tbl.rows(1).Cells
        Set rng = c.Range
        rng.End = rng.End - 1  ' Exclude the cell marker
        
        ' Choose action based on Lp_GP_String_3
        If InStr(Lp_GP_String_3, "P") > 0 Then
            rng.Select
            Sh_Apply_Title_Case_Capitalization
            
        ElseIf InStr(Lp_GP_String_3, "U") > 0 Then
            ' Uppercase only a–z, keep formatting
            For i = 1 To rng.Characters.count
                With rng.Characters(i)
                    If .Text Like "[a-z]" Then .Text = UCase(.Text)
                End With
            Next i
        End If
    Next c

End Sub  '*** end of Lp_Table_Apply_Character_Case_To_Row_Headers***

Sub Lp_Table_Apply_Character_Case_To_Column_Headers()
'
' Version: 1.1 Date: 9/19/1015 - removed lower case setting
' Version: 1.0 Date: 8/9/2025
'
    Dim tbl  As Table
    Dim cel  As cell
    Dim rng  As Range
    Dim i    As Long

    ' Reference the first table in the document
    Set tbl = ActiveDocument.Tables(1)
    
    ' Loop through each cell in column 1
    For Each cel In tbl.Columns(1).Cells
        Set rng = cel.Range
        rng.End = rng.End - 1    ' Exclude end-of-cell marker

        ' Decide action based on Lp_GP_String_3
        If InStr(Lp_GP_String_3, "P") > 0 Then
            rng.Select
            Sh_Apply_Title_Case_Capitalization

        ElseIf InStr(Lp_GP_String_3, "U") > 0 Then
            ' Uppercase only a–z, preserving formatting
            For i = 1 To rng.Characters.count
                With rng.Characters(i)
                    If .Text Like "[a-z]" Then .Text = UCase(.Text)
                End With
            Next i
        End If
    Next cel

End Sub  '*** end of Lp_Table_Apply_Character_Case_To_Column_Headers***

'-----------------------------------------------------------------------------------
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
' ** Top of Project DN=DAISY-NIMAS
'                            Sh=Shared between BANA Macros and Large Print
'                             MS=Microsoft Word **
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
'------------------------------------------------------------------------------------
Sub DN_Menu_Starter()
    '
    ' Version: 1.0  Date: 3/2/2026
    '
    Application.Run MacroName:="Sh_Is_Doc_Open"
    DN_Auto_Or_Manual_Form.Show

End Sub

Sub DN_Add_PgNo_Tags_To_DAISY_or_NIMAS()
'
' Version 1.2  Date: 3/3/2026 - added "Unload DN_Auto_Or_Manual_Form"
' Version 1.1  Date: 4-9-2015
'
    Unload DN_Auto_Or_Manual_Form

    Application.Run MacroName:="Sh_Is_Doc_Open"

    Load DN_Tag_Daisy_Nimas_Form
    DN_Tag_Daisy_Nimas_Form.Show
    
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    ActiveDocument.UndoClear

End Sub   '*** end of DN_Add_PgNo_Tags_To_DAISY_or_NIMAS macro ***

Sub DN_Remove_Para_Formatting_From_Text_Files()
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
' Version 1.1  Date: 4-9-15
'
' Removes paragraph marks which often appear at the end of every line
' in a document (e.g. Gutenburg Project Text Files). The paragraph marks
' often doubled to create a visual presentation of space between paragraphs.
' This macro removes the paragraph marks and preserves the paragraphs.
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    DN_Text_File_Para_Fix_Warning.Show

End Sub '***** End of DN_Remove_Para_Formatting_From_Text_Files Macro *****

Sub Sh_Remove_Hyperlinks(Optional ByVal target As Range)
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Strips the blue-and-underlined look, then deletes the links themselves.
'
' target  the range to work on. Omit it and the WHOLE DOCUMENT is stripped, which is what this
'         macro has always done and what its unconverted callers still expect.
'
'         That default is the point to understand before converting anything that calls this.
'         The old Kill_The_Hyperlinks pair confined it to a selection by copying that selection
'         into a temporary document first - the deletion below reaches every hyperlink in
'         whatever document is active, so the scratch document WAS the scope. That is real work
'         the round trip was doing, not plumbing (Jerry, 8/12/2026).
'
' Version: 2.0  Date: 8/12/2026 - takes a range
' Version: 1.0  Date: 1/31/2016
'
    Dim rng As Range
    Dim guard As Long

    If target Is Nothing Then
        Set rng = ActiveDocument.Content
    Else
        Set rng = target
    End If

    ' the blue and the underline
    With rng.Find
        .ClearFormatting
        .Font.Underline = wdUnderlineSingle
        .Font.Color = wdColorBlue
        .Replacement.ClearFormatting
        .Replacement.Font.Underline = wdUnderlineNone
        .Replacement.Font.Color = wdColorAutomatic
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

    ' the links themselves. Always index 1: deleting renumbers the collection, so walking it
    ' forwards would step over every other one. Bounded, because a collection that will not
    ' shrink must not spin for ever.
    Do While rng.Hyperlinks.count > 0 And guard < 5000
        rng.Hyperlinks(1).Delete
        guard = guard + 1
    Loop

 End Sub  '***** end of Sh_Remove_Hyperlinks macro ***

Sub Sh_Color_Dollar_PG_Red()
'
' Version: 1.3 Date: 7/5/2026 - added normal style to F&R
' Version: 1.3 Date: 2/8/2017
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    ' Apply Normal style + red color to replacement text
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Normal")
    Selection.Find.Replacement.Font.Color = wdColorRed
    
    With Selection.Find
        .Text = "$pg"
        .Replacement.Text = "$pg"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    ActiveDocument.UndoClear
    
End Sub  '*** Sh_Color_Dollar_PG_Red Macro ***

Sub Sh_Apply_Title_Case_Capitalization()
    '
    ' Version: 1.0  Date: 12/31/2019
    '
    ' applies title case capitalization to the first paragraph in the selection
    ' Adapted From: https://www.brainbell.com/tutorials/ms-office/Word/Apply_Proper_Capitalization.htm
    '
    Application.Run MacroName:="Sh_Is_Doc_Open"
    Dim strL As String
    Dim i As Integer
    If Selection.Type <> wdSelectionIP Then Selection.Collapse
    Selection.Paragraphs(1).Range.Select
    Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
    Selection.words(1).Case = wdTitleWord
    Selection.words(Selection.words.count).Case = wdTitleWord
    For i = 2 To Selection.words.count - 1
         strL = LCase(Trim(Selection.words(i)))
         If strL = "a" Or strL = "above" Or strL = "after" Or strL = "an" Or _
               strL = "and" Or strL = "as" Or strL = "at" Or strL = "below" Or _
               strL = "but" Or strL = "by" Or strL = "down" Or strL = "for" Or _
               strL = "from" Or strL = "in" Or strL = "into" Or strL = "of" Or _
               strL = "off" Or strL = "on" Or strL = "onto" Or strL = "or" Or strL = "yet" Or _
               strL = "out" Or strL = "over" Or strL = "the" Or strL = "to" Or strL = "nor" Or _
               strL = "under" Or strL = "up" Or strL = "with" Or strL = "is" Then
               Selection.words(i).Case = wdLowerCase
         Else
           Selection.words(i).Case = wdTitleWord
        End If
  Next i
  Selection.Collapse Direction:=wdCollapseEnd
    
End Sub   ' *** end of Sh_Apply_Title_Case_Capitalization macro ***

Sub Lp_Picture_Color_Change_Menu()
'
'Version 1.0  Date: 12/26/2023
'
    Unload Lp_Bakgrnd_Picture_Menu_Form
    Load Lp_Change_Image_Color_Form
    Lp_Change_Image_Color_Form.Show
    
End Sub '*** end of Lp_Picture_Color_Change_Menu macro ***

Sub Lp_Table_Row_Column_Header_Setup()
    '
    ' these routines are only used when table contains row headings, column headings or both
    ' - not used for tables without row and column headings
    '
    ' Version: 1.1  Date: 8/13/2025 - fixes for rotated tables and fixes for RC tables
    ' Version: 1.0  Date: 7/7/2025 - new code
    '
    Dim tbl As Table
    Dim rng As Range
    Dim col As Column
    Dim para As Paragraph
    
    Application.ScreenUpdating = False
    
     '++++++++++ Begin style setup for table types X and Y +++++++++++++++++++++
    If InStr(Lp_GP_String_3, "Y") > 0 Then
        Set tbl = ActiveDocument.Tables(1)
        
        'set style for table types X and Y
        tbl.Range.Style = "List 2" 'Apply the "List 2" style to the whole table
        
        ' Apply the "List" style to each paragraph in the first column
         With ActiveDocument.Tables(1).Columns(1)
            .Select
            Selection.Style = "List"
        End With
    End If
    '++++++++++ End set style for table types X and Y +++++++++++++++++++++
    
    ' Get the first table in the document
    Set tbl = Selection.Tables(1)

    '++++++++++ Begin Title Case, Lower Case, Upper Case Header changes +++++++++++++++++++++

    If InStr(Lp_GP_String_3, "X") > 0 Then 'table is rotated and both column and row headers
        Application.Run MacroName:="Lp_Table_Apply_Character_Case_To_Column_Headers"
        Application.Run MacroName:="Lp_Table_Apply_Character_Case_To_Row_Headers"
    End If

    If InStr(Lp_GP_String_3, "Y") > 0 Then ' table is rotated and now has column headers only
        Application.Run MacroName:="Lp_Table_Apply_Character_Case_To_Column_Headers"
    End If
        
    '++++++++++ End Title Case, Lower Case, Upper Case Header changes +++++++++++++++++++++

    '++++++++++ Begin Header Color  and Bold Settings +++++++++++++++
    Set tbl = ActiveDocument.Tables(1)
    
    If InStr(Lp_GP_String_3, "C") > 0 Then
        Dim cRed        As Integer        ' 1st digit of RGB color
        Dim cGreen      As Integer        ' 2nd digit of RGB color
        Dim cBlue       As Integer        ' 3rd digit of RGB color
        
        If InStr(Lp_GP_String_3, "2") > 0 Then      'Red
            cRed = 255
            cGreen = 0
            cBlue = 0
        ElseIf InStr(Lp_GP_String_3, "3") > 0 Then  'Orange
            cRed = 228
            cGreen = 90
            cBlue = 45
        ElseIf InStr(Lp_GP_String_3, "4") > 0 Then  'Blue
            cRed = 74
            cGreen = 93
            cBlue = 255
        ElseIf InStr(Lp_GP_String_3, "5") > 0 Then  ' Violet
            cRed = 192
            cGreen = 0
            cBlue = 249
        ElseIf InStr(Lp_GP_String_3, "6") > 0 Then  ' Green
            cRed = 0
            cGreen = 128
            cBlue = 0
        End If
        
        ' ++++++ Type X +++++++
        If InStr(Lp_GP_String_3, "X") > 0 Then 'table has both row and column headers
            tbl.Columns(1).Select
            With Selection.Font
                .Color = wdColorAutomatic
                If InStr(Lp_GP_String_3, "H") > 0 Then 'Bold wanted
                    With Selection.Font
                        .Bold = True
                    End With
                End If
            End With
            tbl.rows(1).Select
            With Selection.Font
                .Color = RGB(cRed, cGreen, cBlue)
            End With
            Exit Sub
        End If

        ' ++++++ Type Y +++++++
        If (InStr(Lp_GP_String_3, "Y") > 0) Then    ' table has column headers only
            tbl.Columns(1).Select
            With Selection.Font
                If InStr(Lp_GP_String_3, "1") > 0 Then
                    .Color = wdColorAutomatic
                Else
                    .Color = RGB(cRed, cGreen, cBlue)
                End If
                If InStr(Lp_GP_String_3, "H") > 0 Then 'Bold wanted
                    With Selection.Font
                        .Bold = True
                    End With
                End If
            End With
            Exit Sub
        End If
    End If
       
    '++++++++++ End Header Color  and Bold Settings +++++++++++++++

End Sub '*** end of Lp_Table_Row_Column_Header_Setup ***

Sub Lp_Table_Convert_NoRC_Table_To_List()
'
' Version: 1.4  Date: 7/26/2026 - returns the user to where the cursor was when the macro started
' Version: 1.2  Date: 8/15/2025 - remove para mark at top placed by Lp_Copy_To_Temp_Doc
'                               - added Application.Run MacroName:="Lp_Table_Style_InCell_Para_And_Image"
Dim su_Prev As Boolean
su_Prev = Application.ScreenUpdating
' Version: 1.1  Date: 814/2025  - remove 40% screen setting - added Application.ScreenUpdating = False
' Version: 1.0  Date: 8/8/2025
'
    Application.ScreenUpdating = False

    Sh_Save_User_Position

    Dim tbl As Table
    Dim tblRange As Range
    Dim TempFileName As String

    '**********  Start put a par above table '*********

    If Selection.Tables.count > 0 Then
        Set tbl = Selection.Tables(1)
        Set tblRange = tbl.Range
        tblRange.Cut
        Selection.TypeParagraph
        Selection.MoveUp Unit:=wdParagraph, count:=1
        Selection.MoveDown Unit:=wdParagraph, count:=1
        DoEvents
        Selection.Paste
    End If

    '**********  convert whole table to list style *********
    ActiveDocument.Tables(1).Select
    Selection.Style = "List"

     ' make any in-cell paragraphs List 4
    Application.Run MacroName:="Lp_Table_Style_InCell_Para_And_Image"
 
    'convert table to list
    ActiveDocument.Tables(1).ConvertToText Separator:=wdSeparateByParagraphs
    
    '********** start create transcriber note ********
    Selection.HomeKey Unit:=wdStory
    Selection.Delete 'remove para mark
    Application.Run "Lp_Table_Insert_Transcriber_Note"

    ' Store full path + file name of the active (temp) document
    Dim TempDocName As String
    TempFileName = ActiveDocument.fullName
    
    ' remove ending para mark
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1
    
    'copy the list
     DoEvents
     Selection.WholeStory
     DoEvents
     Selection.Copy
     DoEvents

     'open original doc and paste
     DoEvents
     Documents(Lp_GP_String_2).Activate
     DoEvents
     ActiveDocument.Tables(Lp_GP_Counter_1).Select
     DoEvents
     ActiveDocument.Tables(Lp_GP_Counter_1).Delete
     DoEvents
     'The table is gone, so the cursor now sits exactly where the converted block will land.
     'The transcriber note was put at the TOP of the temp document, so it is the first
     'paragraph of what we are about to paste - leave the user on it (Jerry, 7/26/2026).
     Sh_Set_Return_Position Selection.Range.start
     Selection.Paste
     DoEvents

     'Screen stays OFF through the temp-file cleanup below. Turning it on here and THEN
     'activating the temp document painted that document on screen - a splash of the table's
     'alternating row color - before it was closed again (Jerry, 7/26/2026).
     'delete temp file
     Documents(TempFileName).Activate
     ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
     DoEvents
     Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"

     Application.ScreenUpdating = su_Prev
     Sh_Return_User_To_Start_Position

End Sub   '*** end of Lp_Table_Convert_NoRC_Table_To_List ***

Sub Lp_Table_Cleanup_For_Roation_And_List()
    '
    ' Verskon: 1.2  Date: 8/13/2025 - removed automatic colors on Headers - added removal of multiple para marks before end of cell marker
    ' Version: 1.1  Date: 8/6/2025 ' added clear headers
    ' Version: 1.0  Date: 7/24/2025
    '
    '********** Begin Remove manual line breaks, tabs and extra spaces from table and clear headers********
    
    Application.ScreenUpdating = False
    Dim tbl As Table
    Dim cell As cell
    Dim txt As String

    ActiveDocument.Tables(1).Select

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = ""
        .Forward = False
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
   Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Tables(1).Select  'select entire table
     
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "^032"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Tables(1).Select  'select entire table
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}"
        .Replacement.Text = "^032"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Tables(1).Select  'select entire table
    
     ' Replace manual line breaks with space
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^l"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '***** begin clear all headers ****
    
    ' Reference the first table in the active document
    Set tbl = ActiveDocument.Tables(1)
    
     If InStr(Lp_GP_String_3, "X") > 0 Or InStr(Lp_GP_String_3, "Y") > 0 Then
        ' Loop through every cell in row 1
        For Each cell In tbl.rows(1).Cells
            With cell.Range.Font
                .Bold = False
                .Italic = False
                .Underline = wdUnderlineNone
            End With
        Next cell
    End If
        
     If InStr(Lp_GP_String_3, "X") > 0 Then
        ' Loop through every cell in row 1
        For Each cell In tbl.Columns(1).Cells
            With cell.Range.Font
                .Bold = False
                .Italic = False
                .Underline = wdUnderlineNone
            End With
        Next cell
    End If
    
    'TrimTrailingParasInTable1()
    Dim ur As UndoRecord
    'Dim tbl As table
    Dim cel As cell
    Dim r As Range, delR As Range
    'Dim txt As String
    Dim i As Long, n As Long

    If ActiveDocument.Tables.count = 0 Then Exit Sub
    Set tbl = ActiveDocument.Tables(1)

    ' Wrap in a single undo step.
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "Trim trailing paragraph marks in Table(1)"

    For Each cel In tbl.Range.Cells
        ' Work inside the cell, excluding the end-of-cell marker.
        Set r = cel.Range
        r.End = r.End - 1

        txt = r.Text
        If Len(txt) = 0 Then GoTo NextCell

        ' Count only trailing paragraph marks (Chr(13)).
        i = Len(txt)
        n = 0
        Do While i > 0
            If Mid$(txt, i, 1) = vbCr Then
                n = n + 1
                i = i - 1
            Else
                Exit Do
            End If
        Loop

        ' Delete all trailing paragraph marks before the end-of-cell marker.
        If n > 0 Then
            Set delR = r.Duplicate
            delR.start = delR.End - n
            delR.Delete
        End If
NextCell:
    Next cel

    ur.EndCustomRecord

    '***** end clear all headers ****
    
    '********** End Remove manual line breaks, tabs and extra spaces from table and clear headers********

End Sub   '*** end of Lp_Table_Cleanup_For_Roation_And_List ***

Sub Lp_Table_Transpose_Table()
    '
    ' Version: 1.0 Date: 9/2/2025 - new code
    '
    ActiveDocument.Tables(1).Select
    
    Dim src As Table, dst As Table
    Dim r As Long, c As Long
    Dim rows As Long, cols As Long
    Dim srcCellRng As Range, dstCellRng As Range
    Dim insertRng As Range
    Dim tblStyle As Style

    Set src = Selection.Tables(1)

    ' Guard: merged cells are not supported for a true transpose
    If src.Range.Cells.count <> (src.rows.count * src.Columns.count) Then
        MsgBox "This table has merged cells. Unmerge before transposing.", vbExclamation
        Exit Sub
    End If

    rows = src.rows.count
    cols = src.Columns.count
    Set tblStyle = src.Style

    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ' Insert destination table immediately after the source
    Set insertRng = src.Range.Duplicate
    insertRng.Collapse wdCollapseEnd
    insertRng.InsertParagraphAfter
    insertRng.Collapse wdCollapseEnd

    Set dst = insertRng.Tables.Add(Range:=insertRng, NumRows:=cols, NumColumns:=rows)
    dst.Style = tblStyle

    ' Fill destination with formatted content (keeps inline images)
    For r = 1 To rows
        For c = 1 To cols
            Set srcCellRng = src.cell(r, c).Range.Duplicate
            ' Trim end-of-cell marker (Chr(13) + Chr(7))
            srcCellRng.End = srcCellRng.End - 1

            Set dstCellRng = dst.cell(c, r).Range.Duplicate
            dstCellRng.End = dstCellRng.End - 1

            ' Copy full formatted content without using the clipboard
            dstCellRng.FormattedText = srcCellRng
        Next c
    Next r

    ' Remove the original table; the transposed one remains in its place
    src.Delete

    ' Clean up the extra paragraph we inserted (optional)
    If dst.Range.Previous Is Nothing Then
        ' nothing to clean
    ElseIf dst.Range.Previous.Text = vbCr Then
        dst.Range.Previous.Delete
    End If

    Application.ScreenUpdating = su_Prev

End Sub   '*** end of Lp_Table_Transpose_Table ***

Sub Lp_Table_Fill_Empty_Cells()

    'Version 1.0 Date: 7/26/2025

    Dim tTable As Table
    Dim cCell As cell
    Dim sTemp As String
    Dim bReplaced As Boolean
    Lp_GP_Boolean_1 = False 'was blank cell replaced?
    
    ActiveDocument.Tables(1).Select
    
    If InStr(Lp_GP_String_3, "M") > 0 Then
        sTemp = Chr(151) 'em dash
    End If
    
    If InStr(Lp_GP_String_3, "N") > 0 Then
        sTemp = "N/A"    'not available
    End If

   If Selection.Information(wdWithInTable) Then
        Set tTable = Selection.Tables(1)
        For Each cCell In tTable.Range.Cells
            'An apparently empty cell contains an end of cell marker
            If Len(Trim(cCell.Range.Text)) < 3 Then
                cCell.Range = sTemp
                bReplaced = True
            End If
        Next
    End If
    
    Lp_GP_Boolean_1 = bReplaced
    
End Sub   '****** End of Lp_Table_Fill_Empty_Cells *******

Sub Lp_ValidateTableIntegrityForListOrRotation()
'
' Version 1.0  Date: 8/4/2025
'
    Dim tbl        As Table
    Dim selRng     As Range
    Dim c          As cell
    Dim rw         As row
    Dim i          As Long
    Dim hasMerge   As Boolean
    Dim nonUniform As Boolean
    Dim totalCols  As Long
    Dim totalRows  As Long

    '----- 1. Identify the innermost table under the cursor -----
    Set selRng = Selection.Range
    If selRng.Information(wdWithInTable) Then
        ' Use Tables.Count to pick the deepest/nested table
        Set tbl = selRng.Tables(selRng.Tables.count)
    ElseIf Selection.Tables.count > 0 Then
        Set tbl = Selection.Tables(Selection.Tables.count)
    Else
        MsgBox "Place the cursor in a table or select one first.", vbExclamation
        Exit Sub
    End If

    totalCols = tbl.Columns.count
    totalRows = tbl.rows.count
    hasMerge = False
    nonUniform = False

    '----- 2. Detect merged cells via hidden interior borders -----
    For Each c In tbl.Range.Cells
        ' Horizontal merge test (right border)
        If c.ColumnIndex < totalCols Then
            If c.Borders(wdBorderRight).LineStyle = wdLineStyleNone Then
                hasMerge = True: Exit For
            End If
        End If
        ' Vertical merge test (bottom border)
        If c.rowIndex < totalRows Then
            If c.Borders(wdBorderBottom).LineStyle = wdLineStyleNone Then
                hasMerge = True: Exit For
            End If
        End If
    Next c

    '----- 3. Check for non-uniform rows by index loop -----
    For i = 1 To totalRows
        On Error Resume Next
        Set rw = tbl.rows(i)
        If Err.Number <> 0 Then
            ' If we can’t access this row, treat as non-uniform and clear error
            nonUniform = True
            Err.Clear
            On Error GoTo 0
            Exit For
        End If
        On Error GoTo 0

        If rw.Cells.count <> totalCols Then
            nonUniform = True
            Exit For
        End If
    Next i

    '----- 4. Invoke badtable if any check failed -----
    If hasMerge Or nonUniform Then
      Application.Run MacroName:="Lp_Table_Convert_Table_Format_Error"
    End If

End Sub  '*** end of Lp_ValidateTableIntegrityForListOrRotation ***

Sub Lp_DoesRangeHaveATOCStyle()
'
' Version: 1.0  Date 8/4/2025
'
    Dim rng As Range
    Dim para As Paragraph
    Dim containsTOC As Boolean
    
    Set rng = Selection.Range
    containsTOC = False
    
    For Each para In rng.Paragraphs
        If Left(para.Style.NameLocal, 3) = "TOC" Then
            containsTOC = True
            Exit For
        End If
    Next para
    
    If containsTOC = False Then
        MsgBox "The selected range contains NO TOC styles.", , "VistaType LP  (203)"
        End
    End If
    
End Sub   '*** end of Lp_DoesRangeHaveATOCStyle ***

Sub Lp_Table_Insert_Transcriber_Note()
'
'  Version: 1.1  Date: 10/29/2025 - added Lp_GP_Boolean_1 = True only when blank cell has been filled wiht "N/A" Or em dash
'  Version: 1.0  Date: 8/12/2025
'
    Selection.Font.Bold = True
    Selection.TypeText Text:="Note"
    Selection.Font.Bold = False
    
    If InStr(Lp_GP_String_3, "L") > 0 Then
        Selection.TypeText Text:=": The original table has been changed into a list."
    Else
        Selection.TypeText Text:=": The original table has been rotated (transposed)."
    End If

    If InStr(Lp_GP_String_3, "M") > 0 And Lp_GP_Boolean_1 = True Then 'empty cells were filled
        Selection.TypeText Text:=" Empty table cells are shown as " + ChrW(34) + ChrW(&H2014) + ChrW(34) + "."
    End If
    If InStr(Lp_GP_String_3, "N") > 0 And Lp_GP_Boolean_1 = True Then 'empty cells were filled
        Selection.TypeText Text:=" Empty table cells are shown as " + ChrW(34) + "N/A" + ChrW(34) + "."
    End If
    
    If Sh_Style_Exists(ActiveDocument, "Box Blue") Then
    Selection.Style = ActiveDocument.Styles("Box Blue")
    End If
    Selection.ParagraphFormat.KeepWithNext = wdToggle
    
End Sub   '*** end of Lp_Table_Insert_Transcriber_Note ***

Sub Lp_Table_Is_R1C1_Empty()
'
' Version 1.0  Date: 8/12/2025

    Dim tbl As Table
    Dim cellText As String

    ' Assume cursor is in or table is selected
    Set tbl = Selection.Tables(1)

    ' Get raw cell text
    cellText = tbl.cell(1, 1).Range.Text

    ' Remove end-of-cell marker and all non-printing characters
    cellText = Replace(cellText, Chr(7), "")
    cellText = Replace(cellText, vbCr, "")
    cellText = Replace(cellText, vbLf, "")
    cellText = Replace(cellText, Chr(160), "") ' Non-breaking space
    cellText = Trim(cellText)

    ' Check if cell is effectively empty
    If Len(cellText) = 0 Then
        MsgBox "The table cell in row 1 column 1 is empty or contains only spaces." & _
        " The cell must contain a header which describes the content of the column it heads.", , "VistaType LP  (205)"
        End
    End If

End Sub   '*** end of Lp_Table_Is_R1C1_Empty ***

Sub Lp_Table_Convert_RC_Table_To_List()
'
' Version: 1.6  Date: 7/26/2026 - returns the user to where the cursor was when the macro started
' version: 1.4  Date: 8/16/2025 - added clear clipboard
' Version: 1.3  Date: 815/2025 - added Application.Run MacroName:="Lp_Table_Style_InCell_Para_And_Image"
' Version: 1.2  Date: 8/15/2025 - remove para mark at top placed by Lp_Copy_To_Temp_Doc
' Version: 1.1  Date: 8/13/2025 - new paste routine - : + space to color automatic
' Version: 1.0  Date: 8/12/2025
'
    Dim tbl As Table
    Dim i As Long
    Dim totalRows As Long
    Dim doc As Document
    Dim nRows As Long, nCols As Long
    Dim r As Long, c As Long
    Dim hdr As Range, tgtCell As Range, ins As Range
    Dim delim As String
    Dim undoOn As Boolean
    Dim lastRow As Long, lastCol As Long
    Dim para As Paragraph
    Set tbl = ActiveDocument.Tables(1)
    Dim tblc As cell
    totalRows = tbl.rows.count
    Dim fileName As String
    Dim TempFileName As String
    Dim d As Document

    Sh_Save_User_Position

    ' Step 1: Modify Row 1 – add ":" to each cell
    For i = 2 To totalRows
        With tbl.cell(i, 1).Range
            ' Move the end back one hit to stay inside the cell (avoiding the end-of-cell marker)
            .MoveEnd Unit:=wdCharacter, count:=-1
            ' This adds the colon without destroying images or formatting
            .InsertAfter ":"
        End With
    Next i

    ' Step 2: Modify Column 1, Rows 2 to last – safely add ":"
        For i = 2 To totalRows
            With tbl.cell(i, 1).Range
                ' 1. Pull the end of the range back by 1 to skip the "End of Cell" marker
                ' If you don't do this, the colon might appear on a new line or outside the cell
                .MoveEnd Unit:=wdCharacter, count:=-1
                
                ' 2. Use InsertAfter instead of .Text = .Text
                ' This adds the colon to the end while leaving existing Images/Formatting alone
                '.InsertAfter ":"
            End With
        Next i

    ' convert to list
    Set doc = ActiveDocument
    If doc.Tables.count = 0 Then Exit Sub
    
    Set tbl = doc.Tables(1)
    nRows = tbl.rows.count
    nCols = tbl.Columns.count
    If nRows < 2 Then Exit Sub ' Need at least a header row + one data row
    
    delim = " " ' Delimiter after the prefixed header text (adjust if needed)
    
    On Error GoTo CleanFail
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    If Not Application.UndoRecord Is Nothing Then
        Application.UndoRecord.StartCustomRecord "Style, Prefix, Remove Header (First Table)"
        undoOn = True
    End If
    
    ' 1) Apply paragraph style "List" to Column 1, Rows 2..n
    For r = 2 To nRows
        With tbl.cell(r, 1).Range
            .End = .End - 1 ' exclude end-of-cell marker
            .Style = wdStyleList
        End With
    Next r
    
    ' 2) Apply paragraph style "List 2" to Columns 2..n, Rows 2..n
    If nCols >= 2 Then
        For r = 2 To nRows
            For c = 2 To nCols
                With tbl.cell(r, c).Range
                    .End = .End - 1
                    .Style = "List 2"
                End With
            Next c
        Next r
    End If
    
    ' 3) Set font color to Automatic for header row (Row 1), Columns 2..n
    If nCols >= 2 Then
        For c = 2 To nCols
            With tbl.cell(1, c).Range
                .End = .End - 1
                .Font.Color = wdColorAutomatic
            End With
        Next c
    End If
   
    ' 4) Prefix each data cell with its column header (preserving header formatting)
    For c = 1 To nCols
        Set hdr = tbl.cell(1, c).Range
        hdr.End = hdr.End - 1 ' exclude end-of-cell marker
        
        If Len(hdr.Text) > 0 Then
            For r = 2 To nRows
                ' Handle possible merged cells safely
                On Error Resume Next
                Set tgtCell = tbl.cell(r, c).Range
                If Err.Number <> 0 Then
                    Err.Clear
                    On Error GoTo CleanFail
                    GoTo NextCell
                End If
                On Error GoTo CleanFail
                
                ' Insert formatted header at the very start of the target cell
                Set ins = tgtCell.Duplicate
                ins.End = ins.start            ' collapse to start of cell
                ins.FormattedText = hdr.FormattedText
                
                ' Optional delimiter after the prefixed header
                ins.Collapse wdCollapseEnd
                ins.Text = delim
NextCell:
            Next r
        End If
    Next c
    
    ' 5) Remove Row 1 (header row)
    tbl.rows(1).Delete

    'set bold on headers
    ActiveDocument.Tables(1).Columns(1).Select
    Selection.Font.Bold = True
    
     Application.Run MacroName:="Lp_Table_Style_InCell_Para_And_Image"
     Application.Run MacroName:="Lp_Table_Mark_Keep_With_Next"

   ' convert table to text
    ActiveDocument.Tables(1).Select
    Selection.rows.ConvertToText Separator:=wdSeparateByParagraphs, NestedTables:=False

    ' delete last to para marks
    Selection.Collapse wdCollapseStart
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1
    
    ' create normal style para mark at top
    Dim rng As Range
    Set doc = ActiveDocument
    Set rng = doc.Paragraphs(1).Range
    rng.InsertParagraphBefore
    With doc.Paragraphs(1).Range
        .Style = wdStyleNormal
    End With
    
    ' Store full path + file name of the active (temp) document
    Dim TempDocName As String
    TempFileName = ActiveDocument.fullName
    
    'transcriber note
     Selection.HomeKey Unit:=wdStory

    'place transcriber note at top in temp file
     Selection.HomeKey Unit:=wdStory 'top of temp doc - move to the single para mark at top
     Selection.Delete ' remove para mark at top

     Application.Run MacroName:="Lp_Table_Insert_Transcriber_Note"

    'make colons and spaces color automatic
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorAutomatic
    With Selection.Find
        .Text = ": "
        .Replacement.Text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
        Selection.Find.Execute Replace:=wdReplaceAll
     DoEvents
     Selection.WholeStory
     DoEvents
     Selection.Copy
     DoEvents

     'open original doc
     DoEvents
     Documents(Lp_GP_String_2).Activate
     DoEvents
     ActiveDocument.Tables(Lp_GP_Counter_1).Select
     DoEvents
     ActiveDocument.Tables(Lp_GP_Counter_1).Delete
     DoEvents
     'The table is gone, so the cursor now sits exactly where the converted block will land.
     'The transcriber note was put at the TOP of the temp document, so it is the first
     'paragraph of what we are about to paste - leave the user on it (Jerry, 7/26/2026).
     Sh_Set_Return_Position Selection.Range.start
     Selection.Paste
     DoEvents

     'Screen stays OFF through the temp-file cleanup below - see the note in the other two
     'converts. Painting the temp document before closing it flashed the table's row color.
     'delete temp file
     Documents(TempFileName).Activate
     ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
     DoEvents

    'delete the temp file
    For Each d In Application.Documents
        If StrComp(d.fullName, TempDocName, vbTextCompare) = 0 Then
            d.Close SaveChanges:=wdDoNotSaveChanges
            Exit For
        End If
    Next d
    DoEvents
    On Error Resume Next
    Kill TempDocName
    On Error GoTo 0
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Application.ScreenUpdating = su_Prev
    Sh_Return_User_To_Start_Position
    Exit Sub

DoEvents
    Application.ScreenRefresh
    Exit Sub

CleanExit:
    If undoOn Then Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = su_Prev
    Sh_Return_User_To_Start_Position
    Exit Sub

CleanFail:
    ' Minimal cleanup; allow user to undo partially if needed
    Resume CleanExit
    
End Sub   '*** end of Lp_Table_Convert_RC_Table_To_List ***

Sub Lp_Table_Style_InCell_Para_And_Image()
'
' Version: 1.1  Date: 9/2/2025
' Version: 1.0  Date: 8/15/2025
'
    
    Dim tbl As Table, tblc As cell
    Dim ils As inlineShape
    Dim imgRange As Range
    Dim pImage As Paragraph, pAbove As Paragraph, p As Paragraph
    Dim startStyling As Boolean
    
    ' === Select the only table in the document ===
    If ActiveDocument.Tables.count = 1 Then
        ActiveDocument.Tables(1).Range.Select
    Else
        MsgBox "This macro expects exactly one table in the document.", vbExclamation
        Exit Sub
    End If
    
    ' === Begin undo record ===
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "Style Above Image as List 2, Image and Following as List 4"
    On Error GoTo 0
    
    Set tbl = Selection.Tables(1)
    
    For Each tblc In tbl.Range.Cells
        For Each ils In tblc.Range.InlineShapes
            Set imgRange = ils.Range
    
            ' === Ensure paragraph mark before image ===
            If imgRange.start > tblc.Range.start Then
                If Mid(tblc.Range.Text, imgRange.start - tblc.Range.start, 1) <> vbCr Then
                    imgRange.InsertBefore vbCr
                End If
            Else
                tblc.Range.InsertBefore vbCr
            End If
    
            ' === Re-identify image paragraph and above paragraph ===
            ' Work from the cell's paragraph collection to avoid stale references
            For Each p In tblc.Range.Paragraphs
                If p.Range.InlineShapes.count > 0 Then
                    If p.Range.InlineShapes(1) Is ils Then
                        Set pImage = p
                        Exit For
                    End If
                End If
            Next p
            If Not pImage Is Nothing Then
                If pImage.Previous Is Nothing Then
                    Set pAbove = Nothing
                Else
                    Set pAbove = pImage.Previous
                End If
            End If
    
            ' === Style paragraph above image as "List 2" ===
            If Not pAbove Is Nothing Then
                With pAbove.Range
                    .ListFormat.RemoveNumbers
                    On Error Resume Next
                    .Style = "List 2"
                    On Error GoTo 0
                End With
            End If
    
            ' === Style image paragraph and everything after as "List 4" ===
            startStyling = False
            For Each p In tblc.Range.Paragraphs
                If Not pImage Is Nothing Then
                    If p.Range.start = pImage.Range.start Then startStyling = True
                End If
                If startStyling Then
                    With p.Range
                        .ListFormat.RemoveNumbers
                        On Error Resume Next
                        .Style = "List 4"
                        On Error GoTo 0
                    End With
                End If
            Next p

            ' Clear for next image
            Set pImage = Nothing
            Set pAbove = Nothing
        Next ils
    Next tblc
    
    ' === End undo record ===
    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    On Error GoTo 0

End Sub   '*** end of Lp_Table_Style_InCell_Para_And_Image ***

Sub Lp_Table_Mark_Keep_With_Next()
'
' Version: 1.1  Date: 9/2/2025 - fixed problem with all paras set to keep - last item h
' Version: 1.0  Date: 8/15/2025
'
    'mark each except cell in the last column with KeepWithNext
    'S=Keep List Group on same page
    
    Dim tbl As Table
    Dim lastRow As Integer
    Dim lastCol As Integer
    Dim r As Integer
    Dim c As Integer
    Dim para As Paragraph

    If Selection.Information(wdWithInTable) And InStr(Lp_GP_String_3, "S") > 0 Then
        Set tbl = Selection.Tables(1)
        lastRow = tbl.rows.count
        lastCol = tbl.Columns.count
        
        For r = 1 To lastRow
            For c = 1 To lastCol
                ' Skip any cell in the last column
                If c <> lastCol Then
                    With tbl.cell(r, c).Range
                        ' Avoid selecting the end-of-cell marker
                        .End = .End - 1
                        For Each para In .Paragraphs
                            para.Format.KeepWithNext = True
                        Next para
                    End With
                End If
            Next c
        Next r
    End If
    
    'put blank line between groups
    If InStr(Lp_GP_String_3, "E") > 0 Then  'E=Create empty para after list group
    
        lastCol = tbl.Columns.count
        
        For r = 1 To tbl.rows.count
            With tbl.cell(r, lastCol).Range
                .End = .End - 1 ' exclude marker
        
                ' Trim trailing blanks/NBSP paragraphs but leave one if styled
                Do While .Paragraphs.count > 1 _
                  And Len(Trim$(Replace(Replace(.Paragraphs(.Paragraphs.count).Range.Text, vbCr, ""), Chr(160), " "))) = 0
                    .Paragraphs(.Paragraphs.count).Range.Delete
                Loop
        
                ' After cleanup, if the last para has content, append a new one
                If Len(Trim$(Replace(Replace(.Paragraphs(.Paragraphs.count).Range.Text, vbCr, ""), Chr(160), " "))) <> 0 Then
                    .Collapse wdCollapseEnd
                    .InsertAfter vbCr
                End If
        
                ' Style only the trailing empty para (if >1 para, it's truly trailing)
                If .Paragraphs.count > 1 Then
                    .Paragraphs(.Paragraphs.count).Range.Style = "Normal"
                End If
            End With
        Next r
        
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "^p^p"
            .Replacement.Text = "^p"
            .Forward = True
            .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
        End With
    End If
End Sub   '*** end of Lp_Table_Mark_Keep_With_Next '***
  
Function Sh_IsValidRomanNumeral(s As String) As Boolean

    Dim validChars As String: validChars = "IVXLCDM"
    Dim i As Long

    s = Trim(UCase(s))
    If Len(s) = 0 Then Exit Function

    ' Reject if any character isn’t a Roman letter
    For i = 1 To Len(s)
        If InStr(validChars, Mid(s, i, 1)) = 0 Then Exit Function
    Next i

    ' Accept if it matches known valid forms
    Select Case s
        Case "I", "II", "III", "IV", "V", "VI", "VII", "VIII", _
             "IX", "X", "XVI", "XVII", "XVIII", "XIX", "XX", _
             "XXI", "XXII", "XXIII", "XXIV", "XXV", "XXX", _
             "XL", "L", "LX", "LXX", "LXXX", _
             "XC", "C", "CC", "CCC", "CD", "D", _
             "DC", "DCC", "DCCC", "CM", "M"
             Sh_IsValidRomanNumeral = True
        Case Else
            Sh_IsValidRomanNumeral = False
    End Select
End Function   '*** end of Function Sh_IsValidRomanNumeral ***

Sub Lp_Normalize_Styles()
    '
    ' Version: 3.6  Date: 8/8/2026 - sets the typeface on the styles as well as their sizes, and
    '                               scales the List Paragraph hanging indents for a face wider than
    '                               Tahoma. Takes the typeface from the DOCUMENT every time, never
    '                               from the public - this macro also runs on an imported selection
    '                               file, where a stale public gave it the wrong indents
    ' Version: 3.5  Date: 7/30/2026 - two progress messages both said they were setting table
    '                                 border weights; each step now has its own. Sh_Color_Dollar_PG_Red
    '                                 moved from near the top to the LAST action, and shows no
    '                                 message - the style updates in between reset the font
    '                                 color of the styles they touch, so the red could be
    '                                 undone before the macro finished
    ' Version: 3.3  Date: 7/26/2026 - stopped the screen flashing - see the note in the changelog header
    ' Version: 3.2  Date: 7/18/2026 - space-after now set once at story level (was a per-paragraph loop)
    ' Version: 3.1  Date: 7/6/2026 - optimized style updates, preserved all status messages and DoEvents
    ' Modifies the font sizes and character spacing of the document based on Lp_Base_Font_Size
    '

    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False   ' both callers already do this, but the sub must stand on its own

    ' Always take the typeface from the document in front of us, never from the Public.
    '
    ' Testing only for an EMPTY Public would not be enough, and the difference is a real bug:
    ' Lp_Import_Exported_Selection_File opens the exported file, activates it and calls this
    ' macro on THAT document. Attach a book as Legible, then import a selection into a Tahoma
    ' book later in the same Word session, and the Public still says Legible - so the imported
    ' paragraphs would get list indents 5.4% deeper than the rest of the book. Those are direct
    ' paragraph formatting, so they travel into the transcriber's book and stay there.
    '
    ' In the attach path this reads back exactly what was just written, so nothing is lost.
    Lp_Base_Font_Name = ActiveDocument.Styles(wdStyleNormal).Font.Name
    If Len(Trim(Lp_Base_Font_Name)) = 0 Then Lp_Base_Font_Name = LP_FONT_TAHOMA

    Sh_NonModalMessageForm.SetActivityMessage "Underlining Italics"
    Application.Run MacroName:="Lp_Italics_To_Dashed_Underline"

    Dim LoopCounter As Integer
    Dim BxClr As String

    ' This message said "Adjusting table border weights" and covered TWO calls, so it
    ' described the second one not at all - and it read as a duplicate of the genuine
    ' "Setting table border weights" further down, leaving no way to tell which was really
    ' doing the job. (Jerry, 7/30/2026.)
    '
    ' Lp_Set_Table_Border_Weights does set the weights, but the loop further down sets them
    ' again and is what the user finally sees - so this message names the part of its work
    ' that only happens here, rather than announcing border weights twice.
    '
    ' Sh_Color_Dollar_PG_Red used to run here too. It is now the last thing the macro does
    ' to the document, and shows no message of its own - see the end of this sub.
    Sh_NonModalMessageForm.SetActivityMessage "Removing table shadows and diagonal borders"
    Application.Run MacroName:="Lp_Set_Table_Border_Weights"

    '***** Begin Setting Automatic List Indention Size ****
    Sh_NonModalMessageForm.SetActivityMessage "Setting list paragraph indention sizes"

    ' The numbers below are Tahoma's, tuned by hand, and stay the authority for it. A wider face
    ' needs a proportionally wider hang, so it scales off them rather than getting a second table
    ' of guesses - see Lp_Indent_Factor_For_Font. Tahoma's factor is 1, so its documents come out
    ' exactly as they always have.
    Dim IndentFactor As Double
    Dim BaseIndent As Double
    IndentFactor = Lp_Indent_Factor_For_Font(Lp_Base_Font_Name)

    Dim oPara As Paragraph
    For Each oPara In ActiveDocument.Paragraphs
        If oPara.Style = "List Paragraph" Then
            BaseIndent = 0
            Select Case Lp_Base_Font_Size
                Case "14": BaseIndent = -0.26
                Case "16": BaseIndent = -0.3
                Case "18": BaseIndent = -0.34
                Case "20": BaseIndent = -0.38
                Case "22": BaseIndent = -0.41
                Case "24": BaseIndent = -0.44
                Case "26": BaseIndent = -0.47
                Case "28": BaseIndent = -0.5
                Case "30": BaseIndent = -0.53
                Case "32": BaseIndent = -0.57
                Case "34": BaseIndent = -0.61
                Case "36": BaseIndent = -0.64
                Case "38": BaseIndent = -0.68
                Case "40": BaseIndent = -0.72
                Case "42": BaseIndent = -0.77
            End Select
            ' Zero means the size is not one of ours. Leave the paragraph's own indent alone,
            ' which is what the old Select Case did too - it had no Case Else.
            If BaseIndent <> 0 Then
                oPara.FirstLineIndent = InchesToPoints(BaseIndent * IndentFactor)
            End If
        End If
    Next oPara
    '***** End Setting Automatic List Indention Size ****

    '***** Begin Setting space after para Size ****
    Sh_NonModalMessageForm.SetActivityMessage "Setting spacing between paragraphs"

    ' Same value on every paragraph => one story-level assignment instead of a per-paragraph
    ' loop. On Error Resume Next preserves the original no-op behavior when Lp_Base_Font_Size
    ' is "" (type mismatch, swallowed); a numeric string like "18" coerces to 18 pt as before.
    On Error Resume Next
    ActiveDocument.Content.ParagraphFormat.SpaceAfter = Lp_Base_Font_Size
    On Error GoTo 0
    '***** End Setting space after para Size ****

    '**** Begin set border weights for box styles, table styles and Print Pg Num Style ****
    Sh_NonModalMessageForm.SetActivityMessage "Setting box border weights"

    Dim tbl As Table
    Dim targetWeight As WdLineWidth
    Dim sName As Variant
    Dim arrBoxStyles As Variant

    ' Same decision Lp_Set_Table_Border_Weights uses, so the two can never disagree about
    ' how heavy a border should be. They did until 7/30/2026.
    targetWeight = Lp_Border_Weight_For_Base_Font

    Sh_NonModalMessageForm.SetActivityMessage "Setting box border spacing"

    arrBoxStyles = Array("Box Black", "Box Blue", "Box Orange", _
                         "Box Red", "Box Violet", "Box White")

    On Error Resume Next
    For Each sName In arrBoxStyles
        If ActiveDocument.Styles(sName).InUse Then
            With ActiveDocument.Styles(sName).ParagraphFormat.Borders
                .Item(wdBorderTop).LineStyle = wdLineStyleSingle
                .Item(wdBorderBottom).LineStyle = wdLineStyleSingle
                .Item(wdBorderLeft).LineStyle = wdLineStyleSingle
                .Item(wdBorderRight).LineStyle = wdLineStyleSingle

                .Item(wdBorderTop).LineWidth = targetWeight
                .Item(wdBorderBottom).LineWidth = targetWeight
                .Item(wdBorderLeft).LineWidth = targetWeight
                .Item(wdBorderRight).LineWidth = targetWeight

                .DistanceFromTop = 2
                .DistanceFromLeft = 4
                .DistanceFromBottom = 4
                .DistanceFromRight = 4
            End With
        End If
    Next sName

    Sh_NonModalMessageForm.SetActivityMessage "Setting reference page border weight"

    If Sh_Style_In_Use(ActiveDocument, "Print Pg Num") Then
        With ActiveDocument.Styles("Print Pg Num").ParagraphFormat
            .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
            .Borders(wdBorderRight).LineStyle = wdLineStyleNone

            .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
            .Borders(wdBorderTop).LineWidth = targetWeight

            .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
            .Borders(wdBorderBottom).LineWidth = targetWeight
        End With
    End If
    On Error GoTo 0

    Sh_NonModalMessageForm.SetActivityMessage "Setting table border weights"

    For Each tbl In ActiveDocument.Tables
        With tbl.Borders
            .InsideLineStyle = wdLineStyleSingle
            .OutsideLineStyle = wdLineStyleSingle
            .InsideColorIndex = wdAuto
            .OutsideColorIndex = wdAuto

            .InsideLineWidth = targetWeight
            .OutsideLineWidth = targetWeight
        End With
    Next tbl
    '**** End border weights ****

    '********* Begin Expand Font Spacing Settings ******
    Sh_NonModalMessageForm.SetActivityMessage "Setting sizes of inter-character spacing for Normal and List Paragraph styles"

    Dim styleName As String
    Dim i As Long
    Dim StyleList As Variant
    Dim SpacingTable As Variant
    Dim targetSpacing As Single
    Dim sty As Style

    StyleList = Array("Normal", "List Paragraph")

    SpacingTable = Array( _
        14, 0.8, _
        16, 1.05, _
        18, 1.3, _
        20, 1.55, _
        22, 1.8, _
        24, 2.05, _
        26, 2.3, _
        28, 2.55, _
        30, 2.8, _
        32, 3.05, _
        34, 3.3, _
        36, 3.55, _
        38, 3.65, _
        40, 4.05, _
        42, 4.3)

    For i = LBound(SpacingTable) To UBound(SpacingTable) Step 2
        If SpacingTable(i) = CLng(Lp_Base_Font_Size) Then
            targetSpacing = SpacingTable(i + 1)
            Exit For
        End If
    Next i

    For i = LBound(StyleList) To UBound(StyleList)
        styleName = StyleList(i)
        Set sty = ActiveDocument.Styles(styleName)

        sty.AutomaticallyUpdate = True
        sty.Font.spacing = targetSpacing
        sty.AutomaticallyUpdate = False
    Next i
    '********* End Expand Font Spacing Settings ******

    '********** Unified font size updates for styles ***********
    Sh_NonModalMessageForm.SetActivityMessage "Setting font sizes for styles"

    Dim base As Long
    Dim UnifiedStyles As Variant

    base = CLng(Lp_Base_Font_Size)

    UnifiedStyles = Array( _
        "Text Blue", "Text Green", "Text Orange", "Text Red", "Text Violet", _
        "Box Blue", "Box Orange", "Box Red", "Box Violet", "Box White", _
        "Words Aqua", "Words Black", "Words Blue", "Words Green", _
        "Words Pink", "Words Tan", "Words Yellow", _
        "Para Aqua", "Para Black", "Para Black Inverted", _
        "Para Blue", "Para Green", "Para Tan", "Para Yellow", _
        "TOC 1", "TOC 2", "TOC 3", "TOC 4", "TOC 5", _
        "Print Pg Num", _
        "Normal", "List Paragraph")

    For i = LBound(UnifiedStyles) To UBound(UnifiedStyles)
        Set sty = ActiveDocument.Styles(UnifiedStyles(i))
        sty.Font.Size = base
    Next i
    '********** End unified font size updates ***********

    '*********** begin typeface for styles **************
    ' Deliberately a separate pass rather than a .Font.Name added to the loop above: that loop's
    ' Styles(...) lookup is unguarded, and widening what depends on it would widen the blast
    ' radius if a style is ever missing. Lp_Apply_Base_Font_To_Styles carries its own guard.
    Sh_NonModalMessageForm.SetActivityMessage "Setting the typeface for styles"
    Lp_Apply_Base_Font_To_Styles ActiveDocument, Lp_Base_Font_Name
    '*********** end typeface for styles **************

    '*********** begin heading styles (size + bold + spacing) **************
    Sh_NonModalMessageForm.SetActivityMessage "Setting font sizes and spacing for heading styles"
    
    Dim StyleNames As Variant
    Dim SizeOffsets As Variant
    Dim headingSpacing As Single

    StyleNames = Array("Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5")
    SizeOffsets = Array(10, 8, 6, 4, 2)

    Select Case base
        Case 14: headingSpacing = 1.8
        Case 16: headingSpacing = 1.9
        Case 18: headingSpacing = 2
        Case 20: headingSpacing = 2.1
        Case 22: headingSpacing = 2.2
        Case 24: headingSpacing = 2.3
        Case 26: headingSpacing = 2.4
        Case 28: headingSpacing = 2.5
        Case 30: headingSpacing = 2.6
        Case 32: headingSpacing = 2.7
        Case 34: headingSpacing = 2.8
        Case 36: headingSpacing = 2.9
        Case 38: headingSpacing = 3
        Case 40: headingSpacing = 3.1
        Case Else: headingSpacing = 3.2
    End Select

    For i = LBound(StyleNames) To UBound(StyleNames)
        Set sty = ActiveDocument.Styles(StyleNames(i))
        sty.Font.Size = base + SizeOffsets(i)
        sty.Font.Bold = True
        sty.Font.Position = 0
        sty.Font.spacing = headingSpacing
    Next i
    '*********** end heading styles **************

    '********** Begin Set Base Font Size for Para Styles (Normal + colored paras) ***********
    Sh_NonModalMessageForm.SetActivityMessage "Ensuring base font size for Normal and color paragraph styles"

    StyleNames = Array( _
        "Normal", _
        "Para Aqua", "Para Black", "Para Black Inverted", _
        "Para Blue", "Para Green", "Para Tan", "Para Yellow")

    For i = LBound(StyleNames) To UBound(StyleNames)
        Set sty = ActiveDocument.Styles(StyleNames(i))
        sty.Font.Size = base
    Next i
    '********** End Set Base Font Size for Para Styles ***********

    ' LAST thing done to the document, and deliberately silent (Jerry, 7/30/2026).
    ' It used to run near the top, before every style update above - and those updates reset
    ' the font color of the styles they touch, so the red could be undone again before the
    ' macro had finished. Doing it last means it survives. No progress message: it is quick,
    ' and it is not a step the user needs narrating.
    Application.Run MacroName:="Sh_Color_Dollar_PG_Red"

    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    ActiveDocument.UndoClear

     Application.ScreenUpdating = su_Prev

End Sub

Sub Lp_TOC_CleanAndFormat_TOC()
    '
    ' Version 1.1  Date: 7/18/2026 - perf: nbsp removal now a single Find pass, and the bold-map
    '                                read/write enumerates characters once each way instead of
    '                                indexed Characters(i) (O(n) vs O(n^2) per paragraph). Same output.
    ' Version 1.0  Date: 9/1/2025
    '

    Dim OriginalDocName As String
    Dim TempDocName As String

    If Not Selection.Type = wdSelectionNormal Then
        MsgBox "Select the entire TOC including any text that does reference page numbers", , "VistaType LP  (206)"
        End
    End If

    OriginalDocName = Lp_GP_String_1 '"Lp_GP_String_1" was filled by "Lp_TOC_Format_And_Color_Form"
    
    Selection.MoveStart Unit:=wdCharacter, count:=0 'move to top of selection
    Sh_Save_User_Position
    
    'copy to temp doc
    Application.Run MacroName:="Lp_Copy_To_Temp_Doc"
    DoEvents
    'Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph 'a space charater as the first char in the temp doc will crash the formatting
                            'this forces the removal of spaces  after a para mark to work later in the cleanup
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1 'get rid of the ending para mark created with the new document
    ActiveDocument.Content.Select

    TempDocName = ActiveDocument.Name
    
    Dim para As Paragraph
    Dim paraRange As Range
    Dim charRange As Range
    Dim i As Long
    Dim endText As String
    Dim numStart As Long
    Dim ch As String
    Dim regexNum As Object
    Dim regexRoman As Object
    Dim hasNumberOrRoman As Boolean
    Dim boldMap() As Boolean
    Dim charCount As Long
    Dim Sel As Range
    Dim paraText As String
    Dim re As Object
    Dim match As Object
    Dim matchStart As Long
    Dim matchLength As Long
    Dim numberRng As Range

    'begin cleanup before formatting
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        
        ' ? U+2666
        .Text = ChrW(&H2666)
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
        
        ' * U+002A
        .Text = ChrW(&H2A)
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
 
        ' Tab
        .Text = "^t"
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        
        ' Ellipsis … U+2026
        .Text = ChrW(&H2026)
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
        
        ' Non-breaking space U+00A0
        '.text = ChrW(&HA0)
        '.Replacement.text = ""
        '.Execute Replace:=wdReplaceAll
        
        ' Soft hyphen U+00AD
        .Text = ChrW(&HAD)
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
        
        ' Zero-width space U+200B
        .Text = ChrW(&H200B)
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
        
        ' Zero-width non-joiner U+200C
        .Text = ChrW(&H200C)
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
        
        ' Zero-width joiner U+200D
        .Text = ChrW(&H200D)
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
        
        ' En dash U+2013
        .Text = ChrW(&H2013)
        .Replacement.Text = "-"
        .Execute Replace:=wdReplaceAll
        
        ' Em dash U+2014
        .Text = ChrW(&H2014)
        .Replacement.Text = "-"
        .Execute Replace:=wdReplaceAll
        
        ' Middle dot · U+00B7
        .Text = ChrW(&HB7)
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
        
        ' Ligature ? U+FB01
        .Text = ChrW(&HFB01)
        .Replacement.Text = "fi"
        .Execute Replace:=wdReplaceAll
        
        ' Ligature ? U+FB02
        .Text = ChrW(&HFB02)
        .Replacement.Text = "fl"
        .Execute Replace:=wdReplaceAll
        
        ' Multiple spaces ? single space
        .MatchWildcards = True
        .Text = "[ ]{2,}"
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        
        ' Any bullet character (common set: • ? ? ? ? ? U+F0B7 etc.)
        .Text = "[" & ChrW(&H2022) & ChrW(&H2023) & ChrW(&H25AA) & ChrW(&H25E6) & ChrW(&H25CF) & ChrW(&H25CB) & ChrW(&HF0B7) & "]"
        .Replacement.Text = ""
        .Execute Replace:=wdReplaceAll
        
        ' Remove space following a paragraph mark
        .MatchWildcards = False
        .Text = "^p "
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
    End With
    
    '*******************************************************
    ' Remove spaces before paragraph marks
    '*******************************************************
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}^013"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue  'replaces all in the document
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '*******************************************************
    ' Remove Spaces following paragraph marks
    '*******************************************************
    With Selection.Find
        .Text = "^013^032{1,}"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue  'replaces all in the document
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ActiveDocument.Content.Select
    Set Sel = ActiveDocument.Content

    'start remove nonbreaking space for all except style "Print Pg Num"
    For Each para In Sel.Paragraphs
    Set paraRange = para.Range
    paraText = paraRange.Text

    If paraRange.Style = "Print Pg Num" Then
        ' Skip this paragraph
        GoTo NextPara
    End If

    ' Remove all non-breaking spaces from the paragraph (single Find pass, not a per-character scan)
    With paraRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = Chr(160)
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

NextPara:
Next para

    'end remove nonbreaking space for all except style "Print Pg Num"
    
    Set regexNum = CreateObject("VBScript.RegExp")
    With regexNum
        .Global = False
        .IgnoreCase = False
        .pattern = "^\d+$"
    End With
    
    Set regexRoman = CreateObject("VBScript.RegExp")
    With regexRoman
        .Global = False
        .IgnoreCase = True
        .pattern = "^[mdclxvi]{2,}$""Section 3"
    End With

    '*** begin formatting ***
    ' === Unified regex pattern ===
    Const unifiedPattern As String = _
        " (\b(M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})|[A-Za-z]?\d+))[\r\n]{1,2}$"

    Set regexNum = CreateObject("VBScript.RegExp")
    With regexNum
        .pattern = unifiedPattern
        .IgnoreCase = True
        .Global = False
    End With

    ' === First loop: detect and style ===
    For Each para In Selection.Paragraphs
        Set paraRange = para.Range
        
        ' Skip paragraphs that begin with "$pg" or have style "Print Pg Num"
        ' Sh_Para_Style_Is, not "= ActiveDocument.Styles(...)": VBA evaluates BOTH sides of Or,
        ' so a document without the style raised 5941 here even when the text test already passed.
        If Left(paraRange.Text, 3) = "$pg" Or Sh_Para_Style_Is(paraRange, "Print Pg Num") Then
            GoTo SkipPara
        End If
        
        ' Test for number/roman/letter+number
        hasNumberOrRoman = regexNum.test(paraRange.Text)
        
        ' If no match, reset style to Normal but preserve bold map
        If Not hasNumberOrRoman Then
            charCount = paraRange.Characters.count - 1
            If charCount > 0 Then
                ' Enumerate characters once each way (O(n)); indexed Characters(i) is O(n^2).
                ReDim boldMap(1 To charCount)
                i = 0
                For Each charRange In paraRange.Characters
                    i = i + 1
                    If i > charCount Then Exit For
                    boldMap(i) = charRange.Font.Bold
                Next charRange
                para.Style = ActiveDocument.Styles("Normal")
                para.SpaceAfter = 0
                i = 0
                For Each charRange In paraRange.Characters
                    i = i + 1
                    If i > charCount Then Exit For
                    charRange.Font.Bold = boldMap(i)
                Next charRange
            Else
                para.Style = ActiveDocument.Styles("Normal")
                para.SpaceAfter = 0
            End If
        End If
        
SkipPara:
    Next para

    ' === Second loop: replace space with tab and apply TOC style ===
    Set Sel = Selection.Range
    For Each para In Sel.Paragraphs
        Set paraRange = para.Range
        paraText = paraRange.Text
        
        If regexNum.test(paraText) Then
            Set match = regexNum.Execute(paraText)(0)
            matchStart = match.FirstIndex + 1 ' space's position (zero-based)
            matchLength = Len(match.SubMatches(0)) ' matched number/roman/letter+number
            
            ' Replace the space before the pattern with a tab
            paraRange.Characters(matchStart).Text = vbTab
            
            ' Select and un-bold the matched number/roman/letter+number
            Set numberRng = paraRange.Duplicate
            numberRng.start = paraRange.start + matchStart
            numberRng.End = numberRng.start + matchLength
            numberRng.Font.Bold = False
            
            ' Apply style
            On Error Resume Next
            paraRange.Style = "TOC 1"
            On Error GoTo 0
        End If
        
    Next para
    '*** End formatting ***

    ' begin Fix reference pages
    Set Sel = ActiveDocument.Content
    
    For Each para In Sel.Paragraphs
        Set paraRange = para.Range
        paraText = paraRange.Text
    
        If paraRange.Style = "Print Pg Num" Then
            ' Insert NonBreakingSpace at start
            paraRange.InsertBefore Chr(160)
            ' Insert NBSP before paragraph mark (end of visible text)
            'paraRange.End = paraRange.End - 1
            'paraRange.InsertAfter Chr(160)
        End If
    Next para
    
    Set Sel = ActiveDocument.Content
    
    For Each para In Sel.Paragraphs
        Set paraRange = para.Range
        paraText = paraRange.Text
    
        If paraRange.Style = "Print Pg Num" Then
            numStart = InStr(paraText, "pn")
            If numStart > 0 Then
                paraRange.start = paraRange.start + numStart - 1
                paraRange.End = paraRange.start + 2
                paraRange.Text = vbTab & "pn"
            End If
        End If
    Next para
    ' end Fix reference pages
    
    Selection.HomeKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, count:=1
    ActiveDocument.Content.Select

    ' begin copy from temp doc and paste into selected text area in the original doc
    Dim srcDoc As Document
    Dim destDoc As Document
    
    ' Reference the open source and destination docs
    Set srcDoc = Documents(TempDocName)
    Set destDoc = Documents(OriginalDocName)
    
    ' Replace the current selection in the destination
    destDoc.Activate
    Selection.FormattedText = srcDoc.Range.FormattedText
    ' end copy from temp doc and paste into selected text area in the original doc

    ' begin kill the temp doc
       'Dim srcDoc As Document
    Set srcDoc = Documents(TempDocName)
    ' Activate it
    srcDoc.Activate
    ' Close without saving (discard changes)
    srcDoc.Close SaveChanges:=wdDoNotSaveChanges
    ' end kill the temp doc
    
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Sh_Return_User_To_Start_Position

    Application.ScreenUpdating = True
    Application.ScreenRefresh
    DoEvents
    ActiveWindow.View.Type = wdNormalView
    ActiveWindow.View.Type = wdPrintView
    Selection.Collapse Direction:=wdCollapseStart
    DoEvents

End Sub   '*** end of Lp_TOC_CleanAndFormat_TOC ***

Sub Lp_Replace_Underline_Tab_With_Underlined_Underscore()
'
' Version: 1.0  Date:3/10/2026
'
' common in scans from AbbyyFineReader

    Selection.Find.ClearFormatting
    Selection.Find.Font.Underline = wdUnderlineSingle
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Underline = wdUnderlineSingle
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "_"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub  '*** end of Lp_Replace_Underline_Tab_With_Underline macro ***


Sub Lp_Delete_Square_Bullet()
'
' Version 1.0:  Date: 11/25/2025
'

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(61623)
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub   '*** end of Lp_Delete_Square_Bullet macro ***


'------------------------------------------------------------------------------------
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
' ** start of  MS macros  ***
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
'-----------------------------------------------------------------------------------

Sub MS_Set_Word_Config_For_New_Install()
    '
    ' MS_Set_Word_Config_For_New_Install Macro
    '
    ' Author: Jerry Whittaker -  jerry@thewhittakers.org
    '
    ' Version: 2.7  Date: 8/22/2026 - ticks "Automatically use suggestions from the spelling checker" here
    '                                 too, with the spell-check-as-you-type gate it gets nowhere without
    '                                 (Jerry, 3.0.222). Both settings leave the ledger with it: forced in
    '                                 all three configurations, they are not hers to give back any more
    ' Version: 2.6  Date: 8/20/2026 - the SCREEN half of Piece 3 of the automatic-configuration plan. Formatting
    '                                 marks off rather than on, the navigation pane closed, and the three Styles
    '                                 pane settings saved inside her file no longer written at all. The twenty
    '                                 Options and AutoCorrect writes below are NOT touched yet - they go when the
    '                                 settings ledger exists to replace them, which is Piece 4
    ' Version: 2.5  Date: 8/18/2026 - the 15 settings all three configurations wrote IDENTICALLY are gone
    '                                 from all three: nothing was ever switching them, so the only thing the
    '                                 writes did was overwrite the transcriber's own choice - which is what
    '                                 the beta tester on 3.0.135 reported. They are hers now, in every kind
    '                                 of document. This sub also PUTS BACK the five spelling and grammar
    '                                 settings the two books switch off (Sh_Restore_Transcriber_Settings), which
    '                                 nothing had ever done. 35 writes -> 20.
    ' Version: 2.4  Date: 8/18/2026 - checked setting by setting against Word's factory defaults (Jerry's table,
    '                                 8/18/2026). Three differences, all here: AutoFormatAsYouTypeApplyFirstIndents
    '                                 was never written and is now True; AutoFormatReplacePlainTextEmphasis was
    '                                 True and a clean install has it False; CorrectKeyboardSetting is no longer
    '                                 written at all. The other 32 already agreed. Note the deliberate asymmetry
    '                                 that is NOT a bug: ApplyHeadings is False as you type and True on demand.
    ' Version: 2.3  Date: 8/18/2026 - stopped DELETING the 18 compact fraction AutoCorrect entries and ADDS
    '                                 them instead, through the shared Sh_Add_Compact_Fractions - the same
    '                                 eighteen the braille configuration uses (Jerry, 8/18/2026). Fifteen of
    '                                 them are not Word's; they now stay in the transcriber's own AutoCorrect
    '                                 list. Large print still deletes them.
    ' Version: 2.2  Date: 7/18/2026 - only write Options/AutoCorrect that actually differ (idempotent),
    '                                 so re-running on every new doc no longer churns Word's roaming
    '                                 settings and triggers Office's "restart to apply your privacy
    '                                 settings" notice; also stops silently re-clobbering user prefs.
    ' Version: 2.1  Date: 10/27/2021 - set for wdShowFilterStylesAll
    ' Version: 2.0  Date: 10/19/2021 - added removal of compact fractions
    ' Version: 1.9  Date: 3/1/2020 - added set view to print view
    ' Version: 1.8  Date: 1/8/2020 - closed reading view and added on error trap
    ' Version: 1.7  Date: 4/14/2017
    '
    
    ActiveDocument.ActiveWindow.View.ReadingLayout = False  'will crash if document is in reading view ... close reading view

    ' 8/20/2026 - three Styles pane settings are no longer written here, and must not be put
    ' back: FormattingShowNextLevel, StyleSortMethod, and FormattingShowFilter, which was being
    ' forced to wdShowFilterStylesAll. All three are DOCUMENT properties, saved inside her file
    ' and carried with it, so writing them on every ordinary document open overwrote whatever
    ' that document was carrying - the 7/24/2026 complaint, in her letters rather than her books.
    ' Decision 6 of the automatic-configuration plan. The pane's VISIBILITY is a separate
    ' question and is still decided here; see the block below.
    
    ' Only write settings that differ from their target, so re-running this on every new
    ' document doesn't hand Word's (roaming) settings store a no-op "change" each time.
    ' 8/20/2026 - the twenty Options and AutoCorrect writes that stood here are GONE, and this is
    ' the typing half of Piece 3 of the automatic-configuration plan. They set Word's factory
    ' value for twenty settings on every ordinary document open, which LOOKED like restoring her
    ' settings and was overwriting them: Jerry, testing 3.0.210, cleared every box in AutoCorrect,
    ' AutoFormat and AutoFormat As You Type, closed Word, came back, and found them checked again.
    ' Nothing had failed to save. This sub was rewriting them.
    '
    ' What puts them back now is Sh_Restore_Transcriber_Settings, one call below, working from the
    ' ledger - HER values, not Word's factory ones, and covering thirty-four settings rather than
    ' the twenty that were written here. Do not add a fixed write back into this sub: a value
    ' written here cannot be told from a value she chose. The ONE exception is the pair of
    ' spelling-checker writes further down, and it is only an exception because those two stopped
    ' being preferences: all three configurations force them from 3.0.222, so there is no value of
    ' hers to tell apart, and both names left Sh_Tracked_Settings on 8/22/2026 to say so.
    '
    ' AutoFormatAsYouTypeApplyFirstIndents went with them and is written by nothing now. No book
    ' configuration ever touched it, so the only thing the write did was overwrite her own choice -
    ' the same reasoning that removed fifteen settings from all three configurations on 8/18/2026.
    '
        ' NOT WRITTEN HERE, DELIBERATELY - do not add them.
        '
        ' The four "AutoAdd" properties - FirstLetterAutoAdd, TwoInitialCapsAutoAdd,
        ' OtherCorrectionsAutoAdd and HangulAndAlphabetAutoAdd - and the four exception
        ' lists they feed (FirstLetterExceptions and the rest) are the transcriber's
        ' accumulated work, not configuration. Everything above can be written back from
        ' the table this sub is built to; an exception list cannot be reconstructed once
        ' it is damaged. Nothing anywhere in this project writes them, and nothing should.
        '
        ' CorrectKeyboardSetting was written False here until 8/18/2026, and by large print
        ' and braille as well. None of the three writes it now. It is the keyboard-follows-
        ' the-language setting, which belongs to a multilingual transcriber's own setup and
        ' to no part of large print or braille. Leaving it in only the two books would have
        ' been worse than leaving it everywhere: one book opened and it was off for good.
    
    ' What a plain letter looks like when it opens, decided by Jerry and the beta tester on
    ' 8/20/2026: no formatting marks, no Styles pane, no navigation pane, both rulers, print
    ' view. Five things stated rather than four, and two of them changed that day.
    '
    ' ShowAll was True here until 8/20/2026 and is now False. Forcing pilcrows ON meant a
    ' transcriber who works without them had to switch them off again on every single letter she
    ' opened - not once. The rulers are the opposite case and stay forced: the vertical one is
    ' buried in File > Options > Advanced > Display and she should not have to go looking.
    '
    ' The navigation pane is new here, and it goes through Sh_Set_Navigation_Pane because it has
    ' two halves that both need writing - a window property and an application-wide command bar.
    ' See that sub; getting it wrong is what left the pane standing after a close in 3.0.207.
    If Not Sh_Config_Skip_Display Then   ' the transcriber's screen is theirs - see Sh_Config_Skip_Display
        ActiveWindow.View.ShowAll = False
        ActiveWindow.DisplayRulers = True
        ActiveWindow.DisplayVerticalRuler = True
        ActiveWindow.ActivePane.View.Type = wdPrintView
        Application.TaskPanes(wdTaskPaneFormatting).Visible = False 'turn off styles pane
        Sh_Set_Navigation_Pane False                                'turn off navigation pane
        Application.ScreenRefresh
    End If
    
    MS_Word_Config = "Word is configured with default settings"
    Sh_ConfiguredAs = "DEF"   ' see Sh_HandleDocumentActivated: record what is ACTUALLY in force

    ' The settings braille and large print switch off - five spelling and grammar ones, and the
    ' tab-indent key. Nothing put them back before 8/18/2026 - see the top of this module.
    Sh_Restore_Transcriber_Settings

    ' Spell-check-as-you-type and the suggestions-from-the-spelling-checker box, both ON - Jerry,
    ' 3.0.222. He asked for the AutoCorrect box ticked in the ordinary configuration; the first
    ' line is what makes the second one possible. ReplaceTextFromSpellingChecker is GATED on
    ' CheckSpellingAsYouType: while that is False, Word greys the box out in the dialog and
    ' refuses the assignment silently - no error raised, the value simply stays False. That is
    ' what made 3.0.218 and 3.0.219 look as though they wrote it. MUST STAY ABOVE IT.
    '
    ' Both books have done this same pair since 3.0.220, so from here all three configurations
    ' agree and the box is ticked wherever she is working. The cost, accepted by Jerry on
    ' 8/22/2026: red squiggles in her ordinary letters as well as in her books, and her own choice
    ' on spell-check-as-you-type overridden rather than handed back.
    '
    ' BELOW Sh_Restore_Transcriber_Settings on purpose - run above it, the restore would put her
    ' stored value straight back over the top of both. And both names came OUT of the ledger the
    ' same day: a setting all three configurations force is no longer a preference, and left in
    ' the list the save would have written this True back as though she had chosen it.
    If Options.CheckSpellingAsYouType <> True Then Options.CheckSpellingAsYouType = True
    If AutoCorrect.ReplaceTextFromSpellingChecker <> True Then AutoCorrect.ReplaceTextFromSpellingChecker = True

    ' The compact fractions. Until 8/18/2026 this sub DELETED all eighteen. Be clear about what
    ' that did and did not cost: Word itself does only three - one half, one quarter and three
    ' quarters - and it does them through AutoFormatAsYouTypeReplaceFractions above, not through
    ' an AutoCorrect entry, so those three went on working. The other fifteen (1/3, 2/3, 1/5 and
    ' the rest) exist ONLY as the entries the braille configuration adds, and these deletes were
    ' written the same day, 10/19/2021, to take them straight back out again.
    ' Jerry, 8/18/2026: an ordinary document should have all eighteen, so the same list the
    ' braille configuration uses is added here instead. They then STAY in the transcriber's own
    ' AutoCorrect list. Large print still deletes them - a large print book keeps 1/2 as typed.
    ' Runs LAST, and below the two lines above on purpose: if it ever raised, Sh_Apply_Word_Config's
    ' handler would blank Sh_ConfiguredAs and Doc Info would report the previous document's setup.
    Sh_Add_Compact_Fractions

End Sub '*** end of MS_Set_Word_Config_For_New_Install ***

Sub MS_Set_Word_Config_For_Large_Print()
    '
    ' MS_Set_Word_Config_For_Large_Print Macro
    '
    ' Author: Jerry Whittaker -  jerry@thewhittakers.org
    '
    ' Version: 2.6  Date: 8/21/2026 - the three AutoFormat As You Type boxes this sub had never written are
    '                                 written now, all False: built-in heading styles, the Markdown box
    '                                 (AutoFormatAsYouTypeReplacePlainTextEmphasis) and define-styles-from-
    '                                 formatting. Jerry found all three checked in a large print document
    '                                 testing 3.0.216. The on-demand AutoFormat half had been correct all
    '                                 along, which is why this went unnoticed. Tracked too - 31 -> 34
    ' Version: 2.5  Date: 8/21/2026 - smart quotes and hyperlinks are written again, both forms, and
    '                                 Sh_Note_Book_Settings moved below the trailing writes (Jerry, 3.0.215)
    ' Version: 2.4  Date: 8/20/2026 - records what it applied, for the ledger - Sh_Note_Book_Settings
    ' Version: 2.3  Date: 8/20/2026 - opens the Styles pane for a large print document, the mirror image of what
    '                                 the braille and default configurations have always done. Visibility only:
    '                                 the Recommended sort and filter stay the transcriber's, see the note below
    ' Version: 2.2  Date: 8/18/2026 - dropped the 15 settings that were identical in all three configurations
    '                                 (see MS_Set_Word_Config_For_New_Install 2.5), stopped writing
    '                                 CorrectKeyboardSetting, and every remaining write is now guarded with
    '                                 If <> - about a dozen places run this sub directly while large print is
    '                                 ALREADY in force, and those wrote all 37 settings blind. 37 -> 22.
    ' Version: 2.1  Date: 8/18/2026 - no code change: the REASON the compact fractions are deleted is
    '                                 now written down beside them (Jerry's cardinal rule for large
    '                                 print - a compact fraction's digits are smaller than the base
    '                                 font size). Nothing else in this sub was touched on 8/18/2026.
    ' Version: 2.0  Date: 8/2/2026 - no longer calls Lp_Turn_on_Styles_Pane, so configuring Word for large print no longer seizes the user's Styles pane; FormattingShowNextLevel, which rode inside that call, is now written here
    ' Version: 1.9  Date: 5/13/2025 - added Application.ShowStylePreviews = True
    ' Version: 1.8  Date: 5/7/2025  -  added Application.RestrictLinkedStyles = True
    ' Version: 1.7  Date: 2/20/2024 - added Application.Run MacroName:="Sh_Is_Doc_Open"
    ' Version: 1.6  Date: 10/19/2021 - added Delete compact fractions
    ' Version: 1.5  Date:  2/25/2020 - added turn on print view and show styles pane
    ' Version: 1.4  Date: 1/8/2020 - closed reading view and added on error trap
    ' Version: 1.3  Date: 4/14/2017
    '
    Application.Run MacroName:="Sh_Is_Doc_Open"
    
    ActiveDocument.ActiveWindow.View.ReadingLayout = False  'will crash if document is in reading view ... close reading view

    ' Note the settings this sub is about to take away, but only if the ORDINARY configuration
    ' is in force - see Sh_Save_Transcriber_Settings.
    If Sh_ConfiguredAs = "DEF" Then Sh_Save_Transcriber_Settings
    
    With Options
        If .AutoFormatAsYouTypeApplyBorders <> False Then .AutoFormatAsYouTypeApplyBorders = False
        If .AutoFormatAsYouTypeApplyBulletedLists <> False Then .AutoFormatAsYouTypeApplyBulletedLists = False
        If .AutoFormatAsYouTypeApplyNumberedLists <> False Then .AutoFormatAsYouTypeApplyNumberedLists = False
        If .AutoFormatAsYouTypeApplyTables <> False Then .AutoFormatAsYouTypeApplyTables = False

        ' Built-in Heading styles, as you type. Jerry, 8/21/2026, testing 3.0.216: this box was
        ' checked in a large print document and should not be. It had never been written by
        ' anything - the on-demand half (AutoFormatApplyHeadings, below) was there from the start
        ' and the as-you-type half was simply missed, so Word's own True stood in a book. Typing
        ' a short line and pressing Enter twice silently applies Heading 1, which in large print
        ' takes the transcriber's body size away and replaces it with the style's own.
        If .AutoFormatAsYouTypeApplyHeadings <> False Then .AutoFormatAsYouTypeApplyHeadings = False

        ' "Hyphens (--) with dash", as you type. ON from 8/21/2026 - Jerry, 3.0.217. Note the
        ' name: there is no Hyphens property, and this is the box Word used to label "Symbol
        ' characters (--) with symbols". ReplaceSymbols is the only candidate in the object
        ' model - AutoFormatAsYouTypeReplaceFarEastDashes is the East Asian one, not this.
        If .AutoFormatAsYouTypeReplaceSymbols <> True Then .AutoFormatAsYouTypeReplaceSymbols = True
        ' "Ordinals (1st) with superscript", as you type. ON from 8/21/2026 - Jerry, 3.0.217.
        ' Deliberately NOT matched on the on-demand tab, which keeps AutoFormatReplaceOrdinals
        ' False: he asked for this tab by name and only this one.
        If .AutoFormatAsYouTypeReplaceOrdinals <> True Then .AutoFormatAsYouTypeReplaceOrdinals = True
        If .AutoFormatAsYouTypeReplaceFractions <> False Then .AutoFormatAsYouTypeReplaceFractions = False

        ' The box the dialog now calls "Markdown for heading, bold, italic and strikethrough".
        ' There is no property of that name: Word relabelled the old "*Bold* and _italic_ with
        ' real formatting" checkbox when it widened the feature, and this is still what backs it.
        ' Confirmed against the 22 AutoFormatAsYouType properties Word 16.0 exposes - none of them
        ' mentions Markdown. Off in a book for the same reason as the heading styles: it rewrites
        ' the transcriber's formatting from punctuation she may have typed deliberately.
        If .AutoFormatAsYouTypeReplacePlainTextEmphasis <> False Then .AutoFormatAsYouTypeReplacePlainTextEmphasis = False

        If .AutoFormatAsYouTypeFormatListItemBeginning <> False Then .AutoFormatAsYouTypeFormatListItemBeginning = False


        ' "Define styles based on your formatting" - never written either, same 8/21/2026 report.
        ' The worst of the three for large print: it lets Word REDEFINE a style out from under the
        ' template because of something typed in one paragraph, which is exactly what attaching an
        ' LP template is meant to prevent.
        If .AutoFormatAsYouTypeDefineStyles <> False Then .AutoFormatAsYouTypeDefineStyles = False

        ' Smart quotes, ON as a BOOK setting since 8/21/2026 - Jerry. Both books forced them on
        ' until 8/18/2026, when they were dropped from all three configurations as "identical
        ' everywhere". Identical is not the same as unimportant: dropping them left Word's own
        ' value standing in a book, and a book needs these on. In Sh_Tracked_Settings, so her own
        ' value comes back in her letters.
        If .AutoFormatAsYouTypeReplaceQuotes <> True Then .AutoFormatAsYouTypeReplaceQuotes = True

        ' Hyperlinks, OFF as you type from 3.0.221 - Jerry. Large print is now off on BOTH tabs,
        ' having gone off on the on-demand tab first and had this half left on because he named
        ' that tab alone. Braille is the opposite and is on for both, decided separately the same
        ' day; the two books genuinely differ here and neither should be aligned to the other.
        If .AutoFormatAsYouTypeReplaceHyperlinks <> False Then .AutoFormatAsYouTypeReplaceHyperlinks = False
    End With
    
    ' Spell-check-as-you-type, ON - Jerry, 3.0.220. It is here for its own sake AND because the
    ' AutoCorrect setting below is GATED on it: while this is False, Word greys the suggestions
    ' box out in the dialog and silently refuses the assignment, raising no error and leaving the
    ' value False. That is why 3.0.218 and 3.0.219 appeared to write it and did nothing. Proved on
    ' the build box 8/21/2026: with this True the very next assignment sticks. MUST STAY ABOVE IT.
    If Options.CheckSpellingAsYouType <> True Then Options.CheckSpellingAsYouType = True

    With AutoCorrect
        If .CorrectSentenceCaps <> False Then .CorrectSentenceCaps = False
        If .CorrectTableCells <> False Then .CorrectTableCells = False

        ' The suggestions-from-the-spelling-checker box. GATED on Options.CheckSpellingAsYouType,
        ' set True just above - see the note there. Do not move this above it.
        If .ReplaceTextFromSpellingChecker <> True Then .ReplaceTextFromSpellingChecker = True
    End With

    With Options
        If .AutoFormatApplyHeadings <> False Then .AutoFormatApplyHeadings = False
        If .AutoFormatApplyLists <> False Then .AutoFormatApplyLists = False
        If .AutoFormatApplyBulletedLists <> False Then .AutoFormatApplyBulletedLists = False
        If .AutoFormatApplyOtherParas <> False Then .AutoFormatApplyOtherParas = False
        ' "Hyphens (--) with dash", on demand - the other half of the as-you-type box above.
        If .AutoFormatReplaceSymbols <> True Then .AutoFormatReplaceSymbols = True
        If .AutoFormatReplaceOrdinals <> False Then .AutoFormatReplaceOrdinals = False
        If .AutoFormatReplaceFractions <> False Then .AutoFormatReplaceFractions = False
        If .AutoFormatPreserveStyles <> True Then .AutoFormatPreserveStyles = True
        If .AutoFormatPlainTextWordMail <> False Then .AutoFormatPlainTextWordMail = False
        ' The on-demand half of the same pair - see the note above. Smart quotes only now:
        ' "Internet and network paths with hyperlinks" is OFF here from 8/21/2026 (Jerry, 3.0.217),
        ' which reverses half of what went in earlier the same day. The AS-YOU-TYPE hyperlink box
        ' stays ON - he named this tab and only this tab, so the two halves differ on purpose.
        If .AutoFormatReplaceQuotes <> True Then .AutoFormatReplaceQuotes = True
        If .AutoFormatReplaceHyperlinks <> False Then .AutoFormatReplaceHyperlinks = False
    End With

    If Not Sh_Config_Skip_Display Then   ' the transcriber's screen is theirs - see Sh_Config_Skip_Display
        ActiveWindow.StyleAreaWidth = 24.5
        'Application.Options.ShowCropMarks = True
        ActiveWindow.View.ShowAll = True
        ActiveWindow.DisplayRulers = True
        ActiveWindow.DisplayVerticalRuler = True
    End If
    ' The Styles pane is OPEN for a large print document. Jerry, 8/20/2026, after attaching the
    ' template to a book, saving it, opening it again and finding no pane. This completes a
    ' pattern the other two configurations have had all along - MS_Set_Word_Config_For_Braille
    ' hides the pane and so does MS_Set_Word_Config_For_New_Install, and large print was the one
    ' of the three that said nothing, so it showed whatever the last document left behind.
    '
    ' VISIBILITY ONLY, and that is the whole of the 8/2/2026 note this replaces. The Recommended
    ' sort and the "Select styles to show" filter are NOT set here, and Lp_Turn_on_Styles_Pane
    ' must not be put back: those two are DOCUMENT properties saved into the file, so a book the
    ' attach set to Recommended reopens that way by itself, while a transcriber who chose "All
    ' Styles" in her own book keeps it. Forcing those on every open was the 7/24 and 8/2/2026
    ' complaint and is not being reopened - only whether the pane is on screen has changed.
    '
    ' Inside the skip-display guard, so merely glancing at another book cannot move it: on a
    ' SWITCH between open documents this block does not run at all.
    '
    ' FormattingShowNextLevel was riding INSIDE the removed Lp_Turn_on_Styles_Pane call and is
    ' nowhere else in this sub, so it is written here explicitly rather than lost with it.
    If Not Sh_Config_Skip_Display Then   ' the transcriber's screen is theirs - see Sh_Config_Skip_Display
        Application.TaskPanes(wdTaskPaneFormatting).Visible = True 'turn on styles pane
        ActiveDocument.FormattingShowNextLevel = False
        Application.ShowStylePreviews = True
        Application.RestrictLinkedStyles = True
        ActiveDocument.FormattingShowUserStyleName = False
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    If Not Sh_Config_Skip_Display Then Application.ScreenRefresh
    
    ' Delete the 18 compact fraction AutoCorrect entries. This is not a preference and it is not
    ' negotiable, so do not remove it to save the writes:
    '
    '   A compact fraction is ONE character, and its digits are drawn far smaller than the base
    '   font size - on the screen and on the printed page alike. Large print's cardinal rule is
    '   that every character is the size the transcriber set. A book at 18 pt containing a single
    '   'one half' character has text in it that the reader it was made for cannot read.
    '   Jerry, 8/18/2026.
    '
    ' So the entries go, and typing 1/2 in a large print book leaves 1/2 on the page. Characters
    ' that arrive some other way - pasted in from a source document - are caught separately, by
    ' Lp_Replace_Compact_Fractions_With_Fraction_Text, inside Lp_Fix_Common_File_Errors - which
    ' Full File Cleanup reaches through its menu form, and which also runs when the LP template is
    ' attached. So a pasted-in one is corrected when the book is cleaned up, not as it lands.
    '
    ' They may be present because the braille configuration adds them, and, from 8/18/2026, so
    ' does the default one (Jerry: an ordinary document should have them).
    On Error Resume Next
    AutoCorrect.Entries("1/2").Delete
    AutoCorrect.Entries("1/3").Delete
    AutoCorrect.Entries("2/3").Delete
    AutoCorrect.Entries("1/4").Delete
    AutoCorrect.Entries("3/4").Delete
    AutoCorrect.Entries("1/5").Delete
    AutoCorrect.Entries("2/5").Delete
    AutoCorrect.Entries("3/5").Delete
    AutoCorrect.Entries("4/5").Delete
    AutoCorrect.Entries("1/6").Delete
    AutoCorrect.Entries("5/6").Delete
    AutoCorrect.Entries("1/7").Delete
    AutoCorrect.Entries("1/8").Delete
    AutoCorrect.Entries("3/8").Delete
    AutoCorrect.Entries("5/8").Delete
    AutoCorrect.Entries("7/8").Delete
    AutoCorrect.Entries("1/9").Delete
    AutoCorrect.Entries("1/10").Delete

    MS_Word_Config = "Word is configured for large print"
    Sh_ConfiguredAs = "LP"    ' see Sh_HandleDocumentActivated: record what is ACTUALLY in force

    ' These come LAST, after the line above that records what is in force. If this sub raises
    ' part way through - the display block can, on an odd window - the five must not already be
    ' switched off while Sh_ConfiguredAs still says "DEF", because the next direct call would then
    ' save VistaType's own switched-off values as the transcriber's. Written here, a sub that dies
    ' early leaves them untouched. See Sh_Save_Transcriber_Settings.
    If Options.LabelSmartTags <> False Then Options.LabelSmartTags = False
    If Options.IgnoreUppercase <> False Then Options.IgnoreUppercase = False

    ' "Set left- and first-indent with tabs and backspaces" - ON from 8/21/2026, Jerry, 3.0.220.
    ' Large print had never written this; braille switched it OFF, and because it is one global
    ' Word setting rather than a per-document one, a braille session left it off underneath large
    ' print too. Sits with the two writes above, below the Sh_ConfiguredAs line, for the same
    ' reason they do. See the braille sub for how this checkbox was identified.
    If Options.TabIndentKey <> True Then Options.TabIndentKey = True

    ' Write down what this configuration has just applied, so that returning to an ordinary
    ' document can tell a setting SHE changed from one this sub set. THE VERY LAST THING THE SUB
    ' DOES, below the two writes above and not above them - see Sh_Note_Book_Settings 1.1.
    Sh_Note_Book_Settings
    
    On Error GoTo 0
End Sub  '*** end of macro MS_Set_Word_Config_For_Large_Print ***

Sub MS_Set_Word_Config_For_Braille()
    
    '
    ' Author: Jerry Whittaker -  jerry@thewhittakers.org
    '
    ' Version: 2.5  Date: 8/21/2026 - smart quotes and hyperlinks are written again, both forms, and
    '                                 Sh_Note_Book_Settings moved below the trailing writes (Jerry, 3.0.215)
    ' Version: 2.4  Date: 8/20/2026 - records what it applied, for the ledger - Sh_Note_Book_Settings
    ' Version: 2.3  Date: 8/18/2026 - "Set left- and first-indent with tabs and backspaces" (TabIndentKey) is
    '                                 switched OFF for braille (Jerry). It was one of the 15 removed earlier the
    '                                 same day as identical everywhere, so it comes back as a SWITCHED setting,
    '                                 noted and put back through the store rather than left off for good.
    ' Version: 2.2  Date: 8/18/2026 - dropped the 15 settings that were identical in all three configurations
    '                                 (see MS_Set_Word_Config_For_New_Install 2.5), stopped writing
    '                                 CorrectKeyboardSetting, guarded every remaining write with If <>, and
    '                                 noted the five spelling and grammar settings on the way in so the
    '                                 default configuration can put them back. 40 -> 25.
    ' Version: 2.1  Date: 8/18/2026 - the 18 compact fraction entries moved out into Sh_Add_Compact_Fractions,
    '                                 now shared with MS_Set_Word_Config_For_New_Install. Same eighteen, same
    '                                 values; braille behavior unchanged.
    ' Version: 2.0  Date:  10/27/2021 - Turned ruler display on - Disable linked styles in styles pane
    ' Version: 1.9  Date:  10/23/2021 - added     ' Options.CheckGrammarAsYouType = False
                                                                    ' Options.IgnoreMixedDigits = False
                                                                    ' Options.ContextualSpeller = False
    ' Version: 1.8  Date: 10/19/2021 - Added fractions to autocorrect
    ' Version: 1.7  Date: 3/22/2021 - added Susan's new config options
    ' Version: 1.6  Date: 1/8/2020 - closed reading view and added on error trap
    ' Version: 1.5  Date: 10/28/2019 - minor changes per Susan Christensen new handout
    ' Version: 1.4  Date: 6/13/2019 - turned off ruler display
    ' Version: 1.4  Date: 8/8/2018
    '
    ActiveDocument.ActiveWindow.View.ReadingLayout = False  'will crash if document is in reading view ... close reading view
    ' Note the settings this sub is about to take away, but only if the ORDINARY configuration
    ' is in force - see Sh_Save_Transcriber_Settings.
    If Sh_ConfiguredAs = "DEF" Then Sh_Save_Transcriber_Settings

    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating off
    If Not Sh_Config_Skip_Display Then   ' the transcriber's screen is theirs - see Sh_Config_Skip_Display
        ActiveDocument.FormattingShowNextLevel = False
        ActiveDocument.StyleSortMethod = wdStyleSortRecommended
        ActiveDocument.FormattingShowFilter = wdShowFilterFormattingRecommended
        Application.RestrictLinkedStyles = True
    End If
    
    With Options
        If .AutoFormatAsYouTypeApplyBorders <> False Then .AutoFormatAsYouTypeApplyBorders = False
        If .AutoFormatAsYouTypeApplyBulletedLists <> False Then .AutoFormatAsYouTypeApplyBulletedLists = False
        If .AutoFormatAsYouTypeApplyNumberedLists <> False Then .AutoFormatAsYouTypeApplyNumberedLists = False
        If .AutoFormatAsYouTypeApplyTables <> False Then .AutoFormatAsYouTypeApplyTables = False
        If .AutoFormatAsYouTypeReplaceSymbols <> True Then .AutoFormatAsYouTypeReplaceSymbols = True
        If .AutoFormatAsYouTypeReplaceOrdinals <> True Then .AutoFormatAsYouTypeReplaceOrdinals = True
        If .AutoFormatAsYouTypeReplaceFractions <> True Then .AutoFormatAsYouTypeReplaceFractions = True
        If .AutoFormatAsYouTypeFormatListItemBeginning <> False Then .AutoFormatAsYouTypeFormatListItemBeginning = False


        ' The "Markdown for heading, bold, italic and strikethrough" box, off - and KNOWN NOT TO
        ' WORK YET. Written here only so the two books stay in step and there is one place to fix.
        ' 3.0.217 shipped this same write for large print; Jerry installed it and the box was still
        ' checked, so ReplacePlainTextEmphasis is NOT what backs that checkbox, and Word exposes no
        ' property matching Markdown at all. Whatever this property really does control should be
        ' off in a book anyway, so it is harmless where it stands. See the note in large print.
        If .AutoFormatAsYouTypeReplacePlainTextEmphasis <> False Then .AutoFormatAsYouTypeReplacePlainTextEmphasis = False

        ' Smart quotes and hyperlinks, back as BOOK settings on 8/21/2026 - Jerry. Both books
        ' forced them on until 8/18/2026, when they were dropped from all three configurations
        ' as "identical everywhere". Identical is not the same as unimportant: dropping them
        ' left Word's own value standing in a book, and a book needs these on. They are in
        ' Sh_Tracked_Settings, so her own value comes back in her letters.
        If .AutoFormatAsYouTypeReplaceQuotes <> True Then .AutoFormatAsYouTypeReplaceQuotes = True
        ' "Internet and network paths with hyperlinks", ON in braille. This box has now been
        ' turned off and on again within one day: off in 3.0.218 at Jerry's request, back on in
        ' 3.0.219 once he saw it in the tab. ON is where it rests. Large print is NOT the same -
        ' there the as-you-type box is on and the on-demand one off, which he has left alone
        ' through both rounds, so the two books genuinely differ here.
        If .AutoFormatAsYouTypeReplaceHyperlinks <> True Then .AutoFormatAsYouTypeReplaceHyperlinks = True
    End With
    
    ' Spell-check-as-you-type, ON - Jerry, 3.0.220. It is here for its own sake AND because the
    ' AutoCorrect setting below is GATED on it: while this is False, Word greys the suggestions
    ' box out in the dialog and silently refuses the assignment, raising no error and leaving the
    ' value False. That is why 3.0.218 and 3.0.219 appeared to write it and did nothing. Proved on
    ' the build box 8/21/2026: with this True the very next assignment sticks. MUST STAY ABOVE IT.
    If Options.CheckSpellingAsYouType <> True Then Options.CheckSpellingAsYouType = True

    With AutoCorrect
        If .CorrectSentenceCaps <> False Then .CorrectSentenceCaps = False
        If .CorrectTableCells <> False Then .CorrectTableCells = False

        ' The suggestions-from-the-spelling-checker box. GATED on Options.CheckSpellingAsYouType,
        ' set True just above - see the note there. Do not move this above it.
        If .ReplaceTextFromSpellingChecker <> True Then .ReplaceTextFromSpellingChecker = True
    End With

    With Options
        If .AutoFormatApplyHeadings <> False Then .AutoFormatApplyHeadings = False
        If .AutoFormatApplyLists <> False Then .AutoFormatApplyLists = False
        If .AutoFormatApplyBulletedLists <> False Then .AutoFormatApplyBulletedLists = False
        If .AutoFormatApplyOtherParas <> False Then .AutoFormatApplyOtherParas = False
        If .AutoFormatReplaceSymbols <> True Then .AutoFormatReplaceSymbols = True
        If .AutoFormatReplaceOrdinals <> True Then .AutoFormatReplaceOrdinals = True
        If .AutoFormatReplaceFractions <> True Then .AutoFormatReplaceFractions = True
        ' The "Styles" box in the AutoFormat tab's Preserve group. ON from 8/21/2026 - Jerry,
        ' 3.0.217. Large print has had this True throughout; braille forced it False and now
        ' agrees with it.
        If .AutoFormatPreserveStyles <> True Then .AutoFormatPreserveStyles = True
        If .AutoFormatPlainTextWordMail <> False Then .AutoFormatPlainTextWordMail = False
        ' The on-demand half of the same pair - see the note above.
        If .AutoFormatReplaceQuotes <> True Then .AutoFormatReplaceQuotes = True
        ' The on-demand half - see the note on the as-you-type box above.
        If .AutoFormatReplaceHyperlinks <> True Then .AutoFormatReplaceHyperlinks = True
    End With
    
    ' add the compact fractions to autocorrect - the list lives in Sh_Add_Compact_Fractions,
    ' shared with MS_Set_Word_Config_For_New_Install from 8/18/2026
    Sh_Add_Compact_Fractions
    
    
    If Not Sh_Config_Skip_Display Then   ' the transcriber's screen is theirs - see Sh_Config_Skip_Display
        ActiveWindow.StyleAreaWidth = 64.5
        'ActiveWindow.View.ShowAll = True
        ActiveWindow.DisplayRulers = True
        ActiveWindow.DisplayVerticalRuler = False
        'Application.Options.ShowCropMarks = False

        Application.TaskPanes(wdTaskPaneFormatting).Visible = False 'turn off styles pane
        ActiveWindow.ActivePane.View.Type = wdNormalView
    End If

    MS_Word_Config = "Word is configured for braille"
    Sh_ConfiguredAs = "BRL"   ' see Sh_HandleDocumentActivated: record what is ACTUALLY in force

    ' These come LAST, after the line above that records what is in force. If this sub raises
    ' part way through - the display block can, on an odd window - the five must not already be
    ' switched off while Sh_ConfiguredAs still says "DEF", because the next direct call would then
    ' save VistaType's own switched-off values as the transcriber's. Written here, a sub that dies
    ' early leaves them untouched. See Sh_Save_Transcriber_Settings.
    If Options.CheckGrammarAsYouType <> False Then Options.CheckGrammarAsYouType = False
    If Options.IgnoreMixedDigits <> False Then Options.IgnoreMixedDigits = False
    If Options.ContextualSpeller <> False Then Options.ContextualSpeller = False
    If Options.LabelSmartTags <> False Then Options.LabelSmartTags = False
    If Options.IgnoreUppercase <> False Then Options.IgnoreUppercase = False
    ' "Set left- and first-indent with tabs and backspaces" - off for braille. Jerry, 8/18/2026.
    ' "Set left- and first-indent with tabs and backspaces" - ON from 8/21/2026, Jerry, 3.0.220.
    ' Braille had switched it off since 8/18/2026. THIS is that checkbox: proved on the build box
    ' by parking the add-in, ticking the box by hand and reading back - TabIndentKey went False to
    ' True and nothing else moved. 3.0.218 and 3.0.219 aimed at AutoFormatAsYouTypeApplyFirstIndents
    ' instead, which held True the whole time while the box sat unticked; that write is gone from
    ' both books. The name at the top of this module said TabIndentKey all along.
    If Options.TabIndentKey <> True Then Options.TabIndentKey = True

    ' Write down what this configuration has just applied, so that returning to an ordinary
    ' document can tell a setting SHE changed from one this sub set. THE VERY LAST THING THE SUB
    ' DOES, below the six writes above and not above them - see Sh_Note_Book_Settings 1.1.
    Sh_Note_Book_Settings
    
    Application.ScreenUpdating = su_Prev ' Turn screen updating on
    If Not Sh_Config_Skip_Display Then Application.ScreenRefresh

End Sub  '*** end of  MS_Set_Word_Config_For_Braille macro***

Sub Sh_Add_Compact_Fractions()
'
' Adds the 18 compact fraction AutoCorrect entries - 1/2 -> ½ and the rest.
'
' Only three of the eighteen are Word's own, and Word does those through the AutoFormat as you
' type option, not through an AutoCorrect entry. The other fifteen are this add-in's.
'
' ONE copy of the list. Called by MS_Set_Word_Config_For_Braille and, from 8/18/2026, by
' MS_Set_Word_Config_For_New_Install as well: two copies of eighteen entries would drift.
' MS_Set_Word_Config_For_Large_Print deletes them instead - a large print book keeps 1/2
' as it was typed - so that list stays where it is.
'
' Writes only an entry that is MISSING or holds something else, for the same reason
' MS_Set_Word_Config_For_New_Install writes only the Options that differ (7/18/2026): that
' sub runs on every ordinary document, and eighteen needless writes to Word's AutoCorrect
' store on each one is exactly the churn that brought on Office's "restart to apply your
' privacy settings" notice.
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version: 1.0  Date: 8/18/2026
'
    Sh_Add_One_Compact_Fraction "1/2", ChrW(189)
    Sh_Add_One_Compact_Fraction "1/3", ChrW(8531)
    Sh_Add_One_Compact_Fraction "2/3", ChrW(8532)
    Sh_Add_One_Compact_Fraction "1/4", ChrW(188)
    Sh_Add_One_Compact_Fraction "3/4", ChrW(190)
    Sh_Add_One_Compact_Fraction "1/5", ChrW(8533)
    Sh_Add_One_Compact_Fraction "2/5", ChrW(8534)
    Sh_Add_One_Compact_Fraction "3/5", ChrW(8535)
    Sh_Add_One_Compact_Fraction "4/5", ChrW(8536)
    Sh_Add_One_Compact_Fraction "1/6", ChrW(8537)
    Sh_Add_One_Compact_Fraction "5/6", ChrW(8538)
    Sh_Add_One_Compact_Fraction "1/7", ChrW(8528)
    Sh_Add_One_Compact_Fraction "1/8", ChrW(8539)
    Sh_Add_One_Compact_Fraction "3/8", ChrW(8540)
    Sh_Add_One_Compact_Fraction "5/8", ChrW(8541)
    Sh_Add_One_Compact_Fraction "7/8", ChrW(8542)
    Sh_Add_One_Compact_Fraction "1/9", ChrW(8529)
    Sh_Add_One_Compact_Fraction "1/10", ChrW(8530)

End Sub  '*** end of Sh_Add_Compact_Fractions ***

Private Sub Sh_Add_One_Compact_Fraction(ByVal fractionTyped As String, ByVal fractionChar As String)
'
' One entry for Sh_Add_Compact_Fractions. Reading an entry that is not there raises an
' error, so the read is trapped and an absent entry simply reads as empty and gets added.
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version: 1.0  Date: 8/18/2026
'
    Dim inForce As String

    On Error Resume Next
    inForce = AutoCorrect.Entries(fractionTyped).Value
    On Error GoTo 0

    If inForce <> fractionChar Then AutoCorrect.Entries.Add Name:=fractionTyped, Value:=fractionChar

End Sub  '*** end of Sh_Add_One_Compact_Fraction ***

' --- The settings store -------------------------------------------------------------------
' One small file, written and read with Word's own System.PrivateProfileString, so no file
' handling has to be written here. RibbonCallbacks.bas already uses the same call.
'
' It lives in %AppData%\VistaType LP Settings and NOT in %AppData%\VistaType LP, deliberately:
' the uninstaller deletes that second folder whole, and a reinstall must not cost the
' transcriber her settings. Same reasoning that puts the font license in its own folder - see
' installer/vistatype.iss. A file rather than the registry because transcribers get new
' computers and a file can be copied to one.
'
' Version: 1.0  Date: 8/18/2026
'

Private Function Sh_Settings_File() As String
'
' The full path of the store, creating its folder the first time. Returns "" if it cannot be
' had at all - every caller treats that as "nothing is remembered" and leaves Word alone,
' which is the safe direction: no store is far better than a store holding the wrong answer.
'
' Version: 1.0  Date: 8/18/2026
'
    Dim fso As Object
    Dim folder As String

    On Error Resume Next
    folder = Environ$("APPDATA")
    If folder = "" Then Exit Function
    folder = folder & "\" & VT_STORE_FOLDER

    ' FileSystemObject rather than Dir(): Dir is stateful, and a Dir loop running anywhere else
    ' in the project would be silently restarted by a call from here.
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    If Not fso.FolderExists(folder) Then Exit Function

    Sh_Settings_File = folder & "\" & VT_STORE_FILE

End Function  '*** end of Sh_Settings_File ***

Public Function Sh_Setting_Read(ByVal storeSection As String, ByVal storeKey As String, _
                                ByVal whenMissing As String) As String
'
' One value out of the store, or whenMissing if there is no store or no such value.
'
' Version: 1.0  Date: 8/18/2026
'
    Dim storePath As String
    Dim held As String

    Sh_Setting_Read = whenMissing

    storePath = Sh_Settings_File()
    If storePath = "" Then Exit Function

    On Error Resume Next
    held = System.PrivateProfileString(storePath, storeSection, storeKey)
    On Error GoTo 0

    If held <> "" Then Sh_Setting_Read = held

End Function  '*** end of Sh_Setting_Read ***

Public Sub Sh_Setting_Write(ByVal storeSection As String, ByVal storeKey As String, _
                            ByVal storeValue As String)
'
' One value into the store. Silently does nothing if there is no store to write to - see
' Sh_Settings_File. Never write "": that DELETES the key rather than storing an empty value.
'
' Version: 1.0  Date: 8/18/2026
'
    Dim storePath As String

    storePath = Sh_Settings_File()
    If storePath = "" Then Exit Sub
    If storeValue = "" Then Exit Sub

    On Error Resume Next
    System.PrivateProfileString(storePath, storeSection, storeKey) = storeValue
    On Error GoTo 0

End Sub  '*** end of Sh_Setting_Write ***

Private Function Sh_Bool_To_Store(ByVal flag As Boolean) As String
    If flag Then Sh_Bool_To_Store = "1" Else Sh_Bool_To_Store = "0"
End Function

Private Function Sh_Store_To_Bool(ByVal held As String) As Boolean
    Sh_Store_To_Bool = (held = "1")
End Function

' EVERY setting a large print or braille configuration takes away from her, in one list, walked
' by all three of the subs below - the save, the compare and the restore. Thirty-four of them,
' and the list is not guesswork: it is the union of what MS_Set_Word_Config_For_Large_Print and
' MS_Set_Word_Config_For_Braille actually write, read out of those two subs on 8/20/2026, plus
' the four smart-quote and hyperlink settings the books took back on 8/21/2026.
'
' ADDING A SETTING TO EITHER BOOK CONFIGURATION MEANS ADDING IT HERE, and to both Select Cases
' below, and bumping VT_STORE_STAMP_NOW. A setting a book writes and this list does not name is
' one she never gets back - which is exactly the fault found on 8/18/2026, when one braille file
' turned grammar checking off for every document she opened, that session and every session after.
'
' Twenty of these the ordinary configuration used to rewrite to Word's factory value, which is
' why nobody noticed: it looked like restoring and was overwriting. Seven more it never put back
' at all - the six spelling and grammar ones found on 8/18, plus ShowStylePreviews and
' RestrictLinkedStyles, which the books switch on and nothing has ever switched off.
'
' The last four are smart quotes and hyperlinks, in both their as-you-type and on-demand forms.
' Both books forced them on until 8/18/2026, when they went with the fifteen settings that were
' identical in all three configurations. Jerry, 8/21/2026, testing 3.0.215: identical is not the
' same as unimportant. Dropping them left WORD's value standing inside a book, where the work
' needs them on. Back in the two book configurations only - never in the ordinary one, which is
' the difference between a book setting and a preference imposed on her Word.
'
' Version: 1.5  Date: 8/22/2026 - CheckSpellingAsYouType and ReplaceTextFromSpellingChecker come OUT:
'                                 the ordinary configuration forces both from 3.0.222, so no configuration
'                                 leaves her a value to be given back. 36 -> 34
' Version: 1.4  Date: 8/21/2026 - AutoFormatAsYouTypeApplyFirstIndents comes back OUT: it is not the
'                                 checkbox anybody thought it was (see MS_Set_Word_Config_For_Braille),
'                                 nothing writes it again, and it is hers. 36 -> 35
' Version: 1.3  Date: 8/21/2026 - AutoFormatAsYouTypeApplyFirstIndents and the AutoCorrect setting
'                                 ReplaceTextFromSpellingChecker, both written for the first time by
'                                 large print (Jerry, 3.0.217). 34 -> 36
' Version: 1.2  Date: 8/21/2026 - the three AutoFormat As You Type settings large print had never
'                                 written: ApplyHeadings, DefineStyles and ReplacePlainTextEmphasis
'                                 (the "Markdown" box). 31 -> 34. Braille does not write these three;
'                                 they are tracked so large print cannot keep her value once it has
'                                 taken it, exactly as with smart quotes below.
' Version: 1.1  Date: 8/21/2026 - the four smart-quote and hyperlink settings, 27 -> 31
' Version: 1.0  Date: 8/20/2026
Private Function Sh_Tracked_Settings() As Variant
    ' Built with Split rather than Array(...) and a line continuation per name: VBA allows at
    ' most 25 continuations in one statement and there are 34 names here, which does not fail
    ' at the line - it fails the whole module at import, with a COM error from the build script
    ' that says nothing about continuations. 8/20/2026, and it cost a build to find.
    Dim nm As String

    nm = "AutoFormatApplyBulletedLists,AutoFormatApplyHeadings,AutoFormatApplyLists,"
    nm = nm & "AutoFormatApplyOtherParas,AutoFormatAsYouTypeApplyBorders,AutoFormatAsYouTypeApplyBulletedLists,"
    nm = nm & "AutoFormatAsYouTypeApplyHeadings,AutoFormatAsYouTypeApplyNumberedLists,AutoFormatAsYouTypeApplyTables,"
    nm = nm & "AutoFormatAsYouTypeDefineStyles,AutoFormatAsYouTypeFormatListItemBeginning,AutoFormatAsYouTypeReplaceFractions,"
    nm = nm & "AutoFormatAsYouTypeReplaceHyperlinks,AutoFormatAsYouTypeReplaceOrdinals,AutoFormatAsYouTypeReplacePlainTextEmphasis,"
    nm = nm & "AutoFormatAsYouTypeReplaceQuotes,AutoFormatAsYouTypeReplaceSymbols,AutoFormatPlainTextWordMail,"
    nm = nm & "AutoFormatPreserveStyles,AutoFormatReplaceFractions,AutoFormatReplaceHyperlinks,"
    nm = nm & "AutoFormatReplaceOrdinals,AutoFormatReplaceQuotes,AutoFormatReplaceSymbols,"
    nm = nm & "CheckGrammarAsYouType,ContextualSpeller,"
    nm = nm & "CorrectSentenceCaps,CorrectTableCells,IgnoreMixedDigits,"
    nm = nm & "IgnoreUppercase,LabelSmartTags,"
    nm = nm & "RestrictLinkedStyles,ShowStylePreviews,TabIndentKey"

    Sh_Tracked_Settings = Split(nm, ",")
End Function  '*** end of Sh_Tracked_Settings ***

' What one tracked setting is set to RIGHT NOW. The three owners - Options, AutoCorrect and
' Application - are why this is a Select Case and not something cleverer: VBA cannot reach a
' property by name without CallByName, which is slower and silently returns Empty on a typo.
'
' Version: 1.0  Date: 8/20/2026
Private Function Sh_Setting_Live(ByVal nm As String) As Boolean
    On Error Resume Next
    Select Case nm
        Case "AutoFormatApplyBulletedLists": Sh_Setting_Live = Options.AutoFormatApplyBulletedLists
        Case "AutoFormatApplyHeadings": Sh_Setting_Live = Options.AutoFormatApplyHeadings
        Case "AutoFormatApplyLists": Sh_Setting_Live = Options.AutoFormatApplyLists
        Case "AutoFormatApplyOtherParas": Sh_Setting_Live = Options.AutoFormatApplyOtherParas
        Case "AutoFormatAsYouTypeApplyBorders": Sh_Setting_Live = Options.AutoFormatAsYouTypeApplyBorders
        Case "AutoFormatAsYouTypeApplyBulletedLists": Sh_Setting_Live = Options.AutoFormatAsYouTypeApplyBulletedLists
        Case "AutoFormatAsYouTypeApplyHeadings": Sh_Setting_Live = Options.AutoFormatAsYouTypeApplyHeadings
        Case "AutoFormatAsYouTypeApplyNumberedLists": Sh_Setting_Live = Options.AutoFormatAsYouTypeApplyNumberedLists
        Case "AutoFormatAsYouTypeApplyTables": Sh_Setting_Live = Options.AutoFormatAsYouTypeApplyTables
        Case "AutoFormatAsYouTypeDefineStyles": Sh_Setting_Live = Options.AutoFormatAsYouTypeDefineStyles
        Case "AutoFormatAsYouTypeFormatListItemBeginning": Sh_Setting_Live = Options.AutoFormatAsYouTypeFormatListItemBeginning
        Case "AutoFormatAsYouTypeReplaceFractions": Sh_Setting_Live = Options.AutoFormatAsYouTypeReplaceFractions
        Case "AutoFormatAsYouTypeReplaceHyperlinks": Sh_Setting_Live = Options.AutoFormatAsYouTypeReplaceHyperlinks
        Case "AutoFormatAsYouTypeReplaceOrdinals": Sh_Setting_Live = Options.AutoFormatAsYouTypeReplaceOrdinals
        Case "AutoFormatAsYouTypeReplacePlainTextEmphasis": Sh_Setting_Live = Options.AutoFormatAsYouTypeReplacePlainTextEmphasis
        Case "AutoFormatAsYouTypeReplaceQuotes": Sh_Setting_Live = Options.AutoFormatAsYouTypeReplaceQuotes
        Case "AutoFormatAsYouTypeReplaceSymbols": Sh_Setting_Live = Options.AutoFormatAsYouTypeReplaceSymbols
        Case "AutoFormatPlainTextWordMail": Sh_Setting_Live = Options.AutoFormatPlainTextWordMail
        Case "AutoFormatPreserveStyles": Sh_Setting_Live = Options.AutoFormatPreserveStyles
        Case "AutoFormatReplaceFractions": Sh_Setting_Live = Options.AutoFormatReplaceFractions
        Case "AutoFormatReplaceHyperlinks": Sh_Setting_Live = Options.AutoFormatReplaceHyperlinks
        Case "AutoFormatReplaceOrdinals": Sh_Setting_Live = Options.AutoFormatReplaceOrdinals
        Case "AutoFormatReplaceQuotes": Sh_Setting_Live = Options.AutoFormatReplaceQuotes
        Case "AutoFormatReplaceSymbols": Sh_Setting_Live = Options.AutoFormatReplaceSymbols
        Case "CheckGrammarAsYouType": Sh_Setting_Live = Options.CheckGrammarAsYouType
        Case "ContextualSpeller": Sh_Setting_Live = Options.ContextualSpeller
        Case "CorrectSentenceCaps": Sh_Setting_Live = AutoCorrect.CorrectSentenceCaps
        Case "CorrectTableCells": Sh_Setting_Live = AutoCorrect.CorrectTableCells
        Case "IgnoreMixedDigits": Sh_Setting_Live = Options.IgnoreMixedDigits
        Case "IgnoreUppercase": Sh_Setting_Live = Options.IgnoreUppercase
        Case "LabelSmartTags": Sh_Setting_Live = Options.LabelSmartTags
        Case "RestrictLinkedStyles": Sh_Setting_Live = Application.RestrictLinkedStyles
        Case "ShowStylePreviews": Sh_Setting_Live = Application.ShowStylePreviews
        Case "TabIndentKey": Sh_Setting_Live = Options.TabIndentKey
    End Select
    Err.Clear
End Function  '*** end of Sh_Setting_Live ***

' Put one tracked setting back. Guarded, like every other write in the three configurations:
' her value is usually already in force and Word does not need telling twice - needless writes to
' Options and AutoCorrect are what brought on Office's "restart to apply your privacy settings"
' notice on 7/18/2026.
'
' Version: 1.0  Date: 8/20/2026
Private Sub Sh_Setting_Put(ByVal nm As String, ByVal wanted As Boolean)
    On Error Resume Next
    Select Case nm
        Case "AutoFormatApplyBulletedLists": If Options.AutoFormatApplyBulletedLists <> wanted Then Options.AutoFormatApplyBulletedLists = wanted
        Case "AutoFormatApplyHeadings": If Options.AutoFormatApplyHeadings <> wanted Then Options.AutoFormatApplyHeadings = wanted
        Case "AutoFormatApplyLists": If Options.AutoFormatApplyLists <> wanted Then Options.AutoFormatApplyLists = wanted
        Case "AutoFormatApplyOtherParas": If Options.AutoFormatApplyOtherParas <> wanted Then Options.AutoFormatApplyOtherParas = wanted
        Case "AutoFormatAsYouTypeApplyBorders": If Options.AutoFormatAsYouTypeApplyBorders <> wanted Then Options.AutoFormatAsYouTypeApplyBorders = wanted
        Case "AutoFormatAsYouTypeApplyBulletedLists": If Options.AutoFormatAsYouTypeApplyBulletedLists <> wanted Then Options.AutoFormatAsYouTypeApplyBulletedLists = wanted
        Case "AutoFormatAsYouTypeApplyHeadings": If Options.AutoFormatAsYouTypeApplyHeadings <> wanted Then Options.AutoFormatAsYouTypeApplyHeadings = wanted
        Case "AutoFormatAsYouTypeApplyNumberedLists": If Options.AutoFormatAsYouTypeApplyNumberedLists <> wanted Then Options.AutoFormatAsYouTypeApplyNumberedLists = wanted
        Case "AutoFormatAsYouTypeApplyTables": If Options.AutoFormatAsYouTypeApplyTables <> wanted Then Options.AutoFormatAsYouTypeApplyTables = wanted
        Case "AutoFormatAsYouTypeDefineStyles": If Options.AutoFormatAsYouTypeDefineStyles <> wanted Then Options.AutoFormatAsYouTypeDefineStyles = wanted
        Case "AutoFormatAsYouTypeFormatListItemBeginning": If Options.AutoFormatAsYouTypeFormatListItemBeginning <> wanted Then Options.AutoFormatAsYouTypeFormatListItemBeginning = wanted
        Case "AutoFormatAsYouTypeReplaceFractions": If Options.AutoFormatAsYouTypeReplaceFractions <> wanted Then Options.AutoFormatAsYouTypeReplaceFractions = wanted
        Case "AutoFormatAsYouTypeReplaceHyperlinks": If Options.AutoFormatAsYouTypeReplaceHyperlinks <> wanted Then Options.AutoFormatAsYouTypeReplaceHyperlinks = wanted
        Case "AutoFormatAsYouTypeReplaceOrdinals": If Options.AutoFormatAsYouTypeReplaceOrdinals <> wanted Then Options.AutoFormatAsYouTypeReplaceOrdinals = wanted
        Case "AutoFormatAsYouTypeReplacePlainTextEmphasis": If Options.AutoFormatAsYouTypeReplacePlainTextEmphasis <> wanted Then Options.AutoFormatAsYouTypeReplacePlainTextEmphasis = wanted
        Case "AutoFormatAsYouTypeReplaceQuotes": If Options.AutoFormatAsYouTypeReplaceQuotes <> wanted Then Options.AutoFormatAsYouTypeReplaceQuotes = wanted
        Case "AutoFormatAsYouTypeReplaceSymbols": If Options.AutoFormatAsYouTypeReplaceSymbols <> wanted Then Options.AutoFormatAsYouTypeReplaceSymbols = wanted
        Case "AutoFormatPlainTextWordMail": If Options.AutoFormatPlainTextWordMail <> wanted Then Options.AutoFormatPlainTextWordMail = wanted
        Case "AutoFormatPreserveStyles": If Options.AutoFormatPreserveStyles <> wanted Then Options.AutoFormatPreserveStyles = wanted
        Case "AutoFormatReplaceFractions": If Options.AutoFormatReplaceFractions <> wanted Then Options.AutoFormatReplaceFractions = wanted
        Case "AutoFormatReplaceHyperlinks": If Options.AutoFormatReplaceHyperlinks <> wanted Then Options.AutoFormatReplaceHyperlinks = wanted
        Case "AutoFormatReplaceOrdinals": If Options.AutoFormatReplaceOrdinals <> wanted Then Options.AutoFormatReplaceOrdinals = wanted
        Case "AutoFormatReplaceQuotes": If Options.AutoFormatReplaceQuotes <> wanted Then Options.AutoFormatReplaceQuotes = wanted
        Case "AutoFormatReplaceSymbols": If Options.AutoFormatReplaceSymbols <> wanted Then Options.AutoFormatReplaceSymbols = wanted
        Case "CheckGrammarAsYouType": If Options.CheckGrammarAsYouType <> wanted Then Options.CheckGrammarAsYouType = wanted
        Case "ContextualSpeller": If Options.ContextualSpeller <> wanted Then Options.ContextualSpeller = wanted
        Case "CorrectSentenceCaps": If AutoCorrect.CorrectSentenceCaps <> wanted Then AutoCorrect.CorrectSentenceCaps = wanted
        Case "CorrectTableCells": If AutoCorrect.CorrectTableCells <> wanted Then AutoCorrect.CorrectTableCells = wanted
        Case "IgnoreMixedDigits": If Options.IgnoreMixedDigits <> wanted Then Options.IgnoreMixedDigits = wanted
        Case "IgnoreUppercase": If Options.IgnoreUppercase <> wanted Then Options.IgnoreUppercase = wanted
        Case "LabelSmartTags": If Options.LabelSmartTags <> wanted Then Options.LabelSmartTags = wanted
        Case "RestrictLinkedStyles": If Application.RestrictLinkedStyles <> wanted Then Application.RestrictLinkedStyles = wanted
        Case "ShowStylePreviews": If Application.ShowStylePreviews <> wanted Then Application.ShowStylePreviews = wanted
        Case "TabIndentKey": If Options.TabIndentKey <> wanted Then Options.TabIndentKey = wanted
    End Select
    Err.Clear
End Sub  '*** end of Sh_Setting_Put ***

Sub Sh_Save_Transcriber_Settings()
'
' Note the settings a book is about to take away, while they are still hers.
'
' Called from AutoExec, so the very first thing a session knows is what Word loaded from her own
' settings; from the top of the two book configurations, but ONLY when the ordinary configuration
' is genuinely in force, because called any other time the live values are ones a book already
' imposed and saving those would record VistaType's own setting as her choice; and from
' Sh_Restore_Transcriber_Settings on the path where no book has run since we were last in an
' ordinary document, which is how a change she makes in a LETTER becomes her preference.
'
' The test for that guard is on Sh_ConfiguredAs and not on who the caller is - about a dozen
' places run the configuration subs directly.
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version: 3.3  Date: 8/22/2026 - 34 settings: the two spelling-checker ones are gone, forced by all
'                                 three configurations from 3.0.222 rather than restored
' Version: 3.2  Date: 8/21/2026 - 34 settings: the three AutoFormat As You Type boxes large print
'                                 had never written (Jerry, 3.0.216)
' Version: 3.1  Date: 8/21/2026 - 31 settings: smart quotes and hyperlinks are book settings again
' Version: 3.0  Date: 8/20/2026 - walks Sh_Tracked_Settings: 27 settings instead of 6, and writes
'                                 the store's version stamp. Only writes a value that has actually
'                                 changed, because this now runs on every ordinary document open
' Version: 2.0  Date: 8/18/2026 - kept in the settings file instead of module variables, which
'                                 VBA's End statement wiped
' Version: 1.0  Date: 8/18/2026
'
    Dim names As Variant
    Dim i As Long
    Dim nm As String
    Dim live As String

    On Error Resume Next

    ' A book's values are standing in Word right now, so nothing here is hers. This is the guard
    ' that makes the ledger survive Word CLOSING: if she quits while working in a book, Word saves
    ' the book's settings as its own and loads them again next session - and AutoExec, which runs
    ' before anything can be configured, would otherwise write those straight into her
    ' preferences and lose the real ones for good. The book record outlives the session in the
    ' same file, which is what makes the question answerable at all.
    '
    ' It costs nothing on the paths that legitimately save: the two book configurations only call
    ' this while the ordinary configuration is in force, and Sh_Restore_Transcriber_Settings only
    ' calls it after establishing that no book record is outstanding.
    If Sh_Setting_Read(VT_STORE_BOOK, "Saved", "") = "1" Then
        Err.Clear
        Exit Sub
    End If

    names = Sh_Tracked_Settings()
    For i = LBound(names) To UBound(names)
        nm = names(i)
        live = Sh_Bool_To_Store(Sh_Setting_Live(nm))
        ' Written only when it differs. Thirty-one unconditional writes on every letter she
        ' opens is a great deal of file work for nothing.
        If Sh_Setting_Read(VT_STORE_MINE, nm, "") <> live Then
            Sh_Setting_Write VT_STORE_MINE, nm, live
        End If
    Next i

    Sh_Setting_Write VT_STORE_MINE, VT_STORE_STAMP, VT_STORE_STAMP_NOW
    ' Written LAST, so a half-finished save is never mistaken for a complete one.
    Sh_Setting_Write VT_STORE_MINE, "Saved", "1"

    Err.Clear

End Sub  '*** end of Sh_Save_Transcriber_Settings ***

' Record what a book configuration just applied. Called at the END of
' MS_Set_Word_Config_For_Large_Print and MS_Set_Word_Config_For_Braille, after every write in them
' has run - which is deliberate and is the answer to a trap.
'
' Every write in those two subs is guarded with If <>, so a setting that already held the book's
' value is never written at all. Recording what the sub INTENDED, write by write, would therefore
' have to know about writes that did not happen. Reading the live values once at the end sidesteps
' it completely: by then the live value IS what the book asked for, whether it was written or was
' already there.
'
' Running at the end also means a configuration that raises part way through records nothing, so
' the next return to an ordinary document reads no book record and treats the live values as hers.
' That is the safe way round: it can lose a restore, never her preferences.
'
' "The end" means the END, and getting that wrong was the first fault this piece had. Both book
' configurations write a handful of spelling and grammar settings BELOW the line that sets
' Sh_ConfiguredAs, deliberately, so that a sub which raises part way through does not leave them
' switched off while the ledger still believes the ordinary configuration is in force. The call
' first went in above those writes, which recorded her values as the book's for two settings in
' large print and six in braille. The next ordinary document then read live against the record,
' found them different, and wrote the BOOK's switched-off values into her preferences as though
' she had chosen them - the 8/18/2026 fault exactly, rebuilt by the thing meant to prevent it.
'
' Version: 1.1  Date: 8/20/2026 - moved below the trailing spelling and grammar writes in both
'                                 book configurations. See above; found before 3.0.214 was tested
' Version: 1.0  Date: 8/20/2026
Sub Sh_Note_Book_Settings()
    Dim names As Variant
    Dim i As Long
    Dim nm As String

    On Error Resume Next

    names = Sh_Tracked_Settings()
    For i = LBound(names) To UBound(names)
        nm = names(i)
        Sh_Setting_Write VT_STORE_BOOK, nm, Sh_Bool_To_Store(Sh_Setting_Live(nm))
    Next i

    ' Written LAST, and it is what Sh_Restore_Transcriber_Settings tests. A half-written record
    ' is never read as a whole one.
    Sh_Setting_Write VT_STORE_BOOK, "Saved", "1"

    Err.Clear

End Sub  '*** end of Sh_Note_Book_Settings ***

Sub Sh_Restore_Transcriber_Settings()
'
' Called by the ordinary configuration only - arriving at an ordinary document is the moment her
' settings become hers again. Despite the name it does two jobs, and the first one is the reason
' this piece exists at all.
'
' LEARNING WHAT SHE CHANGED. Jerry, 8/20/2026: she does far more in Word than large print and
' braille, and she does not stop being an ordinary user when she opens a book - so a setting she
' changes while working IN a book is a real preference and has to follow her out to her letters.
' Word gives us no way to be told about it, and live values alone cannot say whether a setting is
' switched off because she switched it off or because large print did. So the book configurations
' write down what they applied (Sh_Note_Book_Settings) and this compares against that record:
' anything that no longer matches, she changed, and it becomes hers.
'
' No heuristics, and it degrades the safe way. A setting no book writes is not in the list at all
' and is already hers; a book run that raised part way through wrote no record, and then the live
' values are taken as hers, which can lose a restore but can never lose a preference.
'
' The three paths, in the order they are tested:
'
'   The store was written by an older build - a different shape, and half-reading it would put
'   values in the wrong places. Throw it away, note what is in force now, and change nothing.
'
'   No book has been in force since we were last in an ordinary document. Then the live values
'   ARE hers - this covers her changing something in one letter and opening another, which the
'   old sub got wrong by restoring over the top of it. Note them, restore nothing.
'
'   A book has been in force. Learn, restore, and clear the record so the next ordinary document
'   does not compare against a book that is long gone.
'
' Guarded writes throughout, like everything else in that sub: her value is usually already in
' force and Word does not need telling twice.
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Version: 3.3  Date: 8/22/2026 - 34 settings: the two spelling-checker ones are gone, forced by all
'                                 three configurations from 3.0.222 rather than restored
' Version: 3.2  Date: 8/21/2026 - 34 settings: the three AutoFormat As You Type boxes large print
'                                 had never written (Jerry, 3.0.216)
' Version: 3.1  Date: 8/21/2026 - 31 settings: smart quotes and hyperlinks are book settings again
' Version: 3.0  Date: 8/20/2026 - learns what she changed inside a book, and covers all 27 tracked
'                                 settings rather than 6. Piece 4 of the automatic-configuration plan
' Version: 2.0  Date: 8/18/2026 - reads the settings file rather than module variables
' Version: 1.0  Date: 8/18/2026
'
    Dim names As Variant
    Dim i As Long
    Dim nm As String
    Dim held As String
    Dim bookHeld As String
    Dim live As String

    On Error Resume Next

    ' A store from an older build. Do not half-read it. The book half is thrown away FIRST -
    ' it is of the old shape too, and leaving it standing would make Sh_Save_Transcriber_Settings
    ' decline (see its guard), so no new stamp would ever be written and every ordinary document
    ' from then on would take this same path.
    If Sh_Setting_Read(VT_STORE_MINE, VT_STORE_STAMP, "") <> VT_STORE_STAMP_NOW Then
        Sh_Setting_Write VT_STORE_BOOK, "Saved", "0"
        Sh_Save_Transcriber_Settings
        Err.Clear
        Exit Sub
    End If

    ' Nothing has taken anything away since we were last here, so there is nothing to give back
    ' and what is in force is hers - including whatever she has just changed in a letter.
    If Sh_Setting_Read(VT_STORE_BOOK, "Saved", "") <> "1" Then
        Sh_Save_Transcriber_Settings
        Err.Clear
        Exit Sub
    End If

    names = Sh_Tracked_Settings()

    ' 1. Learn. Anything that no longer matches what the book wrote, she changed while working.
    For i = LBound(names) To UBound(names)
        nm = names(i)
        bookHeld = Sh_Setting_Read(VT_STORE_BOOK, nm, "")
        If bookHeld <> "" Then
            live = Sh_Bool_To_Store(Sh_Setting_Live(nm))
            If live <> bookHeld Then Sh_Setting_Write VT_STORE_MINE, nm, live
        End If
    Next i

    ' 2. Give them back. A setting with nothing stored is left alone rather than guessed at.
    For i = LBound(names) To UBound(names)
        nm = names(i)
        held = Sh_Setting_Read(VT_STORE_MINE, nm, "")
        If held <> "" Then Sh_Setting_Put nm, Sh_Store_To_Bool(held)
    Next i

    ' 3. The book record has been spent. Leave it marked as such - Sh_Setting_Write with an empty
    ' string DELETES a key rather than emptying it, so "0" is written instead.
    Sh_Setting_Write VT_STORE_BOOK, "Saved", "0"

    Err.Clear

End Sub  '*** end of Sh_Restore_Transcriber_Settings ***

Sub MS_Clear_F_and_R_Params_and_Clipboard()
'
' Version: 1.2  Date: 8/1/2025 - added clear clipboard
' Version: 1.1  Date: 1/8/2020

    ' Clear R&R Params
    On Error Resume Next
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    ' Clear clipboard
    Application.Run MacroName:="MS_SafeClearClipboard"
    On Error GoTo 0

End Sub  '*** end of MS_Clear_F_and_R_Params_and_Clipboard macro ***

Sub MS_SafeClearClipboard()
'
' Version: 1.0  Date: 8/1/2025

    Dim clip As MSForms.DataObject
    Dim attempt As Integer
    
    Set clip = New MSForms.DataObject
    clip.SetText ""     ' clear text
    
    For attempt = 1 To 5
        On Error Resume Next
        clip.PutInClipboard
        If Err.Number = 0 Then
            Exit For
        End If
        Err.Clear
        DoEvents ' let the OS process pending messages
        Sh_SleepForSeconds 1 ' pause 1 second before retry
    Next attempt
    On Error GoTo 0
    
End Sub   '*** end of MS_SafeClearClipboard ***

'------------------------------------------------------------------------------------
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
' ** start of Shared macros  ***
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
'-----------------------------------------------------------------------------------

Sub Sh_Is_Doc_Open()
'
' Check to see if a document is open

' Version: 1.1 Date: 2/14/2023
' Version: 1.0 Date: 7/31/2016
'
    If Application.Documents.count = 0 Then
        MsgBox "Cannot continue... no document is open."
        End
    End If
    
End Sub  '*** end of Sh_Is_Doc_Open macro ***

Function Sh_Software_Agreement_Text() As String
'
' The Software Agreement shown in BOTH About dialogs, Lp_About_Title_And_Agreement and
' Dx_About_Title_And_Agreement. One copy, here, so the two can never drift apart and so the
' wording lives in git-tracked text rather than inside a form's binary .frx.
'
' Replaces the old permissive agreement, which granted use "at no cost to others" and read
' like an MIT license. It never matched what this software actually ships under. The canonical
' wording is docs/Software-Agreement.md; keep the two in step.
'
' This is a plain-English summary, not the license itself. The full GNU General Public License
' is what the installer displays and what it writes to %AppData%\VistaType LP\LICENSE.txt --
' the "View Full License" button on each dialog opens that copy. See Sh_Show_Full_License.
'
' Version: 1.0  Date: 8/2/2026
'
    Dim s As String

    s = "Software License" & vbCrLf & vbCrLf

    s = s & "VistaType LP and Braille Macros is free software: you may run, copy, study, " _
          & "share, and modify it under the terms of the GNU General Public License, " _
          & "version 3 (GPLv3), as published by the Free Software Foundation." _
          & vbCrLf & vbCrLf

    s = s & "Your freedoms. You may use the Software for any purpose, examine how it works, " _
          & "and redistribute copies. You may modify it and distribute your modified " _
          & "versions - but those versions must also be licensed under the GPLv3, and you " _
          & "must make their source code available. This notice and the license must be " _
          & "included with all copies." & vbCrLf & vbCrLf

    s = s & "No warranty. This program is distributed in the hope that it will be useful, " _
          & "but WITHOUT ANY WARRANTY - without even the implied warranty of MERCHANTABILITY " _
          & "or FITNESS FOR A PARTICULAR PURPOSE. The author makes no guarantees regarding " _
          & "its performance, reliability, or suitability for any task, and is not liable " _
          & "for any loss or damage arising from its use." & vbCrLf & vbCrLf

    s = s & "Full terms. A complete copy of the GNU General Public License is installed with " _
          & "the Software. Click View Full License below to read it, or see " _
          & "https://www.gnu.org/licenses/." & vbCrLf & vbCrLf

    s = s & "By installing, running, or otherwise using this Software you acknowledge these " _
          & "terms." & vbCrLf & vbCrLf

    s = s & ChrW(169) & " 2015-2026 Jerry Whittaker  -  jerry@thewhittakers.org"

    Sh_Software_Agreement_Text = s

End Function   '*** end of Sh_Software_Agreement_Text function ***

Sub Sh_Show_Full_License()
'
' Opens the full GNU General Public License -- the copy the installer writes to
' %AppData%\VistaType LP\LICENSE.txt so the user "receives a copy of the license" as the GPL
' requires -- as a READ-ONLY Word document.
'
' Shared by the "View Full License" button on both About dialogs.
'
' Why a Word document and not a box on a form. Jerry, 8/3/2026: an MSForms text box does not
' respond to the mouse wheel. The control has no wheel handling and there is no property to
' switch on; the only way to add it is a Windows mouse hook, and a VBA callback from a system
' hook is a well known way to crash Word. Reading 674 lines by PageDown is not reasonable in a
' LARGE PRINT product. Opening it in Word gives the wheel, Ctrl+scroll zoom, Find, and screen
' reader support, in the application the transcriber is already in. Notepad, the first attempt,
' gave small fixed type and none of that.
'
' Version: 3.0  Date: 8/3/2026 - opens read-only in Word (was: a scrollable box on a form,
'                               which could not be scrolled with the mouse wheel)
' Version: 2.0  Date: 8/3/2026 - (superseded) shown in Sh_License_Form
' Version: 1.0  Date: 8/2/2026
'
    Dim licensePath As String
    Dim cc_Prev As Boolean

    licensePath = Environ$("APPDATA") & "\VistaType LP\LICENSE.txt"

    On Error GoTo NoLicense
    If Dir$(licensePath) = "" Then GoTo NoLicense

    ' ConfirmConversions off across the open: Word otherwise stops on its "Convert File"
    ' dialog for a .txt, which is a needless question to put in front of the user.
    cc_Prev = Application.Options.ConfirmConversions
    Application.Options.ConfirmConversions = False

    ' Tell the add-in's own document-open handler to leave this one alone. Without it,
    ' Sh_HandleDocumentOpened treats the license as an ordinary document and runs
    ' MS_Set_Word_Config_For_New_Install over it -- which would reset the transcriber's Styles
    ' pane just because they clicked a button to read the license.
    Sh_Skip_Open_Handler = True

    Documents.Open FileName:=licensePath, ReadOnly:=True, AddToRecentFiles:=False, Visible:=True

    Sh_Skip_Open_Handler = False
    Application.Options.ConfirmConversions = cc_Prev
    Exit Sub

NoLicense:
    Sh_Skip_Open_Handler = False
    On Error Resume Next
    Application.Options.ConfirmConversions = cc_Prev
    On Error GoTo 0
    MsgBox "The full license file could not be opened." & vbCr & vbCr _
         & "It is normally installed at:" & vbCr & licensePath & vbCr & vbCr _
         & "You can also read the GNU General Public License version 3 at " _
         & "https://www.gnu.org/licenses/.", vbInformation, "VistaType LP (231)"

End Sub   '*** end of Sh_Show_Full_License macro ***

Sub Sh_Replace_All_Until_Done()
'
' Runs the Find/Replace already set up on Selection.Find over and over until it has nothing
' left to change. The caller builds Selection.Find exactly as before; this only replaces the
' single Execute at the end of it.
'
' For the auto-tag patterns this is not a tidy-up, it is the fix. Those patterns are shaped
' "^013(a page number)^013" and a Word replace consumes BOTH paragraph marks. When two page
' numbers sit in consecutive paragraphs -- which is exactly what a blank print page produces --
' the mark AFTER the first number is the very mark the second number needs in FRONT of it, and
' it has already been eaten. One Execute therefore tags alternate numbers: given 15, 16 and 17
' on three lines it tags 15 and 17 and walks straight past 16. Jerry's sample, 8/2/2026, missed
' 14, 16, G3 and G5 for precisely this reason. Running the same replace again picks up what was
' skipped, because on a fresh Execute every mark is available again. Two or three rounds
' converge; the guard is only a backstop.
'
' Safe to repeat. Once a number is tagged its paragraph reads "$pg16", which cannot match a
' pattern requiring only digits, or only letters, between the two marks -- so nothing is ever
' tagged twice. Roman numerals never came here: they are tagged by a paragraph loop further up
' (Sh_IsValidRomanNumeral), which is why they were the ones that came out right.
'
' Shared: both Lp_AutoTag_Page_Numbers and Dx_AutoTag_Page_Numbers use it, 22 passes in all.
'
' Version: 1.0  Date: 8/2/2026
'
    Dim guard As Long

    Do
        guard = guard + 1
    Loop While Selection.Find.Execute(Replace:=wdReplaceAll) And guard < 20

End Sub   '*** end of Sh_Replace_All_Until_Done macro ***

Sub Sh_Show_Recommended_Styles_Pane()
'
' Opens the Styles pane and puts it back to VistaType's working setup: sorted "As Recommended",
' with "Select styles to show" set to "Recommended".
'
' Lives on the Quick Access Toolbar. As of 8/2/2026 the macros no longer force these settings
' on the user -- only attaching the large print template does -- so this is the button that
' puts them back on demand. Shared rather than Lp_: one toolbar serves both the large print
' and the braille side, and braille uses the same pair.
'
' Sets ONLY the three things it advertises. It deliberately does not call
' Lp_Turn_on_Styles_Pane, which also writes Application.RestrictLinkedStyles -- a Word-wide
' setting that a toolbar button has no business changing behind the user's back.
'
' Note the split: sort order and the show-filter are DOCUMENT properties, saved into the file
' and carried with it. Whether the pane is open is a Word-wide setting.
'
' Version: 1.0  Date: 8/2/2026
'
    ' The toolbar is clickable with no document open, and ActiveDocument would raise 4248.
    Application.Run MacroName:="Sh_Is_Doc_Open"

    Application.TaskPanes(wdTaskPaneFormatting).Visible = True
    ActiveDocument.StyleSortMethod = wdStyleSortRecommended
    ActiveDocument.FormattingShowFilter = wdShowFilterFormattingRecommended

End Sub   '*** end of Sh_Show_Recommended_Styles_Pane macro ***

Sub Sh_RemoveHeadAndFoot()
'
' Sh_RemoveHeadAndFoot Macro
'
' From http://word.tips.net/T001777_Deleting_All_Headers_and_Footers.html
'
    Dim oSec As Section
    Dim oHead As HeaderFooter
    Dim oFoot As HeaderFooter

    For Each oSec In ActiveDocument.Sections
        For Each oHead In oSec.headers
            If oHead.Exists Then oHead.Range.Delete
        Next oHead

        For Each oFoot In oSec.Footers
            If oFoot.Exists Then oFoot.Range.Delete
        Next oFoot
    Next oSec
End Sub '*** End of Sh_RemoveHeadAndFoot Macro ***

Sub Sh_Para_Before_Dollar()
'
' Sh_Para_Before_Dollar macro
'
' Version 1.0 3/2/2016
'
' places a para mark before each $pg then removes
'   extra para marks
' for global cleanup for both braille and LP
' DAISY files often do not have page number is in
'  paragraphs of their own
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False ' Turn screen updating off
    
    ' find $pg and put a para mark before it
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "$pg"
        .Replacement.Text = "^p$pg"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' if there are two para marks before $pg then
    ' replace it with a single para mark
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^013{2,}$pg"
        .Replacement.Text = "^p$pg"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = wdColorRed
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorAutomatic
    With Selection.Find
        .Text = "^013"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ActiveDocument.UndoClear
    Application.Run MacroName:="MS_Clear_F_and_R_Params_and_Clipboard"
    Application.ScreenUpdating = su_Prev ' Turn screen updating on
    
End Sub '*** end of Sh_Para_Before_Dollar macro ***

' The section whose page setup should be reported: the one the cursor is in, falling back to the
' first. Everything that reads a page measurement for display goes through this, because
' ActiveDocument.PageSetup returns wdUndefined (9999999) as soon as two sections disagree - and a
' large-print book with one landscape page for a table disagrees with itself by design.
'
' Version: 1.0  Date: 8/9/2026
Public Function Sh_Target_Section() As Section
    On Error GoTo firstSection
    If Not Selection Is Nothing Then
        If Not Selection.Range Is Nothing Then
            If Selection.Range.Sections.count > 0 Then
                Set Sh_Target_Section = Selection.Range.Sections(1)
                Exit Function
            End If
        End If
    End If
firstSection:
    Set Sh_Target_Section = ActiveDocument.Sections(1)
End Function

Sub Sh_Doc_Info()
'
' Sh_Doc_Info macro
'
' Shows the settings and path info of the current Document
'
' Version: 2.0  Date: 8/9/2026 - page measurements are read from ONE SECTION, not from the
'                               document. ActiveDocument.PageSetup returns wdUndefined - the
'                               literal 9999999 - the moment two sections disagree, and divided
'                               by 72 that printed 138888.88 where an inch measurement belonged.
'                               A single landscape page for a wide table is enough, which is why
'                               it struck only some books. New Sh_Target_Section picks the
'                               section the cursor is in, and the box now says so when the rest
'                               of the document differs (Jerry, 8/9/2026)
' Version: 1.9  Date: 8/9/2026 - reconciles before it reports. This box states "Word is configured
'                               for ...", and that sentence used to be whatever the last configuration
'                               macro happened to set - so after working on a braille file it went on
'                               saying braille in a plain document (Jerry, 8/9/2026). It now asks the
'                               document in front of the transcriber what it wants and puts that in
'                               force first, so the sentence is true by the time it is shown. That
'                               also makes the box right on any machine where Word never raises the
'                               window events at all, which is not hypothetical: they could not be
'                               made to fire once on the build box
' Version: 1.8  Date: 8/8/2026 - reports the typeface three ways: what the Normal style says, what
'                               was chosen when the template was attached, and whether that font is
'                               actually installed on this machine. The last one answers "my
'                               document went small" in one look
' Version: 1.7  Date: 11/3/2025 - added Get Gutter Rounded As String
' Version: 1.6  Date: 3/5/2024 - Fixed EBAN display
' Version: 1.5  Date: 1/19/2024 - bug fixes
' Version: 1.4  Date: 1/4/2024 - changed BANA display to have BANA as the First word
' Version: 1.2  Date: 11/13/2021 - bypassed crash when older braille file has no "BrailleType"
' Version: 1.1  Date: 2/15/2021 - minor revisions
' Version: 1.0  Date: 2/4/2021 - Full Rewrite

    Application.Run MacroName:="Sh_Is_Doc_Open"

    ' Make the "Word is configured for ..." line below TRUE before printing it - see 1.9 above.
    ' Costs nothing when it already agrees, which is the normal case.
    Sh_HandleDocumentActivated

    Dim AttachedTemplate As String
    Dim BrlType As String
    Dim MixedSetup As Boolean
    Dim targetSection As Section
    Dim ps As PageSetup
    Dim SectionNote As String

    AttachedTemplate = ActiveDocument.AttachedTemplate

    ' Read the page setup from ONE section, not from the document.
    '
    ' ActiveDocument.PageSetup answers for the document as a whole, and when its sections do not
    ' agree Word does not average them or pick one - it returns wdUndefined, which is the literal
    ' number 9999999. Divided by 72 that printed "138888.88" where an inch measurement belonged
    ' (Jerry, 8/9/2026: "a very large number rather than a simple decimal"). One landscape page
    ' for a wide table is enough to do it, which is why it only happened to some books.
    '
    ' The gutter code further down already worked this way; now everything does. targetSection is
    ' the section the cursor is in, so the figures describe the page the transcriber is looking
    ' at, and MixedSetup below says so when the rest of the book differs.
    Set targetSection = Sh_Target_Section()
    Set ps = targetSection.PageSetup

    MixedSetup = (ActiveDocument.PageSetup.PageHeight = wdUndefined) _
              Or (ActiveDocument.PageSetup.PageWidth = wdUndefined) _
              Or (ActiveDocument.PageSetup.TopMargin = wdUndefined) _
              Or (ActiveDocument.PageSetup.LeftMargin = wdUndefined)

    If MixedSetup Then
        SectionNote = vbCr & vbCr _
            & "NOTE: this document's sections do not all use the same page setup." & vbCr _
            & "The measurements above are for section " & targetSection.Index _
            & ", the one the cursor is in."
    End If

    PTM = Str(Round(ps.TopMargin / Application.InchesToPoints(1), 2))
    PBM = Str(Round(ps.BottomMargin / Application.InchesToPoints(1), 2))
    PLM = Str(Round(ps.LeftMargin / Application.InchesToPoints(1), 2))
    PRM = Str(Round(ps.RightMargin / Application.InchesToPoints(1), 2))
    PPH = Str(Round(PointsToInches(ps.PageHeight), 2))
    PPW = Str(Round(PointsToInches(ps.PageWidth), 2))
    PMM = ps.MirrorMargins  ' Zero = not mirrored
    Lp_Base_Font_Size = ActiveDocument.Styles(wdStyleNormal).Font.Size
    Lp_Base_Font_Name = ActiveDocument.Styles(wdStyleNormal).Font.Name
    If Len(Trim(Lp_Base_Font_Name)) = 0 Then Lp_Base_Font_Name = LP_FONT_TAHOMA
    TOCTabSetting = Str(Val(PPW) - (Val(PLM) + Val(PRM)))

    If ps.Orientation = 1 Then 'Landscape
        PPO = "L"
    Else
        PPO = "P"
    End If
    
    If PMM = 0 Then
        MirrorString = "No"
        PPG = "0"
    Else
        MirrorString = "Yes"
        ' get gutter size  - no way to get this from active document - is stored in the document xml
        Sh_GP_String_1 = ""
        Sh_Read_Document_Variables "GutterSize", "VarValue"
        PPG = Sh_GP_String_1
        Sh_GP_String_1 = ""
    End If
    
    ' get media type  - no way to get this from active document - is stored in the document/settings xml
    Sh_Read_Document_Variables "Media", "VarValue" 'places "VarValue" into Sh_GP_String_1
    If Sh_GP_String_1 = "" Then '  this document has no media entry
        DM = "Unknown"
    Else
        DM = Sh_GP_String_1
    End If
    Sh_GP_String_1 = ""

    ' The typeface the transcriber ASKED for when the template was attached. Lp_Base_Font_Name
    ' above is what the Normal style says TODAY, and the two can drift apart - "update style to
    ' match selection" on a Normal paragraph is enough to do it. Showing both makes that visible
    ' instead of leaving it to be guessed at.
    Dim IntendedFontName As String
    ' Blank it FIRST, in this block rather than relying on the one above. When the document has
    ' no such variable, Sh_Read_Document_Variables jumps its error label and leaves
    ' Sh_GP_String_1 exactly as it found it - so an unblanked call reports the PREVIOUS
    ' variable's value, which here would mean a Tahoma book claiming the media type as its
    ' typeface.
    Sh_GP_String_1 = ""
    Sh_Read_Document_Variables "BaseFontName", "VarValue"
    IntendedFontName = Sh_GP_String_1
    Sh_GP_String_1 = ""
    If IntendedFontName = "" Then
        ' Deliberately no version number: the third digit is a private build counter that never
        ' ships, so quoting one at a transcriber names something they cannot check.
        IntendedFontName = "not recorded (attached by an earlier version)"
    End If

    If PPO = "L" Then
        Sh_GP_String_1 = "Landscape"
    Else
        Sh_GP_String_1 = "Portrait"
    End If

    ' *** begin Get Gutter Rounded As String
    ' the PPG (print page gutter) is a calculated value and does not appear correctly in the doc info files which have been exported
    Dim gutterPts As Double
    Dim gutterInches As Double
    Dim gutterRounded As Double

    ' targetSection and ps were chosen at the top, with everything else - see the note there.

    ' Gutter is returned in points (72 points = 1 inch)
    gutterPts = ps.Gutter
    gutterInches = gutterPts / 72#

    ' Round to two decimal places
    gutterRounded = Round(gutterInches, 2)
    PPG = Str(gutterRounded)
    '*** end Get Gutter Rounded AsString ***
    
    If ActiveDocument.AttachedTemplate = "LargePrintTemplate.dotx" Then
        MsgBox "Attached Template = " & ActiveDocument.AttachedTemplate & vbCr & vbCr _
                    & MS_Word_Config & vbCr & vbCr _
                    & " Normal Style Font Size      = " + Trim(Lp_Base_Font_Size) & vbCr _
                    & " Typeface (Normal style)     = " + Trim(Lp_Base_Font_Name) & vbCr _
                    & " Typeface recorded at attach = " + IntendedFontName & vbCr _
                    & " Typeface installed here     = " + Sh_Font_Status_Text(Lp_Base_Font_Name, ActiveDocument) & vbCr _
                    & " Paper/Screen Height         = " + PPH & vbCr _
                    & " Paper/ScreenWidth           = " + PPW & vbCr _
                    & " Top Margin                       = " + PTM & vbCr _
                    & " Bottom Margin                 = " + PBM & vbCr _
                    & " Left Margin                       = " + PLM & vbCr _
                    & " Right Margin                     = " + PRM & vbCr _
                    & " Mirrored Margins              = " + MirrorString & vbCr _
                    & " Binding (Gutter) Width       = " + PPG & vbCr _
                    & " Orientation                        = " + Sh_GP_String_1 & vbCr _
                    & " Output Media Type           = " + DM & SectionNote, , "Document Settings"
                    
    ElseIf InStr(UCase(ActiveDocument.AttachedTemplate), "BRAILLE") > 0 Then

        If InStr(UCase(ActiveDocument.AttachedTemplate), "BRAILLE") > 0 Then
                On Error GoTo Unknown
                If ActiveDocument.Variables("BrailleType") = "EBAT" Then
                    BrlType = "Macros will format this document for BANA EBAE translation"
                ElseIf ActiveDocument.Variables("BrailleType") = "EBAN" Then
                    BrlType = "Macros will format this document for BANA EBAE Nemeth translation"
                ElseIf ActiveDocument.Variables("BrailleType") = "UEBT" Then
                    BrlType = "Macros will format this document for BANA UEB translation"
                ElseIf ActiveDocument.Variables("BrailleType") = "UEBN" Then
                    BrlType = "Macros will format this document for BANA UEB Nemeth translation"
                ElseIf ActiveDocument.Variables("BrailleType") = "Undefined" Then
Unknown:
                    BrlType = "Document UEB or EBAE translation settings are undefined"
                End If
            Else
                BrlType = ""
            End If

            MsgBox "Attached Template = " & ActiveDocument.AttachedTemplate & vbCr & vbCr _
                        & MS_Word_Config & vbCr & vbCr _
                        & " Orientation                        = " + Sh_GP_String_1 & vbCr _
                        & " Paper/Screen Height        = " + PPH & vbCr _
                        & " Paper/ScreenWidth          = " + PPW & vbCr _
                        & " Top Margin                        = " + PTM & vbCr _
                        & " Bottom Margin                  = " + PBM & vbCr _
                        & " Left Margin                        = " + PLM & vbCr _
                        & " Right Margin                     = " + PRM & vbCr _
                        & vbCr & BrlType & vbCr & SectionNote, , "Document Settings"
        Else
            MsgBox "Attached Template = " & ActiveDocument.AttachedTemplate & vbCr & vbCr _
                    & MS_Word_Config & vbCr & vbCr _
                    & " Orientation                        = " + Sh_GP_String_1 & vbCr _
                    & " Paper/Screen Height        = " + PPH & vbCr _
                    & " Paper/ScreenWidth          = " + PPW & vbCr _
                    & " Top Margin                        = " + PTM & vbCr _
                    & " Bottom Margin                  = " + PBM & vbCr _
                    & " Left Margin                        = " + PLM & vbCr _
                    & " Right Margin                     = " + PRM & SectionNote, , "Document Settings"
    End If
               
    Sh_GP_String_1 = ""

End Sub  '*** end of Sh_Doc_Info macro ***

Sub Sh_Text_Frame_Warning_To_Red(Optional ByVal target As Range)
'
' Sh_Text_Frame_Warning_To_Red Macro
'
' Author: Jerry Whittaker -  jerry@thewhittakers.org
'
' Colors the markers Remove_Txt_Bxs_And_Frames leaves behind, so rescued text is easy to find.
'
' target  the range to color. Omit it and the whole document is colored, which is what this
'         macro did before it took an argument.
'
' Version: 2.0  Date: 8/13/2026 - takes a range, and works through that range instead of through
'                                 Selection.Find. Its two callers pass their own scope; on the
'                                 old Selection.Find with wdFindContinue, "just the selection"
'                                 was never available to them.
' Version: 1.2  Date: 11/16/2016
'
    Dim rng As Range
    Dim marker As Variant

    If target Is Nothing Then
        Set rng = ActiveDocument.Content
    Else
        Set rng = target
    End If

    For Each marker In Array(SH_TBX_BELOW, SH_TBX_ABOVE)
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Replacement.Font.Color = wdColorRed
            .Text = CStr(marker)
            .Replacement.Text = "^&"
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next marker

End Sub   '*** end of Sh_Text_Frame_Warning_To_Red ***

Sub Sh_Remove_Spaces_Before_Punctuation()
'
' Sh_Remove_Spaces_Before_Punctuation Macro
'
' Version 1.1 Date: 9/21/2019  added additional punctuation
' Version 1.0
'
' Author: Jerry Whittaker jerry@thewhittakers.org
'
' Date 11/2/16
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}\?"
        .Replacement.Text = "?"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}^046"
        .Replacement.Text = "^046"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}\!"
        .Replacement.Text = "!"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,};"
        .Replacement.Text = ";"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}:"
        .Replacement.Text = ":"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8220) & "^032{1,}"
        .Replacement.Text = ChrW(8220)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}" & ChrW(8221)
        .Replacement.Text = ChrW(8221)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}`"
        .Replacement.Text = "`"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "`^032{1,}"
        .Replacement.Text = "`"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}’"
        .Replacement.Text = "’"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "‘^032{1,}"
        .Replacement.Text = "‘"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}’"
        .Replacement.Text = "’"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\[^032{1,}"
        .Replacement.Text = "["
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}\]"
        .Replacement.Text = "]"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}\}"
        .Replacement.Text = "}"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\{^032{1,}"
        .Replacement.Text = "{"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "‘^032{1,}"
        .Replacement.Text = "‘"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^032{1,}’"
        .Replacement.Text = "’"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
 
End Sub  '*** end of Sh_Remove_Spaces_Before_Punctuation macro ***

Sub Sh_Save_User_Position()
'
' Sh_Save_User_Position macro
'
' Records where the user's cursor is, so Sh_Return_User_To_Start_Position can put
' them back when the macro finishes. Pair the two: every call here needs one there.
'
' Records a character offset rather than a bookmark on purpose - see the notes on
' Sh_Start_Pos at the top of this module for why TempPlaceholder cannot do this job.
'
' Nested macros do nothing: the outermost caller owns the saved position, so a
' File_Fix_Sequence running twenty cleanups returns the user once, not twenty times.
'
' Version: 1.0  Date: 7/26/2026
'
' Author: Jerry Whittaker jerry@thewhittakers.org
'
    On Error Resume Next

    Sh_Pos_Depth = Sh_Pos_Depth + 1
    If Sh_Pos_Depth > 1 Then Exit Sub      ' already inside a macro that saved the position

    Sh_Start_Doc = ActiveDocument.Name
    Sh_Start_Pos = Selection.Start
    Sh_Pos_Saved = True

End Sub   '*** end of Sh_Save_User_Position macro ***

Sub Sh_Set_Return_Position(ByVal NewPos As Long)
'
' Sh_Set_Return_Position macro
'
' Overrides where Sh_Return_User_To_Start_Position will leave the user, for macros that
' should finish somewhere more useful than where the user began. The table-to-list converts
' use it to land on the blue "Note:" box that now heads the list, so the transcriber can read
' or edit that note straight away instead of hunting for it.
'
' Only meaningful between a Sh_Save_User_Position and its matching return; ignored otherwise,
' so a stray call cannot send the cursor anywhere on its own.
'
' Version: 1.0  Date: 7/26/2026
'
' Author: Jerry Whittaker jerry@thewhittakers.org
'
    On Error Resume Next

    If Not Sh_Pos_Saved Then Exit Sub      ' nothing pending - do not invent a destination

    Sh_Start_Doc = ActiveDocument.Name     ' the override belongs to whatever document is active now
    Sh_Start_Pos = NewPos

End Sub   '*** end of Sh_Set_Return_Position macro ***

Sub Sh_Return_User_To_Start_Position()
'
' Sh_Return_User_To_Start_Position macro
'
' Puts the cursor back where Sh_Save_User_Position found it and scrolls it into view.
'
' Screen updating goes back on BEFORE the cursor moves. Moving the cursor while the
' screen is frozen leaves the insertion point correct but the window still showing
' wherever the macro was last working - usually the top of the document, which reads
' to the user as "the macro threw me back to page 1".
'
' The position is clamped to the document length because cleanup macros delete text,
' so the document may now be shorter than it was. Landing close is the goal.
'
' Version: 1.1  Date: 7/26/2026 - added Application.ScreenRefresh; without it the insertion point moved but was never drawn, so the user saw no cursor
' Version: 1.0  Date: 7/26/2026
'
' Author: Jerry Whittaker jerry@thewhittakers.org
'
    Dim Return_Pos As Long

    On Error Resume Next

    Sh_Pos_Depth = Sh_Pos_Depth - 1
    If Sh_Pos_Depth > 0 Then Exit Sub      ' an outer macro is still running; it will return the user
    Sh_Pos_Depth = 0

    Application.ScreenUpdating = True      ' must be on before the cursor moves, or the view will not follow

    If Not Sh_Pos_Saved Then Exit Sub      ' nothing was saved - never jump to a leftover position
    Sh_Pos_Saved = False                   ' this position is spent; a second call must do nothing

    If ActiveDocument.Name <> Sh_Start_Doc Then Exit Sub   ' user's document is no longer the active one

    Return_Pos = Sh_Start_Pos
    If Return_Pos > ActiveDocument.Content.End - 1 Then Return_Pos = ActiveDocument.Content.End - 1
    If Return_Pos < 0 Then Return_Pos = 0

    ActiveDocument.Range(Return_Pos, Return_Pos).Select
    ActiveWindow.ScrollIntoView Selection.Range, True

    ' Repaint, or the insertion point is set correctly but never drawn and the user sees no
    ' cursor at all. The macros this replaced each ended with their own Application.ScreenRefresh
    ' for exactly this reason; centralising it here means no macro can forget it again.
    Application.ScreenRefresh

End Sub   '*** end of Sh_Return_User_To_Start_Position macro ***

Sub Sh_Is_End_Paragraph_Mark_Included()
'
' Sh_Is_End_Paragraph_Mark_Included Macro
'
' Validates existance of para mark at end of selected text
'
' Author: Jerry Whittaker - jerry@thewhittakers.org

'
' Version: 1.1  Date: 12/18/20203 - fixed bug when only para marks are included
' Version: 1.0  Date: 11/16/2016
'
    Sh_GP_Boolean_1 = False ' General Pupose boolean
    
    Dim Para_Counter As Integer
    Para_Counter = 0
    
    Dim Loop_Continue As Boolean
    Loop_Continue = True
    
    Selection.EndKey Unit:=wdStory   'go to end of document
    
    On Error GoTo Error1
    Do While Loop_Continue
        If Asc(WordBasic.[Selection$]()) = 13 Then
            Para_Counter = Para_Counter + 1
            Selection.MoveLeft Unit:=wdCharacter, count:=1 'move back one character
        Else
            Loop_Continue = False
        End If
    Loop

    If Para_Counter > 1 Then
        Sh_GP_Boolean_1 = True 'more than one para mark
        Sh_GP_Counter_1 = Para_Counter
    End If
Error1:
On Error GoTo 0
    If Sh_GP_Boolean_1 = False Then
        ActiveDocument.Close SaveChanges:=False
        Application.ScreenUpdating = True ' Turn screen updating on
        MsgBox "Selected text must include the ending paragraph mark."
        End
    End If
    
    Application.ScreenUpdating = True  'turn on screen
        
End Sub '*** end of Sh_Is_End_Paragraph_Mark_Included ***

Sub Sh_Remove_DollarPG_For_Retag()

    ' removes the $pg tag so that manual tagging does not add a second $pg
    '
    ' Application.Run MacroName:="Sh_Remove_DollarPG_For_Retag"
    '
    ' Author: Jerry Whittaker - jerry@thewhittakers.org

    Dim strTemp As String
 
    'select the current line of text not including the para mark at the end
    Selection.HomeKey Unit:=wdLine
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Style = ActiveDocument.Styles("Normal")
    Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend 'don't want the para mark
    
    strTemp = Selection.Text

    If Left(strTemp, 11) = "$pg[[*ii*]]" Then  ' for braille tag of lower case roman for EBAE
        Selection.HomeKey Unit:=wdLine
        Selection.MoveRight Unit:=wdCharacter, count:=11, Extend:=wdExtend
    ElseIf Left(strTemp, 3) = "$pg" Then  ' plain $pg for LP or braille
        Selection.HomeKey Unit:=wdLine
        Selection.MoveRight Unit:=wdCharacter, count:=3, Extend:=wdExtend
    Else
        GoTo MacroEnd
    End If
    
    Selection.Copy
    Selection.Delete Unit:=wdCharacter, count:=1
    
MacroEnd:
    
End Sub  '*** end of Sh_Remove_DollarPG_For_Retag macro ***

Sub Sh_Replace_White_Text_With_Automatic()
    '
    ' remove shading and change white text to automatic
    '
    ' Version: 1.0  Date: 10/23/2018
    '
    ' Replace white text with automatic
    Selection.Find.ClearFormatting
    Selection.Find.Font.Color = -603914241 'color=Background or white
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorAutomatic
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub   '*** end of Sh_Replace_White_Text_With_Automatic macro ***

Sub Sh_Sort_Ascending()
'
' Version: 1.0  Date: 9/27/2018
'
' Author: Jerry Whittaker  jerry@thewhittakers.org

    Selection.Sort ExcludeHeader:=False, FieldNumber:="Paragraphs", _
        SortFieldType:=wdSortFieldAlphanumeric, SortOrder:=wdSortOrderAscending, _
        FieldNumber2:="", SortFieldType2:=wdSortFieldAlphanumeric, SortOrder2:= _
        wdSortOrderAscending, FieldNumber3:="", SortFieldType3:= _
        wdSortFieldAlphanumeric, SortOrder3:=wdSortOrderAscending, Separator:= _
        wdSortSeparateByTabs, SortColumn:=False, CaseSensitive:=False, LanguageID _
        :=wdEnglishUS, SubFieldNumber:="Paragraphs", SubFieldNumber2:= _
        "Paragraphs", SubFieldNumber3:="Paragraphs"
        
End Sub   '*** end of Sh_Sort_Ascending Macro ***

 Function Sh_Write_Document_Variables(VarName, VarValue)

    On Error GoTo TheVariableDoesNotExist
    ActiveDocument.Variables(VarName).Delete
TheVariableDoesNotExist:
    ActiveDocument.Variables.Add Name:=VarName, Value:=VarValue
    On Error GoTo 0
    
    'To view this setting in the .docx file, rename the .docx to .zip, right click and select open
    ' select the folder "word" right click and select open
    ' select settings.xml and right click and select open
 
End Function

Function Sh_Read_Document_Variables(VarName As String, VarValue As String)

    On Error GoTo TheVariableDoesNotExist
    VarValue = ActiveDocument.Variables(VarName) ' returns the value of the VarName
    Sh_GP_String_1 = VarValue
TheVariableDoesNotExist:
    On Error GoTo 0
 
End Function

Function Sh_FileExists(filePath As String) As Boolean
    ' Does a File exist in a specified path
    
    Dim TestStr As String
    TestStr = ""
        On Error Resume Next
        TestStr = Dir(filePath)
        On Error GoTo 0
        If TestStr = "" Then
            Sh_FileExists = False
        Else
            Sh_FileExists = True
        End If
End Function

Sub Sh_Keep_Cursor_In_View()
'
' Sh_Keep_Cursor_In_View macro
'
' Scrolls the window to wherever the cursor already is. Does NOT move the cursor.
'
' The paragraph toggles below set PageBreakBefore / KeepTogether / KeepWithNext, any of which
' can push the paragraph onto the next page. The cursor travels with the paragraph, but the
' WINDOW stays where it was - so the user is left looking at the page the paragraph just left,
' with no cursor in sight (Jerry, 7/27/2026).
'
' Not the same job as Sh_Return_User_To_Start_Position, which MOVES the cursor back to a
' remembered spot. This one leaves the cursor alone and simply follows it.
'
' Version: 1.0  Date: 7/27/2026
'
' Author: Jerry Whittaker jerry@thewhittakers.org
'
    On Error Resume Next

    Application.ScreenUpdating = True
    ActiveWindow.ScrollIntoView Selection.Range, True
    Application.ScreenRefresh

End Sub   '*** end of Sh_Keep_Cursor_In_View macro ***

Sub Sh_Move_Paragraph_To_Next_Page()
'
' Version: 1.3  Date: 7/27/2026 - select via Selection.Paragraphs(1) instead of a character
'                                 offset through ActiveDocument.Range. NOTE: this was first
'                                 written up as a Word 2019 incompatibility - that was WRONG.
'                                 The laptop it "failed" on had an older build installed (Word
'                                 was open during the install, so it silently no-opped). No
'                                 version difference was ever demonstrated. The change stands
'                                 on its own: no offset arithmetic, and no On Error Resume Next
'                                 quietly swallowing a failed move.
' Version: 1.2  Date: 7/27/2026 - land at the START of the moved paragraph, not wherever in it
'                                 the user happened to be clicking (Jerry)
' Version: 1.1  Date: 7/27/2026 - follow the paragraph when it jumps to the next page
'
    Application.Run MacroName:="Sh_Is_Doc_Open"

    Selection.ParagraphFormat.PageBreakBefore = wdToggle

    'Put the user at the top of the paragraph they just moved. Work through the Selection's
    'own paragraph rather than reading a character offset and re-selecting through
    'ActiveDocument.Range(): fewer moving parts, and the old form hid failures behind
    'On Error Resume Next. Paragraphs(1) is the first when several are selected.
    'Note PageBreakBefore does not change character offsets at all - it is a formatting
    'property - so there is nothing to re-read after the toggle.
    Selection.Paragraphs(1).Range.Select
    Selection.Collapse Direction:=wdCollapseStart

    Sh_Keep_Cursor_In_View
End Sub

Sub Sh_Keep_Lines_Of_Para_Together()
'
' Version: 1.1  Date: 7/27/2026 - follow the paragraph if it reflows to the next page
'
    Application.Run MacroName:="Sh_Is_Doc_Open"
    Selection.ParagraphFormat.KeepTogether = wdToggle
    Sh_Keep_Cursor_In_View
End Sub

Sub Sh_Fix_Ref_Pages_Before_and_After_Tables()
'
' The reference page identification procedures (Lp_AutoTab_Page_Numbers and Dx_AutoTab_Page_Numbers)
'   will not identify ref pg numbers immediatly preceeding and immediatly following a table.
' This proceedure will place a temporary paragraph immediatly preceeding and immediatly following a table
'   so that the ref page numbers can to tagged.
' The temporary paragraphs are removed by Lp_AutoTag_Page_Numbers and Dx_AutoTab_Page_Numbers
'
' Version: 1.3 Date: 1/28/2026 - complete rewrite
'
    Dim tbl As Table
    Dim i As Long
    Dim doc As Document
    Set doc = ActiveDocument

    Dim su_Prev As Boolean
    su_Prev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ' We loop backwards to maintain document stability
    For i = doc.Tables.count To 1 Step -1
        Set tbl = doc.Tables(i)
        
        ' --- 1. HANDLE THE TOP (BEFORE) ---
        ' Using SplitTable is the only way to GUARANTEE
        ' a paragraph mark appears ABOVE the table if it's at the top.
        tbl.rows(1).Range.Select
        Selection.SplitTable
        
        ' --- 2. HANDLE THE BOTTOM (AFTER) ---
        ' We target the document's range at the end of the table
        ' and "push" a carriage return after it.
        doc.Range(tbl.Range.End, tbl.Range.End).InsertAfter vbCr
        
        ' --- 3. FIX "FLOATING" TABLES ---
        ' If a table is set to "Around" wrapping, paragraphs
        ' won't stay put. This sets it to "None" (Inline).
        tbl.rows.WrapAroundText = wdWrapNone
    Next i

    ' Move cursor back to the start
    doc.Range(0, 0).Select
    Application.ScreenUpdating = su_Prev
    
End Sub   '*** end of Sh_Fix_Ref_Pages_Before_and_After_Tables macro ***

Sub Sh_Remove_Empty_Para_Before_Tables()
'
' The reference page identification procedures (Lp_AutoTab_Page_Numbers and Dx_AutoTab_Page_Numbers)
'   identifies and tags ref pg numbers immediatly preceeding a table.
' However, these procedures leave an empty para mark before the table
' This procedure will remove the temporary paragraph immediatly preceeding the table
' This procedure is run immediatly after the ref pg tags have been formatted.
'
' Version: 1.2  Date: 3/2/2026 - full rewrite
' Version: 1.1  Date: 11/19/2023 - fix crash for merged cells
' Version: 1.0  Date: 5/1/2023
'
    Dim doc As Document
    Dim i As Long
    Dim tblRng As Range
    Dim prevPara As Paragraph
    Dim s As String
    
    Set doc = ActiveDocument
    
    'Work from last table to first so deletions don't disturb later ranges
    For i = doc.Tables.count To 1 Step -1
        Set tblRng = doc.Tables(i).Range
        
        'If the table is at the very start of the document, there's nothing to check
        If tblRng.start = 0 Then GoTo NextTable
        
        'Keep deleting blank paragraphs immediately before the table
        Do
            'Get the paragraph that contains the character immediately before the table
            Set prevPara = doc.Range(tblRng.start - 1, tblRng.start).Paragraphs(1)
            
            'Paragraph text typically includes a trailing vbCr; remove it and trim spaces/tabs
            s = Replace(prevPara.Range.Text, vbCr, vbNullString)
            s = Replace(s, Chr(7), vbNullString) 'cell marker safety; usually not needed here
            s = Trim$(s)
            
            'Stop if it's not blank
            If Len(s) > 0 Then Exit Do
            
            'Delete the entire paragraph (including its paragraph mark)
            prevPara.Range.Delete
            
            'tblRng.Start may have shifted; refresh table range
            Set tblRng = doc.Tables(i).Range
            
            If tblRng.start = 0 Then Exit Do
        Loop
        
NextTable:
    Next i

End Sub   '*** end of Sh_Remove_Empty_Para_Before_Tables macro ***

Sub Sh_Clear_Multi_Selection()
'
' Version: 1.0  Date: 8/5/2026
'
' Author: Jerry Whittaker - jerry@thewhittakers.org
'
' Collapses the selection to ONE insertion point at its start.
'
' The temp-file validation builds its list by driving Word's Find and Replace dialog with
' SendKeys and choosing Find In > Main Document - see TempFileButton_Click on
' Sh_Validation_Choices_Form. That is Word's Find All, and it leaves a DISCONTIGUOUS selection:
' every $pg paragraph in the document selected at once. The user cannot avoid it, because the
' Navigation pane refuses the job once there are more tags than it can hold.
'
' That selection SURVIVES the validation, and most Selection methods refuse to work across a
' multiple selection - the next macro to touch one dies with run-time error 4605. Jerry hit it
' running AutoTag Ref Pages straight after validating, 8/5/2026.
'
' Re-selecting a single collapsed range is the reliable cure. Selection.Collapse is not: it can
' raise 4605 on the very selection it is meant to fix.

    Dim keepAt As Long

    On Error Resume Next
    keepAt = Selection.Start
    If Err.Number <> 0 Then keepAt = 0
    Err.Clear
    On Error GoTo 0

    If keepAt < 0 Then keepAt = 0
    If keepAt > ActiveDocument.Content.End Then keepAt = 0

    ActiveDocument.Range(keepAt, keepAt).Select

End Sub   '*** end of Sh_Clear_Multi_Selection macro ***

Sub Sh_Copy_Ref_Pg_Tags_To_Temp_File()

' this macro copies the tagged $pg paragraphs to a temporary document
'
' Version: 1.2  Date: 11/10/2025 - added "DoEvents" before and after "Selection.Paste" to avoid crash when MathType MathPage.wll is corrupt
' Version: 1.1  Date: 2/18/2024 - code to set word configuration added
' Version: 1.3  Date: 7/27/2026 - hands off to the modeless validation helper instead of
'                                 telling the user to Alt+Tab (see ShNonModalMessage)
' Version: 1.0  Date: 2/16/2024

    Dim MsgBoxLabel As String
    Dim tmpDoc As Document
    Dim srcDoc As Document

    'The document being validated, captured before Documents.Add makes the list active.
    Set srcDoc = ActiveDocument

    If Dx_Is_The_Attached_Template_BANA_Braille Then
        MsgBoxLabel = "Braille Macros"
    Else
        MsgBoxLabel = "VistaType"
    End If
    
    Selection.Copy

    Set tmpDoc = Documents.Add
    
    With tmpDoc.PageSetup.TextColumns
        If MsgBoxLabel = "Braille Macros" Then
            Application.Run MacroName:="MS_Set_Word_Config_For_Braille"
            .SetCount NumColumns:=2 'Braille
       Else
            Application.Run MacroName:="MS_Set_Word_Config_For_Large_Print"
            MS_Word_Config = "Word is configured for large print"
            .SetCount NumColumns:=3 'Large Print
        End If
    End With

    ActiveWindow.ActivePane.View.Type = wdPrintView
    Application.TaskPanes(wdTaskPaneFormatting).Visible = False
    
    With Selection.PageSetup
        .TopMargin = InchesToPoints(0.5)
        .BottomMargin = InchesToPoints(0.5)
        .LeftMargin = InchesToPoints(0.5)
        .RightMargin = InchesToPoints(0.5)
    End With

    DoEvents
    Selection.Paste
    DoEvents

    Selection.WholeStory
    Selection.Font.Name = "Tahoma"
    Selection.Font.Size = 12
    Selection.HomeKey Unit:=wdStory

    'The Alt+Tab instructions that used to live here are exactly what the helper replaces: it
    'walks the user down the list, finds each tag in the document, and brings them back to
    'the next one. Sh_Valid_Ref_Pg_No_1_Form still carries the directions, on demand.
    Sh_PgVal_Start srcDoc, tmpDoc, MsgBoxLabel

End Sub   '*** end of Sh_Copy_Ref_Pg_Tags_To_Temp_File macro ***

Sub Sh_ReplaceNonBreakingSpacesWithNormalSpace()
    '
    ' Version: 3.0  Date: 8/12/2026 - the ONLY one of these left, and no temporary document.
    '                                There were three doing this job and they disagreed:
    '                                Dx_Replace_NonBreaking_Spaces went through the temp
    '                                document, and Dx_Replace_NonBreaking_Space_With_Space was a
    '                                bare whole-document replace with no scoping at all. Neither
    '                                protected a "Print Pg Num" bar. Both are gone and everything
    '                                runs through here (Jerry, 8/12/2026).
    '
    '                                Scope is the same rule as the rest of the cleanup suite:
    '                                the selection if there is one, the whole document if not.
    '                                Runs are collapsed afterwards on the same range, which is
    '                                what the two braille copies did by replacing ^s{1,} in one
    '                                go - it cannot be done that way here, because the slow path
    '                                below depends on replacing one character with one character.
    ' Version: 2.0  Date: 8/5/2026 - a non-breaking space in a paragraph styled "Print Pg Num" is left alone
    ' Version: 1.0  Date: 3/2/2026
    '
    ' The pink bar built by Lp_Format_Page_Numbers begins and ends with a NON-BREAKING space, and
    ' the color pass in that macro finds those ends by searching for ^s. Flattening them here
    ' took the bar apart. Find can select BY a style but has no way to exclude one, so every hit
    ' has to be looked at one at a time - see the fast path below for why that is affordable.

    Const PgNumStyle As String = "Print Pg Num"

    Dim rng As Range
    Dim hit As Range
    Dim stopAt As Long

    If Selection.Type = wdSelectionNormal Then
        Set rng = Selection.Range
    Else
        Set rng = ActiveDocument.Content
    End If

    ' Fast path. With no Print Pg Num paragraph in the document there is nothing to protect, so
    ' keep the single replace this macro has always done. That is the usual case: File Cleanup
    ' normally runs BEFORE Format $pg Tags has built any bars.
    If Not Sh_Style_Is_In_Use(ActiveDocument, PgNumStyle) Then
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "^s"          'Word wildcard for non-breaking space
            .Replacement.Text = " "
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
        Sh_Remove_Multi_Spaces rng
        Exit Sub
    End If

    ' Slow path. Walk the non-breaking spaces one by one and skip the ones in a bar. Only the
    ' hits are visited, not every paragraph, so this costs no more than the number of
    ' non-breaking spaces in the range.
    stopAt = rng.End
    Set hit = rng.Duplicate

    With hit.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^s"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop    ' MUST be wdFindStop - wdFindContinue would wrap past the end and never finish
        .Format = False
        .MatchWildcards = False
    End With

    Do While hit.Find.Execute
        If hit.Start >= stopAt Then Exit Do
        If Not Sh_Para_Style_Is(hit, PgNumStyle) Then
            hit.Text = " "    ' one character for one, so nothing after this point shifts
        End If

        ' Step past the space just looked at, whether or not it was replaced, and search on to
        ' the end of the range from there.
        hit.Collapse Direction:=wdCollapseEnd
        If hit.Start >= stopAt Then Exit Do
        hit.End = stopAt
    Loop

    ' The runs the two braille copies used to collapse in one replace - done here instead,
    ' because the walk above must replace one character with one character.
    Sh_Remove_Multi_Spaces ActiveDocument.Range(rng.Start, stopAt)

End Sub   '*** end of Sh_ReplaceNonBreakingSpacesWithNormalSpace macro ***

Private Function Sh_Style_Is_In_Use(ByVal targetDoc As Document, ByVal styleName As String) As Boolean
    '
    ' Version: 1.0  Date: 8/5/2026
    '
    ' True if the document has that style AND has used it. A document with no such style at all -
    ' a braille file, or one with no LP template attached - raises 5941 and answers False.

    On Error GoTo NotInUse
    Sh_Style_Is_In_Use = targetDoc.Styles(styleName).InUse
    Exit Function

NotInUse:
    Sh_Style_Is_In_Use = False

End Function   '*** end of Sh_Style_Is_In_Use function ***

Private Function Sh_Para_Style_Is(ByVal r As Range, ByVal styleName As String) As Boolean
    '
    ' Version: 1.0  Date: 8/5/2026
    '
    ' Is the first paragraph of the range in that style? The safe answer to "could not tell" is
    ' False, which leaves the caller doing what it did before this check existed.

    On Error GoTo NoMatch
    Sh_Para_Style_Is = (StrComp(r.Paragraphs(1).Style.NameLocal, styleName, vbTextCompare) = 0)
    Exit Function

NoMatch:
    Sh_Para_Style_Is = False

End Function   '*** end of Sh_Para_Style_Is function ***

Sub Sh_SleepForSeconds(ByVal Seconds As Double)

    ' Purpose: Sleep for a specified number of seconds while keeping Word responsive
    ' Version 1.0  Date: 7/15/2025
    
    Dim startTime As Single
    startTime = Timer
    
    Do While Timer < startTime + Seconds
        DoEvents ' Allows Word to remain responsive
    Loop
End Sub

Sub Sh_Convert_XML_File_To_Word_Document()
'
' Automatically converts an .xml (NIMAS or DAISY) file into a Word Document with reference pages tagged with $pg
'
' Version: 1.7  Date: 7/23/2026 - the "Prodnote" style is set red (EE0000, same as the LP template) in the converted document, which is built from Normal and so would otherwise show prodnotes as plain body text until the LP template is attached
' Version: 1.6  Date: 7/22/2026 - fix black Word workspace behind the "conversion complete" message: ScreenUpdating/Print view are restored (with a ScreenRefresh) BEFORE that message box instead of after the repaginate; background pagination and live spell/grammar stay suppressed through the repaginate
' Version: 1.5  Date: 7/21/2026 - <prodnote> content (body text and in tables) is now emitted as <p class="Prodnote"> so it imports carrying the "Prodnote" paragraph style (see Sh_Tag_Prodnotes_As_Prodnote_Style); requires "Prodnote" in the LpStyles keep-list
' Version: 1.4  Date: 7/21/2026 - perf: ScreenUpdating stays off through the whole import/repaginate region; live spell/grammar check + background pagination silenced during it (restored after); HTML imported in Draft view; fixed DoEvents pauses trimmed (~18s -> ~2s); redundant post-Unlink Fields.Update dropped; images embed via BreakLink without a per-image .Update disk re-fetch
' Version: 1.3  Date: 7/18/2026 - Save As now uses one dialog object so the file saves under the name the user types
' Version: 1.2  Date: 7/18/2026 - stabilize the document before saving; now saves only once (stabilize->save)
' Version: 1.1  Date: 7/8/2026 - added non modal message code - and file stabilization code
' Version: 1.0  Date: 3/3/2026
'
    Sh_Is_Doc_Open
    Sh_Convert_Progress_Form.Hide
    Sh_Spin_DoEvents

    If MsgBox( _
            "A file selection dialog window will appear after this message is closed." & vbCrLf & vbCrLf & _
            "Select or open the folder which contains the DAISY or NIMAS book to be converted into a Word Document, then click OK." & vbCrLf & vbCrLf & _
            "Your screen will show a wait message until the book has been converted into a Word document." & vbCrLf & vbCrLf & _
            "After completion, the folder will contain two new files: The .xml in .txt format (with reference pages tagged with $pg), and an .html file which can be opened with a web browser.", _
            vbOKCancel, _
            "Convert DAISY/NIMAS to Word Document") = vbOK Then
    Else
        Exit Sub
    End If

    Dim fldr As FileDialog
    Dim folderPath As String, fileName As String, baseFolder As String
    Dim xmlFile As String, txtCopy As String, htmlPath As String
    Dim finalDoc As Document
    Dim start As Single
    Dim fileContent As String
    
    Dim origState As Long, origTop As Long, origLeft As Long
    Dim su_Prev As Boolean, pag_Prev As Boolean, spell_Prev As Boolean, gram_Prev As Boolean

    ' --- Step 1: Window Setup ---
    origState = ActiveWindow.WindowState
    If origState = wdWindowStateMaximize Then ActiveWindow.WindowState = wdWindowStateNormal
    origTop = Application.Top: origLeft = Application.Left
    ActiveWindow.WindowState = origState
    
    ' --- Step 2: Folder Selection ---
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    If fldr.Show <> -1 Then Exit Sub
    folderPath = fldr.SelectedItems(1) & "\"

    ' --- Step 3: Check for XML ---
    fileName = Dir(folderPath & "*.xml", vbNormal)
    If fileName = "" Then
        MsgBox "No XML file was found in the selected folder. The process will now end.", vbCritical, "No XML Found"
        Exit Sub
    End If

    ' --- Step 4: User Input ---
    Sh_GP_String_1 = ""
    DN_XML_Type_Form.Show
    If UCase(Sh_GP_String_1) <> "DAISY" And UCase(Sh_GP_String_1) <> "NIMAS" Then
        MsgBox "No valid format selected. The process will now end.", vbExclamation, "Process Aborted"
        Exit Sub
    End If

    ' --- Step 4b: keep the images, or leave them out? ---
    ' Braille almost never wants the pictures and large print always does, so this is asked
    ' rather than guessed (Jerry, 8/5/2026). It is also the single biggest lever on how long
    ' the conversion takes. MEASURED end to end by Jerry on the same book and the same VM:
    ' 5 min 30 sec keeping the images, 2 min 30 sec omitting them. Embedding was 213 seconds
    ' of that, and dropping the pictures shortens the save as well - the file no longer has
    ' 86 MB of JPEGs to write.
    Sh_GP_String_2 = ""
    DN_Keep_Or_Omit_Images_Form.Show
    If Sh_GP_String_2 <> "KEEP" And Sh_GP_String_2 <> "OMIT" Then
        MsgBox "No image choice was made. The process will now end.", vbExclamation, "Process Aborted"
        Exit Sub
    End If

    ' --- Show the progress box before any real work starts ---
    ' It has to be up from the first stage. An earlier rewrite of Step 6 removed the opening
    ' Show along with the temp-document code it was sitting in, so the box did not appear until
    ' the repaginate stage and the user watched a still screen until then (8/3/2026).
    Sh_Convert_Progress_Form.Show vbModeless
    Sh_Convert_Progress_Form.SetProgress 1, "Reading the " & Sh_GP_String_1 & " file."
    DoEvents

    ' --- Step 5: Read XML using UTF-8 Stream
    xmlFile = folderPath & fileName
    txtCopy = folderPath & Left(fileName, InStrRev(fileName, ".") - 1) & ".txt"
    htmlPath = folderPath & Left(fileName, InStrRev(fileName, ".") - 1) & ".html"

    
    Dim tStream As Object
    Set tStream = CreateObject("ADODB.Stream")
    tStream.Charset = "utf-8"
    tStream.Open
    tStream.LoadFromFile xmlFile
    fileContent = tStream.ReadText
    tStream.Close

    fileContent = Replace(fileContent, ChrW(&HC2), "")
    fileContent = Replace(fileContent, ChrW(&HA0), " ")

    ' --- Step 6: Tagging Logic ---
    '
    ' Done on the STRING. This used to open a hidden Word document, push the whole XML into it,
    ' run a wildcard Find/Replace across it and read it all back out - a full document round
    ' trip to perform one substitution. Measured on the build box 8/3/2026 against a real 900KB
    ' NIMAS book: 0.48 seconds that way, below the timer's resolution this way, and the output
    ' is identical character for character apart from the trailing paragraph mark Word appends
    ' to every document.
    '
    ' The patterns are the same ones, translated from Word wildcards to RegExp. In Word
    ' wildcards "<" and ">" are word boundaries and have to be escaped as \< and \>; in a
    ' regular expression they are ordinary characters. Everything else carries over unchanged,
    ' including [0-9A-z], which is an ASCII RANGE from "0" to "z" - it admits the punctuation
    ' between them, and always has.

    Sh_Convert_Progress_Form.SetProgress 2, "Tagging reference pages with '$pg'"
    Sh_Spin_DoEvents

    ' everything before <dtbook is the XML prolog and is dropped, as before
    Dim dtPos As Long
    dtPos = InStr(1, fileContent, "<dtbook", vbTextCompare)
    If dtPos > 1 Then fileContent = Mid$(fileContent, dtPos)

    Dim reTag As Object
    Set reTag = CreateObject("VBScript.RegExp")
    reTag.Global = True

    If UCase(Sh_GP_String_1) = "DAISY" Then
        reTag.Pattern = ">([0-9A-z]{1,})</pagenum>"
        fileContent = reTag.Replace(fileContent, "<p>$pg$1<p></p></pagenum><p></p>")
    Else
        reTag.Pattern = ">([0-9A-z-]{1,}</pagenum>)"
        fileContent = reTag.Replace(fileContent, "<p>$pg$1<p></p>")
    End If
    
    Sh_Convert_Progress_Form.Show vbModeless

    ' Start it turning. Showing the form does not: without this the spinner flag stays False
    ' and every Advance and SpinTick exits at once, so it has never moved during a conversion
    ' (Jerry, 8/3/2026). The DoEvents through the rest of this macro are Sh_Spin_DoEvents for
    ' the same reason - the OnTime tick alone fires once a second at best.
    Sh_Convert_Progress_Form.SetProgress 3, "Creating .txt file with .xml code and creating .html file"
    Sh_Spin_DoEvents
    Sh_PauseSeconds 0.3   'brief tick so the status form paints

    ' --- Step 7: Fix Images (Base Href) ---
    baseFolder = "file:///" & Replace(folderPath, "\", "/")
    baseFolder = Replace(baseFolder, " ", "%20")
    
    ' Mark <prodnote> content (body text and inside tables) so it imports carrying the
    ' "Prodnote" paragraph style. Word's HTML importer strips unknown tags like <prodnote>,
    ' so the content must already be <p class="Prodnote"> by the time the HTML is imported.
    fileContent = Sh_Tag_Prodnotes_As_Prodnote_Style(fileContent)

    ' The mso-style-name rule is what makes Word's HTML importer map class="Prodnote"
    ' onto the Word paragraph style named "Prodnote".
    fileContent = "<html><head><meta charset=""UTF-8""><base href=""" & baseFolder & """>" & _
                  "<style>p.Prodnote{mso-style-name:""Prodnote"";}</style>" & _
                  "</head><body>" & fileContent & "</body></html>"

    ' --- Step 8: Save Outputs using UTF-8 Stream (Fixes Black Diamonds) ---
    Dim outStream As Object
    Set outStream = CreateObject("ADODB.Stream")
    outStream.Type = 2
    outStream.Charset = "utf-8"
    
    outStream.Open
    outStream.WriteText fileContent
    outStream.SaveToFile htmlPath, 2
    outStream.Close
    
    outStream.Open
    outStream.WriteText fileContent
    outStream.SaveToFile txtCopy, 2
    outStream.Close
    
    Set outStream = Nothing

    ' --- Step 9: Final Reveal & Image Embedding ---
    Set finalDoc = Documents.Add
    
    With finalDoc.ActiveWindow
        .WindowState = wdWindowStateNormal: .Top = origTop: .Left = origLeft: .WindowState = origState
    End With
    
    'save the name of the current document
    Dim currentdoc As Document
    Set currentdoc = ActiveDocument 'will work with blank, unsaved documents too

    Sh_Convert_Progress_Form.Show vbModeless
    Sh_Convert_Progress_Form.SetProgress 4, "Importing the book into Word. The bar cannot move during this step."
    Sh_Spin_DoEvents
    Sh_PauseSeconds 0.3   'brief tick so the status form paints

    ' --- Speed: silence background work for the whole heavy region (InsertFile, image
    ' embedding, repaginate). ScreenUpdating stays OFF until every heavy step is done; the
    ' *_Prev locals are restored just before the Save As UI further below. ---
    su_Prev = Application.ScreenUpdating
    pag_Prev = Application.Options.Pagination
    spell_Prev = Application.Options.CheckSpellingAsYouType
    gram_Prev = Application.Options.CheckGrammarAsYouType
    Application.ScreenUpdating = False
    Application.Options.Pagination = False
    Application.Options.CheckSpellingAsYouType = False
    Application.Options.CheckGrammarAsYouType = False

    ' Supress the initial security warning during import
    Application.DisplayAlerts = wdAlertsNone

    ' Import + process in Draft view with proofing marks off -- far less work than Print
    ' view / live spell+grammar checking on a large imported book.
    On Error Resume Next
    finalDoc.ActiveWindow.View.Type = wdNormalView
    finalDoc.ShowSpellingErrors = False
    finalDoc.ShowGrammaticalErrors = False
    On Error GoTo 0

    ' Use InsertFile instead of Copy/Paste to prevent the 0x5 Clipboard Crash
    finalDoc.Range.InsertFile fileName:=htmlPath, ConfirmConversions:=False

    ' Permanently embed all images and break links so the security warning goes away forever.
    ' BreakLink embeds the image InsertFile already loaded, so no per-image .Update re-fetch
    ' from disk is needed -- that re-fetch was the biggest cost on image-heavy books.
    '
    ' THIS LOOP IS THE CONVERSION. Timed on the build box 8/3/2026 against a real NIMAS book
    ' with 1,236 images: 213 seconds here out of 301 for the whole job - 71 per cent of it, at
    ' about 173 ms an image. The HTML import before it takes 14 seconds and the save after it
    ' 71. Everything else together is under 3.
    '
    ' Two ways round it were tried and neither worked. Fields.Unlink does nothing for these
    ' pictures - they arrive as linked InlineShapes, not INCLUDEPICTURE fields, and the saved
    ' file came to 0.3 MB against 43.7 MB, so nothing was embedded. Setting
    ' SavePictureWithDocument without BreakLink was no faster either. The time is the price of
    ' putting 86 MB of JPEGs inside the document, and it is what makes the file standalone.
    '
    ' So it is REPORTED instead. This is the only long stage that is a loop, which makes it the
    ' only one that can show progress at all - the import and the save are single blocking calls
    ' and nothing can move during them. The bar steps every 10 images rather than every one: a
    ' repaint per image would add its own cost to the very loop being measured.
    Dim shp As inlineShape
    Dim nShapes As Long, iShape As Long

    nShapes = finalDoc.InlineShapes.Count
    If Sh_GP_String_2 = "OMIT" Then
        Sh_Convert_Progress_Form.SetProgress 6, "Removing " & Format(nShapes, "#,##0") & " images."
    Else
        Sh_Convert_Progress_Form.SetProgress 6, "Embedding " & Format(nShapes, "#,##0") & _
            " images. This is the longest part of the conversion."
    End If

    If Sh_GP_String_2 = "OMIT" Then

        ' Backwards: deleting a shape renumbers every one after it, so a forward loop would
        ' step over every second picture.
        For iShape = nShapes To 1 Step -1
            finalDoc.InlineShapes(iShape).Delete
            If iShape Mod 10 = 0 And nShapes > 0 Then
                Sh_Convert_Progress_Form.SetProgress 6 + (71 * (nShapes - iShape) / nShapes), _
                    "Removing image " & Format(nShapes - iShape, "#,##0") & " of " & Format(nShapes, "#,##0") & "."
                DoEvents
            End If
        Next iShape

    Else

        For Each shp In finalDoc.InlineShapes
            If Not shp.LinkFormat Is Nothing Then
                shp.LinkFormat.SavePictureWithDocument = True
                shp.LinkFormat.BreakLink
            End If
            iShape = iShape + 1
            If iShape Mod 10 = 0 And nShapes > 0 Then
                Sh_Convert_Progress_Form.SetProgress 6 + (71 * iShape / nShapes), _
                    "Embedding image " & Format(iShape, "#,##0") & " of " & Format(nShapes, "#,##0") & "."
                DoEvents
            End If
        Next shp

    End If

    ' Unlink any remaining field codes Word might complain about
    finalDoc.Fields.Unlink

    ' Turn alerts back on
    Application.DisplayAlerts = wdAlertsAll

    With finalDoc.Range.Font
        .Name = "Courier New": .Size = 10
    End With

    ' Make the "Prodnote" style red in the converted document. This document is created from
    ' Normal (Documents.Add), so Word invents "Prodnote" from the mso-style-name rule in the
    ' generated HTML with no formatting of its own -- prodnotes would otherwise look like
    ' ordinary body text until the LP template is attached. EE0000 = RGB(238, 0, 0) is the
    ' same red the Prodnote style carries in LargePrintTemplate.dotx, so the color does not
    ' shift when that template is attached later. Silently skipped when the book contained no
    ' prodnotes (Word never creates the style, so the lookup fails).
    On Error Resume Next
    finalDoc.Styles("Prodnote").Font.Color = RGB(238, 0, 0)
    Err.Clear
    On Error GoTo 0

    ' Some DAISY and NIMAS books wrap each prodnote in quotation marks. Those belong to the
    ' source markup rather than to the note, so take them off here -- after the prodnotes
    ' carry the Prodnote style, and before the document is stabilized and saved. Silent, and
    ' does nothing at all when the book contained no prodnotes.
    Sh_Strip_Prodnote_Enclosing_Quotes finalDoc

    Sh_Color_Dollar_PG_Red

    ' Repaint BEFORE the "conversion complete" message box: turn ScreenUpdating back on,
    ' come out of Draft into Print view, and force a refresh. Without this the message box
    ' appears over a workspace Word never painted, which shows BLACK instead of the normal
    ' gray surround (the document page itself still looks white).
    ' Background pagination and live spell/grammar stay OFF through the repaginate below --
    ' those are the expensive ones and they cause no painting artifacts.
    finalDoc.Activate
    On Error Resume Next
    finalDoc.ActiveWindow.View.Type = wdPrintView
    On Error GoTo 0
    Application.ScreenUpdating = su_Prev
    Application.ScreenRefresh
    Sh_PauseSeconds 0.3
    
    ' Make the document visible and active on screen
    Sh_Convert_Progress_Form.Hide ' Hide the progress box so it doesn't block the document
    Application.Activate        ' Force the Word Application itself to the front of Windows
    currentdoc.Activate         ' Ensure the specific document is active
    Sh_Spin_DoEvents
    
    MsgBox "Here is the " & Sh_GP_String_1 & " file in Word format." & _
    vbCrLf & vbCrLf & "All reference page numbers have been tagged with $pg tags ready for validation." & _
    vbCrLf & vbCrLf & "The document will now be stabilized and then saved as a Word document with a .docx file type.", vbInformation
    
    Dim doc As Document
    Dim userChoice As VbMsgBoxResult
    Set doc = ActiveDocument
    
    ' Stabilize the document FIRST, then save it exactly once (stabilize -> save).
    Set currentdoc = doc

    ' The whole book goes to Tahoma 12 HERE, ahead of the repaginate below. Changing the font
    ' changes where every line and every page breaks, so it has to happen while a repaginate and
    ' a save still follow it - done after the save, the file on disk would not match the file on
    ' screen, and the document would be left dirty. Color is not touched, so the red $pg tags
    ' and the red prodnotes survive it. (Jerry, 8/5/2026)
    Sh_Convert_Progress_Form.Show vbModeless
    Sh_Convert_Progress_Form.SetProgress 76, "Setting the document font to Tahoma 12"
    Sh_Spin_DoEvents
    Sh_Set_Whole_Document_Font doc, "Tahoma", 12

    Sh_Convert_Progress_Form.Show vbModeless
    Sh_Convert_Progress_Form.SetProgress 78, "Repaginating the document"
    Sh_Spin_DoEvents
    Sh_PauseSeconds 0.3   'brief tick so the status form paints

    doc.Repaginate
    doc.UndoClear

    ' --- Speed: heavy work is done. Restore the background-processing options we silenced.
    ' (ScreenUpdating and Print view were already restored before the message box above, so
    ' the workspace paints correctly there. Fields.Unlink already made every field static
    ' text, so the old Fields.Update pass here was redundant.) ---
    Application.Options.CheckGrammarAsYouType = gram_Prev
    Application.Options.CheckSpellingAsYouType = spell_Prev
    Application.Options.Pagination = pag_Prev

    ' Hide the progress form momentarily so Windows can cleanly shift focus to the Save As dialog
    Sh_Convert_Progress_Form.Hide

    ' Force the Word Application and your specific document to the front
    Application.Activate
    currentdoc.Activate
    Sh_Spin_DoEvents
    Sh_PauseSeconds 0.5   'brief settle so the Save As dialog receives focus cleanly

    Dim dlgSaveAs As Dialog
SaveTheFile:
    ' Now save the stabilized document exactly once.
    If doc.Path = "" Then
        ' Unnamed document: force the Save As dialog.
        ' Capture ONE dialog object and use it for BOTH .Display and .Execute, so the file
        ' is saved under the name the user types. (A separate Dialogs(wdDialogFileSaveAs)
        ' reference for .Execute ignores the typed name and uses the document's default name.)
        Set dlgSaveAs = Dialogs(wdDialogFileSaveAs)
        ' .Display ONLY opens the window to get the file name; it does NOT save yet
        If dlgSaveAs.Display <> -1 Then
            ' User canceled the dialog
            userChoice = MsgBox( _
                "You canceled the Save As." & vbCrLf & vbCrLf & _
                "Continuing without saving may result in an unstable Word document." & vbCrLf & vbCrLf & _
                "Do you want to reconsider and save this file?", _
                vbYesNo + vbExclamation, _
                "Save As Canceled")

            If userChoice = vbNo Then
                Unload Sh_Convert_Progress_Form
                Exit Sub
            Else
                GoTo SaveTheFile
            End If
        Else
            ' --- SUCCESSFUL FILE CHOICE ---
            ' 1. Show the non-modal form BEFORE saving
            Sh_Convert_Progress_Form.Show vbModeless
            Sh_Convert_Progress_Form.SetProgress 80, "Saving the document. On a book with many images this is the second longest step, and the bar cannot move during it."
            Sh_Spin_DoEvents

            ' 2. Execute the save on the SAME dialog object so the typed name is used
            dlgSaveAs.Execute

            ' 3. Keep the message up for a brief moment so they see it finish
            Sh_PauseSeconds 1
        End If
    Else
        ' Named document: save it in place, exactly once.
        Sh_Convert_Progress_Form.Show vbModeless
        Sh_Convert_Progress_Form.SetProgress 80, "Saving the document. On a book with many images this is the second longest step, and the bar cannot move during it."
        Sh_Spin_DoEvents
        Sh_PauseSeconds 0.5   'brief tick so the status form paints

        doc.Save
    End If

    Sh_Convert_Progress_Form.SetProgress 100, "Finished."
    DoEvents
    Sh_PauseSeconds 0.4   ' let the full bar be seen before the box goes

    'Unload the progress form completely
    Unload Sh_Convert_Progress_Form
    Sh_Spin_DoEvents
    
    'Re-assert dominance for your saved document AFTER the form is entirely gone
    Application.Activate
    doc.Activate
    ActiveWindow.View.Type = wdPrintView
    Sh_Spin_DoEvents
    
    MsgBox "Conversion is complete and the file has been stabilized and saved", vbInformation, "Operation Complete"

    ' A converted book that carries prodnotes needs a word of explanation -- what they are,
    ' why they are red, and what to do with them for braille and for large print. Shown after
    ' the file is saved, and only when the document actually has some: the note opens by
    ' saying it "contains one or more prodnotes", which must not appear over a book with none.
    ' The text is too long for a MsgBox (VBA truncates at about 1024 characters), so it lives
    ' on Sh_Prodnote_Info_Form. Closing the form ends the macro.
    If Sh_Document_Has_Prodnotes(doc) Then
        ' Open the Styles pane behind the note, listing EVERY style and sorted As Recommended,
        ' so the reader can see the red Prodnote entry in the list while the note explains it.
        ' Same pair MS_Set_Word_Config_For_New_Install sets. Guarded: the pane is a courtesy
        ' and a failure here must not cost the user the note itself.
        On Error Resume Next
        Application.TaskPanes(wdTaskPaneFormatting).Visible = True
        doc.FormattingShowNextLevel = False
        doc.StyleSortMethod = wdStyleSortRecommended
        doc.FormattingShowFilter = wdShowFilterStylesAll
        Err.Clear
        On Error GoTo 0

        Sh_Prodnote_Info_Form.Show
    End If

End Sub   '*** end of Sh_Convert_XML_File_To_Word_Document macro ***

Function Sh_Tag_Prodnotes_As_Prodnote_Style(ByVal src As String) As String
'
' Rewrites DTBook <prodnote> blocks so the converted Word document carries the "Prodnote"
' paragraph style on the prodnote's content -- in the body text and inside tables.
'
' Word's HTML importer strips unknown tags such as <prodnote> and keeps only their text,
' so the marking has to happen in the string BEFORE Sh_Convert_XML_File_To_Word_Document
' imports the generated .html. Two shapes occur in real books:
'   NIMAS: <prodnote render="optional">bare text</prodnote>
'   DAISY: <prodnote imgref=".." ..><p id="..">text</p><p id="..">text</p></prodnote>
' DAISY prodnotes span several lines, so this is done with string parsing rather than a
' Word wildcard Find (wildcards cannot match across paragraph marks).
'
' For the DAISY shape each inner <p> gets class="Prodnote"; for the NIMAS shape the bare
' text is wrapped in a single <p class="Prodnote"> paragraph. The <prodnote> wrapper itself
' is dropped either way (Word would discard it anyway).
'
' NOTE: "Prodnote" must stay in the LpStyles keep-list in
' Lp_Remove_All_Styles_Except_Lp_Styles, or attaching the LP template will convert these
' paragraphs back to Normal and delete the style.
'
' Version: 1.1  Date: 8/3/2026 - 3.79 s -> 0.00 s on a real 900KB NIMAS book, output identical. The tag search was
'                               case-insensitive InStr over the whole file, 1564 times; it now searches a lowercased
'                               copy with a binary compare. Accumulation changed to Join, which was NOT the problem.
' Version: 1.0  Date: 7/21/2026
'
    Dim parts() As String
    Dim n As Long
    Dim pos As Long, openStart As Long, openEnd As Long, closeStart As Long
    Dim inner As String
    Const CLOSETAG As String = "</prodnote>"   ' lowercase: matched against srcLower

    ' Collected into an array and Joined at the end rather than built up with
    ' outStr = outStr & ... - VBA copies the whole accumulated string on every concatenation.
    ' Honest note: measured, this was NOT what made the routine slow. Replacing it changed
    ' 3.77 s to 3.68 s. It is kept because it is the better shape as files grow, but the real
    ' cost was the case-insensitive search below.
    ReDim parts(0 To 255)

    ' Search a LOWERCASED COPY with a binary compare, and slice from the original. The tags are
    ' matched case-insensitively, as before, but vbTextCompare does locale-aware matching a
    ' character at a time and this walks a whole book looking for 782 of them, twice over.
    ' Measured on the build box 8/3/2026 against a real 900KB NIMAS file: 1.85 s for the opening
    ' tag and 1.83 s for the closing one - 3.68 s, which was this entire routine. The same walk
    ' with vbBinaryCompare is below the timer's resolution, and LCase$ of the whole string costs
    ' nothing and does not change its length, so every position still lines up with src.
    Dim srcLower As String
    srcLower = LCase$(src)

    pos = 1
    Do
        openStart = InStr(pos, srcLower, "<prodnote", vbBinaryCompare)
        If openStart = 0 Then Exit Do
        openEnd = InStr(openStart, src, ">")                      ' end of the opening tag
        If openEnd = 0 Then Exit Do
        closeStart = InStr(openEnd, srcLower, CLOSETAG, vbBinaryCompare)
        If closeStart = 0 Then Exit Do

        If n + 2 > UBound(parts) Then ReDim Preserve parts(0 To UBound(parts) * 2)

        ' everything ahead of this prodnote passes through untouched
        parts(n) = Mid$(src, pos, openStart - pos)
        n = n + 1

        inner = Mid$(src, openEnd + 1, closeStart - openEnd - 1)

        If InStr(1, inner, "<p ", vbTextCompare) > 0 Or InStr(1, inner, "<p>", vbTextCompare) > 0 Then
            ' DAISY shape -- style each paragraph the prodnote already contains
            inner = Replace(inner, "<p ", "<p class=""Prodnote"" ", , , vbTextCompare)
            inner = Replace(inner, "<p>", "<p class=""Prodnote"">", , , vbTextCompare)
            parts(n) = inner
        Else
            ' NIMAS shape -- bare text becomes one Prodnote paragraph
            parts(n) = "<p class=""Prodnote"">" & inner & "</p>"
        End If
        n = n + 1

        pos = closeStart + Len(CLOSETAG)
    Loop

    If n > UBound(parts) Then ReDim Preserve parts(0 To n)
    parts(n) = Mid$(src, pos)

    ReDim Preserve parts(0 To n)
    Sh_Tag_Prodnotes_As_Prodnote_Style = Join(parts, "")
End Function   '*** end of Sh_Tag_Prodnotes_As_Prodnote_Style function ***

Function Sh_Strip_Prodnote_Enclosing_Quotes(ByVal targetDoc As Document) As Long
'
' Removes the quotation marks that ENCLOSE a prodnote -- in the body text and inside tables.
' Some DAISY and NIMAS books wrap the whole note in them, "Illustration of a barn.", where
' the marks are punctuation of the source markup and not part of the note itself.
'
' Called at the end of Sh_Convert_XML_File_To_Word_Document, once the prodnotes carry the
' "Prodnote" style (Sh_Tag_Prodnotes_As_Prodnote_Style put it there) and before the document
' is stabilized and saved. Silent -- the converter reports its own progress.
'
' A note is only unwrapped when it BEGINS and ENDS with a matching pair. Four pairs count:
' straight " ", straight ' ', and the two curly pairs. A stray mark at one end only is left
' exactly as it is, and so is a mismatched pair such as a curly open with a straight close.
'
' Two shapes, and the order they are tried in matters:
'   1. Every Prodnote paragraph is tested on its own first.
'   2. A DAISY prodnote can arrive as SEVERAL Prodnote paragraphs in a row, with the opening
'      mark on the first and the closing mark on the last. So any run of consecutive Prodnote
'      paragraphs in which step 1 removed nothing is then tested as one single note.
' Testing paragraphs first is what stops two individually-quoted notes that happen to sit
' back to back from being read as one quoted block.
'
' Leading and trailing spaces are ignored when deciding, so a note that begins with a space
' before its opening mark still unwraps. Returns the number of pairs removed.
'
' Version: 1.0  Date: 8/1/2026
'
    Dim runs As Collection          ' each item is itself a Collection of paragraph Ranges
    Dim thisRun As Collection
    Dim para As Paragraph
    Dim r As Long, p As Long
    Dim runRemoved As Long
    Dim removed As Long

    ' --- 0. Leave immediately if this book has no prodnotes ---
    ' A style-aware Find settles that far more cheaply than the walk below, which visits every
    ' paragraph in the book -- many thousands of them in a converted title.
    If Not Sh_Document_Has_Prodnotes(targetDoc) Then Exit Function

    ' --- 1. Group the Prodnote paragraphs into runs of consecutive ones ---
    ' Document.Paragraphs walks the body text and the table cells in document order, so
    ' paragraphs adjacent in this collection are adjacent on the page. Comparing style names
    ' is safe even in a document with no Prodnote style at all -- nothing matches.
    Set runs = New Collection

    For Each para In targetDoc.Paragraphs
        If StrComp(para.Style.NameLocal, "Prodnote", vbTextCompare) = 0 Then
            If thisRun Is Nothing Then
                Set thisRun = New Collection
                runs.Add thisRun
            End If
            thisRun.Add para.Range
        Else
            Set thisRun = Nothing       ' a non-prodnote paragraph ends the run
        End If
    Next para

    If runs.count = 0 Then Exit Function      ' this book had no prodnotes

    ' --- 2. Unwrap, working backwards through the document ---
    ' Back to front so that removing a mark can never shift the position of one still to be
    ' examined -- the same reason Sh_Delete_Prodnote_Paragraphs deletes in reverse.
    For r = runs.count To 1 Step -1
        Set thisRun = runs(r)
        runRemoved = 0

        ' 2a. Each paragraph of the run on its own, last one first
        For p = thisRun.count To 1 Step -1
            If Sh_Strip_One_Quote_Pair(thisRun(p), thisRun(p)) Then
                runRemoved = runRemoved + 1
            End If
        Next p

        ' 2b. Nothing matched inside the run, so try the run as ONE note: opening mark on
        '     the first paragraph, closing mark on the last.
        If runRemoved = 0 And thisRun.count >= 2 Then
            If Sh_Strip_One_Quote_Pair(thisRun(1), thisRun(thisRun.count)) Then
                runRemoved = 1
            End If
        End If

        removed = removed + runRemoved
    Next r

    Sh_Strip_Prodnote_Enclosing_Quotes = removed
End Function   '*** end of Sh_Strip_Prodnote_Enclosing_Quotes function ***

Private Function Sh_Strip_One_Quote_Pair(ByVal openRange As Range, ByVal closeRange As Range) As Boolean
'
' Removes one enclosing pair of quotation marks: the first visible character of openRange and
' the last visible character of closeRange, but only when the two are a matching pair.
' Pass the SAME range twice to test a single paragraph; pass the first and last paragraphs of
' a run to test that run as one note. Returns True when a pair was removed.
'
' Version: 1.0  Date: 8/1/2026
'
    Dim openText As String, closeText As String
    Dim openPos As Long, closePos As Long

    openText = Sh_Para_Visible_Text(openRange)
    closeText = Sh_Para_Visible_Text(closeRange)

    openPos = Sh_First_Visible_Char(openText)
    closePos = Sh_Last_Visible_Char(closeText)
    If openPos = 0 Or closePos = 0 Then Exit Function       ' empty, or nothing but spaces

    ' Within ONE paragraph the two marks must have something between them, so a paragraph
    ' holding nothing but "" is left alone instead of being emptied.
    If openRange.Start = closeRange.Start Then
        If closePos - openPos < 2 Then Exit Function
    End If

    If Not Sh_Is_Quote_Pair(Mid$(openText, openPos, 1), Mid$(closeText, closePos, 1)) Then Exit Function

    ' The closing mark goes first: it sits later in the document, so removing it cannot move
    ' the opening one. Doing it the other way round would leave closePos pointing one
    ' character too far along.
    closeRange.Characters(closePos).Delete
    openRange.Characters(openPos).Delete

    Sh_Strip_One_Quote_Pair = True
End Function   '*** end of Sh_Strip_One_Quote_Pair function ***

Private Function Sh_Para_Visible_Text(ByVal r As Range) As String
'
' The text of a paragraph range as the reader sees it. A paragraph range always ends with its
' paragraph mark, and the last paragraph of a table cell ends with the end-of-cell marker as
' well; neither is text. What is left is a prefix of the range, so a character position in it
' is also a valid index into Range.Characters.
'
' Version: 1.0  Date: 8/1/2026
'
    Dim t As String
    t = r.Text

    Do While Len(t) > 0
        Select Case Right$(t, 1)
            Case vbCr, vbLf, Chr$(7)
                t = Left$(t, Len(t) - 1)
            Case Else
                Exit Do
        End Select
    Loop

    Sh_Para_Visible_Text = t
End Function   '*** end of Sh_Para_Visible_Text function ***

Private Function Sh_First_Visible_Char(ByVal t As String) As Long
'
' Position of the first character that is not a space, or 0 when there is none.
'
' Version: 1.0  Date: 8/1/2026
'
    Dim i As Long
    For i = 1 To Len(t)
        If Not Sh_Is_Space_Char(Mid$(t, i, 1)) Then
            Sh_First_Visible_Char = i
            Exit Function
        End If
    Next i
End Function   '*** end of Sh_First_Visible_Char function ***

Private Function Sh_Last_Visible_Char(ByVal t As String) As Long
'
' Position of the last character that is not a space, or 0 when there is none.
'
' Version: 1.0  Date: 8/1/2026
'
    Dim i As Long
    For i = Len(t) To 1 Step -1
        If Not Sh_Is_Space_Char(Mid$(t, i, 1)) Then
            Sh_Last_Visible_Char = i
            Exit Function
        End If
    Next i
End Function   '*** end of Sh_Last_Visible_Char function ***

Private Function Sh_Is_Space_Char(ByVal c As String) As Boolean
'
' A space, a tab, or a non-breaking space -- the three that can sit outside a quotation mark
' after an HTML import and hide it from a plain first-character test.
'
' Version: 1.0  Date: 8/1/2026
'
    Sh_Is_Space_Char = (c = " " Or c = vbTab Or c = ChrW(160))
End Function   '*** end of Sh_Is_Space_Char function ***

Private Function Sh_Is_Quote_Pair(ByVal openCh As String, ByVal closeCh As String) As Boolean
'
' True when the two characters are a matching pair of quotation marks. The straight marks are
' their own partner; the curly ones have a distinct opening and closing character and must be
' matched properly, so a curly open followed by a straight close is NOT a pair. The curly
' marks are written as ChrW codes so they survive every export and re-import of this module.
'
' Version: 1.0  Date: 8/1/2026
'
    Select Case openCh
        Case """"                                       ' straight double  "
            Sh_Is_Quote_Pair = (closeCh = """")
        Case "'"                                        ' straight single  '
            Sh_Is_Quote_Pair = (closeCh = "'")
        Case ChrW(8220)                                 ' curly double     open
            Sh_Is_Quote_Pair = (closeCh = ChrW(8221))   '                  close
        Case ChrW(8216)                                 ' curly single     open
            Sh_Is_Quote_Pair = (closeCh = ChrW(8217))   '                  close
    End Select
End Function   '*** end of Sh_Is_Quote_Pair function ***

Sub Sh_PauseSeconds(ByVal Seconds As Single)
    Dim startTime As Single
    startTime = Timer

    Do While (Timer - startTime + 86400) Mod 86400 < Seconds
        DoEvents
    Loop
End Sub   '*** end of Sh_PauseSeconds(ByVal Seconds As Single)macro ***


'------------------------------------------------------------------------------------
'/ / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
'\ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
'
'                    ************* End of Macros ************************
'
'/ / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
'\ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
'------------------------------------------------------------------------------------
