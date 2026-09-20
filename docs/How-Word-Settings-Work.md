# How VistaType LP Handles Word's Settings

**A guide for transcribers.** You do not have to do anything described here — except one
ten-second habit at the very end, which is worth forming. This explains what VistaType LP is
doing behind the scenes, so that if you ever notice Word behaving differently in one document
than in another, you know why, and you know it is on purpose.

---

## Definitions

Here are some terms used in this document:

Ordinary File/Document:        a document type which is not a braille file or a large print file - think a blank document or a letter to grandma
Braille File/Document:         a document type which is destined to be used in the Duxbury Braille Translator and has a BANA template attached
Large Print File/Document:     a document type created by VistaType LP which has a large print template attached
AutoCorrect Settings:          a set of rules used to control how text looks when typing text in a Word Document. 

---

## The one Word fact everything else rests on

**Word has only one set of typing settings, for the whole program.**

Open **File → Options → Proofing → AutoCorrect Options** and you are looking at settings that
Word uses for the type of document you are looking at... the document in front of you. Change smart quotes while writing a
ordinary file, and you have changed them for every document you will ever open — your letter to grandma, your grocery list, or your master's thesis.

That is how Microsoft built Word. Everything below follows from it.

---

## Why is one set of typing settings a problem for our work

Braille and large print need Word to type differently. Sometimes they need opposite things.

Type `1/2` in a **braille** file and it should become **½**, because that single character is
what the Duxbury translator turns into correct braille.

Type `1/2` in a **large print** book and it must stay exactly as you typed it: `1/2`. A compact
fraction is one character, and its digits are drawn much smaller than the rest of your text. In
an 18-point book, that fraction is not 18 point. A reader who needs large print cannot read it because the characters are too small.

Neither of those is a transcriber preference. One is what the braille translator requires; the other is
what the large print reader requires. And Word can only hold one configuration at a time.

---

## What VistaType LP does about Word's typing settings

**It follows the document type you are actually looking at.**

Click into a braille file, and Word's typing setting are set for braille. Click into a large print book, and
Word is set up for large print. Click into an ordinary document, and Word goes back to the default setting... (your own
everyday settings).

It happens the moment you click, without asking you, and **you cannot see it on your screen** —
your Styles pane, your rulers, your formatting marks and your screen view stays unchanged. Only the typing behavior changes.

Opening or creating a document is different: that is the moment VistaType LP sets the view up
for the kind of document it is. After that, the screen is yours.

So with a large print document, a braille document and an ordinary document all open side by side, you can move between them all afternoon and each
one behaves correctly according to their AutoCorrect option settings. You will not notice it happening.

---

### Which one is in force?

The **Document Settings** macro on the Quick Access Toolbar (it looks like a circle with the letter
"i" in it) tells you what the settings for the current document are. It has a line reading *"Word is
configured for..."* — large print, braille, or default (ordinary document) settings. If Word is ever
behaving oddly, that line is the first thing to look at.

Below it is a second line, *"AutoCorrect list in use:"*, which says which of the three replacement
lists is loaded. See *Your AutoCorrect list is your own again* below.

## How typing settings get created

The typing settings are created when VistaType LP is installed.

The setting for braille and Large print documents are ready to go right after installation. They can be changed but 
only for the amount of time that the document is open in Word. New or opened braille and large print documents 
will default to their original settings.

The typing settings for an ordinary Word file are also created with the installation and they are essentially 
the same setting Word uses when it is first installed. Those settings may not be suitable for many users, 
so VistaType LP provides a method for changing them on a more permanent basis, and that is explained in detail below.

---

## The settings this is all about
  
In the AutoCorrect Options — **File → Options → Proofing → AutoCorrect Options** — there are several
tabs across the top. Three of them matter here:

| Tab | What it decides |
|---|---|
| **AutoCorrect** | Capital letters, and a list that turns a mistyped `teh` into `the` |
| **AutoFormat As You Type** | What Word changes **while you are typing** |
| **AutoFormat** | What Word changes when you run the **AutoFormat** command on a whole document |

Every checkbox on those three tabs is one Word uses — **thirty-four of them** —
along with eight more about spelling, grammar and the display of the Styles pane. **Forty-two in all.**

The **AutoFormat** command is the odd one out. It reformats a whole document in one go, and while
Word does not put it on the ribbon, VistaType has placed it on the full Vista Type LP Quick Access Toolbar... 
it looks like a little rectangle with a lightning bolt. If you let the installer put VistaType's 
own Quick Access then you will see it. You can generally ignore its use... it's there for the power-users.  — nothing runs it on its own.

---

### It learns from you when you use an ordinary document

Here is the part most people would not expect.

You can do far more in Word than just editing braille or large print documents, so, when you are working in an ordinary document,
and when you change one or more of these settings, VistaType LP remembers the changes. When you create or open an
ordinary document those changed settings will be used. Keep in mind that the settings are not part of the document and 
changes in the settings from one edited document to another may not look the same as when they were created with different typing settings.

**You never have to say "save this".** Change a checkbox in any of the settings and it is
yours until you change it again.

---

## What actually differs between braille and large print

Most of the AutoCorrect Options are set the same way for both. Only a few things genuinely differ, and
every one of them is about what the reader receives:

| When you type | In a large print book | In a braille file |
|---|---|---|
| `1/2`, `1/4`, `3/4` | stay as typed | become `½`, `¼`, `¾` |
| any other fraction — `1/3`, `2/3`, `5/8` … | stay as typed | **also stay as typed** |
| a web or network address | stays plain text | becomes a clickable link |

The second row changed on 29 August 2026 and is worth knowing. Those three are Word's own
conversion, which large print switches off and braille switches on. **Every other fraction used to
convert as you typed as well, and no longer does.** Braille still needs them, but it gets them when
you run Full File Cleanup rather than while you are typing — so what you typed stays what you typed
until you ask for it to be changed.

And if you run the **AutoFormat** command on a whole large print document, `1st` is left alone but in a large
braille document it is raised to `1ˢᵗ`.

Neither of those may be your preference... They decide what the
Duxbury translator receives, and what a low-vision reader sees on the page. They belong to the
readers, not to us.

---

## Spelling and grammar

Braille files switch off **grammar checking as you type** and the **contextual spelling
checker**, and tell Word to stop ignoring numbers mixed into words. That is correct: the
grammar checker objects endlessly to braille formatting and would drive you to distraction.

Both braille and large print switch on **"Set left- and first-indent with tabs and backspaces"**, so the Tab
and Backspace keys set your indents while you work. If you prefer it off, switch it off in an
ordinary document and that is what you will get back.

---

## Where your settings are kept

In a small file, in your own Windows profile:

```
%AppData%\VistaType LP Settings\VistaType.ini
```

Three things worth knowing about it:

- **Uninstalling VistaType LP does not delete the settings.** If you reinstall, or install a new version,
  your settings are still there. That is deliberate.
- **It can be copied to a new computer.** If you get a new machine, copy that folder across and
  your settings come with you.
- **You never need to open it**, and nothing is lost if you delete it — VistaType LP simply
  starts learning your settings again.

Three more files sit beside it, one per kind of document:

```
%AppData%\VistaType LP Settings\AutoCorrect-DEF.txt     your ordinary documents
%AppData%\VistaType LP Settings\AutoCorrect-LP.txt      large print
%AppData%\VistaType LP Settings\AutoCorrect-BRL.txt     braille
```

And a log, `VistaType-Errors.log`, holding the last fifty faults, newest first. If a macro ever
fails, the message offers to open that folder for you.

---

## Your AutoCorrect list is your own again

This is the change most worth knowing about, and it arrived on 19 September 2026.

**Word keeps one AutoCorrect replacement list for the whole program.** Not one per document — one,
shared by everything you open. Nothing in VBA can make a second.

That had a consequence nobody intended. Braille and large print both delete nineteen compact
fraction entries — the ones that turn `1/2` into `½` — because a braille or large print book must
keep `1/2` as it was typed. But because the list is shared, those entries came off **your own
letters too, and stayed off**. A transcriber on Office 365 worked the mechanism out for herself and
reported it on that date.

**Each kind of document now has its own list**, in the three files above. Opening a book loads the
book's list; going back to a letter loads yours. Your own entries are not touched by anything a
book does.

Two details:

- **Formatted entries are never touched.** An AutoCorrect entry carrying bold, a picture or special
  characters lives inside `Normal.dotm` and cannot be read back as plain text, so those stay shared
  across all three. They were never the problem.
- **Switching takes about a second.** Measured on a list of 926 entries: only the differences are
  applied, not the whole list.

If the *"AutoCorrect list in use:"* line ever reads **not recorded**, VistaType LP could not reach
the settings folder, and the old shared-list behavior is what you have. That line exists precisely
so a machine in that state says so instead of quietly stripping your entries.

---

## The first time, and what to expect

VistaType LP has to learn your preferences from somewhere, and all it can do is look at how
Word is set the first time it runs.

So whatever your checkboxes happen to say at that moment is what it learns. If grammar checking
was switched off that day, it will believe you want it off.

**The fix takes ten seconds and you do it once.** Open an ordinary document — a letter, not a
book — set the checkboxes the way you like them, and carry on. From that moment they are yours,
and they will keep coming back.

---

## What to do if you've changed typing settings and need to change back

Even while you are in the middle of editing a document, you can reset to the default settings for that document type
by using the **Reset Word Configuration** icon on the Quick Access Toolbar. The icon is a checkmark and is located at the right end of the QAT.

Keep in mind that if you are in an ordinary document, the typing settings will revert to the typing settings of a newly installed version of Word

---

## Things you might notice, and what they mean

**"A checkbox changed and I did not change it."**
Look at which document is in front of you. In a large print or a braille file some of them are
supposed to be different. Click into an ordinary document and look again.

**"Fractions stopped working in my large print book."**
Correct, and deliberate. A compact fraction is smaller than your base font size and cannot be
in a large print document. Type `1/2` and leave it as `1/2`.

**"I pasted a chapter into my large print book and there are ½ signs in it."**
Pasting brings the characters with it. Run **Full File Cleanup → Fix Common File Errors** and
they will be turned back into `1/2`. Attaching the large print template does it too.

**"Grammar checking is off and I did not turn it off."**
You have been working in a braille file. Click into any ordinary document and it comes back. If
it does not, switch it on there once and it will stick.

**"Word capitalizes my sentences in a letter but not in my braille file."**
That is right. Braille needs it off; your letters need it on. It changes as you move between
them.

**"Document Settings says the wrong thing."**
Click into the document you actually want, then check again — it follows whichever document
window is on top. If it still looks wrong, close Word and reopen it.

**"I want to start over."**
Close Word. Delete the folder `%AppData%\VistaType LP Settings`. Open Word, open an ordinary
document, and set your checkboxes the way you want them. VistaType LP learns them again from
there.

---

## The one thing it cannot do

First, the part that is stronger than it sounds. **A setting you change while a large print or
braille file is on the screen stays with that file.** It is never taken to be a change to how you
want Word to behave. Only a change you make in an ordinary document counts as yours, and that one
is remembered from then on. VistaType LP does not learn anything from a book.

**If you close Word while a large print file or a braille file is still open, Word can carry that
file's settings over to the next time it starts.** Can, not will, and the difference is worth
knowing. It happens when Word saves its Normal template on the way out, which it does silently
whenever something in the session has changed that template. Changing these particular boxes does
not change the template by itself — but a good deal of ordinary work does.

Measured on the build box, 9/12/2026, with six settings flipped and Word restarted:

| | what came back |
|---|---|
| Normal template not saved | five of the six went back to your own values |
| Normal template saved | all six carried over |

One of them, *ignore words in UPPERCASE*, carries over every time. Word writes that one down the
moment it changes and does not wait for anything.

All of that is Word's own behavior. None of it is VistaType LP.

VistaType LP is ready for it either way. It writes down that a book was in force, and that note
outlives the session — so next morning it knows the values Word came up holding may be the book's
rather than yours, and the first ordinary document you open puts your own settings back.

**There is one narrow gap left.** It needs three things in a row: you quit Word with a book on the
screen, you then install a version of VistaType LP that changes the shape of the settings file,
and you open an ordinary document before anything else. In that one sequence the book's values can
be written down as yours. Only a few upgrades change the settings file — most do not touch it —
so this is rare, but it has not been closed and it is honest to say so.

**Ten seconds avoids it entirely: before you close Word for the day, click into an ordinary
document first.** That is the only thing on this page that asks anything of you, and it is a
good habit regardless.

---

## The short version

- Word has one set of typing settings for everything you do but VistaType LP handles them differently.

- Braille, large print, and ordinary documents need different typing settings, so VistaType LP switches them to match whichever  document you are looking at.
- For ordinary documents, VistaType LP remembers your settings and uses them the next time you create or open one.
- For braille and large print you can change the settings and they will remain in effect until the document is closed. Any new or opened braille 
or large print documents will revert to the initial settings.
- Typing settings always switch to match the type of document you are working on.
- Typing settings settings can be reset for the document type even when you are looking at it on the screen.
- Your settings live in a file that survives uninstalling and can move to a new computer.

