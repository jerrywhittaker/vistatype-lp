# How VistaType LP Handles Word's Settings

**A guide for transcribers.** You do not have to do anything described here — except one
ten-second habit at the very end, which is worth forming. This explains what VistaType LP is
doing behind the scenes, so that if you ever notice Word behaving differently in one document
than in another, you know why, and you know it is on purpose.

> **A note on versions.** This describes how VistaType LP behaves from **August 2026** onward.
> If you are running something older, the sections about your settings being written down and
> given back will not match what you see — older versions did not do it.

---

## The one fact everything else rests on

**Word has only one set of typing settings, for the whole program.**

Open **File → Options → Proofing → AutoCorrect Options** and you are looking at settings that
belong to Word itself, not to the document in front of you. Change smart quotes while writing a
braille file, and you have changed them for every document you will ever open — your books,
your letters, your grocery list.

That is how Microsoft built Word. Everything below follows from it.

---

## Why that is a problem for this work

Braille and large print need Word to type differently. Sometimes they need opposite things.

Type `1/2` in a **braille** file and it should become **½**, because that single character is
what the Duxbury translator turns into correct braille.

Type `1/2` in a **large print** book and it must stay exactly as you typed it: `1/2`. A compact
fraction is one character, and its digits are drawn much smaller than the rest of your text. In
an 18-point book, that fraction is not 18 point. A reader who needs large print cannot read it.

Neither of those is a preference. One is what the braille translator requires; the other is
what your reader requires. And Word can only hold one answer at a time.

---

## What VistaType LP does about it

**It follows the document you are actually looking at.**

Click into a braille file, and Word is set up for braille. Click into a large print book, and
Word is set up for large print. Click into an ordinary letter, and Word goes back to your own
everyday settings.

It happens the moment you click, without asking you, and **it does not touch your screen** —
your Styles pane, your rulers, your formatting marks and your view stay exactly as you arranged
them. Only the typing behavior changes.

(Opening or creating a document is different: that is the moment VistaType LP sets the view up
for the kind of document it is. After that, the screen is yours.)

So with a book and a letter open side by side, you can move between them all afternoon and each
one behaves correctly. You will not notice it happening.

### Which one is in force right now

**Doc Info** on either ribbon tab tells you. It has a line reading *"Word is configured for..."*
— large print, braille, or default settings. If Word is ever behaving oddly, that line is the
first thing to look at.

---

## The settings this is about

In that same window — **File → Options → Proofing → AutoCorrect Options** — there are several
tabs across the top. Three of them matter here:

| Tab | What it decides |
|---|---|
| **AutoCorrect** | Capital letters, and the list that turns `teh` into `the` |
| **AutoFormat As You Type** | What Word changes **while you are typing** |
| **AutoFormat** | What Word changes when you run the **AutoFormat** command on a whole document |

Every checkbox on those three tabs is one VistaType LP looks after — **thirty-four of them** —
along with eight more about spelling, grammar and the Styles pane. **Forty-two in all.**

The **AutoFormat** command is the odd one out. It reformats a whole document in one go, and
Word does not put it on the ribbon. If you let the installer put VistaType's Quick Access
Toolbar in place, its button is on the small toolbar at the very top of the window. If you have
never used it, you can ignore that tab entirely — nothing runs it on its own.

---

## Your settings are written down, and given back

This is the part that changed in 2026, and it is the important one.

**VistaType LP writes your settings down before it borrows them, and puts them back
afterward.**

That is the whole idea. Everything else is detail.

Think of a neighbor borrowing a chair. Before the chair leaves the room, someone notes where it
was standing. When it comes back it goes to that exact spot — not to the middle of the room,
and not wherever the neighbor happened to leave it.

### When it writes yours down

- **When Word starts up**, before anything else has happened.
- **The moment before you open a book or a braille file** — while the settings are still yours.

### When it gives them back

- **The moment you click into an ordinary document** — a letter, a list, anything that is not a
  large print book or a braille file.

You do not have to do anything, nothing appears on screen, and it takes no noticeable time.

### It learns from you

Here is the part most people would not expect.

You do far more in Word than braille and large print, and you do not stop being an ordinary
Word user just because you have a book open. So if you change one of these settings **while you
are working inside a book**, VistaType LP treats that as a real preference, and it follows you
out to your letters.

It can tell, because it also writes down what the book asked for. When you return to an
ordinary document, anything that no longer matches what the book set must have been changed by
you — so it becomes your setting from then on, everywhere.

**You never have to say "save this".** Change a checkbox anywhere, in any document, and it is
yours.

### What this replaced

Until 2026, VistaType LP set a long list of these the same way for every kind of document. It
was not switching them to suit your work — it was simply overwriting what you had chosen, and
every time you opened a plain document your choices went back to VistaType's.

**What protects you now is not that it leaves them alone.** Opening a book still changes most
of those forty-two settings, because it has to. What protects you is that yours are written
down first and handed back the moment you return to ordinary work.

---

## What actually differs between braille and large print

Most of those three tabs is set the same way for both. Only a few things genuinely differ, and
every one of them is about what the reader receives:

| When you type | In a large print book | In a braille file |
|---|---|---|
| `1/2` | stays `1/2` | becomes `½` |
| a web or network address | stays plain text | becomes a clickable link |

And if you run the **AutoFormat** command on a whole document, `1st` is left alone in a large
print book and raised to `1ˢᵗ` in a braille file.

None of those are your preference to set, and that is not a slight. They decide what the
Duxbury translator receives, and what a low-vision reader sees on the page. They belong to the
reader, not to any of us.

---

## Spelling and grammar

Braille files switch off **grammar checking as you type** and the **contextual spelling
checker**, and tell Word to stop ignoring numbers mixed into words. That is correct: the
grammar checker objects endlessly to braille formatting and would drive you to distraction.

**What was wrong until 2026 is that nothing ever switched them back on.** Open one braille file
and grammar checking was off — in that file, in your letters, for the rest of the day, and
every day after, because Word remembers it. Nothing connected it to VistaType LP, and most
people never worked out why their grammar checking had stopped.

Now VistaType LP notes how you have them set, and puts them back the moment you return to an
ordinary document.

Both books also switch on **"Set left- and first-indent with tabs and backspaces"**, so the Tab
and Backspace keys set your indents while you work. If you prefer it off, switch it off in an
ordinary document and that is what you will get back.

---

## Where your settings are kept

In a small file, in your own Windows profile:

```
%AppData%\VistaType LP Settings\VistaType.ini
```

Three things worth knowing about it:

- **Uninstalling VistaType LP does not delete it.** If you reinstall, or install a new version,
  your settings are still there. That is deliberate.
- **It can be copied to a new computer.** If you get a new machine, copy that folder across and
  your settings come with you.
- **You never need to open it**, and nothing is lost if you delete it — VistaType LP simply
  starts learning your settings again.

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

## Things you might notice, and what they mean

**"A checkbox changed and I did not change it."**
Look at which document is in front of you. In a book or a braille file some of them are
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

**"Doc Info says the wrong thing."**
Click into the document you actually want, then check again — it follows whichever document
window is on top. If it still looks wrong, close Word and reopen it.

**"I want to start over."**
Close Word. Delete the folder `%AppData%\VistaType LP Settings`. Open Word, open an ordinary
document, and set your checkboxes the way you want them. VistaType LP learns them again from
there.

---

## The one thing it cannot do

**If you close Word while a book or a braille file is still open**, Word saves that document's
settings as its own on the way out, and loads them again next time it starts.

VistaType LP handles this in the ordinary case: it remembers between sessions that a book was
in force, and does not mistake the book's settings for yours.

There is one narrow gap it cannot cover — closing Word inside a braille file **and then
installing a new version of VistaType LP** before you next open an ordinary document. In that
one sequence your settings can be replaced by the braille ones.

**Ten seconds avoids it entirely: before you close Word for the day, click into an ordinary
document first.** That is the only thing on this page that asks anything of you, and it is a
good habit regardless.

---

## The short version

- Word has one set of typing settings for everything you do.
- Braille and large print need different ones, so VistaType LP switches them to match whichever
  document you are looking at.
- It writes your own settings down first, and gives them back the moment you return to an
  ordinary document.
- Change a setting anywhere — even inside a book — and it becomes yours from then on.
- The settings it switches while you are in a book belong to your reader, not to you.
- Your settings live in a file that survives uninstalling and can move to a new computer.
- Before closing Word, click into an ordinary document. That is the only habit worth forming.
