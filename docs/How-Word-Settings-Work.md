# How VistaType LP Handles Word's Settings

**A guide for transcribers.** You do not have to do anything described here. This explains
what VistaType LP is doing behind the scenes, so that if you ever notice Word behaving
differently in one document than in another, you know why — and you know it is on purpose.

> **Version note.** This describes how VistaType LP behaves from **version _____ onward**.
> If you are running an older one, the section *What is yours, and stays yours* and the
> section on *Spelling and grammar* will not match what you see.

---

## The one fact everything else rests on

**Word has only one set of typing settings, for the whole program.**

Open **File → Options → Proofing → AutoCorrect Options** and you are looking at settings
that belong to Word itself, not to the document in front of you. Change smart quotes while
writing a braille file, and you have changed them for every document you will ever open —
your books, your letters, your grocery list.

That is how Microsoft built Word. Everything below follows from it.

---

## Why that is a problem for this work

Braille and large print need Word to type differently. Sometimes they need opposite things.

Type `1/2` in a **braille** file and it should become **½**, because that single character
is what the Duxbury translator turns into correct braille.

Type `1/2` in a **large print** book and it must stay exactly as you typed it: `1/2`. A
compact fraction is one character, and its digits are drawn much smaller than the rest of
your text. In an 18-point book, that fraction is not 18-point. A reader who needs large
print cannot read it.

Neither of those is a preference. One is what the braille translator requires; the other is
what your reader requires. And Word can only hold one answer at a time.

---

## What VistaType LP does about it

**It follows the document you are actually looking at.**

Click into a braille file, and Word is set up for braille. Click into a large print book,
and Word is set up for large print. Click into an ordinary letter, and Word goes back to
your own everyday settings.

It happens the moment you click, without asking you, and **it does not touch your screen** —
your Styles pane, your rulers, your formatting marks and your view stay exactly as you
arranged them. Only the typing behavior changes.

(Opening or creating a document is different: that is the moment VistaType LP sets the view
up for the kind of document it is. After that, the screen is yours.)

So with a book and a letter open side by side, you can move between them all afternoon and
each one behaves correctly. You will not notice it happening.

### Which one is in force right now

**Doc Info** on either ribbon tab tells you. It has a line reading *"Word is configured
for..."* — large print, braille, or default settings. If Word is ever behaving oddly, that
line is the first thing to look at.

---

## What is yours, and stays yours

This is the part that changed in 2026, and it is worth knowing.

VistaType LP used to set fifteen settings the same way for every kind of document — which
meant it was not switching them at all, only overwriting whatever you had chosen. Every time
you opened a plain document, your choices went back to VistaType's.

**It no longer touches any of these. They are yours, in every document, permanently:**

- Replace text as you type
- Correct TWo INitial CApitals
- Capitalize names of days
- Correct accidental use of the cAPS LOCK key
- Straight quotes with smart quotes
- Internet and network paths with hyperlinks
- Automatically use suggestions from the spelling checker
- Show AutoCorrect Options buttons
- Built-in heading styles, and define styles based on your formatting
- \*Bold\* and \_italic\_ with real formatting

Set any of those the way you like it. VistaType LP will never change it back, in any
document, in any session, ever again.

---

## What it does still switch, and why

These are the settings where braille and large print genuinely disagree, so they have to
keep changing as you move between documents:

- Whether typing `1/2` gives you `½`
- Whether typing `1st` gives you a raised `st`
- Whether typing `(c)` gives you a copyright sign
- Whether Word capitalizes the first letter of a sentence, and of a table cell
- Whether Word invents bulleted lists, numbered lists, tables and borders as you type
- Whether the Tab and Backspace keys set your left and first-line indents (off in braille)

None of those are your preference to set, and that is not a slight. They decide what the
Duxbury translator receives, and what a low-vision reader sees on the page. They belong to
the reader, not to any of us.

---

## Spelling and grammar

Braille files switch off grammar checking as you type, the contextual spelling checker, a few
related things, and the Tab-and-Backspace indent setting. That is correct: the grammar checker flags braille formatting endlessly
and would drive you to distraction.

**What was wrong until 2026 is that nothing ever switched them back on.** Open one braille
file and grammar checking was off — in that file, in your letters, for the rest of the day,
and every day after, because Word remembers it. Nothing connected it to VistaType LP, and
most people never worked out why their grammar checking had stopped.

Now VistaType LP notes how you have them set, and puts them back the moment you return to an
ordinary document.

**One thing to expect the first time.** VistaType LP has to learn your preference from
somewhere, and all it can do is look at how Word is set the first time it runs. If grammar
checking happened to be switched off at that moment, that is what it will learn. Switch it
back on in an ordinary document once, and it will hold from then on.

---

## Where your settings are kept

In a small file, in your own Windows profile:

```
%AppData%\VistaType LP Settings\VistaType.ini
```

Two things worth knowing about it:

- **Uninstalling VistaType LP does not delete it.** If you reinstall, or install a new
  version, your settings are still there.
- **It can be copied to a new computer.** If you get a new machine, copy that folder across
  and your settings come with you.

You never need to open it, and nothing is lost if you delete it — VistaType LP simply starts
learning your settings again.

---

## Things you might notice, and what they mean

**"Fractions stopped working in my large print book."**
Correct, and deliberate. A compact fraction is smaller than your base font size and cannot
be in a large print document. Type `1/2` and leave it as `1/2`.

**"I pasted a chapter into my large print book and there are ½ signs in it."**
Pasting brings the characters with it. Run **Full File Cleanup → Fix Common File Errors**
and they will be turned back into `1/2`. Attaching the large print template does it too.

**"Grammar checking is off and I did not turn it off."**
You have been working in a braille file. Click into any ordinary document and it comes back.
If it does not, switch it on there once and it will stick.

**"Word capitalizes my sentences in a letter but not in my braille file."**
That is right. Braille needs it off; your letters need it on. It changes as you move between
them.

**"Doc Info says the wrong thing."**
Click into the document you actually want, then check again — it follows whichever document
window is on top. If it still looks wrong, close Word and reopen it.

---

## The short version

- Word has one set of typing settings for everything.
- Braille and large print need different ones, so VistaType LP switches them to match
  whichever document you are looking at.
- The settings that are genuinely yours are never touched.
- The ones it does switch belong to your reader, not to you.
- Your own settings are kept in a file that survives uninstalling and can move to a new
  computer.
