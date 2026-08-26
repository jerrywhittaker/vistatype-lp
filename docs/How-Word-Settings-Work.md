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
one behaves correctly according their AutoCorrect option settings. You will not notice it happening.

---

### Which one is in force?

The **Document Settings** macro on the Quick Access Toolbar (looks like and circle with the letter "i" in it) tells you what the settings for the current are. It has a line reading *"Word is configured for..."*
— large print, braille, or default (ordinary document) settings. If Word is ever behaving oddly, that line is the
first thing to look at.

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
along with eight more about spelling, grammar and the the display of the Styles pane. **Forty-two in all.**

The **AutoFormat** command is the odd one out. It reformats a whole document in one go, and while
Word does not put it on the ribbon, VistaType has placed it on the full Vista Type LP Quick Access Toolbar... 
it looks like a little rectangle with a lightening bolt. If you let the installer put VistaType's 
own Quick Access then you will see it. You can generally ignore its use... it's there for the power-users.  — nothing runs it on its own.

---

### It learns from you when you use an ordinary document

Here is the part most people would not expect.

You can do far more in Word than just editing braille or large print documents, so, when you are working in an ordinary document,
and when you change one or more of these settings, VistaType LP remembers the changes. When you create or open an
ordinary document those changed settings will be used. Keep in mind that the settings are not part of the document and 
changes in the settings from one edited document to another may not look the same when they wee created with different typing settings.

**You never have to say "save this".** Change a checkbox in any of the settings and it is
yours until you change it again.

---

## What actually differs between braille and large print

Most of the AutoCoreect Options are set the same way for both. Only a few things genuinely differ, and
every one of them is about what the reader receives:

| When you type | In a large print book | In a braille file |
|---|---|---|
| `1/2` | stays `1/2` | becomes `½` |
| a web or network address | stays plain text | becomes a clickable link |

And if you run the **AutoFormat** command on a whole large print document, `1st` is left alone but in a large
braille document ist is raised to `1ˢᵗ`.

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

## What do do if you've changed typing settings and need change back to the default settings

Even while you are in the middle of editing a document, you can reset the the default settings for that document type
by using the Reset Document Settings icon on the Quick Access Toolbar. The icon is a checkmark and is located at the right end of the QAT.

Keep in mind that if you are in an ordinary document, the typing settings will revert to the typing settings of a newley installed version of Word

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

**If you close Word while a large print file or a braille file is still open**, Word saves that document's
settings as its own on the way out, and loads them again next time it starts.

VistaType LP handles this in the ordinary case: it remembers between sessions that a large print or braille document was
in force, and does not mistake the book's settings for yours.

There is one narrow gap it cannot cover — closing Word inside a braille file **and then
installing a new version of VistaType LP** before you next open an ordinary document. In that
one sequence your settings can be replaced by the braille ones.

**Ten seconds avoids it entirely: before you close Word for the day, click into an ordinary
document first.** That is the only thing on this page that asks anything of you, and it is a
good habit regardless.

---

## The short version

- Word has one set of typing settings for everything you do but VistaType LP handles them differently.

- Braille, large print, and ordinary documents need different typing settings, so VistaType LP switches them to match whichever  document you are looking at.
- For ordinary documents, VistaType LP remembers your settings and uses them the next time you create or open one.
- For braille and large print you can chane the settings and they will remain in effect until the document is closed. Any new or opened braille 
or large print documents will revert to the initial settings.
- Typing settings always switch to match the type of document you are working on.
- Typing settings settings can be reset for the document type even when you are looking at it on the screen.
- Your settings live in a file that survives uninstalling and can move to a new computer.

