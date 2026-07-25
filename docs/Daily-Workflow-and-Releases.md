# Daily Workflow and Releases

**A short guide for Jerry.** How day-to-day work happens, how a release gets built and
published, and what to type so Claude does the right thing.

You do not need to memorize any git commands. Claude runs them for you. This guide is so
you know *what is happening* and *what to say*.

---

## The idea in one picture

There are two branches — think of them as two shelves.

```
dev  ──●──●──●──●──●   <-- everything you're working on now
       \
master ●                <-- the last version you actually released
       │
     v3.0.5             <-- a tag: a permanent bookmark of exactly what shipped
```

- **`dev`** is your workbench. All edits, experiments, and fixes go here. It can be messy.
- **`master`** is the shelf holding **the last version you released to transcribers**. It
  only ever changes on release day.
- A **tag** (`v3.0.5`, `v3.0.6`, …) is a permanent bookmark. Tags are how you go back.

**Why bother?** So that if a new version turns out to be broken, the last good one is still
sitting there, untouched and clearly labeled — and you can put a transcriber back on it in
about a minute.

> **You are on `dev` essentially all the time.** Claude will never commit to `master` and
> will tell you if you somehow end up there.

---

## Day to day

### 1. Ask for the change

Just describe what you want in plain language. Claude edits the files under `src/`.

### 2. Look at what changed

Say **"show me what changed"** and Claude will show you the diff.

### 3. Commit it

Say **"commit that"**. Claude writes the commit message and commits to `dev`.

Commit often and in small pieces — one fix per commit. That is what makes it possible to
undo *one* thing later without losing everything else.

### 4. Push — only when you say so

Nothing goes to GitHub until you say the word **"push"**. Claude will not push on its own,
ever. Pushing just backs up your work to GitHub; it does **not** release anything.

> Pushing `dev` is safe and a good habit. It's your off-site backup.

---

## Building and testing

| You want to… | Type this in the terminal |
|---|---|
| Build the add-in from source | `make build` |
| Put the built add-in in the repo root | `make deploy` |
| Build the actual installer `.exe` | `make installer` |
| Read the compiled code (no Windows needed) | `make read` |

Or just ask Claude: **"build the installer"**.

`make installer` does everything — rebuilds the add-in from `src/` on the Windows box,
compiles `VistaType-LP-Setup-<version>.exe`, and drops it in `dist/` and on the VM Desktop.

### Testing it

You test from the **installer**, the way a transcriber would.

> ### ⚠️ Close Word completely first
>
> Word holds the add-in file locked while it's running. If you install over an open Word,
> the installer **silently does nothing** and you'll still be running the old version — the
> About box will show the old number and you'll think the build failed.
>
> **Close every Word window (and Outlook), then run the installer.**

---

## Cutting a release

A release is: move `master` forward, bookmark it with a tag, and publish the `.exe` on
GitHub. Claude walks you through it — you can simply say **"let's release 3.0.7"** — but
here is what's happening.

### Step 1 — Make sure `dev` is ready

Everything committed, and you're happy with the change-set.

### Step 2 — Bump the version number

It lives in **four** places that must all agree. Say **"bump the version to 3.0.7"** and
Claude updates all four:

1. `Makefile` — names the installer file
2. `installer/vistatype.iss` — what Windows shows in *Programs & Features*
3. The **About dialog** caption — both the LP and the Braille About boxes
4. `CLAUDE.md` — Claude's own notes

### Step 3 — Build the installer

`make installer`

### Step 4 — Install it and test it

Close Word → run the `.exe` → open Word → check the About box shows the new number → try
the things you changed.

### Step 5 — Say the word

When you're happy: **"looks good, release it and push"**.

Claude then:
- moves `master` forward to match `dev`
- creates the tag `v3.0.7`
- puts you back on `dev`
- pushes to GitHub
- **publishes the release with the `.exe` attached**
- confirms the `.exe` really is attached

### Step 6 — Send people the link

The GitHub release page has the installer on it. That link is what you give transcribers.

---

## ⚠️ The one rule that matters most

**A release must have the `Setup.exe` attached to it on GitHub. Always.**

GitHub automatically puts a *"Source code (zip)"* download on every release. That is **not**
the program — it's a folder of VBA text files, and nobody can install it. So a release
without the `.exe` looks perfectly official and gives people **nothing usable**.

There's a second reason. The add-in **cannot be rebuilt identically**. Word regenerates
part of the file every time it compiles, so building version 3.0.5 again next year will
*not* produce the same file you shipped. **The `.exe` you upload is the only real copy of
that release.** If it isn't on GitHub, it effectively doesn't exist.

Claude checks the `.exe` is there before and after publishing, every time.

---

## If something goes wrong

**A transcriber says the new version broke something.**
Send them the previous release's `.exe` from its GitHub release page and have them
reinstall (Word closed). That's the fastest fix — nothing to rebuild.

**You want the code back the way it was at a release.**
Say **"go back to what we shipped in 3.0.6"**.

**You need to fix a released version without shipping everything on `dev`.**
Say **"we need a hotfix on 3.0.6"**. Claude branches from the tag, fixes just that, and
releases it — your in-progress work on `dev` stays out of it.

**You undid too much / something looks lost.**
Say so. Git keeps almost everything for a long time, including things that look deleted.
Don't try to fix it by hand — just describe what happened.

---

## Talking to Claude — prompts that work

### Describe the *symptom*, not the fix

This is the single most useful habit. You know Word and the documents; let Claude find the
cause.

> ✅ *"Every time I open the Fill-In Line dialog, the Styles pane jumps back to Recommended
> and loses my alphabetical sort."*

That one sentence led to finding the same redundant call copy-pasted in 14 dialogs. A
request like *"add a check to the styles pane code"* would have patched one symptom.

### Say which side you're on

There are two tabs and a lot of near-identical macros. Naming it saves a round trip.

> ✅ *"On the **Braille** tab, in the File Cleanup group…"*
> ✅ *"This is a **large print** document with the LP template attached."*

### Point at the button or the macro when you know it

> ✅ *"The `Dx_Remove_Bullets` macro"*
> ✅ *"The button labeled 'Prodnote → TN'"*

### Ask before you commit to a direction

> ✅ *"Why is this even being called here?"*
> ✅ *"Is this needed at all, or can it just come out?"*

Often the right fix is deleting something rather than adding a condition around it.

### Real examples from past work

> *"Add a Delete Prodnotes button to the Braille Macros File Cleanup group."*
>
> *"The Word workspace goes black behind the braille template form — fix it."*
>
> *"The DAISY/NIMAS conversion takes forever. Can it be sped up?"*
>
> *"Shorten that ribbon label, it's too wide."*
>
> *"Bump the version and build me an installer."*

### Checking on things

> *"What branch am I on?"*
> *"What's on `dev` that hasn't been released yet?"*
> *"Show me the last few commits."*
> *"What changed since 3.0.5?"*

### Controlling what happens

> *"Commit that."*
> *"Don't push yet."* / *"Push."*
> *"Let's release 3.0.7."*
> *"Explain what you're about to do before you do it."*

---

## Things Claude will always / never do

**Always**
- Work on `dev`
- Show you what changed before committing, if you ask
- Attach the `.exe` when publishing a release, and verify it's attached
- Walk you through release steps one at a time

**Never**
- Commit to `master`
- Push to GitHub unless you say "push"
- Delete a version tag
- Force-push or rewrite published history

If Claude ever seems about to break one of these, stop it and say so — something has gone
wrong.

---

## Handy things to know

- **Run a command yourself inside Claude Code** by starting the line with `!` —
  for example `!git status` or `!make installer`. The output goes straight into the
  conversation so Claude can see it too.
- **Hand-editing a form's layout** (dragging controls around in the Word VBA editor) is the
  one thing Claude can't do for you. Ask **"remind me how to round-trip a form edit"** and
  you'll get the steps — including the gotcha that you must open the `.dotm` in the STARTUP
  folder *directly*, not the loaded add-in.
- **Text-only changes to a form** (a caption, a version label) Claude *can* do without you
  touching the VBA editor.

---

## Quick reference

| Situation | What you type |
|---|---|
| Start a change | Describe it in plain language |
| See the changes | *"show me what changed"* |
| Save the work | *"commit that"* |
| Back it up to GitHub | *"push"* |
| Build the installer | *"build the installer"* or `make installer` |
| Ship it | *"looks good, release it and push"* |
| Undo a bad release | *"go back to what we shipped in 3.0.6"* |
| Fix a shipped version | *"we need a hotfix on 3.0.6"* |
| Find out where you are | *"what branch am I on?"* |

---

*The technical version of all this — exact commands, why fast-forward-only merges, the
hotfix procedure — lives in `DEVELOPMENT.md` and `CLAUDE.md` at the top of the repo.*
