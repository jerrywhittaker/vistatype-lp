# VistaType LP -- Linux-side build orchestration.
#
# Source of truth is the text tree under src/. The .dotm cannot be compiled on
# Linux, so `make build` and `make pull` drive a remote Windows+Word box over SSH
# (see build.config). You stay in this terminal; Word is invisible infrastructure.
#
#   make pull    Export VBA from LPandBRL.dotm INTO src/  (canonical; run once to seed,
#                or after anyone edits in the Word VBE). Overwrites src/vba + src/forms.
#   make build   Import src/ into dist/LPandBRL.dotm via Word, then embed the ribbon
#                (customUI14.xml) and stage the .dotx with its keyboard shortcuts
#                (src/keymap/). Smoke-test before deploying.
#   make ribbon  Regenerate src/ribbon/customUI14.xml from the legacy Word.officeUI.
#   make qat     Check the curated toolbar (installer/qat-template.officeUI) against the
#                ribbon's hidden QAT tab, and regenerate installer/qat-icons-only.officeUI.
#                Runs as part of `make build`; a mismatch stops the build.
#   make read    Refresh reference/ from LPandBRL.dotm using the Linux-only decompressor
#                (read/diff aid; does NOT need Windows and is NOT import-ready).
#   make try     Bump, build, and put the new add-in straight into Word's STARTUP folder on
#                the build box - no installer, no wizard. The fastest loop for VBA and dialog
#                work: change code, `make try`, open Word. Word must be CLOSED on the box.
#                Stops FIRST if a form was hand-edited on the box and not pulled into src/ yet,
#                because the copy at the end would silently go over it.
#                Says afterwards what it cannot test, and when a real installer is needed.
#   make deploy  Promote dist/LPandBRL.dotm (embedded ribbon) to the repo root.
#   make branding  Regenerate the installer's icon and wizard artwork (installer/branding/)
#                from assets/branding/. Only needed after the artwork changes; `make installer`
#                refuses to run if the generated files are missing.
#   make stage   Copy the license into dist/ alongside the built shipping files.
#   make installer  Compile the Inno Setup installer on the Windows box; copies the
#                Setup.exe back to dist/. Ships the .dotm, the .dotx and the GPL.
#   make font-installer  Build the standalone typeface installer (and a .zip of the fonts for
#                anyone whose antivirus eats unsigned installers), and put both on the build
#                box Desktop. Not part of `make installer`; build it when a tester needs the
#                font on its own. See installer/vistatype-font-only.iss.
#   make scan    Upload the newest dist/ Setup.exe to VirusTotal and report which of its
#                ~70 engines flag it, and as what. Fails if Microsoft flags it (that is
#                Defender, which is what a transcriber has) or if more than 3 do. Needs
#                VT_API_KEY in build.config. NOTE: uploads are PUBLIC. See docs/Code-Signing.md.
#   make vba-test  Run the VBA unit tests (tests/vba) on the build box, invisibly, against a
#                throwaway document. Nothing is installed and nothing is left behind. Word must
#                be CLOSED on the box. It tests the SOURCE in src/vba, not the last build.
#   make vm-up   Start the Windows build VM if it is not already up and answering. Runs as
#                part of every target that touches the box, so you should never have to think
#                about it. It always comes up WITH A WINDOW, so it is ready to be driven by
#                hand - which is how every setup here gets tested.
#   make test    Run the test suite in tests/ over the build's own checkers and file
#                generators, against fixtures. No Word, no build box, about half a second.
#                Needs pytest:  sudo apt install python3-pytest
#   make hooks   Point git at .githooks so the pre-push secrets check runs. Once per clone.
#                It refuses a push that would publish a card PIN, a private key, a tracked
#                .pfx/.p12/.pem, or build.config - and runs gitleaks over the whole history,
#                which takes about 1.6 seconds. Better than GitHub's own scanning for
#                prevention: that alerts AFTER the secret is published.
#   make clean   Remove dist/ and build/ scratch.

-include build.config
WIN_PWSH ?= powershell
WIN_ISCC ?= "C:/Program Files (x86)/Inno Setup 6/ISCC.exe"
DOTM      := LPandBRL.dotm
DOTX      := LargePrintTemplate.dotx
RIBBON    := Word.officeUI
PROJNAME  := LPandBRL
APPVER    := 3.0.470
SETUP_EXE := VistaType LP and Braille Macros Setup $(APPVER).exe
VERDATE   := $(shell date +%-m/%-d/%Y)

# Everything ISCC reads out of dist/. VistaType LP bundled a typeface, VistaTypeLP Legible,
# from 3.0.101 to 3.0.196; four .ttf files and an OFL.txt travelled this road as well, and a
# `make fonts` target rescaled them from assets/fonts/. All of it went on 8/20/2026 - the face
# has no Greek, no IPA and not enough mathematics, and a character it has not got is filled in
# silently from somewhere else at some other size. See installer/vistatype.iss.
# The bundled typeface. Four faces because Word does not fall back inside a family - a bold
# character the Bold face lacks comes out of another typeface at another size - and three
# licenses because the face is Noto Sans with Noto Sans Math and Noto Sans Symbols folded in,
# and condition 2 of the OFL wants each one's text beside the font on disk.
FONTDIR   := assets/fonts/vistatypelp-sans
FONTFILES := VistaTypeLPSans-Regular.ttf VistaTypeLPSans-Bold.ttf \
             VistaTypeLPSans-Italic.ttf VistaTypeLPSans-BoldItalic.ttf \
             OFL.txt OFL-NotoSansMath.txt OFL-NotoSansSymbols.txt

SHIPFILES := $(DOTM) $(DOTX) LICENSE.txt $(FONTFILES)

SSH := ssh $(WIN_HOST)
WSCRIPTS := $(WIN_DIR)/tools/windows

.DEFAULT_GOAL := help

help:
	@sed -n 's/^#   //p' Makefile

check-config:
	@test -f build.config || { echo "ERROR: copy build.config.example -> build.config and edit it."; exit 1; }
	@test -n "$(WIN_HOST)" || { echo "ERROR: WIN_HOST not set in build.config."; exit 1; }

# --- push the working tree (source + scripts + shell .dotm) to the Windows box ---
# scp only ADDS and OVERWRITES - it never deletes. So a file removed from src/ stayed on the
# build box for ever, and Import-Vba.ps1, which imports whatever it finds there, put it back
# into every later .dotm. Three deleted UserForms had been shipping that way, one of them
# (Sh_License_Form) since 8/3/2026, and nothing said so: the build log lists what it REMOVES
# from the .dotm, not what it re-imports. Wipe the two pushed trees first so the box is always
# an exact copy of src/. Only src and tools go - dist/ holds the .dotm being built.
# Everything that touches the box comes through here, so this is where the VM gets started.
# Without it an off VM fails at the SSH with a connection error that says nothing about why.
# It does nothing at all when the box already answers - see the script.
vm-up:
	@python3 tools/lib/vm_up.py --host "$(WIN_HOST)" $(if $(VM_NAME),--name "$(VM_NAME)",) $(if $(VBOXMANAGE),--vboxmanage "$(VBOXMANAGE)",)

push-src: check-config vm-up check-dotm-unsigned
	$(SSH) "if not exist \"$(WIN_DIR)\" mkdir \"$(WIN_DIR)\""
	$(SSH) '$(WIN_PWSH) -NoProfile -Command "Remove-Item -Recurse -Force \"$(WIN_DIR)/src\", \"$(WIN_DIR)/tools\" -ErrorAction SilentlyContinue; exit 0"'
	scp -q -r src tools $(DOTM) "$(WIN_HOST):$(WIN_DIR)/"

# --- canonical export: Word writes .bas/.cls/.frm/.frx from the .dotm back to src/ ---
pull: check-config push-src
	$(SSH) '$(WIN_PWSH) -ExecutionPolicy Bypass -File $(WSCRIPTS)/Export-Vba.ps1 -Dotm "$(WIN_DIR)/$(DOTM)" -SrcRoot "$(WIN_DIR)/src"'
	rm -rf src/vba src/forms && mkdir -p src/vba src/forms
	scp -q -r "$(WIN_HOST):$(WIN_DIR)/src/vba/*"   src/vba/   || true
	scp -q -r "$(WIN_HOST):$(WIN_DIR)/src/forms/*" src/forms/ || true
	@echo "Pulled canonical VBA source into src/ (review with 'git diff')."

# Point git at the tracked hooks directory. Needed once per clone - .git/hooks is not
# version controlled, so the pre-push secrets check cannot live there and be shared.
hooks:
	@git config core.hooksPath .githooks
	@chmod +x .githooks/* 2>/dev/null || true
	@echo "git hooks -> .githooks (pre-push runs the secrets check)"
	@command -v gitleaks >/dev/null || echo "  NOTE: gitleaks is not installed - the history pass will be skipped."

# The repo-root .dotm is the base every build starts from, and `make deploy` is what puts a
# file there. Once signing is wired in, dist/ holds a SIGNED .dotm - and promoting that would
# seed every later build with a signature that stops matching the moment the modules are
# re-imported. Nothing else here would notice. See docs/Code-Signing.md.
check-dotm-unsigned:
	@python3 tools/lib/check_dotm_unsigned.py $(DOTM)

# --- build the shipping .dotm from src/ via Word, then embed the ribbon ---
# A .frm rewritten with LF endings still builds, but Word mis-parses the designer header and
# the form fails to compile on the USER's machine. Cheap to check, expensive to miss.
check-frm-eol:
	@python3 tools/lib/check_frm_eol.py

# A VBA line over 1023 characters does not fail the build - it HANGS it. Word chokes on the
# line silently, with no error and no visible dialog, and the ~$ lock file it leaves behind
# while stuck looks like the cause. Cost three hung builds on 7/30/2026 before the real
# reason turned up. Two seconds to check.
check-vba-lines:
	@python3 tools/lib/check_vba_line_length.py

# A UserForm's LAYOUT is edited by hand in the VBA editor, against the add-in Word actually
# loads on the build box - a .frx is binary and cannot be written from Linux. That edit lives in
# that ONE file until it is exported back into src/forms/, and the copy at the end of `make try`
# goes straight over it. The build then SUCCEEDS, says nothing, and the dialog quietly goes back
# to how it looked before; Word does not complain either, since the file it loads is merely
# newer. Nothing detected that until 8/26/2026, when Jerry asked what stops it.
#
# Compares by CONTENT, not by date: installing a real Setup.exe rewrites that file and its date
# with no edit involved, and a date test would cry wolf every time. Runs BEFORE the bump so a
# stop costs no version number. The strict form also refuses while Word is open, which `make try`
# needs anyway - saying so here saves a two-minute build. The soft form only warns about that,
# for `make installer`, which does not care whether Word is running.
check-startup-unpulled:
	@python3 tools/lib/check_startup_unpulled.py

check-startup-unpulled-soft:
	@python3 tools/lib/check_startup_unpulled.py --warn-only

# The toolbar's idQ="x1:btn_*" entries resolve against the hidden tab in customUI14.xml.
# If they drift apart the buttons render blank on the user's machine and nothing warns you,
# so the build refuses to proceed. Also regenerates the append-safe icons-only toolbar.
# Nothing in the build compiles VBA, so a form calling a macro that has been renamed or removed
# builds perfectly and fails on the transcriber's machine as "Compile error in hidden module:
# <form name>" - which names the form and not the missing macro. Cost an install on 8/3/2026.
check-form-calls:
	@python3 tools/lib/check_form_calls.py

# ActiveDocument.Styles("X") raises run-time error 5941 when the document has no such style,
# and a document need not: RefPageNemeth comes from the Nemeth braille templates, Print Pg Num
# from the LP template. The macro then dies where it stands, often before doing anything the
# user can see. Cost Jerry an AutoTag run on 8/5/2026. Every lookup must be guarded.
check-style-guards:
	@python3 tools/lib/check_style_guards.py

# Nothing in this build compiles VBA, so a structural mistake - a module-level Const or Dim
# written beside the procedure that uses it instead of in the declarations section at the top,
# a duplicated procedure name, a With with no End With - imports perfectly, ships, and fails on
# the transcriber's machine as "Compile error in hidden module", with no line number and nothing
# naming the cause. Cost a build on 8/18/2026.
check-vba-structure:
	@python3 tools/lib/check_vba_structure.py

# The VBA unit tests. Nothing on Linux can run VBA, so these go to the build box - but they do
# NOT need the add-in installed and do not touch Word's STARTUP folder. build_vba_test_bundle.py
# lifts the procedures the tests name out of src/vba (following what they call), and the runner
# imports that plus tests/vba into a throwaway invisible document and calls VtRunAllTests.
#
# So it tests the source you are looking at, and it can reach the Private helpers, which a test
# document that merely referenced the add-in could not call at all.
#
# Only pure helpers belong here. A procedure that reaches a Document, a Selection, a UserForm or
# a MsgBox HANGS an invisible Word - see tests/vba/README.md for the list of what cannot be
# tested this way and why.
vba-test: check-config push-src
	@python3 tools/lib/build_vba_test_bundle.py
	$(SSH) '$(WIN_PWSH) -NoProfile -Command "Remove-Item -Recurse -Force \"$(WIN_DIR)/vbatests\" -ErrorAction SilentlyContinue; exit 0"'
	scp -q -r tests/vba "$(WIN_HOST):$(WIN_DIR)/vbatests"
	scp -q build/vbatests/VtFunctionsUnderTest.bas "$(WIN_HOST):$(WIN_DIR)/vbatests/"
	$(SSH) '$(WIN_PWSH) -ExecutionPolicy Bypass -File $(WSCRIPTS)/Run-VbaTests.ps1 -TestRoot "$(WIN_DIR)/vbatests"'

# The checks above are the only thing between a VBA mistake and a transcriber's machine, and
# until 9/20/2026 nothing tested THEM. tests/ does, against fixtures built to look like the
# real source - plus the file generators, whose output goes into the user's own Word.officeUI.
#
# Deliberately NOT part of `make build`: it needs pytest, which a fresh clone may not have, and
# a build that fails for a missing test tool teaches people to skip the tool. Run it after
# changing anything under tools/lib.
test:
	@python3 -c "import pytest" 2>/dev/null || { \
	  echo "ERROR: pytest is not installed, so the tests cannot run."; \
	  echo "       sudo apt install python3-pytest"; exit 1; }
	@python3 -m pytest tests/ -q

# The ribbon dispatch table (src/vba/RibbonDispatch.bas), generated from customUI14.xml.
# MUST run before the VBA checks and before push-src, because it WRITES a file into src/vba.
#
# It exists because Word does not pass an error back out of Application.Run - it shows its own
# Run-time error dialog instead, and RibbonAction's handler is never entered. Measured on the
# build box 8/26/2026. A direct call propagates, so the buttons are dispatched through a
# generated Select Case. See the header of the script.
build-dispatch:
	@python3 tools/lib/build_ribbon_dispatch.py

check-qat:
	@python3 tools/lib/build_qat.py

# Regenerates the ribbon tabs in the form the installer writes into the user's own
# Word.officeUI (so Word lets them hide/reorder/rename them), and refuses to build if a
# button id that has already shipped has gone: every installed tab and toolbar references
# those by id and would render blank with no error.
check-tabs:
	@python3 tools/lib/build_ribbon_tabs.py

# The installer's icon and wizard artwork, generated from assets/branding/. `branding`
# regenerates; `check-branding` only refuses to build an installer when they are missing, so a
# machine without Pillow can still build everything else.
branding:
	@python3 tools/lib/build_branding.py

check-branding:
	@python3 tools/lib/build_branding.py --check-only

build: check-config test build-dispatch check-frm-eol check-vba-lines check-vba-structure check-form-calls check-style-guards check-qat check-tabs push-src
	$(SSH) '$(WIN_PWSH) -ExecutionPolicy Bypass -File $(WSCRIPTS)/Import-Vba.ps1 -Shell "$(WIN_DIR)/$(DOTM)" -SrcRoot "$(WIN_DIR)/src" -OutDotm "$(WIN_DIR)/dist/$(DOTM)" -ProjectName $(PROJNAME) -AppVer $(APPVER) -VerDate "$(VERDATE)"'
	@# The two About forms carry the version in their binary .frx, so bring them home if the
	@# stamp changed them. -u: only if newer, so an unchanged build copies nothing.
	@for f in Lp_About_Title_And_Agreement Dx_About_Title_And_Agreement; do \
	  scp -q "$(WIN_HOST):$(WIN_DIR)/src/forms/$$f.frm" src/forms/ 2>/dev/null || true; \
	  scp -q "$(WIN_HOST):$(WIN_DIR)/src/forms/$$f.frx" src/forms/ 2>/dev/null || true; \
	done
	@# Word inserts one more blank line at the top of a form's code section, and one more at
	@# the end, every single time it exports one. These two are re-exported on EVERY build, so
	@# the runs grow without limit - they had reached 50 and 43 lines by 8/1/2026 - and make a
	@# one-word caption change look like a real edit. Collapse them back to one each.
	@python3 tools/lib/trim_frm_blanks.py src/forms/Lp_About_Title_And_Agreement.frm src/forms/Dx_About_Title_And_Agreement.frm
	mkdir -p dist
	scp -q "$(WIN_HOST):$(WIN_DIR)/dist/$(DOTM)" dist/$(DOTM)
	python3 tools/lib/inject_customui.py dist/$(DOTM) src/ribbon/customUI14.xml
	@# The .dotx at the repo root is the shell; the keyboard shortcuts are text under
	@# src/keymap/ and get injected here, the same way the ribbon goes into the .dotm.
	cp $(DOTX) dist/$(DOTX)
	python3 tools/lib/inject_keymap.py dist/$(DOTX) src/keymap/lp-template-keymap.xml
	@echo "Built dist/$(DOTM) (embedded ribbon) + dist/$(DOTX) (keymap). Smoke-test in Word before deploying."

# --- regenerate customUI14.xml from the legacy Word.officeUI (one-off / reference) ---
ribbon:
	python3 tools/lib/officeui_to_customui.py $(RIBBON) src/ribbon/customUI14.xml

# --- QAT: check the curated toolbar against the ribbon, and regenerate the icons-only one ---
# installer/qat-template.officeUI stays HAND-MAINTAINED: its ordering, separators and the
# visible="false" entries that suppress Word's own default buttons are all deliberate and
# cannot be derived from the ribbon. Only its x1:btn_* entries are validated.
# installer/qat-icons-only.officeUI IS generated - VistaType's icons and nothing else, safe
# to append to a user's own toolbar. (tools/lib/extract_qat.py remains obsolete/unused.)
qat: check-qat

# --- Linux-only read helper (no Windows) ---
read:
	python3 tools/lib/decompress_vba.py

# --- put the freshly built add-in straight into Word on the build box, no installer ---
#
# The loop this shortens is: change code, build an installer, run the wizard, test. The .dotm
# copied here is the SAME FILE the installer packages, so this is not a lesser test of the code -
# it is the same code, without the wizard. What it does not carry is everything the installer puts
# somewhere ELSE, which check_try_scope.py reports at the end.
#
# IT BUMPS, and that is not tidiness. Every third number has meant "an installer you can install",
# and the About box is how you tell one build from the next - three builds shared 3.0.6 on
# 7/26/2026 and there was no way to know whether an install had taken. Swapping .dotm files
# without bumping brings that straight back. So a `make try` takes a number too; that number
# simply never gets a Setup.exe, which is what a private build counter is for.
#
# WORD MUST BE CLOSED ON THE BOX. Word holds the STARTUP .dotm open, so the copy would fail or,
# worse, half-succeed. Checked rather than assumed, by process name - a stray automation Word from
# a headless test run counts, and it is invisible on the desktop.
# Two steps, and the split is load-bearing - `make installer` is built the same way and for the
# same reason. Make expands $(APPVER) ONCE, when it reads this file, so a recipe that bumps and
# then builds in the same run stamps the OLD number into the About dialogs and announces the old
# number too. Caught by running it: the file said 3.0.243 and the build said 3.0.242. The sub-make
# re-reads the Makefile and picks up what bump just wrote.
try: check-startup-unpulled
	@$(MAKE) --no-print-directory bump
	@$(MAKE) --no-print-directory try-build

try-build: check-config build
	@echo "--- checking Word is closed on $(WIN_HOST) ---"
	$(SSH) '$(WIN_PWSH) -NoProfile -Command "$$w = @(Get-Process WINWORD -ErrorAction SilentlyContinue); if ($$w.Count -gt 0) { Write-Host \"ERROR: Word is running on the build box - close it and run make try again.\"; exit 1 }; exit 0"'
	@# The Linux copy, not the box's: the ribbon is injected on THIS side after the build, so
	@# $(WIN_DIR)/dist/$(DOTM) has no ribbon in it. Staged under a space-free name for the same
	@# quoting reason as the Setup.exe. Stale ~$$ lock files need no sweep here - Import-Vba.ps1
	@# has already cleared them during the build this target depends on.
	scp -q dist/$(DOTM) "$(WIN_HOST):$(WIN_DIR)/dist/startup-staged.dotm"
	$(SSH) '$(WIN_PWSH) -NoProfile -Command "Copy-Item \"$(WIN_DIR)/dist/startup-staged.dotm\" \"$$env:APPDATA\\Microsoft\\Word\\STARTUP\\$(DOTM)\" -Force; Remove-Item \"$(WIN_DIR)/dist/startup-staged.dotm\" -Force -ErrorAction SilentlyContinue"'
	@# Record what was just put there, so the next run can tell "the installer moved dist/ on"
	@# from "a form was edited on the box". See the script.
	@python3 tools/lib/check_startup_unpulled.py --stamp
	@echo ""
	@echo "$(DOTM) $(APPVER) is now in Word's STARTUP folder on $(WIN_HOST). Open Word there and test."
	@python3 tools/lib/check_try_scope.py

# --- the typeface on its own, for someone who needs the font and not the add-in ---
#
# Two things are produced, on purpose. The Setup.exe is what was asked for; the .zip is the answer
# to "it must not trip antivirus", which an UNSIGNED .exe can never actually promise - every build
# of the main installer was deleted on sight for two days in August 2026 by a machine-learning
# guess. A .zip of four font files has no program in it to judge, and Windows installs a font from
# right-click > Install for the current user, with no administrator rights.
#
# Neither needs `make build`: no add-in is involved. The four faces and three licenses come
# straight from assets/, which is the source of truth for them.
FONT_SETUP := VistaTypeLP-Sans-Font-Setup-1.1
FONT_ZIP   := VistaTypeLP-Sans-Fonts.zip
FONT_DESK  := VistaTypeLP Sans Font

font-installer: check-config check-branding
	mkdir -p dist
	@for f in $(FONTFILES); do \
	  test -f "$(FONTDIR)/$$f" || { echo "MISSING $(FONTDIR)/$$f"; exit 1; }; \
	  cp "$(FONTDIR)/$$f" dist/; \
	done
	python3 tools/lib/build_font_zip.py "dist/$(FONT_ZIP)"
	$(SSH) "if not exist \"$(WIN_DIR)\\dist\" mkdir \"$(WIN_DIR)\\dist\""
	@# Wiped and re-copied for the same reason installer-build does it: scp only adds and
	@# overwrites, so a file deleted here would linger there and keep being compiled in.
	$(SSH) '$(WIN_PWSH) -NoProfile -Command "Remove-Item -Recurse -Force \"$(WIN_DIR)/installer\" -ErrorAction SilentlyContinue; exit 0"'
	scp -q -r installer "$(WIN_HOST):$(WIN_DIR)/"
	scp -q $(addprefix dist/,$(FONTFILES)) "$(WIN_HOST):$(WIN_DIR)/dist/"
	$(SSH) '$(WIN_ISCC) "/DSrcDir=$(WIN_DIR)/dist" "$(WIN_DIR)/installer/vistatype-font-only.iss"'
	scp -q "$(WIN_HOST):$(WIN_DIR)/dist/$(FONT_SETUP).exe" "dist/$(FONT_SETUP).exe"
	scp -q "dist/$(FONT_ZIP)" "$(WIN_HOST):$(WIN_DIR)/dist/$(FONT_ZIP)"
	$(SSH) '$(WIN_PWSH) -NoProfile -Command "$$d = Join-Path ([Environment]::GetFolderPath(\"Desktop\")) \"$(FONT_DESK)\"; if (-not (Test-Path $$d)) { New-Item -ItemType Directory -Force -Path $$d | Out-Null }; Copy-Item \"$(WIN_DIR)/dist/$(FONT_SETUP).exe\" $$d -Force; Copy-Item \"$(WIN_DIR)/dist/$(FONT_ZIP)\" $$d -Force"'
	@echo ""
	@echo "Put both in the \"$(FONT_DESK)\" folder on the $(WIN_HOST) Desktop:"
	@echo "  $(FONT_SETUP).exe   - the installer (UNSIGNED: antivirus can still object)"
	@echo "  $(FONT_ZIP)         - the fonts to install by right-click (nothing to flag)"

deploy:
	@test -f dist/$(DOTM) || { echo "ERROR: run 'make build' first."; exit 1; }
	@# Check the file being PROMOTED, not the one already at the root. This is the single
	@# place a signed .dotm can get in, so stopping it here beats noticing afterwards.
	@python3 tools/lib/check_dotm_unsigned.py dist/$(DOTM)
	cp dist/$(DOTM) $(DOTM)
	@echo "Promoted dist/$(DOTM) (embedded ribbon) to repo root."

# --- stage the shipping files into dist/ for packaging ---
# `build` already produces dist/$(DOTM) (embedded ribbon) + dist/$(DOTX); the ribbon is
# embedded in the .dotm now, so only those two files ship (no Word.officeUI). Add the license.
stage: build
	cp LICENSE dist/LICENSE.txt
	@# The faces are BUILT by `make fonts` and tracked under assets/, not rebuilt here - the
	@# build reaches the network for the Noto releases, and a release must not depend on that.
	@for f in $(FONTFILES); do \
	  test -f "$(FONTDIR)/$$f" || { echo "MISSING $(FONTDIR)/$$f - run 'make fonts'"; exit 1; }; \
	  cp "$(FONTDIR)/$$f" dist/; \
	done
	@echo "Staged dist/$(DOTM) + dist/$(DOTX) + LICENSE.txt + 4 font faces for packaging."

# --- rebuild the bundled typeface from the Noto releases -------------------------------
# Reaches the network, so it is deliberately NOT part of `make installer`. Run it when a new
# Noto release is worth picking up; the built faces are tracked under assets/ and it is those
# that ship. Twelve checks run at the end and it refuses to claim success if any fail.
fonts:
	python3 tools/lib/build_vistatypelp_sans.py --out $(FONTDIR)
	@echo "Rebuilt $(FONTDIR). Commit the faces if they changed."

# --- compile the Inno Setup installer on the Windows box ---
# The finished Setup.exe is copied both back to local dist/ and into "VT Installer" on the
# build box's Desktop, so it's one double-click away when you test the install on the VM.
# The Desktop path is resolved on the box (GetFolderPath handles OneDrive-redirected Desktops).
#
# It goes in that FOLDER and not on the Desktop itself because Windows Defender deletes the
# Setup.exe within seconds of it landing - Trojan:Win32/Bearfoos.B!ml, a machine-learning
# guess, on an unsigned installer that drops a macro-enabled template into Word's STARTUP
# folder. "VT Installer" is a Defender exclusion on the build box, so builds survive there.
# Jerry, 8/13/2026. The exclusion is on that machine only; a transcriber downloading from
# GitHub has no such thing, and a code-signing certificate is the actual fix.
# Bump the third digit and write it to all four places the version must agree.
# Runs before every installer build so each Setup.exe is distinguishable (Jerry, 7/26/2026).
# The .frx captions are handled by Import-Vba.ps1 during the build itself.
bump: check-config
	@cur=$$(sed -n 's/^APPVER    := //p' Makefile); \
	 maj=$${cur%.*}; pat=$${cur##*.}; new="$$maj.$$((pat+1))"; \
	 sed -i "s/^APPVER    := .*/APPVER    := $$new/" Makefile; \
	 sed -i "s/^\( *\)#define AppVer      \".*\"/\1#define AppVer      \"$$new\"/" installer/vistatype.iss; \
	 echo "version $$cur -> $$new (Makefile + vistatype.iss; both About dialogs stamped during the build)"

# make installer bumps, then re-enters make so the recipe below sees the NEW APPVER.
installer: check-startup-unpulled-soft
	@$(MAKE) --no-print-directory bump
	@$(MAKE) --no-print-directory installer-build

installer-build: check-config check-branding stage vba-test
	$(SSH) "if not exist \"$(WIN_DIR)\" mkdir \"$(WIN_DIR)\""
	$(SSH) "if not exist \"$(WIN_DIR)\\dist\" mkdir \"$(WIN_DIR)\\dist\""
	@# Wipe installer/ on the box first, for the same reason push-src does it: scp only ADDS
	@# and OVERWRITES, so a file deleted here would linger there and keep being compiled into
	@# the Setup.exe. Nothing has been bitten by this yet, but the identical trap in push-src
	@# shipped three deleted UserForms for days without a word (see the note there).
	$(SSH) '$(WIN_PWSH) -NoProfile -Command "Remove-Item -Recurse -Force \"$(WIN_DIR)/installer\" -ErrorAction SilentlyContinue; exit 0"'
	scp -q -r installer "$(WIN_HOST):$(WIN_DIR)/"
	@# Send only what ISCC actually reads. Copying the whole dist/ meant re-uploading every
	@# previously built Setup.exe on every build - slow, and it FAILS outright if Windows is
	@# still holding one of them (a killed installer leaves the .exe locked at the OS level
	@# until reboot). The .exe is produced ON the box and copied back, so it never needs to
	@# travel in this direction.
	@# dist/ on the box is deliberately NOT wiped (see the note above), which is safe because
	@# vistatype.iss names every Source: file explicitly - a leftover we no longer ship is never
	@# compiled in, and one we DO ship is overwritten by this copy. Do not "fix" it by wiping.
	scp -q $(addprefix dist/,$(SHIPFILES)) "$(WIN_HOST):$(WIN_DIR)/dist/"
	$(SSH) '$(WIN_ISCC) "/DSrcDir=$(WIN_DIR)/dist" "/DAppVer=$(APPVER)" "$(WIN_DIR)/installer/vistatype.iss"'
	@# The installer name contains spaces. Quoting a spaced remote path through scp means
	@# satisfying BOTH the local shell and Windows cmd, which does not strip single quotes -
	@# they end up as part of the filename and the copy fails. Stage it under a space-free
	@# name instead and rename on arrival; PowerShell quoting we can rely on.
	$(SSH) '$(WIN_PWSH) -NoProfile -Command "Copy-Item \"$(WIN_DIR)/dist/$(SETUP_EXE)\" \"$(WIN_DIR)/dist/setup-staged.exe\" -Force"'
	scp -q "$(WIN_HOST):$(WIN_DIR)/dist/setup-staged.exe" "dist/$(SETUP_EXE)"
	$(SSH) '$(WIN_PWSH) -NoProfile -Command "Remove-Item \"$(WIN_DIR)/dist/setup-staged.exe\" -Force -ErrorAction SilentlyContinue"'
	$(SSH) '$(WIN_PWSH) -NoProfile -Command "$$d = Join-Path ([Environment]::GetFolderPath(\"Desktop\")) \"VT Installer\"; if (-not (Test-Path $$d)) { New-Item -ItemType Directory -Force -Path $$d | Out-Null }; Copy-Item \"$(WIN_DIR)/dist/$(SETUP_EXE)\" $$d -Force"'
	@echo "Built dist/$(SETUP_EXE)  (also copied to 'VT Installer' on the build box Desktop)"

# --- scan the built installer with about seventy antivirus engines ---
# Deliberately NOT part of `make installer`: the upload is public and permanent, so it has to
# be a decision, not a side effect. Kept separate from the guards that gate `make build` for
# the same reason - this one reaches the network and can be slow.
#
# The one that matters in the output is Microsoft: VirusTotal's other engines flagging an
# unsigned installer is ordinary background noise, but Defender is what a transcriber has.
# Even so, VirusTotal runs Defender WITHOUT its cloud, and the Bearfoos verdicts of 8/13-14
# were cloud verdicts, so a clean line there is not proof. Confirm on the build box with
#  "C:\Program Files\Windows Defender\MpCmdRun.exe" -Scan -ScanType 3 -File <path> -DisableRemediation
# (-DisableRemediation reports the verdict without deleting the file). docs/Code-Signing.md.
# SCAN_FILE names the build EXACTLY rather than letting the script pick the newest file by
# timestamp - a copy without -p, or a restore, would otherwise make it scan the wrong build and
# still look right. Quoted, because every installer name has spaces in it. To scan an older one:
#   make scan SCAN_FILE="dist/VistaType LP and Braille Macros Setup 3.0.135.exe"
SCAN_FILE ?= dist/$(SETUP_EXE)

scan:
	@python3 tools/lib/scan_virustotal.py "$(SCAN_FILE)" $(SCAN_ARGS)

clean:
	rm -rf dist build

.PHONY: help check-config push-src pull build build-dispatch ribbon qat check-qat check-tabs check-frm-eol check-vba-lines read try try-build check-startup-unpulled check-startup-unpulled-soft deploy stage branding check-branding bump font-installer installer installer-build scan test vba-test vm-up clean
