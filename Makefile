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
#   make deploy  Promote dist/LPandBRL.dotm (embedded ribbon) to the repo root.
#   make fonts   Regenerate the bundled typeface from the pristine upstream files. Only
#                needed after changing the scale factor or rescale_font.py — then check it
#                against a font ruler.
#   make stage   Copy the licenses and the bundled fonts into dist/ alongside the built
#                shipping files.
#   make installer  Compile the Inno Setup installer on the Windows box; copies the
#                Setup.exe back to dist/. Ships the .dotm, the .dotx, the GPL, and the
#                four VistaTypeLP Legible font files with their own license (OFL.txt).
#   make clean   Remove dist/ and build/ scratch.

-include build.config
WIN_PWSH ?= powershell
WIN_ISCC ?= "C:/Program Files (x86)/Inno Setup 6/ISCC.exe"
DOTM      := LPandBRL.dotm
DOTX      := LargePrintTemplate.dotx
RIBBON    := Word.officeUI
PROJNAME  := LPandBRL
APPVER    := 3.0.159
SETUP_EXE := VistaType LP and Braille Macros Setup $(APPVER).exe
VERDATE   := $(shell date +%-m/%-d/%Y)

# The bundled typeface. TrueType, NOT the .otf faces alongside them: Word embeds TrueType
# outlines and skips PostScript ones silently, so an .otf would save into a document without
# a word and set at the wrong size on any machine that has not got the font installed.
# See tools/lib/rescale_font.py and assets/fonts/atkinson-hyperlegible/README.md.
FONTSRC   := assets/fonts/atkinson-hyperlegible/scaled
# Must match LP_FONT_LEGIBLE in src/vba/LPandBrlMacros.bas and FontFamily in
# installer/vistatype.iss. The add-in looks the font up by this name to decide whether to
# offer it, so a mismatch grays the choice out on a machine that HAS it installed.
FONT_FAMILY := VistaTypeLP Legible
FONTS     := VistaTypeLPLegible-Regular.ttf VistaTypeLPLegible-Bold.ttf \
             VistaTypeLPLegible-Italic.ttf VistaTypeLPLegible-BoldItalic.ttf
# Everything ISCC reads out of dist/. The fonts and OFL.txt travel this road, and NOT by a
# hand-copy of assets/ to the build box: that folder is never wiped there, so a corrected
# font would keep shipping stale for ever. Same trap that shipped three deleted UserForms.
SHIPFILES := $(DOTM) $(DOTX) LICENSE.txt OFL.txt $(FONTS)

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
push-src: check-config
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

check-qat:
	@python3 tools/lib/build_qat.py

# Regenerates the ribbon tabs in the form the installer writes into the user's own
# Word.officeUI (so Word lets them hide/reorder/rename them), and refuses to build if a
# button id that has already shipped has gone: every installed tab and toolbar references
# those by id and would render blank with no error.
check-tabs:
	@python3 tools/lib/build_ribbon_tabs.py

build: check-config check-frm-eol check-vba-lines check-form-calls check-style-guards check-qat check-tabs push-src
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

deploy:
	@test -f dist/$(DOTM) || { echo "ERROR: run 'make build' first."; exit 1; }
	cp dist/$(DOTM) $(DOTM)
	@echo "Promoted dist/$(DOTM) (embedded ribbon) to repo root."

# --- stage the shipping files into dist/ for packaging ---
# `build` already produces dist/$(DOTM) (embedded ribbon) + dist/$(DOTX); the ribbon is
# embedded in the .dotm now, so only those two files ship (no Word.officeUI). Add the license.
stage: build
	cp LICENSE dist/LICENSE.txt
	@# The font is under the SIL Open Font License, not the GPL, and that license requires the
	@# text to travel with every copy of the font. Both go to the same place LICENSE.txt does.
	cp assets/fonts/atkinson-hyperlegible/OFL.txt dist/OFL.txt
	cp $(addprefix $(FONTSRC)/,$(FONTS)) dist/
	@echo "Staged dist/$(DOTM) + dist/$(DOTX) + LICENSE.txt + OFL.txt + $(words $(FONTS)) font files for packaging."

# --- compile the Inno Setup installer on the Windows box ---
# The finished Setup.exe is copied both back to local dist/ and onto the build box's
# Desktop, so it's one double-click away when you test the install on the VM. The Desktop
# path is resolved on the box (GetFolderPath handles OneDrive-redirected Desktops).
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
installer:
	@$(MAKE) --no-print-directory bump
	@$(MAKE) --no-print-directory installer-build

installer-build: check-config stage
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
	$(SSH) '$(WIN_PWSH) -NoProfile -Command "Copy-Item \"$(WIN_DIR)/dist/$(SETUP_EXE)\" ([Environment]::GetFolderPath(\"Desktop\")) -Force"'
	@echo "Built dist/$(SETUP_EXE)  (also copied to the build box Desktop)."

# --- regenerate the bundled typeface from the pristine upstream files ---
# Nothing else regenerates $(FONTSRC), so it is the one build output that can silently go
# stale: change FONT_SCALE and forget this, and `make stage` ships the OLD font under the NEW
# version number with nothing said. Run it after any change to the scale or to
# tools/lib/rescale_font.py, and put the result in front of the font ruler afterwards - the
# whole point of the scale is that 18 pt MEASURES 18 pt, and only a ruler can confirm that.
#
# Feeds on upstream/ttf, not upstream/otf: Word embeds TrueType outlines and skips PostScript
# ones silently. See the note by FONTS above.
FONT_SCALE := 1.094
fonts:
	rm -f $(FONTSRC)/*.ttf $(FONTSRC)/*.otf
	python3 tools/lib/rescale_font.py --factor $(FONT_SCALE) --family "$(FONT_FAMILY)" \
	    --out $(FONTSRC) assets/fonts/atkinson-hyperlegible/upstream/ttf/*.ttf
	@echo "Regenerated $(FONTSRC) at x$(FONT_SCALE). CHECK IT ON THE FONT RULER."

clean:
	rm -rf dist build

.PHONY: help check-config push-src pull build ribbon qat check-qat check-tabs check-frm-eol check-vba-lines read deploy stage fonts bump installer installer-build clean
