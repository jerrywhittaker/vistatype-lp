# VistaType LP -- Linux-side build orchestration.
#
# Source of truth is the text tree under src/. The .dotm cannot be compiled on
# Linux, so `make build` and `make pull` drive a remote Windows+Word box over SSH
# (see build.config). You stay in this terminal; Word is invisible infrastructure.
#
#   make pull    Export VBA from LPandBRL.dotm INTO src/  (canonical; run once to seed,
#                or after anyone edits in the Word VBE). Overwrites src/vba + src/forms.
#   make build   Import src/ into dist/LPandBRL.dotm via Word, then embed the ribbon
#                (customUI14.xml) and stage the .dotx. Smoke-test before deploying.
#   make ribbon  Regenerate src/ribbon/customUI14.xml from the legacy Word.officeUI.
#   make qat     Regenerate installer/qat-controls.xml from the legacy Word.officeUI.
#   make read    Refresh reference/ from LPandBRL.dotm using the Linux-only decompressor
#                (read/diff aid; does NOT need Windows and is NOT import-ready).
#   make deploy  Promote dist/LPandBRL.dotm (embedded ribbon) to the repo root.
#   make stage   Copy the license into dist/ alongside the built shipping files.
#   make installer  Compile the Inno Setup installer on the Windows box; copies the
#                Setup.exe back to dist/. Ships two files (.dotm + .dotx).
#   make clean   Remove dist/ and build/ scratch.

-include build.config
WIN_PWSH ?= powershell
WIN_ISCC ?= "C:/Program Files (x86)/Inno Setup 6/ISCC.exe"
DOTM      := LPandBRL.dotm
DOTX      := LargePrintTemplate.dotx
RIBBON    := Word.officeUI
PROJNAME  := LPandBRL
APPVER    := 3.0.16
SETUP_EXE := VistaType-LP-Setup-$(APPVER).exe
VERDATE   := $(shell date +%-m/%-d/%Y)

SSH := ssh $(WIN_HOST)
WSCRIPTS := $(WIN_DIR)/tools/windows

.DEFAULT_GOAL := help

help:
	@sed -n 's/^#   //p' Makefile

check-config:
	@test -f build.config || { echo "ERROR: copy build.config.example -> build.config and edit it."; exit 1; }
	@test -n "$(WIN_HOST)" || { echo "ERROR: WIN_HOST not set in build.config."; exit 1; }

# --- push the working tree (source + scripts + shell .dotm) to the Windows box ---
push-src: check-config
	$(SSH) "if not exist \"$(WIN_DIR)\" mkdir \"$(WIN_DIR)\""
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

build: check-config check-frm-eol push-src
	$(SSH) '$(WIN_PWSH) -ExecutionPolicy Bypass -File $(WSCRIPTS)/Import-Vba.ps1 -Shell "$(WIN_DIR)/$(DOTM)" -SrcRoot "$(WIN_DIR)/src" -OutDotm "$(WIN_DIR)/dist/$(DOTM)" -ProjectName $(PROJNAME) -AppVer $(APPVER) -VerDate "$(VERDATE)"'
	@# The two About forms carry the version in their binary .frx, so bring them home if the
	@# stamp changed them. -u: only if newer, so an unchanged build copies nothing.
	@for f in Lp_About_Title_And_Agreement Dx_About_Title_And_Agreement; do \
	  scp -q "$(WIN_HOST):$(WIN_DIR)/src/forms/$$f.frm" src/forms/ 2>/dev/null || true; \
	  scp -q "$(WIN_HOST):$(WIN_DIR)/src/forms/$$f.frx" src/forms/ 2>/dev/null || true; \
	done
	mkdir -p dist
	scp -q "$(WIN_HOST):$(WIN_DIR)/dist/$(DOTM)" dist/$(DOTM)
	python3 tools/lib/inject_customui.py dist/$(DOTM) src/ribbon/customUI14.xml
	cp $(DOTX) dist/$(DOTX)
	@echo "Built dist/$(DOTM) (embedded ribbon) + dist/$(DOTX). Smoke-test in Word before deploying."

# --- regenerate customUI14.xml from the legacy Word.officeUI (one-off / reference) ---
ribbon:
	python3 tools/lib/officeui_to_customui.py $(RIBBON) src/ribbon/customUI14.xml

# --- OBSOLETE: the QAT is now a hand-maintained full toolbar in installer/qat-template.officeUI
# --- (imposed by installer/scripts/Merge-Qat.ps1). extract_qat.py / this target are no longer
# --- part of the pipeline; kept only for reference. Edit qat-template.officeUI by hand.
qat:
	python3 tools/lib/extract_qat.py $(RIBBON) installer/qat-controls.xml

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
	@echo "Staged dist/$(DOTM) + dist/$(DOTX) + dist/LICENSE.txt for packaging."

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
	scp -q -r installer dist "$(WIN_HOST):$(WIN_DIR)/"
	$(SSH) '$(WIN_ISCC) "/DSrcDir=$(WIN_DIR)/dist" "/DAppVer=$(APPVER)" "$(WIN_DIR)/installer/vistatype.iss"'
	scp -q "$(WIN_HOST):$(WIN_DIR)/dist/$(SETUP_EXE)" dist/
	$(SSH) '$(WIN_PWSH) -NoProfile -Command "Copy-Item \"$(WIN_DIR)/dist/$(SETUP_EXE)\" ([Environment]::GetFolderPath(\"Desktop\")) -Force"'
	@echo "Built dist/$(SETUP_EXE) (also copied to the build box Desktop)."

clean:
	rm -rf dist build

.PHONY: help check-config push-src pull build ribbon qat read deploy stage installer clean
