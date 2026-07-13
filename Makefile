# VistaType LP -- Linux-side build orchestration.
#
# Source of truth is the text tree under src/. The .dotm cannot be compiled on
# Linux, so `make build` and `make pull` drive a remote Windows+Word box over SSH
# (see build.config). You stay in this terminal; Word is invisible infrastructure.
#
#   make pull    Export VBA from Normal.dotm INTO src/  (canonical; run once to seed,
#                or after anyone edits in the Word VBE). Overwrites src/vba + src/forms.
#   make build   Import src/ into dist/Normal.dotm via Word, then embed the ribbon
#                (customUI14.xml) and stage the .dotx. Smoke-test before deploying.
#   make ribbon  Regenerate src/ribbon/customUI14.xml from the legacy Word.officeUI.
#   make qat     Regenerate installer/qat-controls.xml from the legacy Word.officeUI.
#   make read    Refresh reference/ from Normal.dotm using the Linux-only decompressor
#                (read/diff aid; does NOT need Windows and is NOT import-ready).
#   make deploy  Promote dist/Normal.dotm (embedded ribbon) to the repo root.
#   make stage   Copy the shipping files into dist/ under their shipped names.
#   make installer  Compile the Inno Setup installer on the Windows box; copies the
#                Setup.exe back to dist/. Ships two files (.dotm + .dotx).
#   make clean   Remove dist/ and build/ scratch.

-include build.config
WIN_PWSH ?= powershell
WIN_ISCC ?= "C:/Program Files (x86)/Inno Setup 6/ISCC.exe"
DOTM      := Normal.dotm
DOTX      := LargePrintTemplate.dotx
RIBBON    := Word.officeUI
SHIP_DOTM := LPandBRL.dotm
PROJNAME  := LPandBRL
APPVER    := 2.2.3

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
build: check-config push-src
	$(SSH) '$(WIN_PWSH) -ExecutionPolicy Bypass -File $(WSCRIPTS)/Import-Vba.ps1 -Shell "$(WIN_DIR)/$(DOTM)" -SrcRoot "$(WIN_DIR)/src" -OutDotm "$(WIN_DIR)/dist/$(DOTM)" -ProjectName $(PROJNAME)'
	mkdir -p dist
	scp -q "$(WIN_HOST):$(WIN_DIR)/dist/$(DOTM)" dist/$(DOTM)
	python3 tools/lib/inject_customui.py dist/$(DOTM) src/ribbon/customUI14.xml
	cp $(DOTX) dist/$(DOTX)
	@echo "Built dist/$(DOTM) (embedded ribbon) + dist/$(DOTX). Smoke-test in Word before deploying."

# --- regenerate customUI14.xml from the legacy Word.officeUI (one-off / reference) ---
ribbon:
	python3 tools/lib/officeui_to_customui.py $(RIBBON) src/ribbon/customUI14.xml

# --- regenerate the installer's QAT icon list from the legacy Word.officeUI ---
qat:
	python3 tools/lib/extract_qat.py $(RIBBON) installer/qat-controls.xml

# --- Linux-only read helper (no Windows) ---
read:
	python3 tools/lib/decompress_vba.py

deploy:
	@test -f dist/$(DOTM) || { echo "ERROR: run 'make build' first."; exit 1; }
	cp dist/$(DOTM) $(DOTM)
	@echo "Promoted dist/$(DOTM) (embedded ribbon) to repo root."

# --- stage the shipping files into dist/ under their SHIPPED names ---
# Ribbon is embedded in the .dotm now, so only two files ship (no Word.officeUI).
stage: build
	cp dist/$(DOTM) dist/$(SHIP_DOTM)
	cp LICENSE dist/LICENSE.txt
	@echo "Staged dist/$(SHIP_DOTM) + dist/$(DOTX) + dist/LICENSE.txt for packaging."

# --- compile the Inno Setup installer on the Windows box ---
installer: check-config stage
	$(SSH) "if not exist \"$(WIN_DIR)\" mkdir \"$(WIN_DIR)\""
	scp -q -r installer dist "$(WIN_HOST):$(WIN_DIR)/"
	$(SSH) '$(WIN_ISCC) "/DSrcDir=$(WIN_DIR)/dist" "/DAppVer=$(APPVER)" "$(WIN_DIR)/installer/vistatype.iss"'
	scp -q "$(WIN_HOST):$(WIN_DIR)/dist/VistaType-LP-Setup-$(APPVER).exe" dist/
	@echo "Built dist/VistaType-LP-Setup-$(APPVER).exe"

clean:
	rm -rf dist build

.PHONY: help check-config push-src pull build ribbon qat read deploy stage installer clean
