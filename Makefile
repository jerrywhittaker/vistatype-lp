# VistaType LP -- Linux-side build orchestration.
#
# Source of truth is the text tree under src/. The .dotm cannot be compiled on
# Linux, so `make build` and `make pull` drive a remote Windows+Word box over SSH
# (see build.config). You stay in this terminal; Word is invisible infrastructure.
#
#   make pull    Export VBA from Normal.dotm INTO src/  (canonical; run once to seed,
#                or after anyone edits in the Word VBE). Overwrites src/vba + src/forms.
#   make build   Import src/ into a fresh dist/Normal.dotm via Word, copy it back.
#                Also assembles the non-VBA artifacts (ribbon, .dotx) into dist/.
#   make read    Refresh reference/ from Normal.dotm using the Linux-only decompressor
#                (read/diff aid; does NOT need Windows and is NOT import-ready).
#   make deploy  Promote dist/Normal.dotm to the repo root (the deployable copy).
#   make installer  Stage shipping files and compile the Inno Setup installer on
#                the Windows box; copies Setup.exe back to dist/.
#   make clean   Remove dist/ and build/ scratch.

-include build.config
WIN_PWSH ?= powershell
WIN_ISCC ?= "C:/Program Files (x86)/Inno Setup 6/ISCC.exe"
DOTM      := Normal.dotm
DOTX      := LargePrintTemplate.dotx
RIBBON    := Word.officeUI
SHIP_DOTM := LPandBRL.dotm
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

# --- build the shipping .dotm from src/ via Word, then assemble dist/ ---
build: check-config push-src
	$(SSH) '$(WIN_PWSH) -ExecutionPolicy Bypass -File $(WSCRIPTS)/Import-Vba.ps1 -Shell "$(WIN_DIR)/$(DOTM)" -SrcRoot "$(WIN_DIR)/src" -OutDotm "$(WIN_DIR)/dist/$(DOTM)"'
	mkdir -p dist
	scp -q "$(WIN_HOST):$(WIN_DIR)/dist/$(DOTM)" dist/$(DOTM)
	cp src/ribbon/$(RIBBON) dist/$(RIBBON)
	cp $(DOTX) dist/$(DOTX)
	@echo "Built dist/$(DOTM) (+ ribbon, .dotx). Smoke-test in Word before deploying."

# --- Linux-only read helper (no Windows) ---
read:
	python3 tools/lib/decompress_vba.py

deploy:
	@test -f dist/$(DOTM) || { echo "ERROR: run 'make build' first."; exit 1; }
	cp dist/$(DOTM) $(DOTM)
	cp dist/$(RIBBON) $(RIBBON)
	@echo "Promoted dist/ artifacts to repo root."

# --- stage the three shipping files into dist/ under their SHIPPED names ---
stage: build
	cp dist/$(DOTM) dist/$(SHIP_DOTM)
	@echo "Staged dist/$(SHIP_DOTM), dist/$(DOTX), dist/$(RIBBON) for packaging."

# --- compile the Inno Setup installer on the Windows box ---
installer: check-config stage
	$(SSH) "if not exist \"$(WIN_DIR)\" mkdir \"$(WIN_DIR)\""
	scp -q -r installer dist "$(WIN_HOST):$(WIN_DIR)/"
	$(SSH) '$(WIN_ISCC) "/DSrcDir=$(WIN_DIR)/dist" "/DAppVer=$(APPVER)" "$(WIN_DIR)/installer/vistatype.iss"'
	scp -q "$(WIN_HOST):$(WIN_DIR)/dist/VistaType-LP-Setup-$(APPVER).exe" dist/
	@echo "Built dist/VistaType-LP-Setup-$(APPVER).exe"

clean:
	rm -rf dist build

.PHONY: help check-config push-src pull build read deploy stage installer clean
