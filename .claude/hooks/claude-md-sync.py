#!/usr/bin/env python3
"""PostToolUse hook: nudge Claude to keep CLAUDE.md in sync with the build pipeline.

Fires after a Write/Edit. If the edited file is one that defines the project's
architecture or build pipeline (Makefile, DEVELOPMENT.md, the build/ribbon/installer
tooling, build config), it injects a reminder telling Claude to reconcile CLAUDE.md in
the same change. Emits nothing for unrelated files, so it is silent on normal edits.

Detection is by path only -- it does not judge whether CLAUDE.md is actually stale; it
just guarantees the question gets asked whenever the architecture surface is touched.
"""
import json
import sys

# Substrings / basenames that indicate an architecture- or build-pipeline-defining file.
# Keep this list in step with CLAUDE.md's "Repo layout" and "Build & edit workflow".
ARCH_MARKERS = (
    "/Makefile",
    "/DEVELOPMENT.md",
    "/tools/",           # Export/Import PS1, ribbon + QAT python, decompressor
    "/src/ribbon/",      # customUI14.xml / Word.officeUI -- the ribbon source
    "/installer/",       # Inno Setup script + QAT merge/remove scripts
    "/build.config",     # build.config / build.config.example
    "/.gitattributes",   # CRLF/binary rules the pipeline depends on
)


def main() -> int:
    try:
        payload = json.load(sys.stdin)
    except (json.JSONDecodeError, ValueError):
        return 0  # not our job to police malformed input

    path = (payload.get("tool_input") or {}).get("file_path") or ""
    if not path:
        return 0

    # Never recurse on CLAUDE.md itself.
    if path.endswith("/CLAUDE.md") or path.endswith("CLAUDE.md"):
        return 0

    norm = path.replace("\\", "/")
    if not any(marker in norm for marker in ARCH_MARKERS):
        return 0

    reminder = (
        f"You just edited `{path}`, which helps define this project's build "
        "pipeline / architecture. Before finishing this turn, re-read CLAUDE.md "
        "(especially the 'Repo layout' and 'Build & edit workflow' sections) and, if "
        "this change made anything there inaccurate -- build paths, make targets, file "
        "roles, tool names -- update CLAUDE.md in the SAME change so it stays in sync. "
        "If CLAUDE.md is already accurate, do nothing."
    )
    print(json.dumps({
        "hookSpecificOutput": {
            "hookEventName": "PostToolUse",
            "additionalContext": reminder,
        }
    }))
    return 0


if __name__ == "__main__":
    sys.exit(main())
