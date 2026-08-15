#!/usr/bin/env python3
"""Scan the built Setup.exe with about seventy antivirus engines, via VirusTotal (Linux-only).

Why this exists: on 8/13/2026 Windows Defender began deleting every built Setup.exe within
seconds, as Trojan:Win32/Bearfoos.B!ml -- a machine-learning GUESS, not a match against
anything known. Two days of arguing about which part of the installer caused it produced four
wrong answers. This turns the question into a measurement: upload the file, see which engines
actually complain and what they call it, and keep a record so the next build can be compared
with the last. The background is in docs/Code-Signing.md.

WHAT IT WILL AND WILL NOT UPLOAD

  It refuses anything that is not a .exe sitting in dist/. That is not fussiness. A VirusTotal
  upload is PUBLIC AND PERMANENT: the file is kept for ever, shared with every antivirus
  vendor, and downloadable by VirusTotal's paying customers. There is no way to withdraw one.
  The likeliest reason anyone would ever type a filename here is a transcriber's document that
  Defender ate -- which is precisely the thing that must never go. Upload of a build is
  harmless (VistaType is given away under the GPL) and useful, because it puts the sample in
  front of the vendors who need to stop flagging it.

  To scan something else deliberately, copy it into dist/ first. That is meant to take a
  conscious act.

WHAT THE ANSWER IS WORTH

  VirusTotal's "Microsoft" engine is NOT the Defender on a real machine. VirusTotal runs the
  engines without their cloud, and an !ml verdict is a cloud verdict. This can come back clean
  while a real machine still deletes the file. Pair it with an on-machine check, where
  -DisableRemediation reports the verdict without destroying the file:

      "C:\\Program Files\\Windows Defender\\MpCmdRun.exe" -Scan -ScanType 3 \\
          -File <path> -DisableRemediation

ACTING ON THE FINDINGS, which is the point of the exit code:

    exit 0   nothing flagged it, or fewer than --max did and Microsoft was not among them
    exit 1   Microsoft flagged it, or more than --max engines did
    exit 2   no answer to act on -- no key, network, quota, or a report with no verdicts in it

Microsoft is singled out because Defender is what a transcriber actually has. One or two
obscure engines flagging an unsigned installer is ordinary background noise; Microsoft doing it
is what stops someone installing. A report that comes back with no engine results at all is
treated as exit 2, never as a pass: a gate that fails open is worse than no gate.

A free API key comes from a virustotal.com account (sign in -> your name -> API key). Put it in
build.config, which is gitignored, as:

    VT_API_KEY = <the key>

or set VT_API_KEY in the environment. The free key allows 4 requests a minute and 500 a day;
one run of this script uses three or four.

Usage:  python3 tools/lib/scan_virustotal.py [FILE] [--max N] [--rescan] [--json]

With no FILE, it takes the newest installer in dist/. Prefer `make scan`, which names the file
exactly rather than guessing from timestamps.
"""
import argparse
import hashlib
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
import uuid
from datetime import datetime
from pathlib import Path

API = "https://www.virustotal.com/api/v3"
GUI = "https://www.virustotal.com/gui/file/"
DIST = Path("dist")
INSTALLER_GLOB = "VistaType LP and Braille Macros Setup *.exe"

# NOT under dist/: `make clean` removes that whole folder, and this file is the only record of
# how previous builds scored. Gitignored at the repo root instead.
HISTORY = Path("virustotal-history.tsv")

# VirusTotal takes a direct upload up to this size; ours is under 4 MB and has no business
# growing past it. Above it the API wants a one-time upload URL fetched separately, which is
# not worth carrying until something actually needs it.
MAX_DIRECT_UPLOAD = 32 * 1024 * 1024

# 25s, not 20: the free key allows four requests a minute, and a run also spends one on the
# initial lookup, one on the upload and one on the final report. At 20s the first minute could
# ask for five and earn a 429 with the file ALREADY uploaded and no report to show for it.
POLL_SECONDS = 25
POLL_ATTEMPTS = 14

# The buckets that mean an engine actually delivered a verdict. The stats object also carries
# type-unsupported, failure, timeout and confirmed-timeout, and counting those would overstate
# how many engines really looked at the file.
VERDICT_BUCKETS = ("harmless", "undetected", "malicious", "suspicious")


def die(msg, code=2):
    print(f"ERROR: {msg}", file=sys.stderr)
    sys.exit(code)


class NoRedirect(urllib.request.HTTPRedirectHandler):
    """Refuse redirects outright.

    urllib copies every header onto a redirected request, the API key included, and will
    happily follow one to another host. Against virustotal.com that is theoretical, but it is
    the one route by which the key could leave the address it was meant for.
    """

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise urllib.error.URLError(f"refused a redirect to {newurl}")


OPENER = urllib.request.build_opener(NoRedirect)


def read_api_key():
    key = os.environ.get("VT_API_KEY", "").strip()
    if key:
        return key
    cfg = Path("build.config")
    if cfg.exists():
        m = re.search(r"^\s*VT_API_KEY\s*=\s*(\S+)", cfg.read_text(), re.MULTILINE)
        if m:
            return m.group(1).strip()
    die(
        "no VirusTotal API key. Add a line to build.config:\n"
        "    VT_API_KEY = <your key>\n"
        "A free key comes from a virustotal.com account: sign in, click your name, API key.\n"
        "build.config is gitignored, so the key never reaches the repository."
    )


def newest_installer():
    if not DIST.is_dir():
        die("no dist/ folder. Run `make installer` first, or name a file under dist/.")
    found = sorted(DIST.glob(INSTALLER_GLOB), key=lambda p: p.stat().st_mtime)
    if not found:
        die(f"no installer in dist/ matching {INSTALLER_GLOB!r}. Run `make installer` first.")
    return found[-1]


def approved_target(arg):
    """Resolve what to upload, refusing anything that is not a .exe inside dist/.

    A VirusTotal upload cannot be taken back, so this is a hard refusal with no override flag.
    The escape hatch is to copy the file into dist/, which is a deliberate act rather than a
    mistyped argument.
    """
    if not arg:
        return newest_installer()
    path = Path(arg)
    if not path.is_file():
        die(f"no such file: {path}")
    resolved = path.resolve()
    dist = DIST.resolve()
    if resolved.parent != dist:
        die(
            f"refusing to upload {path}.\n"
            f"  A VirusTotal upload is PUBLIC and PERMANENT, so this only ever sends a built\n"
            f"  installer: a .exe directly inside {dist}.\n"
            f"  Copy the file there first if you really mean to publish it."
        )
    if resolved.suffix.lower() != ".exe":
        die(
            f"refusing to upload {path.name}: only a .exe is scanned.\n"
            f"  A VirusTotal upload is PUBLIC and PERMANENT. Never send a document."
        )
    return path


def request(url, key, method="GET", body=None, content_type=None, tolerate_429=False):
    headers = {"x-apikey": key, "accept": "application/json"}
    if content_type:
        headers["content-type"] = content_type
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with OPENER.open(req, timeout=180) as resp:
            raw = resp.read().decode("utf-8")
        try:
            return resp.status, json.loads(raw)
        except json.JSONDecodeError:
            die("VirusTotal returned something that is not JSON. Try again in a minute.")
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:400]
        if e.code == 404:
            return 404, None
        if e.code == 401:
            die("VirusTotal rejected the API key (401). Check VT_API_KEY in build.config.")
        if e.code == 429:
            if tolerate_429:
                return 429, None
            die(
                "VirusTotal quota reached (429). A free key allows 4 requests a minute and "
                "500 a day. Wait a minute and try again."
            )
        die(f"VirusTotal returned HTTP {e.code}: {detail}")
    except urllib.error.URLError as e:
        die(f"could not reach VirusTotal: {e.reason}")


def multipart(path):
    """Encode one file as multipart/form-data. Kept by hand so this stays stdlib-only."""
    boundary = "----VistaType" + uuid.uuid4().hex
    # approved_target() has already limited this to a .exe in dist/, but a quote or a line
    # break in a filename would still corrupt the header, so strip anything that could.
    safe = re.sub(r'[^\w .()+-]', "_", path.name)
    head = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="file"; filename="{safe}"\r\n'
        f"Content-Type: application/octet-stream\r\n\r\n"
    ).encode("utf-8")
    tail = f"\r\n--{boundary}--\r\n".encode("utf-8")
    return head + path.read_bytes() + tail, f"multipart/form-data; boundary={boundary}"


def sha256_of(path):
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def analysis_id_of(data):
    return ((data or {}).get("data") or {}).get("id")


def wait_for_analysis(analysis_id, key):
    for attempt in range(POLL_ATTEMPTS):
        time.sleep(POLL_SECONDS)
        status, data = request(f"{API}/analyses/{analysis_id}", key, tolerate_429=True)
        if status == 429:
            # Already uploaded; a rate limit here is worth waiting out, not dying on.
            print("  ... rate-limited, waiting")
            continue
        attrs = ((data or {}).get("data") or {}).get("attributes") or {}
        state = attrs.get("status", "?")
        if state == "completed":
            return True
        print(f"  ... {state} ({(attempt + 1) * POLL_SECONDS}s)")
    return False


def fetch_report(digest, key):
    status, data = request(f"{API}/files/{digest}", key)
    if status == 404 or not data:
        return None
    return ((data.get("data") or {}).get("attributes")) or None


def record_history(path, digest, stats, microsoft, permalink):
    new = not HISTORY.exists()
    with HISTORY.open("a", encoding="utf-8") as fh:
        if new:
            fh.write("when\tfile\tsha256\tmalicious\tsuspicious\tmicrosoft\tpermalink\n")
        fh.write(
            "\t".join(
                [
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    path.name,
                    digest,
                    str(stats.get("malicious", 0)),
                    str(stats.get("suspicious", 0)),
                    microsoft or "-",
                    permalink,
                ]
            )
            + "\n"
        )


def previous_run(digest):
    """The most recent earlier line in the history file, for a different file."""
    if not HISTORY.exists():
        return None
    rows = [r.split("\t") for r in HISTORY.read_text(encoding="utf-8").splitlines()[1:] if r]
    for row in reversed(rows):
        if len(row) >= 6 and row[2] != digest:
            return row
    return None


def main():
    ap = argparse.ArgumentParser(description="Scan a built installer with VirusTotal's engines.")
    ap.add_argument("file", nargs="?", help="a .exe in dist/ (default: the newest installer)")
    ap.add_argument(
        "--max",
        type=int,
        default=3,
        metavar="N",
        help="how many engines may flag it before this fails (default 3)",
    )
    ap.add_argument(
        "--rescan",
        action="store_true",
        help="ask VirusTotal to run the engines again, even if it already knows the file",
    )
    ap.add_argument(
        "--json",
        action="store_true",
        help="also write the full report to dist/virustotal-<hash>.json",
    )
    args = ap.parse_args()

    key = read_api_key()
    path = approved_target(args.file)
    if path.stat().st_size > MAX_DIRECT_UPLOAD:
        die(f"{path.name} is over {MAX_DIRECT_UPLOAD // (1024 * 1024)} MB; direct upload only.")

    digest = sha256_of(path)
    permalink = GUI + digest
    print(f"File   : {path.name}  ({path.stat().st_size:,} bytes)")
    print(f"SHA-256: {digest}")

    # A file VirusTotal has already seen needs no upload, which keeps the free daily allowance
    # for the builds that are genuinely new. --rescan forces fresh engine runs on a known file.
    attrs = fetch_report(digest, key)
    if attrs and not args.rescan:
        seen = attrs.get("last_analysis_date")
        when = datetime.fromtimestamp(seen).strftime("%Y-%m-%d %H:%M") if seen else "unknown"
        print(f"Status : already known to VirusTotal, last scanned {when}")
    else:
        if attrs and args.rescan:
            print("Status : asking VirusTotal to run the engines again ...")
            _, data = request(f"{API}/files/{digest}/analyse", key, method="POST")
        else:
            print("Status : uploading (this file is new to VirusTotal) ...")
            body, ctype = multipart(path)
            _, data = request(f"{API}/files", key, method="POST", body=body, content_type=ctype)
        analysis_id = analysis_id_of(data)
        if not analysis_id:
            die("VirusTotal accepted the file but returned no analysis id.")
        print("         waiting for the engines to report ...")
        if not wait_for_analysis(analysis_id, key):
            die("VirusTotal did not finish in time. The report will appear at:\n  " + permalink)
        attrs = fetch_report(digest, key)
        if not attrs:
            die("VirusTotal finished but returned no report. Try:\n  " + permalink)

    stats = attrs.get("last_analysis_stats") or {}
    results = attrs.get("last_analysis_results") or {}

    # Never treat "no data" as "nothing wrong". This happens transiently right after a rescan,
    # while VirusTotal re-consolidates the file, and on any partial payload. Exiting 0 here
    # would print a green light with nothing behind it.
    if not results:
        die(
            "VirusTotal returned a report with no engine results in it. That is not a pass.\n"
            "  Wait a minute and run it again, or read it directly at:\n  " + permalink
        )

    malicious = stats.get("malicious", 0)
    suspicious = stats.get("suspicious", 0)
    ran = sum(stats.get(b, 0) for b in VERDICT_BUCKETS)
    flagged = {
        n: r for n, r in results.items() if r.get("category") in ("malicious", "suspicious")
    }
    ms = results.get("Microsoft")
    ms_present = ms is not None
    ms_flagged = ms_present and ms.get("category") in ("malicious", "suspicious")
    ms_label = ms.get("result") if ms_flagged else None

    print()
    print(f"Engines: {ran} returned a verdict, {malicious} malicious, {suspicious} suspicious")
    if flagged:
        print()
        for name in sorted(flagged):
            r = flagged[name]
            mark = "<-- this is the one that matters" if name == "Microsoft" else ""
            print(f"  {name:<24} {r.get('category',''):<11} {r.get('result') or '-'} {mark}")
    else:
        print("  Nothing flagged it.")

    print()
    if not ms_present:
        print("Microsoft (Defender's engine): DID NOT REPORT")
    else:
        print(f"Microsoft (Defender's engine): {ms_label or 'clean'}")
        print("  Remember VirusTotal runs Defender WITHOUT its cloud, and Bearfoos-style")
        print("  verdicts are cloud verdicts. A clean line here is not proof a real machine")
        print("  will keep the file. Confirm with MpCmdRun on the build box.")
    print()
    print(f"Full report: {permalink}")

    prev = previous_run(digest)
    if prev:
        was = int(prev[3]) + int(prev[4])
        now = malicious + suspicious
        arrow = "same as" if now == was else ("better than" if now < was else "WORSE than")
        print(f"Previous   : {prev[1]} had {was} detections on {prev[0]} -- this is {arrow} that.")

    record_history(path, digest, stats, ms_label, permalink)
    if args.json:
        out = DIST / f"virustotal-{digest[:12]}.json"
        out.write_text(json.dumps(attrs, indent=2), encoding="utf-8")
        print(f"Wrote      : {out}")

    print()
    if not ms_present:
        print("NO ANSWER: Microsoft's engine did not report on this file, and Microsoft is the")
        print("           one verdict this check exists to get. Run it again shortly.")
        return 2
    if ms_flagged:
        print("FAIL: Microsoft flagged this build. That is what a transcriber's machine runs.")
        print("      Report it at https://www.microsoft.com/en-us/wdsi/filesubmission")
        print("      as 'Software developer', priority Medium. See docs/Code-Signing.md.")
        return 1
    if malicious + suspicious > args.max:
        print(f"FAIL: {malicious + suspicious} engines flagged this build (limit {args.max}).")
        print("      See docs/Code-Signing.md for what is worth doing about it.")
        return 1
    if flagged:
        print("OK, with noise: a couple of engines flagging an unsigned installer is ordinary.")
        return 0
    print("OK: no engine flagged this build.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
