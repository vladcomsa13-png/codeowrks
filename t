#!/usr/bin/env python3
r"""
obda_pipeline.py  (v3)
------------------------------------------------------------------
Full start-to-finish pipeline:

  STEP 1  EXTRACT  - Unzips every .zip in root via PowerShell.
  STEP 2  RENAME   - Reads open Chrome /obda/ tabs, grabs the
                     six-digit OBDA ID, renames folders
                     (case-insensitive matching).
  STEP 3  RUN      - Runs run-me.bat files in chunks of N,
                     waiting for each chunk to finish.
  STEP 4  PREP     - Goes tab by tab: opens the Action dropdown,
                     clicks "Upload Report" so the modal is open,
                     and PRINTS the path of that tab's
                     final-report zip. You pick the file and click
                     Submit manually, press Enter, and it moves to
                     the next tab.

------------------------------------------------------------------
SETUP (one time):
    pip install requests websocket-client

CHROME:
    Close Chrome fully, then relaunch with:
    & "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222

USAGE:
    python obda_pipeline.py --root "C:\Users\111067427\Downloads\Jul 13"

    # Folders already extracted + renamed, bats already run:
    python obda_pipeline.py --root "..." --skip-extract --skip-rename --skip-run

    # See what the script sees (no changes):
    python obda_pipeline.py --root "..." --diagnose
------------------------------------------------------------------
"""

import argparse
import json
import re
import subprocess
import sys
import time
from pathlib import Path

DEFAULT_PREFIX = "automation-"


# ----------------------------------------------------------------
# CHROME HELPERS
# ----------------------------------------------------------------
def cdp_eval(ws_url: str, expression: str):
    from websocket import create_connection
    ws = create_connection(ws_url, timeout=15)
    try:
        ws.send(json.dumps({
            "id": 1,
            "method": "Runtime.evaluate",
            "params": {
                "expression": expression,
                "returnByValue": True,
                "awaitPromise": True,
            },
        }))
        resp = json.loads(ws.recv())
        return resp.get("result", {}).get("result", {}).get("value")
    finally:
        ws.close()


def activate_tab(port: int, tab_id: str):
    """Bring a tab to the foreground so you can see the modal."""
    import requests
    try:
        requests.get(f"http://localhost:{port}/json/activate/{tab_id}", timeout=5)
    except Exception:
        pass


def list_obda_tabs(port: int, url_filter: str):
    import requests
    tabs = requests.get(f"http://localhost:{port}/json/list", timeout=10).json()
    return [t for t in tabs
            if t.get("type") == "page" and url_filter in t.get("url", "")]


# ----------------------------------------------------------------
# STEP 1: EXTRACT
# ----------------------------------------------------------------
def extract_all(root: Path):
    zips = sorted(root.glob("*.zip"))
    if not zips:
        print("  no .zip files found - nothing to extract.")
        return
    print(f"  found {len(zips)} zip file(s).")
    for z in zips:
        dest = root / z.stem
        if dest.exists():
            print(f"  - skip (already extracted): {z.name}")
            continue
        print(f"  extracting {z.name} -> {dest.name}")
        cmd = [
            "powershell", "-NoProfile", "-Command",
            f'Expand-Archive -LiteralPath "{z}" -DestinationPath "{dest}" -Force'
        ]
        r = subprocess.run(cmd, capture_output=True, text=True)
        if r.returncode != 0:
            print(f"  ! FAILED: {r.stderr.strip()}")


# ----------------------------------------------------------------
# STEP 2: RENAME (case-insensitive)
# ----------------------------------------------------------------
def collect_mapping(port: int, url_filter: str) -> dict:
    """{slug_lower: (slug_original, six_digit_id)}"""
    tabs = list_obda_tabs(port, url_filter)
    mapping = {}
    for tab in tabs:
        url = tab["url"]
        slug = url.rstrip("/").split("/")[-1]
        if not slug:
            continue
        try:
            text = cdp_eval(tab["webSocketDebuggerUrl"], "document.body.innerText")
        except Exception as e:
            print(f"  ! could not read tab {slug}: {e}")
            continue
        if not text:
            print(f"  ! empty page text on tab {slug}")
            continue
        m = re.search(r"(\d{4,8})\s*/\s*" + re.escape(slug), text, re.IGNORECASE)
        if not m:
            print(f"  ! no OBDA number found on tab {slug}")
            continue
        mapping[slug.lower()] = (slug, m.group(1))
        print(f"  found  {slug}  ->  {m.group(1)}")
    return mapping


def rename_folders(root: Path, mapping: dict, prefix: str):
    existing = {}
    for p in root.iterdir():
        if p.is_dir() and p.name.lower().startswith(prefix.lower()):
            existing[p.name[len(prefix):].lower()] = p

    count = 0
    for slug_lower, (slug_original, number) in mapping.items():
        src = existing.get(slug_lower)
        if not src:
            print(f"  - skip: no folder for slug {slug_original}")
            continue
        dst = root / f"{prefix}{number}"
        if dst.exists():
            print(f"  - skip: {dst.name} already exists")
            continue
        src.rename(dst)
        print(f"  RENAMED  {src.name}  ->  {dst.name}")
        count += 1
    print(f"  renamed {count} folder(s).")


def step_rename(root: Path, port: int, url_filter: str, prefix: str):
    try:
        import requests  # noqa
        import websocket  # noqa
    except ImportError:
        print("  ! Missing deps. Run: pip install requests websocket-client")
        return
    try:
        mapping = collect_mapping(port, url_filter)
    except Exception as e:
        print(f"  ! Can't reach Chrome on port {port}: {e}")
        return
    if not mapping:
        print("  ! No OBDA IDs collected.")
        return
    rename_folders(root, mapping, prefix)


# ----------------------------------------------------------------
# STEP 3: RUN BATS IN CHUNKS
# ----------------------------------------------------------------
def run_bats(root: Path, pattern: str, chunk_size: int,
             click_delay: float, chunk_pause: float):
    bats = sorted(root.rglob(pattern))
    if not bats:
        print(f"  no files matching '{pattern}' found.")
        return
    print(f"  found {len(bats)} bat file(s). Running in chunks of {chunk_size}.\n")

    chunks = [bats[i:i + chunk_size] for i in range(0, len(bats), chunk_size)]
    for ci, chunk in enumerate(chunks, 1):
        print(f"  --- chunk {ci}/{len(chunks)} ({len(chunk)} file(s)) ---")
        procs = []
        for bat in chunk:
            print(f"    launching: {bat}")
            p = subprocess.Popen(
                ["cmd", "/c", str(bat)],
                cwd=str(bat.parent),
                creationflags=subprocess.CREATE_NEW_CONSOLE,
            )
            procs.append((bat, p))
            time.sleep(click_delay)

        print(f"    waiting for chunk {ci} to finish...")
        for bat, p in procs:
            code = p.wait()
            print(f"    finished ({code}): {bat.name}  [{bat.parent.name}]")

        if ci < len(chunks):
            print(f"    pausing {chunk_pause}s before next chunk...\n")
            time.sleep(chunk_pause)

    print(f"\n  all {len(bats)} bat file(s) completed.")


# ----------------------------------------------------------------
# STEP 4: OPEN "UPLOAD REPORT" MODAL ON EACH TAB (semi-manual)
# ----------------------------------------------------------------
OPEN_MODAL_JS = r"""
(async () => {
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  // If the modal is already open, we're done
  if (/upload report/i.test(document.body.innerText) &&
      document.querySelector('input[type="file"]')) {
    return 'MODAL_ALREADY_OPEN';
  }

  // 1. Open the Action dropdown ("Select action")
  const candidates = Array.from(document.querySelectorAll('div,button,span'))
      .filter(el => el.offsetParent !== null)
      .filter(el => /^\s*select action\s*$/i.test(el.textContent || ''));
  const dropdown = candidates[candidates.length - 1];
  if (!dropdown) return 'NO_ACTION_DROPDOWN';
  dropdown.click();
  await sleep(700);

  // 2. Click the "Upload Report" option
  const option = Array.from(document.querySelectorAll('li,[role="option"],[role="menuitem"],div,span'))
      .filter(el => el.offsetParent !== null)
      .find(el => /^\s*upload report\s*$/i.test(el.textContent || ''));
  if (!option) return 'NO_UPLOAD_REPORT_OPTION';
  option.click();
  await sleep(700);

  // 3. Verify the modal opened
  if (document.querySelector('input[type="file"]')) return 'MODAL_OPENED';
  return 'CLICKED_BUT_NO_MODAL_DETECTED';
})()
"""


def find_report_zip(root: Path, prefix: str, number: str):
    """Find the final-report zip inside the folder for this OBDA number."""
    folder = root / f"{prefix}{number}"
    if not folder.is_dir():
        return None
    hits = sorted(folder.rglob("final-report*.zip")) or \
           sorted(folder.rglob("*final*report*.zip")) or \
           sorted(folder.rglob("*.zip"))
    return hits[0] if hits else None


def step_prep_uploads(root: Path, port: int, url_filter: str, prefix: str):
    try:
        tabs = list_obda_tabs(port, url_filter)
    except Exception as e:
        print(f"  ! Can't reach Chrome: {e}")
        return
    if not tabs:
        print("  ! No /obda/ tabs open.")
        return

    print(f"  found {len(tabs)} tab(s).")
    print("  For each tab: I open the Upload Report modal and show you the")
    print("  matching final-report zip path. You Choose File + Submit,")
    print("  then press Enter here to go to the next tab.\n")

    for i, tab in enumerate(tabs, 1):
        slug = tab["url"].rstrip("/").split("/")[-1]

        # get the six-digit number from the page so we can locate the folder
        number = None
        try:
            text = cdp_eval(tab["webSocketDebuggerUrl"], "document.body.innerText")
            m = re.search(r"(\d{4,8})\s*/\s*" + re.escape(slug), text or "", re.IGNORECASE)
            if m:
                number = m.group(1)
        except Exception:
            pass

        zip_path = find_report_zip(root, prefix, number) if number else None

        activate_tab(port, tab.get("id", ""))
        try:
            result = cdp_eval(tab["webSocketDebuggerUrl"], OPEN_MODAL_JS)
        except Exception as e:
            result = f"ERROR: {e}"

        print(f"  [{i}/{len(tabs)}] tab {slug}" + (f"  (OBDA {number})" if number else ""))
        print(f"      modal: {result}")
        if zip_path:
            print(f"      FILE TO UPLOAD:  {zip_path}")
        else:
            print(f"      ! no final-report zip found for this tab" +
                  (f" (folder {prefix}{number})" if number else ""))
        input("      -> Choose File + Submit in Chrome, then press Enter here...")

    print("\n  all tabs processed.")


# ----------------------------------------------------------------
# DIAGNOSE
# ----------------------------------------------------------------
def diagnose(root: Path, port: int, url_filter: str, prefix: str):
    print("\n--- FOLDERS ON DISK ---")
    for p in sorted(root.iterdir()):
        if p.is_dir():
            print(f"  {p.name}")

    print("\n--- OPEN CHROME /obda/ TABS ---")
    try:
        tabs = list_obda_tabs(port, url_filter)
    except Exception as e:
        print(f"  ! Can't reach Chrome on port {port}: {e}")
        return
    for t in tabs:
        slug = t["url"].rstrip("/").split("/")[-1]
        print(f"  slug={slug}")

    print("\n--- MATCH CHECK ---")
    existing = {p.name[len(prefix):].lower(): p.name
                for p in root.iterdir()
                if p.is_dir() and p.name.lower().startswith(prefix.lower())}
    for t in tabs:
        slug = t["url"].rstrip("/").split("/")[-1]
        match = existing.get(slug.lower())
        print(f"  {slug}  ...  {'MATCH -> ' + match if match else 'NO FOLDER FOUND'}")


# ----------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--pattern", default="run*me*.bat")
    ap.add_argument("--prefix", default=DEFAULT_PREFIX)
    ap.add_argument("--port", type=int, default=9222)
    ap.add_argument("--url-filter", default="/obda/")
    ap.add_argument("--chunk-size", type=int, default=5)
    ap.add_argument("--click-delay", type=float, default=3)
    ap.add_argument("--chunk-pause", type=float, default=10)
    ap.add_argument("--skip-extract", action="store_true")
    ap.add_argument("--skip-rename", action="store_true")
    ap.add_argument("--skip-run", action="store_true")
    ap.add_argument("--skip-submit", action="store_true")
    ap.add_argument("--diagnose", action="store_true")
    args = ap.parse_args()

    root = Path(args.root)
    if not root.is_dir():
        sys.exit(f"Root folder not found: {root}")

    if args.diagnose:
        diagnose(root, args.port, args.url_filter, args.prefix)
        return

    if not args.skip_extract:
        print("\n=== STEP 1: EXTRACT ===")
        extract_all(root)

    if not args.skip_rename:
        print("\n=== STEP 2: RENAME FROM CHROME TABS ===")
        step_rename(root, args.port, args.url_filter, args.prefix)

    if not args.skip_run:
        print(f"\n=== STEP 3: RUN BAT FILES (chunks of {args.chunk_size}) ===")
        run_bats(root, args.pattern, args.chunk_size,
                 args.click_delay, args.chunk_pause)

    if not args.skip_submit:
        print("\n=== STEP 4: OPEN UPLOAD MODALS (you pick file + submit) ===")
        step_prep_uploads(root, args.port, args.url_filter, args.prefix)

    print("\nPipeline complete.")


if __name__ == "__main__":
    main()
