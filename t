#!/usr/bin/env python3
"""
obda_pipeline.py
------------------------------------------------------------------
Full start-to-finish pipeline:

  STEP 1  EXTRACT  - Unzips every .zip in the root folder using
                     Windows PowerShell's Expand-Archive (so the
                     extracted folders actually show up on disk).

  STEP 2  RENAME   - Reads every open Chrome tab on the sage-admin
                     /obda/ pages (via remote debugging), grabs the
                     six-digit OBDA ID, and renames each extracted
                     folder from  automated-<slug>  to
                     automated-<sixdigit>.

  STEP 3  RUN      - Finds every run-me.bat inside the folders and
                     launches them like double-clicking, in CHUNKS
                     OF 5: start 5, wait for all 5 to finish, pause,
                     then the next 5. Small delay between clicks too.

------------------------------------------------------------------
SETUP (one time):
    pip install requests websocket-client

CHROME (only needed for the rename step):
    Close Chrome fully, then relaunch it with:
    "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222
    and restore your sage-admin tabs.

USAGE:
    # Everything, start to finish:
    python obda_pipeline.py --root "C:\Users\111067427\Downloads\222"

    # Skip steps you don't need:
    python obda_pipeline.py --root "..." --skip-extract
    python obda_pipeline.py --root "..." --skip-rename
    python obda_pipeline.py --root "..." --skip-run

    # Tune the run behaviour:
    python obda_pipeline.py --root "..." --chunk-size 5 --click-delay 3 --chunk-pause 15
------------------------------------------------------------------
"""

import argparse
import json
import re
import subprocess
import sys
import time
from pathlib import Path

# ----------------------------------------------------------------
# STEP 1 - EXTRACT ALL ZIPS (via PowerShell Expand-Archive)
# ----------------------------------------------------------------
def extract_all(root: Path):
    zips = sorted(root.glob("*.zip"))
    if not zips:
        print("  no .zip files found - nothing to extract.")
        return

    print(f"  found {len(zips)} zip file(s).")
    for z in zips:
        dest = root / z.stem          # extract into folder named after the zip
        if dest.exists():
            print(f"  - skip (already extracted): {z.name}")
            continue
        print(f"  extracting {z.name} -> {dest.name}\\")
        cmd = [
            "powershell", "-NoProfile", "-Command",
            f'Expand-Archive -LiteralPath "{z}" -DestinationPath "{dest}" -Force'
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"  ! FAILED to extract {z.name}: {result.stderr.strip()}")
        else:
            print(f"  done: {dest.name}")


# ----------------------------------------------------------------
# STEP 2 - RENAME FOLDERS USING OBDA IDs FROM OPEN CHROME TABS
# ----------------------------------------------------------------
def get_tab_text(ws_url):
    from websocket import create_connection
    ws = create_connection(ws_url, timeout=10)
    try:
        ws.send(json.dumps({
            "id": 1,
            "method": "Runtime.evaluate",
            "params": {"expression": "document.body.innerText", "returnByValue": True},
        }))
        resp = json.loads(ws.recv())
        return resp.get("result", {}).get("result", {}).get("value", "") or ""
    finally:
        ws.close()


def collect_mapping(port: int, url_filter: str) -> dict:
    import requests
    tabs = requests.get(f"http://localhost:{port}/json/list", timeout=10).json()
    mapping = {}
    for tab in tabs:
        url = tab.get("url", "")
        if tab.get("type") != "page" or url_filter not in url:
            continue
        slug = url.rstrip("/").split("/")[-1]
        if not slug:
            continue
        try:
            text = get_tab_text(tab["webSocketDebuggerUrl"])
        except Exception as e:
            print(f"  ! could not read tab {slug}: {e}")
            continue
        m = re.search(r"(\d{4,8})\s*/\s*" + re.escape(slug), text)
        if not m:
            print(f"  ! no OBDA number found on tab {slug}")
            continue
        mapping[slug] = m.group(1)
        print(f"  found  {slug}  ->  {m.group(1)}")
    return mapping


def rename_folders(root: Path, mapping: dict, prefix: str):
    count = 0
    for slug, number in mapping.items():
        src = root / f"{prefix}{slug}"
        if not src.is_dir():
            print(f"  - skip: no folder named {src.name}")
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
        print("  ! Skipping rename step.")
        return
    try:
        mapping = collect_mapping(port, url_filter)
    except Exception:
        print(f"  ! Can't reach Chrome on port {port}.")
        print(f"  ! Launch Chrome with --remote-debugging-port={port} and re-run,")
        print(f"  ! or re-run with --skip-rename if folders are already named right.")
        return
    if not mapping:
        print("  ! No OBDA IDs collected - nothing renamed.")
        return
    rename_folders(root, mapping, prefix)


# ----------------------------------------------------------------
# STEP 3 - RUN EVERY run-me.bat IN CHUNKS OF N
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
            # Same as double-clicking: cmd runs the bat, cwd = its own folder
            p = subprocess.Popen(
                ["cmd", "/c", str(bat)],
                cwd=str(bat.parent),
                creationflags=subprocess.CREATE_NEW_CONSOLE,
            )
            procs.append((bat, p))
            time.sleep(click_delay)          # delay between "clicks"

        # wait for all 5 in this chunk to finish before moving on
        print(f"    waiting for chunk {ci} to finish...")
        for bat, p in procs:
            code = p.wait()
            print(f"    finished ({code}): {bat.name}  [{bat.parent.name}]")

        if ci < len(chunks):
            print(f"    pausing {chunk_pause}s before next chunk...\n")
            time.sleep(chunk_pause)

    print(f"\n  all {len(bats)} bat file(s) completed.")


# ----------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Extract -> Rename -> Run pipeline")
    ap.add_argument("--root", required=True, help="Folder with the zips / extracted folders")
    ap.add_argument("--pattern", default="run*me*.bat", help="Bat filename pattern")
    ap.add_argument("--prefix", default="automated-", help="Extracted folder prefix")
    ap.add_argument("--port", type=int, default=9222, help="Chrome debug port")
    ap.add_argument("--url-filter", default="/obda/", help="Only read tabs containing this")
    ap.add_argument("--chunk-size", type=int, default=5, help="Bats per chunk")
    ap.add_argument("--click-delay", type=float, default=3, help="Seconds between launches inside a chunk")
    ap.add_argument("--chunk-pause", type=float, default=10, help="Seconds between chunks")
    ap.add_argument("--skip-extract", action="store_true")
    ap.add_argument("--skip-rename", action="store_true")
    ap.add_argument("--skip-run", action="store_true")
    args = ap.parse_args()

    root = Path(args.root)
    if not root.is_dir():
        sys.exit(f"Root folder not found: {root}")

    if not args.skip_extract:
        print("\n=== STEP 1: EXTRACT ZIPS ===")
        extract_all(root)

    if not args.skip_rename:
        print("\n=== STEP 2: RENAME FROM CHROME TABS ===")
        step_rename(root, args.port, args.url_filter, args.prefix)

    if not args.skip_run:
        print("\n=== STEP 3: RUN BAT FILES (chunks of {}) ===".format(args.chunk_size))
        run_bats(root, args.pattern, args.chunk_size,
                 args.click_delay, args.chunk_pause)

    print("\nPipeline complete.")


if __name__ == "__main__":
    main()
