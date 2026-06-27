"""
OBDA Automation Pipeline
──────────────────────────────────────────────────────────────────────────────
Mirrors your two PowerShell commands, plus scraping + uploading + renaming.

Step 1: Extract all automation zips from Downloads\\Unextracted into folders
Step 2: Run each run-me.bat (with delay between them)
Step 3: Scrape numeric IDs from OBDA portal
Step 4: Upload final-report.zip for each
Step 5: Rename folders from slug → numeric ID

USAGE:
    python obdash.py                         Extract, run, scrape, upload, rename
    python obdash.py --batch-size 10         Process 10 at a time (default)
    python obdash.py --skip-extract          Already extracted — just run + scrape + rename
    python obdash.py --skip-run              Already ran bat files — just scrape + rename
    python obdash.py --dry-run               Preview only
    python obdash.py --delay 5               Seconds between bat file launches (default: 5)

SETUP:
    pip install playwright
    playwright install chromium
"""

import argparse
import logging
import os
import re
import shutil
import subprocess
import time
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

from playwright.sync_api import sync_playwright

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s │ %(levelname)-7s │ %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("obdash")


# ── Configuration ──────────────────────────────────────────────────────────────
@dataclass
class OBDAConfig:
    obda_base_url: str = (
        "https://sage-admin-sw20-private-prod.yul.rbc105.prod.iac0.rbc.com/obda"
    )
    login_domain: str = "id.ownr.co"
    admin_domain: str = "sage-admin-sw20-private-prod"
    login_email: str = "vlad.comsa@ownr.co"

    browser_profile_dir: str = field(
        default_factory=lambda: str(Path.home() / ".obdash-browser")
    )

    # The script looks for automation-*.zip files directly in Downloads
    downloads_path: str = field(
        default_factory=lambda: str(Path.home() / "Downloads")
    )
    documents_path: str = field(
        default_factory=lambda: str(Path.home() / "Documents")
    )

    nav_timeout_ms: int = 30_000
    idle_timeout_ms: int = 15_000
    login_timeout_s: int = 1200
    max_nav_retries: int = 3
    batch_size: int = 10
    delay_between_bats: int = 5  # seconds between running each bat file

    @property
    def chrome_path(self) -> Optional[str]:
        candidates = [
            Path(os.environ.get("PROGRAMFILES", ""))
            / "Google/Chrome/Application/chrome.exe",
            Path(os.environ.get("PROGRAMFILES(X86)", ""))
            / "Google/Chrome/Application/chrome.exe",
            Path(os.environ.get("LOCALAPPDATA", ""))
            / "Google/Chrome/Application/chrome.exe",
        ]
        return next((str(p) for p in candidates if p.exists()), None)


CFG = OBDAConfig()


# ── Data ───────────────────────────────────────────────────────────────────────
@dataclass
class FolderResult:
    slug: str
    folder_path: str
    batch_exit_code: Optional[int] = None
    numeric_id: Optional[str] = None
    upload_ok: Optional[bool] = None
    error: Optional[str] = None

    @property
    def success(self) -> bool:
        return self.error is None


# ── Helpers ────────────────────────────────────────────────────────────────────
def slug_from_folder(folder_name: str) -> str:
    """automation-cool-slug → cool-slug"""
    return folder_name.replace("automation-", "")


def create_date_folder() -> str:
    """Create Documents/OBDA/YYYY/MONTH/MonDD/ and return path."""
    now = datetime.now()
    folder = (
        Path(CFG.documents_path) / "OBDA"
        / now.strftime("%Y") / now.strftime("%B").upper() / now.strftime("%b%d")
    )
    folder.mkdir(parents=True, exist_ok=True)
    return str(folder)


# ══════════════════════════════════════════════════════════════════════════════
# STEP 1: EXTRACT
# Mirrors your PowerShell:
#   Get-ChildItem -Path "...\Unextracted" -Filter *.zip |
#     ForEach-Object { Expand-Archive -Path $_.FullName
#       -DestinationPath ($_.FullName -replace '\.zip$','') -Force }
# ══════════════════════════════════════════════════════════════════════════════
def extract_all_zips(target_dir: str) -> list[str]:
    """
    Extract every automation-*.zip from Downloads into target_dir.

    Each zip extracts preserving its internal folder structure.
    After extraction, moves the zip into a "Done" subfolder in Downloads.

    Returns list of extracted folder paths.
    """
    dl = Path(CFG.downloads_path)

    zips = sorted(dl.glob("automation-*.zip"))
    if not zips:
        log.info("No automation-*.zip files in %s", dl)
        return []

    log.info("Found %d zip(s) to extract:", len(zips))
    for i, z in enumerate(zips, 1):
        log.info("  %d. %s", i, z.name)

    extracted_folders = []

    for zip_path in zips:
        slug = slug_from_folder(zip_path.stem)
        log.info("[%s] Extracting...", slug)

        try:
            with zipfile.ZipFile(str(zip_path), "r") as zf:
                names = zf.namelist()
                log.info("[%s]   %d entries in zip", slug, len(names))

                if not names:
                    log.warning("[%s]   Empty zip — skipping.", slug)
                    continue

                # Extract into target directory
                zf.extractall(target_dir)

                # Figure out what folder was created
                first_parts = [n.split("/")[0] for n in names if "/" in n]
                has_wrapper = first_parts and all(
                    p == first_parts[0] for p in first_parts
                )

                if has_wrapper:
                    result_folder = os.path.join(target_dir, first_parts[0])
                else:
                    # Flat zip — move files into a subfolder
                    subfolder = os.path.join(target_dir, zip_path.stem)
                    os.makedirs(subfolder, exist_ok=True)
                    root_files = [n for n in names if "/" not in n and n]
                    for rf in root_files:
                        src = os.path.join(target_dir, rf)
                        if os.path.exists(src):
                            shutil.move(src, os.path.join(subfolder, rf))
                    result_folder = subfolder

            # Verify
            if os.path.isdir(result_folder):
                contents = os.listdir(result_folder)
                log.info("[%s]   Extracted: %d items in %s", slug, len(contents), Path(result_folder).name)
                extracted_folders.append(result_folder)
            else:
                log.error("[%s]   Folder not found after extraction: %s", slug, result_folder)
                continue

            # Move zip to Done subfolder in Downloads
            done_dir = dl / "Done"
            done_dir.mkdir(exist_ok=True)
            dest = done_dir / zip_path.name
            if dest.exists():
                dest.unlink()
            shutil.move(str(zip_path), str(dest))
            log.info("[%s]   Zip moved to %s", slug, done_dir)

        except Exception as e:
            log.error("[%s]   Extraction failed: %s", slug, e)

    log.info("Extraction complete: %d/%d succeeded.", len(extracted_folders), len(zips))
    return extracted_folders


# ══════════════════════════════════════════════════════════════════════════════
# STEP 2: RUN BAT FILES
# Mirrors your PowerShell:
#   $batFiles = Get-ChildItem -LiteralPath $Root -Recurse -File
#       -Filter $Pattern | Sort-Object FullName
#   foreach ($bat in $batFiles) {
#       Start-Process -FilePath $bat.FullName
#           -WorkingDirectory $bat.DirectoryName
#       if ($i -lt $batFiles.Count) { Start-Sleep -Seconds $DelaySeconds }
#   }
# ══════════════════════════════════════════════════════════════════════════════
def find_all_bat_files(folders: list[str]) -> list[str]:
    """Find all run*me.bat files across the given folders, sorted."""
    bat_files = []
    for folder in folders:
        for root, dirs, files in os.walk(folder):
            for f in files:
                if re.match(r"run.*me\.bat", f, re.IGNORECASE):
                    bat_files.append(os.path.join(root, f))
    bat_files.sort()
    return bat_files


def run_all_bat_files(folders: list[str], delay: int = 5) -> dict[str, int]:
    """
    Run each run*me.bat sequentially with a delay between them.

    Uses Start-Process style (launches each in its own window)
    so you can see them running, with a delay between launches.

    Returns dict of folder_name → exit_code.
    """
    bat_files = find_all_bat_files(folders)

    if not bat_files:
        log.warning("No run*me.bat files found.")
        return {}

    log.info("Found %d bat file(s) to run:", len(bat_files))
    for i, b in enumerate(bat_files, 1):
        log.info("  %d. %s", i, b)

    results = {}

    for i, bat_path in enumerate(bat_files, 1):
        bat_dir = os.path.dirname(bat_path)
        bat_name = os.path.basename(bat_path)
        folder_name = Path(bat_dir).name

        log.info("[%d/%d] Running: %s", i, len(bat_files), bat_path)

        try:
            # Create auto-answer file
            answer_path = os.path.join(bat_dir, "_answer.txt")
            with open(answer_path, "w") as f:
                f.write("yes\n")

            # Launch in its own window (like Start-Process)
            # Using Popen so it launches the window, then we wait
            process = subprocess.Popen(
                ["cmd.exe", "/c", bat_name],
                cwd=bat_dir,
                stdin=open(answer_path, "r"),
                creationflags=(
                    getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
                ),
            )

            # Wait for it to finish
            process.wait(timeout=600)
            log.info("[%d/%d] %s finished (exit code %d)", i, len(bat_files), folder_name, process.returncode)
            results[folder_name] = process.returncode

            # Clean up answer file
            try:
                os.remove(answer_path)
            except OSError:
                pass

        except subprocess.TimeoutExpired:
            log.error("[%d/%d] %s timed out after 600s", i, len(bat_files), folder_name)
            process.kill()
            results[folder_name] = -2
        except Exception as e:
            log.error("[%d/%d] Error running %s: %s", i, len(bat_files), folder_name, e)
            results[folder_name] = -1

        # Delay between bat files (like Start-Sleep -Seconds 5)
        if i < len(bat_files):
            log.info("  Waiting %ds before next...", delay)
            time.sleep(delay)

    log.info("Batch files complete: %d/%d ran.", len(results), len(bat_files))
    return results


# ══════════════════════════════════════════════════════════════════════════════
# STEP 3-5: SCRAPE + UPLOAD + RENAME (Playwright)
# ══════════════════════════════════════════════════════════════════════════════
def _auto_fill_email(page) -> bool:
    selectors = [
        'input[name="identifier"]',
        'input[type="email"], input[name="email"], input[id="email"]',
        'input[type="text"][placeholder*="mail" i]',
        'input[type="text"]',
    ]
    try:
        email_input = None
        for sel in selectors:
            email_input = page.query_selector(sel)
            if email_input:
                break
        if not email_input:
            return False

        email_input.fill(CFG.login_email)
        log.info("  Email auto-filled: %s", CFG.login_email)
        page.wait_for_timeout(500)

        try:
            cb = page.query_selector('input[name="rememberMe"], input[type="checkbox"]')
            if cb and not cb.is_checked():
                cb.check(force=True)
            elif not cb:
                lbl = page.query_selector(
                    'label:has-text("Keep me signed in"), label:has-text("Remember")'
                )
                if lbl:
                    lbl.click()
        except Exception:
            pass

        page.wait_for_timeout(300)
        for sel in [
            'input[type="submit"][value="Next"]',
            'button:has-text("Next")', 'button:has-text("Continue")',
            'button:has-text("Sign in")', 'button[type="submit"]',
            'input[type="submit"]',
        ]:
            btn = page.query_selector(sel)
            if btn:
                btn.click()
                log.info("  Clicked Next.")
                return True
        return True
    except Exception:
        return False


def _detect_admin_page(page, context):
    for p in [page] + list(context.pages):
        try:
            if CFG.admin_domain in p.evaluate("window.location.href"):
                return p
        except Exception:
            continue
    return None


def _wait_for_login(page, context):
    log.info("Login required...")
    email_filled = False
    for _ in range(0, CFG.login_timeout_s, 2):
        time.sleep(2)
        if not email_filled:
            try:
                if CFG.login_domain in page.evaluate("window.location.href"):
                    email_filled = _auto_fill_email(page)
            except Exception:
                pass
        admin = _detect_admin_page(page, context)
        if admin:
            log.info("  Login successful!")
            return admin
        if email_filled:
            break

    print("\n" + "=" * 60)
    print("  Type your password in the browser and click 'Verify'.")
    print("  Wait for the Admin dashboard, then press Enter here.")
    print("=" * 60)
    input("\n>>> Press Enter after logging in... ")

    for waited in range(0, 80, 2):
        admin = _detect_admin_page(page, context)
        if admin:
            log.info("  Login successful!")
            return admin
        time.sleep(2)
    log.warning("  Login timed out — continuing.")
    return page


def _navigate(page, url: str):
    for attempt in range(1, CFG.max_nav_retries + 1):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=CFG.nav_timeout_ms)
            page.wait_for_load_state("networkidle", timeout=CFG.idle_timeout_ms)
            return
        except Exception as e:
            if attempt == CFG.max_nav_retries:
                raise
            time.sleep(2 ** attempt)


def _extract_numeric_id(page, slug: str) -> Optional[str]:
    strategies = [
        lambda: next(
            (m.group(1) for el in page.query_selector_all("div")
             for text in [(el.text_content() or "").strip()]
             for m in [re.match(r"^(\d{5,})\s*/\s*\S+$", text)] if m),
            None,
        ),
        lambda: (
            m.group(1) if (m := re.search(
                r"(\d{5,})\s*/\s*" + re.escape(slug),
                page.text_content("body") or "")) else None
        ),
        lambda: (
            m.group(1) if (m := re.search(
                r"OBDA\s*ID\s*(\d{5,})",
                page.text_content("body") or "")) else None
        ),
    ]
    for s in strategies:
        try:
            r = s()
            if r:
                return r
        except Exception:
            continue
    return None


def _upload_report(page, report_path: str) -> bool:
    try:
        page.click("text=Select action", timeout=5000)
        page.wait_for_timeout(500)
        page.click("text=Upload Report", timeout=5000)
        page.wait_for_selector('input[type="file"]', state="attached", timeout=5000)
        page.wait_for_timeout(500)
        page.set_input_files('input[type="file"]', report_path)
        page.evaluate("""() => {
            const i = document.querySelector('input[type="file"]');
            if (i) i.dispatchEvent(new Event('change', { bubbles: true }));
        }""")
        page.wait_for_timeout(1500)
        submit = page.locator('button:has-text("Submit")')
        submit.wait_for(state="visible", timeout=5000)
        submit.click()

        for _ in range(60):
            page.wait_for_timeout(1000)
            body = (page.text_content("body") or "").lower()
            if any(kw in body for kw in ["complete", "successfully", "uploaded"]):
                return True
            if not page.query_selector('input[type="file"]') and "no data yet" not in body:
                return True

        page.wait_for_load_state("networkidle", timeout=30000)
        return "no data yet" not in (page.text_content("body") or "").lower()
    except Exception as e:
        log.error("    Upload error: %s", e)
        return False


def scrape_upload_rename(results: list[FolderResult]) -> None:
    """Single browser session: scrape IDs, upload reports, rename folders."""
    valid = [r for r in results if r.success]
    if not valid:
        return

    slug_map = {r.slug: r for r in valid}
    slugs = list(slug_map.keys())
    log.info("Opening browser to scrape %d OBDA(s)...", len(slugs))

    try:
        with sync_playwright() as p:
            launch_args = {"user_data_dir": CFG.browser_profile_dir, "headless": False}
            if CFG.chrome_path:
                launch_args["executable_path"] = CFG.chrome_path

            context = p.chromium.launch_persistent_context(**launch_args)
            try:
                page = context.pages[0] if context.pages else context.new_page()
                first_url = f"{CFG.obda_base_url}/{slugs[0]}"
                _navigate(page, first_url)

                cur = page.evaluate("window.location.href")
                if CFG.login_domain in cur or CFG.admin_domain not in cur:
                    page = _wait_for_login(page, context)
                    if not page:
                        log.error("Cannot scrape without login.")
                        return
                    _navigate(page, first_url)

                for i, slug in enumerate(slugs):
                    result = slug_map[slug]
                    if i > 0:
                        try:
                            _navigate(page, f"{CFG.obda_base_url}/{slug}")
                        except Exception as e:
                            log.error("[%s] Nav error: %s", slug, e)
                            continue

                    # Scrape ID
                    try:
                        obda_id = _extract_numeric_id(page, slug)
                        result.numeric_id = obda_id
                        log.info("[%s] → %s", slug, obda_id or "(not found)")
                    except Exception as e:
                        log.error("[%s] Scrape error: %s", slug, e)
                        continue

                    # Upload report
                    report = os.path.join(result.folder_path, "final-report.zip")
                    if os.path.isfile(report):
                        log.info("[%s] Uploading report...", slug)
                        result.upload_ok = _upload_report(page, report)
                        log.info("[%s] Upload: %s", slug, "OK" if result.upload_ok else "FAILED")

                    # Rename folder
                    if result.numeric_id:
                        parent = str(Path(result.folder_path).parent)
                        new_path = os.path.join(parent, result.numeric_id)
                        if os.path.exists(new_path):
                            log.warning("[%s] '%s' already exists — keeping current name.", slug, result.numeric_id)
                        else:
                            try:
                                os.rename(result.folder_path, new_path)
                                log.info("[%s] Renamed → %s", slug, result.numeric_id)
                                result.folder_path = new_path
                            except Exception as e:
                                log.error("[%s] Rename failed: %s", slug, e)

            finally:
                context.close()
    except Exception as e:
        log.error("Playwright error: %s", e)


# ══════════════════════════════════════════════════════════════════════════════
# PIPELINE
# ══════════════════════════════════════════════════════════════════════════════
def _chunk(items: list, size: int) -> list[list]:
    return [items[i : i + size] for i in range(0, len(items), size)]


def run_pipeline(
    batch_size: int = 10,
    delay: int = 5,
    dry_run: bool = False,
    skip_extract: bool = False,
    skip_run: bool = False,
    target_dir: str = None,
):
    # Determine where extracted folders live
    if target_dir:
        work_dir = target_dir
    else:
        work_dir = create_date_folder()

    log.info("Working directory: %s", work_dir)

    # ── Step 1: Extract ────────────────────────────────────────────────────
    if not skip_extract:
        log.info("═══ STEP 1: Extracting zips from %s ═══", CFG.downloads_path)
        if dry_run:
            zips = sorted(Path(CFG.downloads_path).glob("automation-*.zip"))
            log.info("Would extract %d zip(s). Dry run — skipping.", len(zips))
        else:
            extracted = extract_all_zips(work_dir)
            if not extracted and not skip_run:
                log.info("Nothing extracted — checking for existing folders...")
    else:
        log.info("═══ Skipping extraction (--skip-extract) ═══")

    # ── Find all automation folders in work_dir ────────────────────────────
    all_folders = sorted(
        str(f) for f in Path(work_dir).iterdir()
        if f.is_dir() and f.name.startswith("automation-")
    )

    if not all_folders:
        log.info("No automation-* folders found in %s", work_dir)
        return

    log.info("Found %d automation folder(s):", len(all_folders))
    for i, f in enumerate(all_folders, 1):
        log.info("  %d. %s", i, Path(f).name)

    if dry_run:
        log.info("Dry run — stopping.")
        return

    # ── Process in batches ─────────────────────────────────────────────────
    batches = _chunk(all_folders, batch_size)
    all_results: list[FolderResult] = []

    for batch_num, batch in enumerate(batches, 1):
        log.info(
            "━━ Batch %d/%d (%d folders) ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
            batch_num, len(batches), len(batch),
        )

        # Step 2: Run bat files
        if not skip_run:
            log.info("═══ STEP 2: Running bat files ═══")
            run_all_bat_files(batch, delay=delay)
        else:
            log.info("═══ Skipping bat files (--skip-run) ═══")

        # Build results for scraping
        results = []
        for folder in batch:
            folder_name = Path(folder).name
            slug = slug_from_folder(folder_name)
            results.append(FolderResult(slug=slug, folder_path=folder))

        # Steps 3-5: Scrape + upload + rename
        log.info("═══ STEP 3-5: Scrape IDs, upload reports, rename ═══")
        scrape_upload_rename(results)
        all_results.extend(results)

        if batch_num < len(batches):
            log.info("Pausing 3s before next batch...")
            time.sleep(3)

    # ── Summary ────────────────────────────────────────────────────────────
    print()
    log.info("=" * 60)
    log.info("DONE: %d folders processed in %d batch(es)", len(all_results), len(batches))
    for r in all_results:
        status = "✓" if r.numeric_id else "–"
        log.info("  %s %s%s%s", status, r.slug,
                 f" → {r.numeric_id}" if r.numeric_id else "",
                 " [uploaded]" if r.upload_ok else "")
    log.info("=" * 60)


# ── CLI ────────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="OBDA Pipeline: extract → run → scrape → upload → rename",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  python obdash.py                              Full pipeline
  python obdash.py --batch-size 10              10 at a time
  python obdash.py --skip-extract               Already extracted manually
  python obdash.py --skip-extract --skip-run    Just scrape + rename
  python obdash.py --target-dir "C:\\path\\222"   Point at a specific folder
  python obdash.py --dry-run                    Preview only
        """,
    )
    parser.add_argument("--batch-size", type=int, default=10, metavar="N",
                        help="Folders per batch (default: 10)")
    parser.add_argument("--delay", type=int, default=5, metavar="SEC",
                        help="Seconds between bat file launches (default: 5)")
    parser.add_argument("--skip-extract", action="store_true",
                        help="Skip extraction (folders already extracted)")
    parser.add_argument("--skip-run", action="store_true",
                        help="Skip running bat files (just scrape + rename)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Preview without doing anything")
    parser.add_argument("--target-dir", type=str, default=None,
                        help="Folder containing automation-* subfolders (overrides auto date folder)")
    args = parser.parse_args()

    run_pipeline(
        batch_size=args.batch_size,
        delay=args.delay,
        dry_run=args.dry_run,
        skip_extract=args.skip_extract,
        skip_run=args.skip_run,
        target_dir=args.target_dir,
    )


if __name__ == "__main__":
    main()
