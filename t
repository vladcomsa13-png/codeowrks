$Root = "C:\Users\111067427\Downloads\222"

# Find all automation-* folders
$folders = Get-ChildItem -LiteralPath $Root -Directory -Filter "automation-*"

if ($folders.Count -eq 0) { Write-Warning "No automation-* folders found."; return }

Write-Host "Found $($folders.Count) folder(s) to process." -ForegroundColor Cyan

# Install Playwright module if needed (one-time)
# pip install playwright && playwright install chromium

$i = 0
foreach ($folder in $folders) {
    $i++
    $slug = $folder.Name -replace '^automation-', ''
    Write-Host "[$i/$($folders.Count)] $slug" -ForegroundColor Yellow

    # Run run-me.bat if it exists
    $bat = Get-ChildItem -LiteralPath $folder.FullName -Recurse -File -Filter "run*me.bat" | Select-Object -First 1
    if ($bat) {
        Write-Host "  Running $($bat.Name)..." -ForegroundColor Gray
        Start-Process -FilePath $bat.FullName -WorkingDirectory $bat.DirectoryName -Wait
    }

    if ($i -lt $folders.Count) { Start-Sleep -Seconds 5 }
}

Write-Host "`nBat files done. Starting rename..." -ForegroundColor Green

# Now run the Python renamer
python -c @"
import re, sys
from pathlib import Path
from playwright.sync_api import sync_playwright

ROOT = r'$Root'
BASE_URL = 'https://sage-admin-sw20-private-prod.yul.rbc105.prod.iac0.rbc.com/obda'
ADMIN = 'sage-admin-sw20-private-prod'
LOGIN = 'id.ownr.co'
EMAIL = 'vlad.comsa@ownr.co'
PROFILE = str(Path.home() / '.obdash-browser')

folders = sorted([f for f in Path(ROOT).iterdir() if f.is_dir() and f.name.startswith('automation-')])
if not folders:
    print('No automation-* folders to rename.')
    sys.exit(0)

slugs = [(f, f.name.replace('automation-', '')) for f in folders]
print(f'Renaming {len(slugs)} folder(s)...')

with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(user_data_dir=PROFILE, headless=False)
    page = ctx.pages[0] if ctx.pages else ctx.new_page()

    page.goto(f'{BASE_URL}/{slugs[0][1]}', wait_until='domcontentloaded', timeout=30000)
    page.wait_for_load_state('networkidle', timeout=15000)

    url = page.evaluate('window.location.href')
    if LOGIN in url or ADMIN not in url:
        # Auto-fill email
        for sel in ['input[name=\"identifier\"]', 'input[type=\"email\"]', 'input[type=\"text\"]']:
            inp = page.query_selector(sel)
            if inp:
                inp.fill(EMAIL)
                page.wait_for_timeout(500)
                for bs in ['input[type=\"submit\"]', 'button[type=\"submit\"]', 'button:has-text(\"Next\")']:
                    btn = page.query_selector(bs)
                    if btn:
                        btn.click()
                        break
                break
        print('\n' + '='*50)
        print('  Type your password and click Verify.')
        print('  Press Enter here after you see the dashboard.')
        print('='*50)
        input('>>> ')
        page.wait_for_timeout(3000)
        page.goto(f'{BASE_URL}/{slugs[0][1]}', wait_until='domcontentloaded', timeout=30000)
        page.wait_for_load_state('networkidle', timeout=15000)

    for i, (folder, slug) in enumerate(slugs):
        if i > 0:
            page.goto(f'{BASE_URL}/{slug}', wait_until='domcontentloaded', timeout=30000)
            page.wait_for_load_state('networkidle', timeout=15000)

        # Scrape numeric ID
        obda_id = None
        body = page.text_content('body') or ''
        m = re.search(r'(\d{5,})\s*/\s*' + re.escape(slug), body)
        if m:
            obda_id = m.group(1)
        else:
            m = re.search(r'OBDA\s*ID\s*(\d{5,})', body)
            if m:
                obda_id = m.group(1)
            else:
                for el in page.query_selector_all('div'):
                    txt = (el.text_content() or '').strip()
                    m2 = re.match(r'^(\d{5,})\s*/\s*\S+$', txt)
                    if m2:
                        obda_id = m2.group(1)
                        break

        # Upload report if exists
        report = folder / 'final-report.zip'
        if report.is_file():
            try:
                page.click('text=Select action', timeout=5000)
                page.wait_for_timeout(500)
                page.click('text=Upload Report', timeout=5000)
                page.wait_for_selector('input[type=\"file\"]', state='attached', timeout=5000)
                page.set_input_files('input[type=\"file\"]', str(report))
                page.wait_for_timeout(1500)
                sub = page.locator('button:has-text(\"Submit\")')
                sub.wait_for(state='visible', timeout=5000)
                sub.click()
                for _ in range(30):
                    page.wait_for_timeout(1000)
                    b = (page.text_content('body') or '').lower()
                    if any(k in b for k in ['complete','successfully','uploaded']):
                        break
                    if not page.query_selector('input[type=\"file\"]'):
                        break
                print(f'  {slug}: uploaded report')
            except Exception as e:
                print(f'  {slug}: upload failed - {e}')

        # Rename
        if obda_id:
            new_path = folder.parent / obda_id
            if new_path.exists():
                print(f'  {slug} -> {obda_id} (already exists, skipped)')
            else:
                folder.rename(new_path)
                print(f'  {slug} -> {obda_id}')
        else:
            print(f'  {slug} -> (ID not found, kept as-is)')

    ctx.close()

print('Done!')
"@

Write-Host "`nAll done!" -ForegroundColor Green
