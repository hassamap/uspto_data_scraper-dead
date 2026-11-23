import hashlib
import io
import os
import re
import secrets
import sqlite3
import tempfile
import threading
import time
import webbrowser
from contextlib import contextmanager
from pathlib import Path
from urllib.parse import quote, urlparse

from flask import Flask, abort, jsonify, render_template, request, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "scraper_state.sqlite3"
PORT = 8765
RETRY_DELAY_SECONDS = 3
REQUEST_INTERVAL_SECONDS = 0.5
CYBERGHOST_ID = "ffbkglfijbcbgblgflchnbphjdllaogb"
CSRF_TOKEN = secrets.token_urlsafe(32)
REQUIRED_COLUMNS = ("name", "contact", "email", "mark", "serials")
GMAIL = re.compile(r"(?<![a-z0-9.!#$%&'*+/=?^_`{|}~-])[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@gmail\.com(?![a-z0-9.-])", re.I)
CASE_URL = "https://tsdr.uspto.gov/#caseNumber={}&caseSearchType=US_APPLICATION&caseType=DEFAULT&searchType=statusSearch"

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024
state_lock = threading.RLock()
stop_event = threading.Event()
worker_thread = None
current_job = None
message = "Choose an Excel file to get started."
headless_mode = True
vpn_mode = True


@contextmanager
def connect():
    connection = sqlite3.connect(DB_PATH, timeout=30)
    connection.row_factory = sqlite3.Row
    try:
        with connection:
            yield connection
    finally:
        connection.close()


def initialize():
    global current_job, message
    with connect() as db:
        db.executescript("""
            CREATE TABLE IF NOT EXISTS jobs (
                id TEXT PRIMARY KEY, filename TEXT NOT NULL, workbook BLOB NOT NULL,
                selected_at REAL NOT NULL
            );
            CREATE TABLE IF NOT EXISTS rows (
                job_id TEXT NOT NULL, row_num INTEGER NOT NULL, serial TEXT NOT NULL,
                status TEXT NOT NULL, name TEXT NOT NULL DEFAULT '',
                phone TEXT NOT NULL DEFAULT '', email TEXT NOT NULL DEFAULT '',
                error TEXT NOT NULL DEFAULT '', updated_at REAL NOT NULL DEFAULT 0,
                PRIMARY KEY (job_id, row_num)
            );
            CREATE INDEX IF NOT EXISTS rows_job_status ON rows(job_id, status);
        """)
        row = db.execute("SELECT id FROM jobs ORDER BY selected_at DESC LIMIT 1").fetchone()
        current_job = row["id"] if row else None
        message = "Previous progress restored. Start or resume when ready." if row else "Choose an Excel file to get started."


def read_spreadsheet(data):
    workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True, keep_links=False)
    try:
        sheet = workbook.worksheets[0]
        iterator = sheet.iter_rows(values_only=True)
        headers = next(iterator, ())
        indices = {}
        for index, header in enumerate(headers):
            key = str(header).strip().casefold() if header is not None else ""
            if key in REQUIRED_COLUMNS:
                if key in indices:
                    raise ValueError(f"Duplicate '{key}' column.")
                indices[key] = index
        missing = [name.title() for name in REQUIRED_COLUMNS if name not in indices]
        if missing:
            raise ValueError("Missing required columns: " + ", ".join(missing))
        records = []
        for row_num, row in enumerate(iterator, start=2):
            if not any(value is not None and str(value).strip() for value in row):
                continue
            value = row[indices["serials"]] if indices["serials"] < len(row) else None
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            serial = str(value).strip() if value is not None else ""
            valid = bool(re.fullmatch(r"[0-9]{1,8}", serial))
            records.append((row_num, serial, "pending" if valid else "invalid",
                            "" if valid else "Serial must contain 1–8 digits."))
        if not records:
            raise ValueError("The first worksheet does not contain any data rows.")
        return indices, records
    finally:
        workbook.close()


def is_running():
    return worker_thread is not None and worker_thread.is_alive()


def status_data():
    with state_lock:
        job_id = current_job
        running = is_running()
        result = {"job": None, "running": running, "stopping": running and stop_event.is_set(),
                  "headless": headless_mode, "vpn": vpn_mode, "message": message, "recent": []}
    if not job_id:
        return result
    with connect() as db:
        job = db.execute("SELECT id, filename FROM jobs WHERE id = ?", (job_id,)).fetchone()
        counts = {row["status"]: row["count"] for row in db.execute(
            "SELECT status, COUNT(*) AS count FROM rows WHERE job_id = ? GROUP BY status", (job_id,))}
        contacts = db.execute(
            "SELECT COUNT(*) FROM rows WHERE job_id = ? AND (phone != '' OR email != '')", (job_id,)).fetchone()[0]
        recent = db.execute("""SELECT row_num AS row, serial, status, name, phone, email, error
            FROM rows WHERE job_id = ? AND status != 'pending'
            ORDER BY updated_at DESC, row_num DESC LIMIT 10""", (job_id,)).fetchall()
    if job:
        total = sum(counts.values())
        result["job"] = {"id": job["id"], "filename": job["filename"], "total": total,
                         "pending": counts.get("pending", 0), "matched": counts.get("matched", 0),
                         "no_gmail": counts.get("no_gmail", 0), "errors": counts.get("error", 0),
                         "invalid": counts.get("invalid", 0), "contacts": contacts,
                         "processed": total - counts.get("pending", 0)}
        result["recent"] = [dict(row) for row in recent]
    return result


@app.before_request
def protect_local_app():
    if request.host.split(":", 1)[0].lower() not in {"127.0.0.1", "localhost"}:
        abort(403)
    if request.method == "POST" and not secrets.compare_digest(request.headers.get("X-CSRF-Token", ""), CSRF_TOKEN):
        abort(403)


@app.errorhandler(413)
def too_large(_error):
    return jsonify(error="The Excel file must be 30 MB or smaller."), 413


@app.get("/")
def index():
    return render_template("index.html", csrf_token=CSRF_TOKEN)


@app.get("/api/status")
def status():
    return jsonify(status_data())


@app.post("/api/upload")
def upload():
    global current_job, message
    with state_lock:
        if is_running():
            return jsonify(error="Stop the current run before selecting another file."), 409
        file = request.files.get("file")
        if not file or not file.filename or not file.filename.lower().endswith(".xlsx"):
            return jsonify(error="Please choose an .xlsx Excel file."), 400
        data = file.read()
        try:
            _, records = read_spreadsheet(data)
        except Exception as exc:
            return jsonify(error=f"Could not read the first worksheet: {exc}"), 400
        job_id = hashlib.sha256(data).hexdigest()
        with connect() as db:
            existing = db.execute("SELECT id FROM jobs WHERE id = ?", (job_id,)).fetchone()
            if not existing:
                db.execute("INSERT INTO jobs VALUES (?, ?, ?, ?)",
                           (job_id, Path(file.filename).name, data, time.time()))
                db.executemany("INSERT INTO rows (job_id, row_num, serial, status, error, updated_at) VALUES (?, ?, ?, ?, ?, ?)",
                               [(job_id, row, serial, state, error, time.time() if state == "invalid" else 0)
                                for row, serial, state, error in records])
            else:
                db.execute("UPDATE jobs SET selected_at = ? WHERE id = ?", (time.time(), job_id))
        current_job = job_id
        message = "Previous progress restored." if existing else "File ready. Choose a browser mode and start."
    return jsonify(status_data())


@app.post("/api/clear")
def clear():
    global current_job, message
    with state_lock:
        if is_running():
            return jsonify(error="Stop the current run before clearing saved data."), 409
        with connect() as db:
            db.execute("DELETE FROM rows")
            db.execute("DELETE FROM jobs")
        current_job = None
        message = "Saved workbook data cleared. Choose an Excel file to get started."
    return jsonify(status_data())


EXTRACT_CORRESPONDENT = """() => {
    const section = document.querySelector('span[data-sectiontitle="Attorney/Correspondence Information"]')
        ?.closest('.expand_wrapper')?.querySelector('.sectionContainer');
    if (!section) return {error: 'Correspondence section was not found.'};
    const children = Array.from(section.children);
    let active = false, found = false;
    const fields = {};
    for (const child of children) {
        const caption = child.querySelector('.caption');
        if (caption) {
            const text = caption.textContent.trim().replace(/\\s+/g, ' ');
            if (text === 'Correspondent') { active = true; found = true; continue; }
            if (text.startsWith('Correspondent -')) { found = true; break; }
            if (active) break;
        }
        if (!active) continue;
        for (const key of child.querySelectorAll('.key')) {
            const label = key.textContent.trim().replace(/\\s+/g, ' ');
            const value = key.nextElementSibling;
            if (!value || !value.classList.contains('value')) continue;
            if (label === 'Correspondent Name/Address:')
                fields.name = (value.querySelector('div')?.textContent || value.textContent).trim();
            if (label === 'Phone:') fields.phone = value.textContent.trim();
            if (label === 'Correspondent e-mail:')
                fields.emails = Array.from(value.childNodes, node => node.textContent).join(' ');
        }
    }
    if (!found) return {error: 'Unexpected correspondence section format.'};
    return {name: fields.name || '', phone: fields.phone || '', emails: fields.emails || ''};
}"""


class AccessRestricted(Exception):
    def __init__(self, status):
        self.status = status
        super().__init__(f"USPTO returned HTTP {status}")


def open_case(page, serial):
    url = CASE_URL.format(quote(serial, safe=""))
    with page.expect_response(lambda response: urlparse(response.url).hostname == "tsdr.uspto.gov" and urlparse(response.url).path == f"/statusview/sn{serial}", timeout=35000) as response_info:
        page.goto(url, wait_until="domcontentloaded", timeout=35000)
    response = response_info.value
    if response.status in (403, 429) or response.status >= 500:
        raise AccessRestricted(response.status)
    if not response.ok:
        raise ValueError(f"USPTO returned HTTP {response.status}.")


def scrape_case(page, serial):
    page.wait_for_function("""serial => Array.from(document.querySelectorAll('.key')).some(key =>
        key.textContent.trim() === 'US Serial Number:' && key.nextElementSibling?.textContent.trim() === serial)
    """, arg=serial, timeout=20000)
    section = page.locator('span[data-sectiontitle="Attorney/Correspondence Information"] a')
    section.wait_for(state="visible", timeout=20000)
    section.click(timeout=10000)
    values = page.evaluate(EXTRACT_CORRESPONDENT)
    if values.get("error"):
        raise ValueError(values["error"])
    match = GMAIL.search(values["emails"])
    return {"status": "matched" if match else "no_gmail", "name": values["name"],
            "phone": values["phone"], "email": match.group(0) if match else ""}


def launch_browser(playwright, headless):
    failures = []
    for channel in ("chrome", "msedge", None):
        try:
            return playwright.chromium.launch(headless=headless, **({"channel": channel} if channel else {}))
        except Exception as exc:
            failures.append(str(exc))
    raise RuntimeError("Could not launch Chrome or Edge. Install Google Chrome, Microsoft Edge, or run 'python -m playwright install chromium'.")


class CyberGhostUnavailable(Exception):
    pass


BLOCKED_RESOURCE_TYPES = ("image", "media", "font")


def block_heavy_resources(context):
    # Images, fonts and video are not needed to read the case data, so skipping
    # them makes each page load faster. Stylesheets and scripts still load.
    def handle(route):
        if route.request.resource_type in BLOCKED_RESOURCE_TYPES:
            route.abort()
        else:
            route.fallback()
    context.route("**/*", handle)


def cyberghost_extension():
    chrome_data = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")) / "Google" / "Chrome" / "User Data"
    manifests = list(chrome_data.glob(f"*/Extensions/{CYBERGHOST_ID}/*/manifest.json"))
    if not manifests:
        raise CyberGhostUnavailable("CyberGhost Chrome extension was not found. Install it or disable the CyberGhost option.")
    return max(manifests, key=lambda path: tuple(map(int, re.findall(r"\d+", path.parent.name)))).parent


def check_cyberghost(page):
    try:
        connected = page.get_by_test_id("power-button-connected").is_visible()
        settings = page.evaluate("""async () => {
            const proxy = await chrome.proxy.settings.get({});
            const saved = await chrome.storage.local.get('CG_proxy');
            const location = JSON.parse(saved.CG_proxy.value).state.selectedLocation;
            const pac = proxy.value.pacScript?.data || '';
            return {country: location.countryCode, control: proxy.levelOfControl,
                mode: proxy.value.mode, selectedNode: !!location.dnsname && pac.includes('HTTPS ' + location.dnsname)};
        }""")
    except PlaywrightError as exc:
        raise CyberGhostUnavailable("CyberGhost proxy could not be verified; the run was stopped.") from exc
    if not (connected and settings.get("country") == "US" and settings.get("control") == "controlled_by_this_extension"
            and settings.get("mode") == "pac_script" and settings.get("selectedNode")):
        raise CyberGhostUnavailable("CyberGhost is not connected to the United States; the run was stopped.")


def connect_cyberghost(page):
    try:
        page.goto(f"chrome-extension://{CYBERGHOST_ID}/src/popup.html", wait_until="domcontentloaded", timeout=15000)
        selected = page.get_by_test_id("selected-location")
        selected.wait_for(state="visible", timeout=15000)
        if selected.inner_text().strip() != "United States":
            selected.click()
            page.get_by_test_id("location-menu").get_by_text("United States", exact=True).click(timeout=10000)
        if not page.get_by_test_id("power-button-connected").is_visible():
            page.get_by_test_id("power-button-ready").click(timeout=10000)
            page.get_by_test_id("power-button-connected").wait_for(state="visible", timeout=15000)
    except PlaywrightError as exc:
        raise CyberGhostUnavailable("Could not connect CyberGhost to the United States; the run was stopped.") from exc
    check_cyberghost(page)


@contextmanager
def browser_session(playwright, headless, vpn):
    if not vpn:
        browser = launch_browser(playwright, headless)
        try:
            context = browser.new_context()
            block_heavy_resources(context)
            yield context, None
        finally:
            browser.close()
        return
    extension = cyberghost_extension()
    with tempfile.TemporaryDirectory(prefix="uspto-cyberghost-") as profile:
        try:
            context = playwright.chromium.launch_persistent_context(
                profile, channel="chromium", headless=headless, ignore_default_args=["--disable-extensions"],
                args=[f"--disable-extensions-except={extension}", f"--load-extension={extension}"])
        except PlaywrightError as exc:
            raise CyberGhostUnavailable("Could not launch isolated Chromium. Run 'python -m playwright install chromium --no-shell'.") from exc
        block_heavy_resources(context)
        try:
            popup = context.pages[0] if context.pages else context.new_page()
            connect_cyberghost(popup)
            yield context, popup
        finally:
            context.close()


def run_worker(job_id, headless, vpn):
    global message
    last_request = 0.0
    vpn_popup = None

    def attempt_action(page, serial, action, paced=False):
        global message
        nonlocal last_request
        for attempt in range(3):
            if stop_event.is_set():
                return False, None, "", False
            if paced:
                remaining = REQUEST_INTERVAL_SECONDS - (time.monotonic() - last_request)
                if remaining > 0 and stop_event.wait(remaining):
                    return False, None, "", False
                if vpn_popup is not None:
                    check_cyberghost(vpn_popup)
                last_request = time.monotonic()
            try:
                return True, action(page, serial), "", False
            except (AccessRestricted, PlaywrightTimeoutError) as exc:
                reason = str(exc) if isinstance(exc, AccessRestricted) else "USPTO did not respond in time"
                if attempt < 2:
                    with state_lock:
                        message = f"{reason} for {serial}. Retrying in {RETRY_DELAY_SECONDS} seconds ({attempt + 2}/3)."
                    if stop_event.wait(RETRY_DELAY_SECONDS):
                        return False, None, "", False
                    continue
                return False, None, f"{reason} after 3 attempts; the run was paused. Try again later.", True
            except PlaywrightError as exc:
                return False, None, f"Browser or connection error; the run was paused: {str(exc)[:350]}", True
            except Exception as exc:
                return False, None, str(exc)[:500], page.is_closed()

    def save_result(serial, values, error):
        with connect() as db:
            db.execute("""UPDATE rows SET status = ?, name = ?, phone = ?, email = ?,
                error = ?, updated_at = ? WHERE job_id = ? AND serial = ? AND status = 'pending'""",
                (values["status"], values["name"], values["phone"], values["email"],
                 error, time.time(), job_id, serial))

    try:
        with sync_playwright() as playwright:
            with browser_session(playwright, headless, vpn) as session:
                context, vpn_popup = session
                page = None
                while not stop_event.is_set():
                    with connect() as db:
                        row = db.execute("""SELECT serial FROM rows WHERE job_id = ? AND status = 'pending'
                            GROUP BY serial ORDER BY MIN(CASE WHEN updated_at = 0 THEN 0 ELSE 1 END), MIN(row_num)
                            LIMIT 1""", (job_id,)).fetchone()
                        errors = db.execute("SELECT COUNT(*) FROM rows WHERE job_id = ? AND status = 'error'", (job_id,)).fetchone()[0] if not row else 0
                    if not row:
                        with state_lock:
                            message = (f"Finished with {errors} errors. Download the results or retry." if errors
                                       else "Finished. Download the Excel results.")
                        break
                    serial = row["serial"]
                    with state_lock:
                        message = f"Scraping serial {serial}..."
                    if page is None or page.is_closed():
                        page = context.new_page()
                    ok, _, error, restricted = attempt_action(page, serial, open_case, paced=True)
                    if ok:
                        ok, values, error, restricted = attempt_action(page, serial, scrape_case)
                        if ok:
                            save_result(serial, values, "")
                    if not ok and not stop_event.is_set():
                        save_result(serial, {"status": "error", "name": "", "phone": "", "email": ""}, error)
                        if restricted:
                            with state_lock:
                                message = error
                            return
                if stop_event.is_set():
                    with state_lock:
                        message = "Stopped. Progress is saved; resume whenever you are ready."
    except Exception as exc:
        with state_lock:
            message = f"Scraper stopped: {exc}"


@app.post("/api/start")
def start():
    global worker_thread, message, headless_mode, vpn_mode
    with state_lock:
        if is_running():
            return jsonify(error="A scrape is already running."), 409
        if not current_job:
            return jsonify(error="Choose an Excel file first."), 400
        data = request.get_json(silent=True) or {}
        if not isinstance(data.get("headless"), bool) or not isinstance(data.get("vpn"), bool):
            return jsonify(error="Choose valid browser and VPN options."), 400
        headless_mode = data["headless"]
        vpn_mode = data["vpn"]
        with connect() as db:
            db.execute("UPDATE rows SET status = 'pending', error = '' WHERE job_id = ? AND status = 'error'", (current_job,))
            pending = db.execute("SELECT COUNT(*) FROM rows WHERE job_id = ? AND status = 'pending'", (current_job,)).fetchone()[0]
        if not pending:
            return jsonify(error="There are no remaining serials to process."), 400
        stop_event.clear()
        message = "Connecting CyberGhost in an isolated browser..." if vpn_mode else "Starting browser..."
        worker_thread = threading.Thread(target=run_worker, args=(current_job, headless_mode, vpn_mode), daemon=True)
        worker_thread.start()
    return jsonify(status_data())


@app.post("/api/stop")
def stop():
    global message
    with state_lock:
        if is_running():
            stop_event.set()
            message = "Stopping after the current serial..."
    return jsonify(status_data())


def safe_append(sheet, values):
    sheet.append(values)
    for cell in sheet[sheet.max_row]:
        if isinstance(cell.value, str):
            cell.data_type = "s"


@app.get("/api/export")
def export():
    with state_lock:
        job_id = current_job
        if not job_id:
            return jsonify(error="Choose an Excel file first."), 400
        with connect() as db:
            job = db.execute("SELECT filename, workbook FROM jobs WHERE id = ?", (job_id,)).fetchone()
            rows = db.execute("SELECT * FROM rows WHERE job_id = ? AND status IN ('matched', 'no_gmail', 'error', 'invalid') ORDER BY row_num", (job_id,)).fetchall()
    completed = {row["row_num"]: row for row in rows if row["status"] in ("matched", "no_gmail")}
    source = load_workbook(io.BytesIO(job["workbook"]), read_only=True, data_only=True, keep_links=False)
    try:
        iterator = source.worksheets[0].iter_rows(values_only=True)
        headers = next(iterator)
        columns = {str(value).strip().casefold(): index for index, value in enumerate(headers) if value is not None}
        output_indices = [index for index, value in enumerate(headers)
                          if value is None or str(value).strip().casefold() != "status"]
        output_headers = ["Phone" if index == columns["contact"] else headers[index] for index in output_indices]
        workbook = Workbook()
        results = workbook.active
        results.title = "Results"
        safe_append(results, output_headers)
        for row_num, row in enumerate(iterator, start=2):
            if row_num not in completed:
                continue
            values = list(row) + [None] * max(0, len(headers) - len(row))
            entry = completed[row_num]
            values[columns["name"]] = entry["name"]
            values[columns["contact"]] = entry["phone"]
            values[columns["email"]] = entry["email"] or "No Gmail"
            safe_append(results, [values[index] for index in output_indices])
        results.freeze_panes = "A2"
        results.auto_filter.ref = results.dimensions
        for index in range(1, len(output_headers) + 1):
            results.column_dimensions[get_column_letter(index)].width = 24
        errors = workbook.create_sheet("Errors")
        safe_append(errors, ["Input row", "Serial", "Issue"])
        for entry in rows:
            if entry["status"] in ("error", "invalid"):
                safe_append(errors, [entry["row_num"], entry["serial"], entry["error"]])
        errors.freeze_panes = "A2"
        errors.column_dimensions["A"].width = 14
        errors.column_dimensions["B"].width = 20
        errors.column_dimensions["C"].width = 80
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        filename = Path(job["filename"]).stem + "_gmail_results.xlsx"
        return send_file(output, as_attachment=True, download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    finally:
        source.close()


initialize()

if __name__ == "__main__":
    threading.Timer(1.0, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}/")).start()
    app.run(host="127.0.0.1", port=PORT, debug=False, use_reloader=False, threaded=True)
