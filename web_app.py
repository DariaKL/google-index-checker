#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Веб-версія перевірки індексації Google
Flask + Selenium + SSE (Server-Sent Events)
"""

import csv
import io
import json
import os
import re
import time
import uuid
import threading
from datetime import datetime
from urllib.parse import quote

from flask import Flask, render_template, request, jsonify, Response, send_file
from bs4 import BeautifulSoup

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    SELENIUM_OK = True
except ImportError:
    SELENIUM_OK = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

app = Flask(__name__, template_folder="templates")

# ── In-memory task storage ──
tasks = {}
tasks_lock = threading.Lock()


class CheckTask:
    def __init__(self, urls, delay_ms, lang):
        self.id = str(uuid.uuid4())[:8]
        self.urls = urls
        self.delay_ms = delay_ms
        self.lang = lang
        self.results = []
        self.status = "pending"  # pending | running | done | stopped
        self.current = 0
        self.total = len(urls)
        self.events = []  # SSE event queue
        self.stop_event = threading.Event()

    def push_event(self, data):
        self.events.append(data)


# ── Selenium helpers ──

def make_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-background-timer-throttling")
    options.add_argument("--disable-renderer-backgrounding")
    options.add_argument("--window-size=1280,800")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/130.0.0.0 Safari/537.36"
    )

    # Use system chromedriver if available (Docker), otherwise try webdriver_manager
    chromedriver_path = os.environ.get("CHROMEDRIVER_PATH", "")
    chrome_bin = os.environ.get("CHROME_BIN", "")

    if chrome_bin and os.path.isfile(chrome_bin):
        options.binary_location = chrome_bin

    if chromedriver_path and os.path.isfile(chromedriver_path):
        service = Service(chromedriver_path)
    else:
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            service = Service(ChromeDriverManager().install())
        except ImportError:
            service = Service()  # hope chromedriver is in PATH

    driver = webdriver.Chrome(service=service, options=options)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    driver.set_page_load_timeout(30)
    return driver


def check_indexed(driver, url, lang="uk"):
    clean = url.strip()
    if not clean.startswith(("http://", "https://")):
        clean = "https://" + clean

    query = f"site:{clean}"
    search_url = f"https://www.google.com/search?q={quote(query)}&num=10&hl={lang}"

    driver.get(search_url)

    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
    except Exception:
        pass

    time.sleep(1.5)

    html = driver.page_source
    soup = BeautifulSoup(html, "html.parser")
    plain = soup.get_text(" ", strip=True).lower()

    # CAPTCHA / consent
    if (
        "/sorry/" in driver.current_url
        or "consent.google" in driver.current_url
        or "captcha" in plain
        or "unusual traffic" in plain
        or "before you continue" in plain
    ):
        return "blocked", "", "CAPTCHA — зупиніть і спробуйте пізніше"

    def first_h3():
        h3 = soup.find("h3")
        return h3.get_text(strip=True) if h3 else ""

    domain = clean.split("//")[-1].split("/")[0].replace("www.", "")

    # No results
    no_phrases = [
        "did not match any documents",
        "your search did not match",
        "no results found",
        "не збігається жодному",
        "жодних результатів",
        "жодного документа",
        "не знайдено жодного",
        "не відповідає жодному",
        "немає результатів",
    ]
    for ph in no_phrases:
        if ph in plain:
            return "not_indexed", "", "Сторінка не в індексі Google"

    # result-stats
    stats_el = soup.find(id="result-stats")
    if stats_el:
        return "indexed", first_h3(), stats_el.get_text(strip=True)[:120]

    # result cards
    cards = soup.select("div.g, div.tF2Cxc, div[data-hveid]")
    real = []
    for c in cards:
        h3 = c.find("h3")
        a = c.find("a", href=True)
        if h3 and a and domain in a.get("href", ""):
            real.append(c)
    if real:
        h3 = real[0].find("h3")
        return "indexed", h3.get_text(strip=True) if h3 else "", \
               f"Знайдено {len(real)} результат(ів)"

    # cite contains domain
    for cite in soup.find_all("cite"):
        if domain in cite.get_text():
            parent_text = ""
            p = cite.parent
            for _ in range(5):
                if p is None:
                    break
                parent_text = p.get_text(" ", strip=True).lower()
                if "webmasters" in parent_text or "search console" in parent_text:
                    break
                p = p.parent
            else:
                return "indexed", first_h3(), "Знайдено cite з доменом"

    # result count in text
    if re.search(r"about\s[\d,]+\sresult|[\d,]+\sresult", plain):
        return "indexed", first_h3(), "Google показав результати"

    # no signal
    if "google" in plain and len(plain) > 500:
        return "not_indexed", "", "Результатів не знайдено"

    return "error", "", "Не вдалося завантажити сторінку Google"


# ── Background worker ──

def run_check(task):
    task.status = "running"
    task.push_event({"type": "started", "total": task.total})

    try:
        driver = make_driver()
    except Exception as e:
        task.status = "done"
        task.push_event({"type": "error", "message": f"Chrome error: {e}"})
        task.push_event({"type": "done"})
        return

    try:
        for i, url in enumerate(task.urls):
            if task.stop_event.is_set():
                break

            task.current = i + 1

            try:
                status, title, comment = check_indexed(driver, url, task.lang)
            except Exception as exc:
                status, title, comment = "error", "", str(exc)[:120]

            row = {
                "num": i + 1,
                "url": url,
                "status": status,
                "title": title,
                "comment": comment,
            }
            task.results.append(row)
            task.push_event({"type": "result", "row": row, "current": i + 1, "total": task.total})

            if i < task.total - 1 and not task.stop_event.is_set():
                time.sleep(task.delay_ms / 1000)
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    task.status = "done"
    task.push_event({"type": "done"})


# ── Routes ──

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/start", methods=["POST"])
def start_check():
    if not SELENIUM_OK:
        return jsonify({"error": "Selenium not installed"}), 500

    data = request.get_json()
    urls_text = data.get("urls", "")
    delay_ms = int(data.get("delay", 3000))
    lang = data.get("lang", "uk")

    lines = [ln.strip() for ln in urls_text.splitlines() if ln.strip()]
    urls = [u for u in lines if u.startswith(("http://", "https://")) or "." in u]

    if not urls:
        return jsonify({"error": "No valid URLs"}), 400

    task = CheckTask(urls, delay_ms, lang)

    with tasks_lock:
        tasks[task.id] = task

    thread = threading.Thread(target=run_check, args=(task,), daemon=True)
    thread.start()

    return jsonify({"task_id": task.id, "total": len(urls)})


@app.route("/stop/<task_id>", methods=["POST"])
def stop_check(task_id):
    with tasks_lock:
        task = tasks.get(task_id)
    if task:
        task.stop_event.set()
        return jsonify({"ok": True})
    return jsonify({"error": "Task not found"}), 404


@app.route("/events/<task_id>")
def events(task_id):
    with tasks_lock:
        task = tasks.get(task_id)
    if not task:
        return jsonify({"error": "Task not found"}), 404

    def stream():
        idx = 0
        while True:
            while idx < len(task.events):
                evt = task.events[idx]
                yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"
                idx += 1
                if evt.get("type") == "done":
                    return
            time.sleep(0.3)

    return Response(stream(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/export/<task_id>/<fmt>")
def export(task_id, fmt):
    with tasks_lock:
        task = tasks.get(task_id)
    if not task or not task.results:
        return jsonify({"error": "No results"}), 404

    label_map = {
        "indexed": "В індексі",
        "not_indexed": "Не в індексі",
        "blocked": "Заблоковано",
        "error": "Помилка",
    }

    if fmt == "csv":
        buf = io.StringIO()
        writer = csv.writer(buf)
        writer.writerow(["№", "URL", "Статус", "Заголовок", "Коментар"])
        for r in task.results:
            writer.writerow([r["num"], r["url"], label_map.get(r["status"], r["status"]),
                             r["title"], r["comment"]])
        output = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
        output.seek(0)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return send_file(output, mimetype="text/csv", as_attachment=True,
                         download_name=f"index_check_{ts}.csv")

    elif fmt == "xlsx" and OPENPYXL_OK:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Результати"
        headers = ["№", "URL", "Статус", "Заголовок", "Коментар"]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color="1A237E", end_color="1A237E", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

        fill_map = {
            "indexed": "C8E6C9", "not_indexed": "FFCDD2",
            "blocked": "FFE0B2", "error": "FFF9C4",
        }
        for r in task.results:
            ws.append([r["num"], r["url"], label_map.get(r["status"], r["status"]),
                       r["title"], r["comment"]])
            color = fill_map.get(r["status"], "F5F5F5")
            for cell in ws[ws.max_row]:
                cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

        col_widths = [6, 70, 15, 50, 50]
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True, download_name=f"index_check_{ts}.xlsx")

    return jsonify({"error": "Unsupported format"}), 400


if __name__ == "__main__":
    os.makedirs("templates", exist_ok=True)
    port = int(os.environ.get("PORT", 5000))
    print("=" * 50)
    print("  Перевірка індексації Google — Веб-версія")
    print(f"  Відкрийте: http://127.0.0.1:{port}")
    print("=" * 50)
    app.run(debug=False, host="0.0.0.0", port=port, threaded=True)
