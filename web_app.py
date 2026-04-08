#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Веб-версія перевірки індексації Google
Flask + Google Custom Search API + SSE (Server-Sent Events)
"""

import csv
import io
import json
import os
import time
import uuid
import threading
from datetime import datetime

import requests as http_requests
from flask import Flask, render_template, request, jsonify, Response, send_file

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

app = Flask(__name__, template_folder="templates")

# ── Config ──
GOOGLE_API_KEY = os.environ.get("GOOGLE_API_KEY", "")
GOOGLE_CX = os.environ.get("GOOGLE_CX", "")

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
        self.status = "pending"
        self.current = 0
        self.total = len(urls)
        self.events = []
        self.stop_event = threading.Event()

    def push_event(self, data):
        self.events.append(data)


# ── Google Custom Search API ──

def check_indexed_api(url, lang="uk"):
    """Check if URL is indexed using Google Custom Search API."""
    clean = url.strip()
    if not clean.startswith(("http://", "https://")):
        clean = "https://" + clean

    query = f"site:{clean}"
    api_url = "https://www.googleapis.com/customsearch/v1"
    params = {
        "key": GOOGLE_API_KEY,
        "cx": GOOGLE_CX,
        "q": query,
        "num": 1,
        "hl": lang,
    }

    try:
        r = http_requests.get(api_url, params=params, timeout=15)

        if r.status_code == 429:
            return "blocked", "", "API ліміт вичерпано (100/день)"

        if r.status_code == 403:
            data = r.json()
            msg = data.get("error", {}).get("message", "Forbidden")
            if "billing" in msg.lower() or "quota" in msg.lower():
                return "blocked", "", "API ліміт вичерпано"
            return "error", "", f"API помилка: {msg[:120]}"

        if r.status_code != 200:
            return "error", "", f"HTTP {r.status_code}"

        data = r.json()

        total_results = int(data.get("searchInformation", {}).get("totalResults", "0"))
        items = data.get("items", [])

        if total_results == 0 and not items:
            return "not_indexed", "", "Сторінка не в індексі Google"

        title = items[0].get("title", "") if items else ""

        return "indexed", title, f"Знайдено результати ({total_results})"

    except http_requests.Timeout:
        return "error", "", "Таймаут запиту до Google API"
    except Exception as e:
        return "error", "", f"Помилка: {str(e)[:120]}"


# ── Background worker ──

def run_check(task):
    task.status = "running"
    task.push_event({"type": "started", "total": task.total})

    if not GOOGLE_API_KEY or not GOOGLE_CX:
        task.status = "done"
        task.push_event({"type": "error", "message": "GOOGLE_API_KEY або GOOGLE_CX не налаштовані"})
        task.push_event({"type": "done"})
        return

    for i, url in enumerate(task.urls):
        if task.stop_event.is_set():
            break

        task.current = i + 1

        try:
            status, title, comment = check_indexed_api(url, task.lang)
        except Exception as exc:
            status, title, comment = "error", "", str(exc)[:120]

        # Stop on API limit
        if status == "blocked" and "ліміт" in comment:
            row = {
                "num": i + 1, "url": url, "status": status,
                "title": title, "comment": comment,
            }
            task.results.append(row)
            task.push_event({"type": "result", "row": row, "current": i + 1, "total": task.total})
            task.push_event({"type": "error", "message": "API ліміт вичерпано. Зупинено."})
            break

        row = {
            "num": i + 1, "url": url, "status": status,
            "title": title, "comment": comment,
        }
        task.results.append(row)
        task.push_event({"type": "result", "row": row, "current": i + 1, "total": task.total})

        if i < task.total - 1 and not task.stop_event.is_set():
            time.sleep(task.delay_ms / 1000)

    task.status = "done"
    task.push_event({"type": "done"})


# ── Routes ──

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/start", methods=["POST"])
def start_check():
    data = request.get_json()
    urls_text = data.get("urls", "")
    delay_ms = int(data.get("delay", 500))
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


@app.route("/save-results", methods=["POST"])
def save_results():
    """Save manually checked results for export."""
    data = request.get_json()
    results = data.get("results", [])
    if not results:
        return jsonify({"error": "No results"}), 400

    task = CheckTask([], 0, "uk")
    task.results = results
    task.status = "done"

    with tasks_lock:
        tasks[task.id] = task

    return jsonify({"task_id": task.id})


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
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=False, host="0.0.0.0", port=port, threaded=True)
