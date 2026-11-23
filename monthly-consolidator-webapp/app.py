"""Flask front-end for the monthly consolidator.  Run:  python app.py  ->  http://127.0.0.1:5055"""
from __future__ import annotations

import os
import string
import threading
import time
import uuid
from pathlib import Path

from flask import Flask, abort, jsonify, render_template, request, send_file

import consolidator as cz

app = Flask(__name__)
JOBS: dict = {}
LOCK = threading.Lock()
OUTPUT_FILES: set = set()      # only files produced by this server may be downloaded
OUTPUT_DIRS: set = set()


def _opts():
    return cz.Options.from_dict((request.get_json(silent=True) or {}).get("options"))


def _bad(msg, code=400):
    return jsonify({"error": msg}), code


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/browse")
def browse():
    p = request.args.get("path", "").strip()
    if not p:
        if os.name == "nt":
            drives = [f"{d}:\\" for d in string.ascii_uppercase if os.path.exists(f"{d}:\\")]
            return jsonify({"path": "", "parent": None, "dirs": [{"name": d, "path": d} for d in drives]})
        p = "/"
    path = Path(p)
    if not path.is_dir():
        return _bad("Not a folder")
    dirs = []
    try:
        for d in sorted(path.iterdir(), key=lambda x: x.name.lower()):
            if d.is_dir() and not d.name.startswith((".", "$")):
                dirs.append({"name": d.name, "path": str(d)})
    except PermissionError:
        return _bad("Permission denied")
    parent = str(path.parent) if path.parent != path else ("" if os.name == "nt" else None)
    return jsonify({"path": str(path), "parent": parent, "dirs": dirs})


@app.post("/api/scan")
def scan():
    root = (request.get_json(silent=True) or {}).get("root", "").strip().strip('"')
    if not root or not Path(root).is_dir():
        return _bad("Root folder not found")
    months = cz.discover(root, _opts())
    return jsonify({"root": root, "months": months})


@app.post("/api/preview")
def preview():
    path = (request.get_json(silent=True) or {}).get("path", "")
    if not path or not Path(path).is_file():
        return _bad("File not found")
    try:
        return jsonify({"file": Path(path).name, "sheets": cz.analyze_file(path, _opts())})
    except Exception as e:  # noqa: BLE001
        return _bad(f"Could not analyse file: {e}")


@app.post("/api/run")
def run():
    body = request.get_json(silent=True) or {}
    root = body.get("root", "").strip().strip('"')
    if not root or not Path(root).is_dir():
        return _bad("Root folder not found")
    out = (body.get("out") or "").strip().strip('"') or str(Path(root) / "_consolidated")
    months = body.get("months") or None
    opts = cz.Options.from_dict(body.get("options"))
    jid = uuid.uuid4().hex[:10]
    job = {"id": jid, "status": "running", "log": [], "done": 0, "total": 0, "result": None, "error": None,
           "stop": False, "out": out, "started": time.time()}
    with LOCK:
        JOBS[jid] = job

    def log(level, msg):
        job["log"].append({"t": time.strftime("%H:%M:%S"), "level": level, "msg": msg})

    def progress(d, t):
        job["done"], job["total"] = d, t

    def work():
        try:
            res = cz.consolidate(root, out, months, opts, log, progress, lambda: job["stop"])
            for m in res["months"]:
                for o in m["outputs"]:
                    OUTPUT_FILES.add(os.path.normcase(os.path.abspath(o["path"])))
            OUTPUT_DIRS.add(os.path.normcase(os.path.abspath(out)))
            job["result"], job["status"] = res, "done"
            log("OK", "Finished.")
        except cz.Cancelled:
            job["status"] = "cancelled"
            log("WARN", "Cancelled.")
        except Exception as e:  # noqa: BLE001
            job["status"], job["error"] = "error", str(e)
            log("ERROR", f"Failed: {e}")

    threading.Thread(target=work, daemon=True).start()
    return jsonify({"job": jid})


@app.get("/api/job/<jid>")
def job_status(jid):
    job = JOBS.get(jid)
    if not job:
        return _bad("Unknown job", 404)
    since = int(request.args.get("since", 0))
    return jsonify({"status": job["status"], "done": job["done"], "total": job["total"], "log": job["log"][since:],
                    "next": len(job["log"]), "result": job["result"], "error": job["error"], "out": job["out"]})


@app.post("/api/cancel/<jid>")
def cancel(jid):
    if jid in JOBS:
        JOBS[jid]["stop"] = True
    return jsonify({"ok": True})


@app.get("/download")
def download():
    p = os.path.normcase(os.path.abspath(request.args.get("path", "")))
    if p not in OUTPUT_FILES or not os.path.isfile(p):
        abort(404)
    return send_file(p, as_attachment=True)


@app.post("/api/open-folder")
def open_folder():
    p = os.path.normcase(os.path.abspath((request.get_json(silent=True) or {}).get("path", "")))
    if p not in OUTPUT_DIRS or not os.path.isdir(p):
        return _bad("Not an output folder")
    if os.name == "nt":
        os.startfile(p)  # noqa: S606 - local convenience, path whitelisted above
    return jsonify({"ok": True})


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5055, debug=False, threaded=True)
