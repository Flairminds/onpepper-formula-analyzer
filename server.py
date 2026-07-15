"""Flask backend: serves the formula-diff viewer and runs workbook comparisons on demand.

Usage:
    python server.py
    -> open http://127.0.0.1:5000/
"""

import json
import os
import re
import tempfile

from flask import Flask, abort, jsonify, request, send_file, send_from_directory
from werkzeug.utils import secure_filename

from src.compare_service import run_comparison
from src.excel_exporter import export_report_bytes

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORTS_DIR = os.path.join(BASE_DIR, "reports")
VIEWER_FILE = "formula-viewer 2.html"
REPORT_NAME_RE = re.compile(r"^comparison_report_(\d+)\.json$")

os.makedirs(REPORTS_DIR, exist_ok=True)

app = Flask(__name__)


def _existing_versions():
    versions = []
    for name in os.listdir(REPORTS_DIR):
        m = REPORT_NAME_RE.match(name)
        if m:
            versions.append(int(m.group(1)))
    return sorted(versions)


def _next_version():
    versions = _existing_versions()
    return versions[-1] + 1 if versions else 1


@app.route("/")
def index():
    return send_from_directory(BASE_DIR, VIEWER_FILE)


@app.route("/api/reports", methods=["GET"])
def list_reports():
    reports = []
    for v in _existing_versions():
        path = os.path.join(REPORTS_DIR, f"comparison_report_{v}.json")
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, json.JSONDecodeError):
            continue
        reports.append({
            "version": v,
            "file": f"comparison_report_{v}.json",
            "generated_at": data.get("generated_at"),
            "old_file": data.get("old_file"),
            "new_file": data.get("new_file"),
        })
    reports.sort(key=lambda r: r["version"], reverse=True)
    return jsonify(reports)


@app.route("/api/reports/<int:version>", methods=["GET"])
def get_report(version):
    name = f"comparison_report_{version}.json"
    if not os.path.isfile(os.path.join(REPORTS_DIR, name)):
        abort(404)
    return send_from_directory(REPORTS_DIR, name)


@app.route("/api/compare", methods=["POST"])
def compare():
    old_file = request.files.get("old_file")
    new_file = request.files.get("new_file")
    if not old_file or not new_file:
        return jsonify({"error": "Both old_file and new_file are required."}), 400

    for f in (old_file, new_file):
        if not f.filename.lower().endswith(".xlsx"):
            return jsonify({"error": f"'{f.filename}' is not an .xlsx file."}), 400

    with tempfile.TemporaryDirectory() as tmp_dir:
        old_path = os.path.join(tmp_dir, secure_filename(old_file.filename))
        new_path = os.path.join(tmp_dir, secure_filename(new_file.filename))
        old_file.save(old_path)
        new_file.save(new_path)

        try:
            report = run_comparison(old_path, new_path)
        except Exception as e:
            return jsonify({"error": f"Comparison failed: {e}"}), 500

    version = _next_version()
    out_name = f"comparison_report_{version}.json"
    with open(os.path.join(REPORTS_DIR, out_name), "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2)

    return jsonify({"version": version, "file": out_name, "report": report})


@app.route("/api/export", methods=["POST"])
def export():
    report = request.get_json(force=True, silent=True)
    if not report:
        return jsonify({"error": "No report data provided."}), 400

    try:
        buf = export_report_bytes(report)
    except Exception as e:
        return jsonify({"error": f"Export failed: {e}"}), 500

    old_name = report.get("old_file", "old").replace(".xlsx", "")
    new_name = report.get("new_file", "new").replace(".xlsx", "")
    filename = f"comparison_{old_name}_vs_{new_name}.xlsx"

    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
