"""licenses.mhpwebserver.com - Waters Licensing Wizard PDF -> cleaned Excel.

Stateless by design: the uploaded PDF is parsed in memory and never written to
disk, and nothing is logged beyond gunicorn's access line.
"""
from __future__ import annotations

import pathlib
import re
from dataclasses import asdict

from flask import Flask, jsonify, request, send_file, send_from_directory
import io

from excel import build_csv, build_workbook
from licenses import parse_upload, remove_sqt, remove_zero_qty

MAX_BYTES = 10 * 1024 * 1024

app = Flask(__name__, static_folder="www/static", static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = MAX_BYTES


def _get_result():
    f = request.files.get("pdf")
    if f is None or not f.filename:
        return None, ("No PDF uploaded.", 400)
    data = f.read()
    try:
        res = parse_upload(data)
    except ValueError as e:
        return None, (str(e), 400)
    if request.form.get("remove_sqt") in ("1", "true", "on"):
        remove_sqt(res)
    if request.form.get("remove_zero") in ("1", "true", "on"):
        remove_zero_qty(res)
    if not res.licenses and not res.removed:
        return None, ("No license lines found - is this a Licensing Wizard PDF or Checksum .txt?", 422)
    return (res, f.filename), None


def _safe(part: str) -> str:
    return re.sub(r"[^A-Za-z0-9.-]+", "_", part.strip()).strip("_")


def _base_name(res, original: str) -> str:
    """Company_SupportID if either was supplied, else the uploaded file's stem."""
    parts = [_safe(request.form.get(k, "")) for k in ("company", "support_id", "order_number", "sold_to")]
    parts = [p for p in parts if p]
    if parts:
        return "_".join(parts)
    return _safe(re.sub(r"\.(pdf|txt)$", "", original, flags=re.I)) or "licenses"


def _details() -> list[tuple[str, str]]:
    """The optional (label, value) pairs typed on the page, blanks skipped."""
    pairs = [("Company", "company"), ("Support ID", "support_id"), ("Order Number", "order_number"), ("Sold To", "sold_to")]
    return [(label, request.form.get(k, "").strip()) for label, k in pairs if request.form.get(k, "").strip()]


def _xlsx_name(res, original):
    return f"{_base_name(res, original)}_Detailed_ESLC.xlsx"


@app.get("/")
def index():
    return send_from_directory("www", "index.html")


@app.get("/offline/ESLC.html")
def offline():
    """Single-file offline build (offline/build.py) - same parser, runs entirely in the browser."""
    return send_from_directory("www/offline", "ESLC.html", as_attachment=True, download_name=f"EmpowerSmartLicenseCleaner_v{_version()}.html")


@app.get("/favicon.ico")
def favicon():
    return send_from_directory("www/static", "favicon.svg", mimetype="image/svg+xml")


@app.post("/api/preview")
def preview():
    got, err = _get_result()
    if err:
        return jsonify(error=err[0]), err[1]
    res, name = got
    return jsonify(
        installation=res.installation,
        printed=res.printed,
        prefill={"company": res.company, "support_id": res.support_id},
        filename=_xlsx_name(res, name),
        licenses=[asdict(l) for l in res.licenses],
        removed=[asdict(r) for r in res.removed],
        unparsed=res.unparsed,
    )


@app.post("/api/convert")
def convert():
    got, err = _get_result()
    if err:
        return jsonify(error=err[0]), err[1]
    res, name = got
    return send_file(
        io.BytesIO(build_workbook(res, _details())),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=_xlsx_name(res, name),
    )


@app.post("/api/csv")
def csv_export():
    got, err = _get_result()
    if err:
        return jsonify(error=err[0]), err[1]
    res, name = got
    fname = f"{_base_name(res, name)}_Simple_ESLC.csv"
    body = build_csv(res, simple=True, details=_details()).encode("utf-8")
    return send_file(io.BytesIO(body), mimetype="text/csv", as_attachment=True, download_name=fname)


@app.errorhandler(413)
def too_large(_):
    return jsonify(error="File is larger than 10 MB."), 413


@app.get("/api/version")
def version():
    return jsonify(version=_version())


def _version() -> str:
    return (pathlib.Path(__file__).parent / "VERSION").read_text().strip()


@app.get("/healthz")
def healthz():
    return "ok"
