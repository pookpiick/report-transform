#!/usr/bin/env python3
"""
Web UI: upload a CSV (Page, Text) and download the transformed Excel file.
"""

import io
from pathlib import Path

from flask import Flask, render_template, request, send_file, flash, redirect, url_for

from transform import TEMPLATE_PATH, transform_csv_to_workbook

app = Flask(__name__)
app.secret_key = "report-transform-secret"
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/transform", methods=["POST"])
def transform():
    if "file" not in request.files:
        flash("No file selected.", "error")
        return redirect(url_for("index"))

    file = request.files["file"]
    if not file or file.filename == "":
        flash("No file selected.", "error")
        return redirect(url_for("index"))

    if not file.filename.lower().endswith(".csv"):
        flash("Please upload a CSV file.", "error")
        return redirect(url_for("index"))

    try:
        content = file.read().decode("utf-8-sig")
    except UnicodeDecodeError:
        flash("File could not be decoded as UTF-8. Please use a UTF-8 encoded CSV.", "error")
        return redirect(url_for("index"))

    if not TEMPLATE_PATH.exists():
        flash("Template file not found. Please add comment_response_template.xlsx to the output folder.", "error")
        return redirect(url_for("index"))

    revision = request.form.get("revision") or None

    try:
        wb = transform_csv_to_workbook(io.StringIO(content), TEMPLATE_PATH, revision=revision)
    except ValueError as e:
        flash(str(e), "error")
        return redirect(url_for("index"))
    except FileNotFoundError as e:
        flash(str(e), "error")
        return redirect(url_for("index"))

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    out_name = Path(file.filename).stem + ".xlsx"
    return send_file(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=out_name,
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
