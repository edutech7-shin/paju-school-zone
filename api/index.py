from __future__ import annotations

from flask import Flask, Response, jsonify, redirect, request, send_from_directory

from app.datasets import ensure_datasets
from app.server import render_index, render_progress_page
from app.services import (
    get_download,
    get_job,
    get_result_data,
    start_cleanup_worker,
    start_job,
)


app = Flask(__name__)
start_cleanup_worker()


@app.get("/")
def index() -> Response:
    try:
        ensure_datasets()
        return Response(render_index(), mimetype="text/html")
    except Exception as exc:  # pragma: no cover - runtime safe fallback
        return Response(render_index(error=f"초기화 실패: {exc}"), status=500, mimetype="text/html")


@app.post("/process")
def process_files() -> Response:
    files = []
    for item in request.files.getlist("files"):
        if not item:
            continue
        files.append((item.filename or "students.xlsx", item.read()))

    if not files:
        return Response(render_index(error="학생 주소록 엑셀 파일을 하나 이상 올려주세요."), status=400, mimetype="text/html")

    school_name = (request.form.get("school_name") or "").strip()
    base_date = (request.form.get("base_date") or "").strip()
    try:
        job_id = start_job(files, school_name=school_name, base_date=base_date)
        return Response(render_progress_page(job_id), mimetype="text/html")
    except Exception as exc:
        return Response(render_index(error=str(exc)), status=400, mimetype="text/html")


@app.get("/result/<result_id>")
def result_page(result_id: str) -> Response:
    result = get_result_data(result_id)
    if not result:
        return Response("Not Found", status=404)
    return Response(render_index(result=result), mimetype="text/html")


@app.get("/download/<result_id>/<kind>")
def download(result_id: str, kind: str):
    payload = get_download(result_id, kind)
    if not payload:
        return Response("Not Found", status=404)

    if kind == "report":
        filename = "school-zone-report.xlsx"
        display_name = "통학구역_보고서.xlsx"
    elif kind == "school_report":
        filename = "requested-school-zone-report.xlsx"
        display_name = "우리학교_통학구역별_재학생_현황.xlsx"
    elif kind == "details":
        filename = "student-processing-details.xlsx"
        display_name = "학생_처리내역.xlsx"
    else:
        filename = "school-zone-issues.xlsx"
        display_name = "학구불일치_미분류_명단.xlsx"

    response = Response(
        payload,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response.headers["Content-Disposition"] = f"attachment; filename={filename}; filename*=UTF-8''{display_name}"
    return response


@app.get("/api/jobs/<job_id>")
def get_job_state(job_id: str):
    job = get_job(job_id)
    if not job:
        return jsonify({"error": "작업을 찾을 수 없습니다."}), 404
    return jsonify(job)


@app.get("/static/<path:resource_path>")
def static_files(resource_path: str):
    return send_from_directory("static", resource_path)


@app.errorhandler(404)
def not_found(_error):
    return redirect("/")
