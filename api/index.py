from __future__ import annotations

from urllib.parse import quote

from flask import Flask, Response, jsonify, request

from app.config import save_admin_config_override
from app.datasets import ensure_datasets
from app.server import (
    render_address_convert_page,
    render_address_result_page,
    render_index,
    render_progress_page,
)
from app.services import (
    clear_context_cache,
    get_download,
    get_job,
    get_result_data,
    start_cleanup_worker,
    start_job,
    start_job_address_conversion,
)
import os
from uuid import uuid4


app = Flask(__name__, static_folder="../static", static_url_path="/static")
start_cleanup_worker()


@app.get("/")
def index() -> Response:
    try:
        ensure_datasets()
        return Response(render_index(), mimetype="text/html")
    except Exception as exc:  # pragma: no cover - runtime safe fallback
        return Response(render_index(error=f"초기화 실패: {exc}"), status=500, mimetype="text/html")


@app.get("/address-convert")
def address_convert() -> Response:
    return Response(render_address_convert_page(), mimetype="text/html")


@app.post("/process-address-convert")
def process_address_convert() -> Response:
    files = []
    for item in request.files.getlist("files"):
        if not item:
            continue
        files.append((item.filename or "students.xlsx", item.read()))

    if not files:
        return Response(
            render_address_convert_page(error="학생 명렬표 엑셀 파일을 하나 이상 올려주세요."),
            status=400,
            mimetype="text/html",
        )

    try:
        job_id = start_job_address_conversion(files)
        return Response(
            render_progress_page(
                job_id,
                result_redirect_base="/result-address",
                active_nav="convert",
                shell="address",
            ),
            mimetype="text/html",
        )
    except Exception as exc:
        return Response(render_address_convert_page(error=str(exc)), status=400, mimetype="text/html")


@app.get("/result-address/<result_id>")
def address_result_page(result_id: str) -> Response:
    result = get_result_data(result_id)
    if not result or result.get("mode") != "address_convert":
        return Response("Not Found", status=404)
    return Response(render_address_result_page(result), mimetype="text/html")


@app.post("/process")
def process_files() -> Response:
    # (선택) 붙임 파일 업로드가 있으면 런타임 오버라이드로 반영한다.
    overrides: dict[str, str] = {}
    runtime_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "runtime", "sources")
    os.makedirs(runtime_dir, exist_ok=True)

    report_template = request.files.get("report_template")
    if report_template and report_template.filename:
        dest = os.path.join(runtime_dir, f"{uuid4().hex}__report_template.xlsx")
        report_template.save(dest)
        overrides["report_template_source"] = dest

    admin_district = request.files.get("admin_district")
    if admin_district and admin_district.filename:
        dest = os.path.join(runtime_dir, f"{uuid4().hex}__admin_district.xlsx")
        admin_district.save(dest)
        overrides["admin_source"] = dest

    if overrides:
        save_admin_config_override(overrides)
        clear_context_cache()
        # 다음 요청부터도 반영되도록 데이터셋을 갱신
        ensure_datasets()

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
    if result.get("mode") == "address_convert":
        return Response(render_address_result_page(result), mimetype="text/html")
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
    elif kind == "converted":
        filename = "jibun-converted.xlsx"
        display_name = "지번주소_변환결과.xlsx"
    else:
        filename = "school-zone-issues.xlsx"
        display_name = "학구불일치_미분류_명단.xlsx"

    response = Response(
        payload,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response.headers["Content-Disposition"] = (
        f"attachment; filename=\"{filename}\"; filename*=UTF-8''{quote(display_name)}"
    )
    return response


@app.get("/api/jobs/<job_id>")
def get_job_state(job_id: str):
    job = get_job(job_id)
    if not job:
        return jsonify({"error": "작업을 찾을 수 없습니다."}), 404
    return jsonify(job)


@app.errorhandler(404)
def not_found(_error):
    return Response("Not Found", status=404, mimetype="text/plain")
