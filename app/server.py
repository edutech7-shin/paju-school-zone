from __future__ import annotations

try:
    import cgi
except ModuleNotFoundError:  # Python 3.13+ compatibility
    cgi = None
from html import escape
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import os
import traceback
from urllib.parse import parse_qs, urlparse
from urllib.parse import quote

from .config import STATIC_DIR, TEMPLATE_DIR, load_admin_config, save_admin_config, AdminConfig, DataSources, RuntimeSettings
from .datasets import ensure_datasets
from .services import (
    clear_context_cache,
    clear_runtime_data,
    get_context,
    get_download,
    get_job,
    get_result_data,
    start_cleanup_worker,
    start_job,
)


INDEX_TEMPLATE = os.path.join(TEMPLATE_DIR, "index.html")


def esc(value) -> str:
    return escape("" if value is None else str(value))


def render_index(result: dict | None = None, error: str = "") -> bytes:
    with open(INDEX_TEMPLATE, "r", encoding="utf-8") as file:
        html = file.read()
    config = load_admin_config()
    schools = sorted({record["school_name"] for record in get_context().school_zone_records})
    school_options = "".join(f'<option value="{esc(name)}"></option>' for name in schools)

    error_html = f'<div class="error">{esc(error)}</div>' if error else ""
    result_html = ""

    if result:
        summary = result["summary"]
        summary_html = f"""
        <section class="result-card">
          <h2>처리 결과</h2>
          <div class="summary-grid">
            <div><strong>{summary["total_students"]}</strong><span>전체 학생</span></div>
            <div><strong>{summary["exact_matches"]}</strong><span>정확 매칭</span></div>
            <div><strong>{summary["candidate_matches"]}</strong><span>후보 매칭</span></div>
            <div><strong>{summary.get("school_mismatches", 0)}</strong><span>학구 불일치</span></div>
            <div><strong>{summary["unmatched_students"]}</strong><span>미분류</span></div>
          </div>
          <div class="download-row">
            <a class="button-link" href="/download/{result["result_id"]}/report">서식 엑셀 다운로드</a>
            <a class="button-link secondary" href="/download/{result["result_id"]}/school_report">우리학교 통학구역별 재학생 현황 다운로드</a>
            <a class="button-link secondary" href="/download/{result["result_id"]}/details">학생 처리내역 다운로드</a>
            <a class="button-link secondary" href="/download/{result["result_id"]}/issues">학구 불일치/미분류 다운로드</a>
          </div>
        </section>
        """

        rows_html = "".join(
            f"""
            <tr>
              <td>{esc(row["school_name"])}</td>
              <td>{esc(row["admin_area"])}</td>
              <td>{esc(row["tong_ri"])}</td>
              <td>{esc(row["extra_area"])}</td>
              <td>{row["1학년"]}</td>
              <td>{row["2학년"]}</td>
              <td>{row["3학년"]}</td>
              <td>{row["4학년"]}</td>
              <td>{row["5학년"]}</td>
              <td>{row["6학년"]}</td>
              <td>{row["계"]}</td>
            </tr>
            """
            for row in result["table_rows"][:400]
        )
        unmatched_html = "".join(
            f"""
            <tr>
              <td>{esc(row.get("name", ""))}</td>
              <td>{esc(row.get("grade", ""))}</td>
              <td>{esc(row.get("road_address", ""))}</td>
              <td>{esc(row.get("jibun_address", ""))}</td>
              <td>{esc(row.get("admin_area", ""))}</td>
              <td>{esc(row.get("tong_ri", ""))}</td>
              <td>{esc(row.get("zone_match_status", ""))}</td>
              <td>{esc(row.get("match_status", ""))}</td>
            </tr>
            """
            for row in result["unmatched_students"]
        )
        result_html = f"""
        {summary_html}
        <section class="panel">
          <h2>브라우저 미리보기</h2>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>학교명</th><th>읍·면·동</th><th>리·통</th><th>기타</th>
                  <th>1</th><th>2</th><th>3</th><th>4</th><th>5</th><th>6</th><th>계</th>
                </tr>
              </thead>
              <tbody>{rows_html}</tbody>
            </table>
          </div>
        </section>
        <section class="panel">
          <h2>미분류 학생</h2>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>성명</th><th>학년</th><th>원본 주소</th><th>지번주소</th><th>행정동</th><th>통·리</th><th>학교매칭</th><th>행매칭</th>
                </tr>
              </thead>
              <tbody>{unmatched_html}</tbody>
            </table>
          </div>
        </section>
        """

    html = (
        html.replace("{{APP_TITLE}}", esc(config.runtime.app_title))
        .replace("{{ERROR}}", error_html)
        .replace("{{RESULT}}", result_html)
        .replace("{{SCHOOL_OPTIONS}}", school_options)
    )
    return html.encode("utf-8")


def render_progress_page(job_id: str) -> bytes:
    with open(INDEX_TEMPLATE, "r", encoding="utf-8") as file:
        html = file.read()
    config = load_admin_config()
    schools = sorted({record["school_name"] for record in get_context().school_zone_records})
    school_options = "".join(f'<option value="{escape(name)}"></option>' for name in schools)
    progress_html = f"""
    <section class="result-card">
      <h2>처리 중</h2>
      <div class="progress-meta">
        <div id="job-stage">작업을 시작했습니다.</div>
        <div id="job-count">0 / 0</div>
      </div>
      <div class="progress-bar">
        <div id="job-bar" class="progress-fill" style="width:0%"></div>
      </div>
      <p class="hint compact">처리 상태를 자동으로 새로고침합니다.</p>
    </section>
    <script>
      const jobId = "{job_id}";
      async function pollJob() {{
        const res = await fetch(`/api/jobs/${{jobId}}`);
        const data = await res.json();
        if (data.error) {{
          document.getElementById("job-stage").textContent = data.error;
          return;
        }}
        const total = data.total || 0;
        const current = data.current || 0;
        const percent = total ? Math.min(100, Math.round((current / total) * 100)) : (data.status === "completed" ? 100 : 5);
        document.getElementById("job-stage").textContent = data.stage || data.status;
        document.getElementById("job-count").textContent = `${{current}} / ${{total}}`;
        document.getElementById("job-bar").style.width = `${{percent}}%`;
        if (data.status === "completed" && data.result_id) {{
          window.location.href = `/result/${{data.result_id}}`;
          return;
        }}
        if (data.status === "failed") {{
          document.getElementById("job-stage").textContent = data.error || "처리에 실패했습니다.";
          return;
        }}
        setTimeout(pollJob, 1200);
      }}
      pollJob();
    </script>
    """
    html = (
        html.replace("{{APP_TITLE}}", esc(config.runtime.app_title))
        .replace("{{ERROR}}", "")
        .replace("{{RESULT}}", progress_html)
        .replace("{{SCHOOL_OPTIONS}}", school_options)
    )
    return html.encode("utf-8")


def render_admin_page(message: str = "", error: str = "") -> bytes:
    config = load_admin_config()
    is_vercel = bool(os.getenv("VERCEL"))
    source_readonly = "readonly" if is_vercel else ""
    source_note = "Vercel 배포에서는 로컬 원본 경로 대신 저장소의 사전 생성 JSON 데이터셋을 사용합니다." if is_vercel else "내부망 로컬 서버에서는 원본 엑셀 경로를 지정해 데이터셋을 다시 생성할 수 있습니다."
    note = f'<div class="error">{esc(error)}</div>' if error else (f'<div class="hint">{esc(message)}</div>' if message else "")
    html = f"""
    <!doctype html>
    <html lang="ko">
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>{esc(config.runtime.app_title)} 관리자</title>
        <link rel="stylesheet" href="/static/styles.css">
      </head>
      <body>
        <main class="shell">
          <section class="panel">
            <h1>관리자 설정</h1>
            <p class="hint">로그인 없이 내부망에서 쓰는 배포용 설정 화면입니다. 저장 후 다음 처리부터 즉시 반영됩니다.</p>
            <p class="hint">{esc(source_note)}</p>
            {note}
            <form method="post" action="/admin" class="stack">
              <label for="app_title">프로그램 제목</label>
              <input id="app_title" type="text" name="app_title" value="{esc(config.runtime.app_title)}" required>
              <label for="bind_host">바인드 호스트</label>
              <input id="bind_host" type="text" name="bind_host" value="{esc(config.runtime.bind_host)}" required>
              <label for="port">포트</label>
              <input id="port" type="number" name="port" value="{config.runtime.port}" required>
              <label for="result_ttl_minutes">결과 자동 삭제(분)</label>
              <input id="result_ttl_minutes" type="number" name="result_ttl_minutes" value="{config.runtime.result_ttl_minutes}" required>
              <label for="job_ttl_minutes">작업 기록 자동 삭제(분)</label>
              <input id="job_ttl_minutes" type="number" name="job_ttl_minutes" value="{config.runtime.job_ttl_minutes}" required>
              <label for="cleanup_interval_seconds">정리 주기(초)</label>
              <input id="cleanup_interval_seconds" type="number" name="cleanup_interval_seconds" value="{config.runtime.cleanup_interval_seconds}" required>
              <label for="admin_source">통·리·반 원본 파일</label>
              <input id="admin_source" type="text" name="admin_source" value="{esc(config.data_sources.admin_source)}" required {source_readonly}>
              <label for="report_template_source">보고서 서식 파일</label>
              <input id="report_template_source" type="text" name="report_template_source" value="{esc(config.data_sources.report_template_source)}" required {source_readonly}>
              <label for="school_zone_source">학교 통학구역 파일</label>
              <input id="school_zone_source" type="text" name="school_zone_source" value="{esc(config.data_sources.school_zone_source)}" required {source_readonly}>
              <label for="legacy_district_source">보조 관할구역 파일(선택)</label>
              <input id="legacy_district_source" type="text" name="legacy_district_source" value="{esc(config.data_sources.legacy_district_source)}" {source_readonly}>
              <button type="submit">설정 저장</button>
            </form>
            <hr class="admin-separator">
            <form method="post" action="/admin/clear-data" class="stack">
              <p class="hint">서버에 임시 저장된 처리 결과/다운로드 데이터/작업 기록을 즉시 삭제합니다.</p>
              <button type="submit" class="danger">서버 데이터 즉시 삭제</button>
            </form>
          </section>
        </main>
      </body>
    </html>
    """
    return html.encode("utf-8")


class AppHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            try:
                ensure_datasets()
                self._send_html(render_index())
            except Exception as exc:
                self._send_html(render_index(error=f"초기화 실패: {exc}"), status=500)
            return

        if parsed.path == "/admin":
            self._send_html(render_admin_page())
            return

        if parsed.path.startswith("/static/"):
            self._serve_static(parsed.path.replace("/static/", "", 1))
            return

        if parsed.path.startswith("/download/"):
            _, _, result_id, kind = parsed.path.split("/", 3)
            payload = get_download(result_id, kind)
            if not payload:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
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
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header(
                "Content-Disposition",
                f"attachment; filename=\"{filename}\"; filename*=UTF-8''{quote(display_name)}",
            )
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
            return

        if parsed.path.startswith("/result/"):
            _, _, result_id = parsed.path.split("/", 2)
            result = get_result_data(result_id)
            if not result:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            self._send_html(render_index(result=result))
            return

        if parsed.path.startswith("/api/jobs/"):
            _, _, _, job_id = parsed.path.split("/", 3)
            job = get_job(job_id)
            if not job:
                self._send_json({"error": "작업을 찾을 수 없습니다."}, status=404)
                return
            self._send_json(job)
            return

        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        if cgi is None:
            self._send_html(
                render_index(error="현재 Python 런타임에서 cgi 모듈이 없어 로컬 POST 처리를 사용할 수 없습니다."),
                status=500,
            )
            return
        if self.path == "/admin":
            self._handle_admin_post()
            return
        if self.path == "/admin/clear-data":
            self._handle_clear_data_post()
            return

        if self.path != "/process":
            self.send_error(HTTPStatus.NOT_FOUND)
            return

        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
            },
        )

        files = []
        if "files" in form:
            field = form["files"]
            if isinstance(field, list):
                for item in field:
                    files.append((item.filename or "students.xlsx", item.file.read()))
            else:
                files.append((field.filename or "students.xlsx", field.file.read()))

        if not files:
            self._send_html(render_index(error="학생 주소록 엑셀 파일을 하나 이상 올려주세요."), status=400)
            return

        school_name = clean_field(form.getfirst("school_name", ""))
        base_date = clean_field(form.getfirst("base_date", ""))
        try:
            job_id = start_job(files, school_name=school_name, base_date=base_date)
            self._send_html(render_progress_page(job_id))
        except Exception as exc:
            traceback.print_exc()
            self._send_html(render_index(error=str(exc)), status=400)

    def _handle_admin_post(self) -> None:
        if cgi is None:
            self._send_html(render_admin_page(error="현재 Python 런타임에서 cgi 모듈을 사용할 수 없습니다."), status=500)
            return
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
            },
        )
        try:
            config = AdminConfig(
                runtime=RuntimeSettings(
                    app_title=clean_field(form.getfirst("app_title", "")),
                    bind_host=clean_field(form.getfirst("bind_host", "")) or "0.0.0.0",
                    port=int(clean_field(form.getfirst("port", "8000")) or 8000),
                    result_ttl_minutes=int(clean_field(form.getfirst("result_ttl_minutes", "30")) or 30),
                    job_ttl_minutes=int(clean_field(form.getfirst("job_ttl_minutes", "60")) or 60),
                    cleanup_interval_seconds=int(clean_field(form.getfirst("cleanup_interval_seconds", "300")) or 300),
                ),
                data_sources=DataSources(
                    admin_source=clean_field(form.getfirst("admin_source", "")),
                    report_template_source=clean_field(form.getfirst("report_template_source", "")),
                    school_zone_source=clean_field(form.getfirst("school_zone_source", "")),
                    legacy_district_source=clean_field(form.getfirst("legacy_district_source", "")),
                ),
            )
            save_admin_config(config)
            clear_context_cache()
            ensure_datasets()
            self._send_html(render_admin_page(message="설정을 저장했습니다. 새 네트워크 주소/포트는 서버 재시작 후 반영됩니다."))
        except Exception as exc:
            traceback.print_exc()
            self._send_html(render_admin_page(error=f"설정 저장 실패: {exc}"), status=400)

    def _handle_clear_data_post(self) -> None:
        try:
            stats = clear_runtime_data()
            self._send_html(
                render_admin_page(
                    message=(
                        f"서버 데이터를 삭제했습니다. "
                        f"(결과 {stats['results']}건, 다운로드 {stats['downloads']}건, 작업기록 {stats['jobs']}건)"
                    )
                )
            )
        except Exception as exc:
            traceback.print_exc()
            self._send_html(render_admin_page(error=f"서버 데이터 삭제 실패: {exc}"), status=400)

    def _serve_static(self, relative_path: str) -> None:
        normalized = os.path.normpath(relative_path).lstrip(os.sep)
        path = os.path.join(STATIC_DIR, normalized)
        if not path.startswith(STATIC_DIR) or not os.path.isfile(path):
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        content_type = "text/css; charset=utf-8" if path.endswith(".css") else "text/plain; charset=utf-8"
        with open(path, "rb") as file:
            data = file.read()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_html(self, body: bytes, status: int = 200) -> None:
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, payload: dict, status: int = 200) -> None:
        import json
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def clean_field(value: str) -> str:
    return str(value or "").strip()


def run() -> None:
    config = load_admin_config()
    start_cleanup_worker()
    server = ThreadingHTTPServer((config.runtime.bind_host, config.runtime.port), AppHandler)
    print(f"Web app running at http://{config.runtime.bind_host}:{config.runtime.port}")
    server.serve_forever()
