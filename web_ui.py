#!/usr/bin/env python3
"""Local web UI for copying webpage hyperlinks into a translated DOCX."""

from __future__ import annotations

import cgi
import html
import os
import re
import uuid
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Dict, Optional, Tuple
from urllib.parse import unquote, urlparse

from link_webpage_to_docx import process_docx_links


BASE_DIR = Path(__file__).resolve().parent
JOBS_DIR = BASE_DIR / "jobs"
API_KEY_FILE = BASE_DIR / ".openai_api_key"
PUBLIC_DEPLOYMENT = os.getenv("PUBLIC_DEPLOYMENT", "").lower() in {"1", "true", "yes"}
MAX_UPLOAD_BYTES = 60 * 1024 * 1024


def escape(value: object) -> str:
    return html.escape(str(value), quote=True)


def safe_filename(name: str) -> str:
    name = Path(name or "translated.docx").name
    name = re.sub(r"[^A-Za-z0-9가-힣._ -]+", "_", name).strip(" .")
    return name or "translated.docx"


def linked_filename(name: str) -> str:
    original = safe_filename(name)
    path = Path(original)
    stem = path.stem or "document"
    return f"{stem}_linked.docx"


def read_saved_api_key() -> str:
    if PUBLIC_DEPLOYMENT:
        return ""
    if API_KEY_FILE.exists():
        return API_KEY_FILE.read_text(encoding="utf-8").strip()
    return ""


def save_api_key(api_key: str) -> None:
    if PUBLIC_DEPLOYMENT:
        return
    API_KEY_FILE.write_text(api_key.strip(), encoding="utf-8")
    try:
        API_KEY_FILE.chmod(0o600)
    except OSError:
        pass


def clear_saved_api_key() -> None:
    if PUBLIC_DEPLOYMENT:
        return
    if API_KEY_FILE.exists():
        API_KEY_FILE.unlink()


def page(title: str, body: str) -> bytes:
    return f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{escape(title)}</title>
  <style>
    :root {{
      color-scheme: light;
      --ink: #17202a;
      --muted: #627080;
      --line: #d8dee7;
      --panel: #f6f8fb;
      --accent: #0f766e;
      --accent-dark: #115e59;
      --danger: #b42318;
      --ok: #087443;
      --white: #ffffff;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: #eef2f6;
      color: var(--ink);
    }}
    main {{
      width: min(1040px, calc(100% - 32px));
      margin: 32px auto;
    }}
    .shell {{
      background: var(--white);
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      box-shadow: 0 18px 45px rgba(30, 42, 59, 0.10);
    }}
    header {{
      display: grid;
      gap: 8px;
      padding: 28px 32px 22px;
      border-bottom: 1px solid var(--line);
      background: linear-gradient(180deg, #ffffff 0%, #f9fbfd 100%);
    }}
    h1 {{
      margin: 0;
      font-size: 28px;
      line-height: 1.2;
      letter-spacing: 0;
    }}
    p {{
      margin: 0;
      line-height: 1.6;
      color: var(--muted);
    }}
    form {{
      display: grid;
      gap: 20px;
      padding: 28px 32px 32px;
    }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 18px;
    }}
    label {{
      display: grid;
      gap: 8px;
      font-weight: 700;
      font-size: 14px;
    }}
    .hint {{
      font-weight: 500;
      color: var(--muted);
      font-size: 13px;
    }}
    input, select {{
      width: 100%;
      min-height: 44px;
      border: 1px solid #c9d2df;
      border-radius: 6px;
      padding: 10px 12px;
      font: inherit;
      color: var(--ink);
      background: #fff;
    }}
    input[type="file"] {{
      padding: 9px 12px;
    }}
    input:focus, select:focus {{
      outline: 3px solid rgba(15, 118, 110, 0.18);
      border-color: var(--accent);
    }}
    .checkrow {{
      display: flex;
      align-items: center;
      gap: 10px;
      font-weight: 700;
    }}
    .checkrow input {{
      width: 18px;
      min-height: 18px;
    }}
    .actions {{
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      align-items: center;
      padding-top: 2px;
    }}
    button, .button {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 44px;
      border: 0;
      border-radius: 6px;
      padding: 0 18px;
      background: var(--accent);
      color: white;
      font-weight: 800;
      text-decoration: none;
      cursor: pointer;
    }}
    button:hover, .button:hover {{ background: var(--accent-dark); }}
    .button.secondary {{
      background: #27384a;
    }}
    .button.secondary:hover {{
      background: #1f2d3b;
    }}
    .key-saved {{
      min-height: 82px;
      display: grid;
      align-content: center;
      gap: 6px;
      border: 1px solid #c9d2df;
      border-radius: 6px;
      padding: 12px;
      background: #f9fbfd;
      font-size: 14px;
    }}
    .key-saved strong {{
      color: var(--ink);
    }}
    .key-saved span {{
      color: var(--muted);
      font-weight: 600;
    }}
    .key-saved a {{
      color: var(--accent-dark);
      font-weight: 800;
      text-decoration: none;
    }}
    .status {{
      margin: 28px 32px 32px;
      padding: 18px;
      border-radius: 8px;
      border: 1px solid var(--line);
      background: var(--panel);
    }}
    .status.error {{
      border-color: #f4b7b0;
      background: #fff4f2;
      color: var(--danger);
    }}
    .stats {{
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      margin-top: 16px;
    }}
    .stat {{
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px;
      background: white;
    }}
    .stat strong {{
      display: block;
      font-size: 26px;
      color: var(--ok);
      line-height: 1.1;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      margin-top: 18px;
      font-size: 13px;
    }}
    th, td {{
      border-bottom: 1px solid var(--line);
      padding: 10px 8px;
      text-align: left;
      vertical-align: top;
    }}
    th {{
      color: #344256;
      background: #f9fbfd;
    }}
    .full {{ grid-column: 1 / -1; }}
    @media (max-width: 720px) {{
      main {{ width: min(100% - 20px, 1040px); margin: 16px auto; }}
      header, form {{ padding-left: 18px; padding-right: 18px; }}
      .grid, .stats {{ grid-template-columns: 1fr; }}
      .status {{ margin-left: 18px; margin-right: 18px; }}
    }}
  </style>
</head>
<body>
<main>
  <section class="shell">
    <header>
      <h1>웹페이지 링크를 Word 번역본에 넣기</h1>
      <p>영문 웹페이지 주소와 한국어 Word 파일을 넣으면, 원문의 하이퍼링크를 번역 문서의 대응 위치에 자동으로 삽입합니다.</p>
    </header>
    {body}
  </section>
</main>
</body>
</html>""".encode("utf-8")


def form_html(message: str = "") -> bytes:
    notice = f'<div class="status error">{escape(message)}</div>' if message else ""
    has_saved_key = bool(read_saved_api_key())
    has_env_key = bool(os.getenv("OPENAI_API_KEY"))
    if PUBLIC_DEPLOYMENT:
        api_key_block = """<label>
      OpenAI API 키
      <input name="api_key" type="password" placeholder="sk-..." autocomplete="off" data-api-key-input>
      <span class="hint">이 브라우저에만 저장됩니다. 서버에는 저장하지 않습니다.</span>
    </label>"""
    elif has_saved_key:
        api_key_block = """<div class="key-saved">
      <strong>OpenAI API 키</strong>
      <span>저장된 키를 사용합니다.</span>
      <a href="/api-key/reset">키 다시 입력</a>
    </div>"""
    elif has_env_key:
        api_key_block = """<div class="key-saved">
      <strong>OpenAI API 키</strong>
      <span>환경변수 OPENAI_API_KEY를 사용합니다.</span>
    </div>"""
    else:
        api_key_block = """<label>
      OpenAI API 키
      <input name="api_key" type="password" placeholder="sk-..." autocomplete="off">
      <span class="hint">처음 한 번 입력하면 이 컴퓨터에 저장됩니다.</span>
    </label>"""
    return page(
        "Web DOCX Linker",
        f"""{notice}
<form method="post" action="/process" enctype="multipart/form-data">
  <label class="full">
    영문 웹페이지 주소
    <input name="url" type="url" placeholder="https://example.com/original-page" required>
  </label>
  <label class="full">
    번역된 Word 문서
    <input name="docx" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" required>
  </label>
  <div class="grid">
    {api_key_block}
    <label>
      모델
      <input name="model" value="gpt-4.1-mini">
      <span class="hint">정확도가 더 필요하면 더 강한 모델명으로 바꿀 수 있습니다.</span>
    </label>
    <label>
      최소 신뢰도
      <input name="min_confidence" type="number" min="0" max="1" step="0.01" value="0.62">
      <span class="hint">낮추면 더 많이 삽입되고, 높이면 보수적으로 삽입됩니다.</span>
    </label>
    <label>
      최대 링크 수
      <input name="max_links" type="number" min="1" placeholder="전체 처리">
      <span class="hint">처음 테스트할 때만 5 또는 10으로 제한해보세요.</span>
    </label>
  </div>
  <label class="checkrow">
    <input name="dry_run" type="checkbox" value="1">
    Word 파일 수정 없이 리포트만 만들기
  </label>
  <div class="actions">
    <button type="submit">링크 삽입 시작</button>
    <span class="hint">처리 시간은 웹페이지 링크 수와 문서 길이에 따라 달라집니다.</span>
  </div>
</form>
<script>
(() => {{
  const keyName = "web_docx_linker_openai_key";
  const input = document.querySelector("[data-api-key-input]");
  const form = document.querySelector("form");
  if (!input || !form) return;
  const saved = localStorage.getItem(keyName);
  if (saved) {{
    input.value = saved;
    input.placeholder = "저장된 키 사용 중";
  }}
  form.addEventListener("submit", () => {{
    if (input.value.trim()) {{
      localStorage.setItem(keyName, input.value.trim());
    }}
  }});
}})();
</script>""",
    )


def result_html(job_id: str, result: Dict[str, object], dry_run: bool) -> bytes:
    rows = result.get("report_rows", [])
    preview_rows = "".join(
        f"""<tr>
  <td>{escape(row.get("source_text", ""))}</td>
  <td>{escape(row.get("target_text", ""))}</td>
  <td>{escape(row.get("confidence", ""))}</td>
  <td>{'예' if row.get("inserted") else '아니오'}</td>
</tr>"""
        for row in rows[:12]
    )
    docx_button = (
        ""
        if dry_run
        else f'<a class="button" href="/download/{escape(job_id)}/{escape(result.get("download_name", "result_linked.docx"))}">Word 파일 다운로드</a>'
    )
    return page(
        "처리 완료",
        f"""<div class="status">
  <p>처리가 완료되었습니다. 아래 파일을 내려받아 Word에서 확인하세요.</p>
  <div class="stats">
    <div class="stat"><strong>{escape(result.get("link_count", 0))}</strong>추출된 링크</div>
    <div class="stat"><strong>{escape(result.get("match_count", 0))}</strong>매칭 결과</div>
    <div class="stat"><strong>{escape(result.get("inserted_count", 0))}</strong>삽입 완료</div>
  </div>
  <div class="actions" style="margin-top:16px">
    {docx_button}
    <a class="button secondary" href="/download/{escape(job_id)}/report.csv">리포트 다운로드</a>
    <a class="button secondary" href="/">새 문서 처리</a>
  </div>
  <table>
    <thead>
      <tr><th>원문 링크 문구</th><th>번역 문서 문구</th><th>신뢰도</th><th>삽입</th></tr>
    </thead>
    <tbody>{preview_rows}</tbody>
  </table>
</div>""",
    )


def parse_number(value: Optional[str], default: float) -> float:
    try:
        return float(value) if value not in (None, "") else default
    except ValueError:
        return default


class WebDocxHandler(BaseHTTPRequestHandler):
    server_version = "WebDocxLinker/1.0"

    def do_GET(self) -> None:
        try:
            parsed = urlparse(self.path)
            if parsed.path in {"/", "/process"}:
                self.respond(HTTPStatus.OK, form_html())
                return
            if parsed.path == "/api-key/reset":
                clear_saved_api_key()
                if PUBLIC_DEPLOYMENT:
                    self.respond(HTTPStatus.OK, form_html("공개 서버에서는 API 키가 서버에 저장되지 않습니다. 브라우저 저장 키는 브라우저 설정에서 지울 수 있습니다."))
                else:
                    self.respond(HTTPStatus.OK, form_html("저장된 API 키를 지웠습니다. 새 키를 한 번만 입력해주세요."))
                return
            if parsed.path == "/favicon.ico":
                self.send_response(HTTPStatus.NO_CONTENT)
                self.end_headers()
                return
            if parsed.path.startswith("/download/"):
                self.handle_download(parsed.path)
                return
            self.respond(HTTPStatus.NOT_FOUND, form_html("요청한 페이지를 찾을 수 없습니다."))
        except Exception as exc:
            self.safe_error_response(exc)

    def do_POST(self) -> None:
        try:
            parsed = urlparse(self.path)
            if parsed.path != "/process":
                self.respond(HTTPStatus.NOT_FOUND, form_html("요청한 페이지를 찾을 수 없습니다."))
                return
            self.handle_process()
        except Exception as exc:
            self.safe_error_response(exc)

    def handle_process(self) -> None:
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length > MAX_UPLOAD_BYTES:
            self.respond(HTTPStatus.BAD_REQUEST, form_html("업로드 파일이 너무 큽니다."))
            return

        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                "CONTENT_LENGTH": str(content_length),
            },
        )
        url = self.field_value(form, "url")
        model = self.field_value(form, "model") or "gpt-4.1-mini"
        submitted_api_key = self.field_value(form, "api_key")
        if submitted_api_key:
            save_api_key(submitted_api_key)
        api_key = submitted_api_key or read_saved_api_key() or None
        min_confidence = parse_number(self.field_value(form, "min_confidence"), 0.62)
        max_links_raw = self.field_value(form, "max_links")
        max_links = int(max_links_raw) if max_links_raw and max_links_raw.isdigit() else None
        dry_run = bool(self.field_value(form, "dry_run"))

        file_item = form["docx"] if "docx" in form else None
        if not url or file_item is None or not getattr(file_item, "filename", ""):
            self.respond(HTTPStatus.BAD_REQUEST, form_html("웹페이지 주소와 Word 파일을 모두 넣어주세요."))
            return
        if not safe_filename(file_item.filename).lower().endswith(".docx"):
            self.respond(HTTPStatus.BAD_REQUEST, form_html("Word 문서는 .docx 파일이어야 합니다."))
            return
        if not api_key and not os.getenv("OPENAI_API_KEY"):
            self.respond(HTTPStatus.BAD_REQUEST, form_html("OpenAI API 키를 입력해주세요. 공개 서버에서는 키가 이 브라우저에만 저장됩니다."))
            return

        job_id = uuid.uuid4().hex
        job_dir = JOBS_DIR / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        input_docx = job_dir / safe_filename(file_item.filename)
        output_docx = job_dir / "result.docx"
        download_name = linked_filename(file_item.filename)
        report_path = job_dir / "report.csv"
        with input_docx.open("wb") as handle:
            while True:
                chunk = file_item.file.read(1024 * 1024)
                if not chunk:
                    break
                handle.write(chunk)

        try:
            result = process_docx_links(
                url=url,
                docx_path=input_docx,
                output_docx=output_docx,
                model=model,
                min_confidence=min_confidence,
                max_links=max_links,
                report_path=report_path,
                dry_run=dry_run,
                api_key=api_key,
            )
        except Exception as exc:
            self.respond(HTTPStatus.OK, form_html(f"처리 중 오류가 발생했습니다: {exc}"))
            return

        result["download_name"] = download_name
        self.respond(HTTPStatus.OK, result_html(job_id, result, dry_run))

    def handle_download(self, path: str) -> None:
        parts = [unquote(part) for part in path.split("/") if part]
        if len(parts) != 3 or parts[0] != "download":
            self.respond(HTTPStatus.NOT_FOUND, form_html("다운로드 경로가 올바르지 않습니다."))
            return
        job_id, filename = parts[1], parts[2]
        if not re.fullmatch(r"[a-f0-9]{32}", job_id):
            self.respond(HTTPStatus.NOT_FOUND, form_html("다운로드 경로가 올바르지 않습니다."))
            return
        is_report = filename == "report.csv"
        is_docx = filename.endswith(".docx") and safe_filename(filename) == filename
        if not is_report and not is_docx:
            self.respond(HTTPStatus.NOT_FOUND, form_html("다운로드 경로가 올바르지 않습니다."))
            return
        stored_name = "report.csv" if is_report else "result.docx"
        file_path = JOBS_DIR / job_id / stored_name
        if not file_path.exists():
            self.respond(HTTPStatus.NOT_FOUND, form_html("파일을 찾을 수 없습니다."))
            return
        content_type = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            if filename.endswith(".docx")
            else "text/csv; charset=utf-8"
        )
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(file_path.stat().st_size))
        self.end_headers()
        with file_path.open("rb") as handle:
            self.wfile.write(handle.read())

    @staticmethod
    def field_value(form: cgi.FieldStorage, name: str) -> str:
        item = form[name] if name in form else None
        if item is None or getattr(item, "filename", None):
            return ""
        return str(item.value).strip()

    def respond(self, status: HTTPStatus, content: bytes) -> None:
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(content)

    def safe_error_response(self, exc: Exception) -> None:
        message = f"서버 처리 중 오류가 발생했습니다: {exc}"
        print(message)
        try:
            self.respond(HTTPStatus.INTERNAL_SERVER_ERROR, form_html(message))
        except (BrokenPipeError, ConnectionResetError):
            pass

    def log_message(self, format: str, *args: object) -> None:
        print(f"{self.address_string()} - {format % args}")


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    default_host = "0.0.0.0" if PUBLIC_DEPLOYMENT else "127.0.0.1"
    parser.add_argument("--host", default=os.getenv("HOST", default_host))
    parser.add_argument("--port", default=int(os.getenv("PORT", "8765")), type=int)
    args = parser.parse_args()

    JOBS_DIR.mkdir(parents=True, exist_ok=True)
    ThreadingHTTPServer.allow_reuse_address = True
    server = ThreadingHTTPServer((args.host, args.port), WebDocxHandler)
    print(f"웹 UI가 열렸습니다: http://{args.host}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n서버를 종료합니다.")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
