from __future__ import annotations

import argparse
import base64
import binascii
import json
import re
import threading
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, unquote, urlparse

from residential_calculator import (
    ResidentialCalculationError,
    calculate_residential_comparison,
)


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
TEMPLATES_DIR = ROOT / "templates"
FORMAL_REPORT_TEMPLATE = TEMPLATES_DIR / "formal-report-template.docx"
MAX_REQUEST_BYTES = 16 * 1024 * 1024
MAX_TEMPLATE_BYTES = 10 * 1024 * 1024


def _safe_filename(value: str, fallback: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip(" .")
    return value[:100] or fallback


class ReportStudioHandler(BaseHTTPRequestHandler):
    server_version = "ResidentialValuationWorkbench/1.0"

    def log_message(self, fmt: str, *args: object) -> None:
        print(f"[{self.log_date_time_string()}] {fmt % args}")

    def _send(
        self,
        status: int,
        body: bytes,
        content_type: str,
        extra_headers: dict[str, str] | None = None,
    ) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if extra_headers:
            for key, value in extra_headers.items():
                self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, status: int, payload: dict | list) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self._send(status, body, "application/json; charset=utf-8")

    def _read_json(self) -> dict:
        content_length = int(self.headers.get("Content-Length") or 0)
        if content_length <= 0:
            raise ResidentialCalculationError("请求内容为空")
        if content_length > MAX_REQUEST_BYTES:
            raise ResidentialCalculationError("请求内容过大")
        try:
            payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ResidentialCalculationError("请求不是有效的 JSON") from exc
        if not isinstance(payload, dict):
            raise ResidentialCalculationError("请求内容格式不正确")
        return payload

    def do_GET(self) -> None:
        path = unquote(urlparse(self.path).path)
        if path in ("/", "/index.html"):
            self._serve_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
            return
        if path == "/api/health":
            self._send_json(HTTPStatus.OK, {"status": "ok", "version": "1.0"})
            return
        if path.startswith("/static/"):
            requested = (STATIC_DIR / path.removeprefix("/static/")).resolve()
            if STATIC_DIR.resolve() not in requested.parents:
                self._send_json(HTTPStatus.FORBIDDEN, {"error": "禁止访问"})
                return
            content_types = {
                ".js": "text/javascript; charset=utf-8",
                ".css": "text/css; charset=utf-8",
                ".svg": "image/svg+xml",
                ".png": "image/png",
            }
            self._serve_file(
                requested,
                content_types.get(requested.suffix.lower(), "application/octet-stream"),
            )
            return
        self._send_json(HTTPStatus.NOT_FOUND, {"error": "页面不存在"})

    def _serve_file(self, path: Path, content_type: str) -> None:
        try:
            body = path.read_bytes()
        except FileNotFoundError:
            self._send_json(HTTPStatus.NOT_FOUND, {"error": "文件不存在"})
            return
        self._send(HTTPStatus.OK, body, content_type)

    def _attachment_headers(self, filename: str) -> dict[str, str]:
        return {
            "Content-Disposition": f"attachment; filename=\"report\"; filename*=UTF-8''{quote(filename)}",
            "X-Prototype-Notice": "Calculation draft; professional review required",
        }

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        try:
            payload = self._read_json()
            result = calculate_residential_comparison(payload)
            if path == "/api/calculate":
                self._send_json(HTTPStatus.OK, result)
                return

            subject = payload.get("subject") or {}
            report_number = str(subject.get("report_number") or subject.get("report_title") or "住宅估价")
            file_base = _safe_filename(report_number, "住宅估价")

            if path == "/api/export/residential-xlsx":
                from xlsx_template import TemplateFormatError, build_filled_workbook

                encoded = payload.get("template_base64")
                if not encoded:
                    raise ResidentialCalculationError("请先选择 Excel 原表")
                try:
                    template_bytes = base64.b64decode(encoded, validate=True)
                except (binascii.Error, ValueError) as exc:
                    raise ResidentialCalculationError("Excel 原表数据无效") from exc
                if len(template_bytes) > MAX_TEMPLATE_BYTES:
                    raise ResidentialCalculationError("Excel 原表超过 10MB 限制")
                try:
                    content = build_filled_workbook(template_bytes, payload, result)
                except TemplateFormatError as exc:
                    raise ResidentialCalculationError(str(exc)) from exc
                filename = f"{file_base}-住宅测算表-已填写.xlsx"
                content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            elif path == "/api/export/residential-docx":
                if FORMAL_REPORT_TEMPLATE.exists():
                    from report_template import build_template_report

                    content = build_template_report(
                        FORMAL_REPORT_TEMPLATE.read_bytes(),
                        payload,
                        result,
                    )
                else:
                    from office_export import build_residential_docx

                    content = build_residential_docx(payload, result)
                filename = f"{file_base}-正式估价报告.docx"
                content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            else:
                self._send_json(HTTPStatus.NOT_FOUND, {"error": "接口不存在"})
                return

            self._send(
                HTTPStatus.OK,
                content,
                content_type,
                self._attachment_headers(filename),
            )
        except ResidentialCalculationError as exc:
            self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(exc)})
        except (BrokenPipeError, ConnectionResetError):
            return
        except Exception as exc:
            print(f"Unhandled server error: {exc!r}")
            self._send_json(
                HTTPStatus.INTERNAL_SERVER_ERROR,
                {"error": "处理失败，请查看启动窗口中的错误信息"},
            )


def main() -> None:
    parser = argparse.ArgumentParser(description="住宅市场比较法测算工作台")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()

    server = ThreadingHTTPServer((args.host, args.port), ReportStudioHandler)
    url = f"http://{args.host}:{args.port}"
    print()
    print("住宅市场比较法测算工作台已启动")
    print(f"访问地址：{url}")
    print("按 Ctrl+C 停止服务")
    print()
    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
