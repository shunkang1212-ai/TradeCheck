"""本地网页界面服务（仅监听 127.0.0.1，无外部 CDN / 第三方依赖）。

职责：把文件读取、提取、核对、报告串成一次“选择模板 → 上传两份文件 →
预览并修正 → 确认 → 核对 → 查看/保存报告”的会话流程。所有业务逻辑都委托给
tradecheck 包，本模块只做 HTTP 编排与临时文件生命周期管理。

安全要点：
  * 只绑定 127.0.0.1，外部无法访问；
  * 上传文件仅在内存解码后落到本项目的临时会话目录，运行结束后清理；
  * 任何路径都不回写原始输入文件（见 corrections / pipeline）。
"""

from __future__ import annotations

import base64
import binascii
import json
import os
import shutil
import tempfile
import threading
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any, Dict, List, Optional

from ..corrections import Correction
from ..excelio import MAX_FILE_BYTES, WorkbookReadError
from ..models import DocumentExtraction
from ..pipeline import (
    PipelineError,
    build_html_report,
    check_pair,
    extract_pair,
    render_report_file,
)
from ..templates import TemplateError, available_templates

OUTPUT_DIR = os.path.abspath(
    os.environ.get("TRADECHECK_OUTPUT_DIR", os.path.join(os.getcwd(), "output"))
)
SESSIONS_DIR = os.path.join(OUTPUT_DIR, "sessions")

# 单用户本地工具：用模块级会话保存“当前这次上传”的提取结果与文件路径。
_SESSION_LOCK = threading.Lock()
SESSION: Dict[str, Any] = {}
MAX_REQUEST_BYTES = 2 * ((MAX_FILE_BYTES + 2) // 3 * 4) + 65536


# ---------------------------------------------------------------- 序列化


def _field_json(f) -> Dict[str, Any]:
    cell = f.cell
    return {
        "text": f.text,
        "coord": cell.coord if cell else None,
        "sheet": cell.sheet if cell else None,
        "file": cell.file_name if cell else None,
        "state": f.state,
        "formula": f.formula,
        "note": f.note,
    }


def _doc_json(doc: DocumentExtraction) -> Dict[str, Any]:
    return {
        "role": doc.role,
        "role_label": doc.role_label,
        "sheet": doc.sheet,
        "file_name": doc.file_name,
        "sha256": doc.sha256,
        "table_found": doc.table_found,
        "header": {k: _field_json(v) for k, v in doc.header.items()},
        "rows": [
            {
                "row_index": r.row_index,
                "flags": list(r.flags),
                "fields": {k: _field_json(v) for k, v in r.fields.items()},
            }
            for r in doc.rows
        ],
        "total": _field_json(doc.total) if doc.total else None,
        "issues": [i.to_dict() for i in doc.issues],
        "available_columns": list(doc.available_columns),
    }


# ---------------------------------------------------------------- 会话管理


def _new_session_dir() -> str:
    os.makedirs(SESSIONS_DIR, exist_ok=True)
    return tempfile.mkdtemp(prefix="sess_", dir=SESSIONS_DIR)


def _clear_session() -> None:
    old = SESSION.get("dir")
    SESSION.clear()
    if old and os.path.isdir(old):
        shutil.rmtree(old, ignore_errors=True)


# ---------------------------------------------------------------- 业务端点


def do_extract(payload: Dict[str, Any]) -> Dict[str, Any]:
    template_id = str(payload.get("template") or "A").strip()
    inv_b64 = payload.get("invoice_b64") or ""
    pac_b64 = payload.get("packing_b64") or ""
    inv_name = os.path.basename(str(payload.get("invoice_name") or "invoice.xlsx").replace("\\", "/"))
    pac_name = os.path.basename(str(payload.get("packing_name") or "packing.xlsx").replace("\\", "/"))

    with _SESSION_LOCK:
        _clear_session()
        for name in (inv_name, pac_name):
            if not name.casefold().endswith(".xlsx"):
                raise PipelineError("只支持 .xlsx 文件，请重新选择。​")
        if not inv_b64 or not pac_b64:
            raise PipelineError("请同时提供商业发票与装箱单两个文件。")
        sdir = _new_session_dir()
        inv_path = os.path.join(sdir, "invoice.xlsx")
        pac_path = os.path.join(sdir, "packing.xlsx")
        try:
            for path, encoded in ((inv_path, inv_b64), (pac_path, pac_b64)):
                if not isinstance(encoded, str) or len(encoded) > (MAX_FILE_BYTES + 2) // 3 * 4:
                    raise PipelineError("单份文件超过 20 MB 或上传格式无效。")
                content = base64.b64decode(encoded, validate=True)
                if len(content) > MAX_FILE_BYTES:
                    raise PipelineError("单份文件超过 20 MB。")
                with open(path, "wb") as fh:
                    fh.write(content)
            inv, pac, tpl = extract_pair(inv_path, pac_path, template_id, inv_name, pac_name)
        except Exception:
            shutil.rmtree(sdir, ignore_errors=True)
            raise
        session_id = uuid.uuid4().hex
        SESSION.update(
            session_id=session_id,
            dir=sdir,
            inv_path=inv_path,
            pac_path=pac_path,
            template_id=template_id,
            template=tpl,
            inv=inv,
            pac=pac,
        )

    return {
        "session_id": session_id,
        "template_id": template_id,
        "invoice": _doc_json(inv),
        "packing": _doc_json(pac),
        "templates": available_templates(),
    }


def _require_session(payload: Dict[str, Any]) -> None:
    if "inv" not in SESSION or payload.get("session_id") != SESSION.get("session_id"):
        raise PipelineError("当前预览已失效（可能在其他标签页重新提取过），请重新提取文件。")


def _corrections(payload: Dict[str, Any]) -> List[Correction]:
    values = payload.get("corrections") or []
    if not isinstance(values, list) or any(
        not isinstance(c, dict) or not isinstance(c.get("target"), str)
        or not isinstance(c.get("new_text", ""), str) for c in values
    ):
        raise PipelineError("修正数据格式无效，请重新提取后重试。")
    return [Correction(target=c["target"], new_text=c.get("new_text", ""), old_text="") for c in values]


def do_check(payload: Dict[str, Any]) -> Dict[str, Any]:
    with _SESSION_LOCK:
        _require_session(payload)
        inv = SESSION["inv"]
        pac = SESSION["pac"]
        tpl = SESSION["template"]
        corrections = _corrections(payload)
        res, extra = check_pair(inv, pac, tpl, corrections)

    out = res.to_dict()
    out["currency"] = extra
    return out


def do_report(payload: Dict[str, Any]) -> Dict[str, Any]:
    with _SESSION_LOCK:
        _require_session(payload)
        inv = SESSION["inv"]
        pac = SESSION["pac"]
        tpl = SESSION["template"]
        corrections = _corrections(payload)
        res, extra = check_pair(inv, pac, tpl, corrections)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    from datetime import datetime

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    out_path = os.path.join(OUTPUT_DIR, f"TradeCheck_报告_{stamp}.html")
    html = build_html_report(inv, pac, res, tpl, extra, corrections)
    with open(out_path, "x", encoding="utf-8") as fh:
        fh.write(html)
    result = res.to_dict()
    result["currency"] = extra
    return {"path": out_path, "html": html, "result": result}


# ---------------------------------------------------------------- HTTP 处理


class Handler(BaseHTTPRequestHandler):
    server_version = "TradeCheck/0.1"
    protocol_version = "HTTP/1.1"

    # 静默默认日志，避免把文件名刷到控制台
    def log_message(self, *args):  # noqa: D401
        pass

    def _send_json(self, obj: Dict[str, Any], status: int = 200) -> None:
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self) -> Dict[str, Any]:
        if self.headers.get("Content-Type", "").split(";")[0].strip() != "application/json":
            raise PipelineError("请求必须使用 JSON 格式。")
        length = int(self.headers.get("Content-Length") or 0)
        if not 0 < length <= MAX_REQUEST_BYTES:
            raise PipelineError("请求体为空或超过大小限制。")
        raw = self.rfile.read(length) if length else b""
        if not raw:
            return {}
        data = json.loads(raw.decode("utf-8"))
        if not isinstance(data, dict):
            raise PipelineError("请求数据格式无效。")
        return data

    def _local_request(self) -> bool:
        port = self.server.server_address[1]
        host = self.headers.get("Host", "")
        if host not in (f"127.0.0.1:{port}", f"localhost:{port}"):
            return False
        origin = self.headers.get("Origin")
        return origin is None or origin == f"http://{host}"

    def do_GET(self):
        if not self._local_request():
            self._send_json({"error": "只允许本机同源访问。"}, status=403)
            return
        if self.path.split("?")[0] in ("/", "/index.html"):
            self._serve_app()
        else:
            self._send_json({"error": "not found"}, status=404)

    def _serve_app(self):
        html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.html")
        try:
            with open(html_path, "r", encoding="utf-8") as fh:
                body = fh.read().encode("utf-8")
        except FileNotFoundError:
            self._send_json({"error": "app.html 缺失"}, status=500)
            return
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        self.close_connection = True
        if not self._local_request():
            self._send_json({"error": "只允许本机同源访问。"}, status=403)
            return
        path = self.path.split("?")[0]
        try:
            data = self._read_json()
            if path == "/api/extract":
                result = do_extract(data)
            elif path == "/api/check":
                result = do_check(data)
            elif path == "/api/report":
                result = do_report(data)
            else:
                self._send_json({"error": "unknown endpoint"}, status=404)
                return
            self._send_json(result)
        except (PipelineError, TemplateError, WorkbookReadError) as exc:
            self._send_json({"error": str(exc)}, status=400)
        except (ValueError, UnicodeError, binascii.Error):
            self._send_json({"error": "上传内容或修正数据无效，请检查输入。"}, status=400)
        except Exception as exc:  # pragma: no cover - 兜底，避免服务崩溃
            self._send_json({"error": "处理失败，请检查输入是否符合模板及数值范围。"}, status=500)


def run(host: str = "127.0.0.1", port: int = 8753) -> ThreadingHTTPServer:
    """启动本地服务。端口被占用时向后试探若干次。"""
    if host != "127.0.0.1":
        raise RuntimeError("本工具只允许监听 127.0.0.1。")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    last_exc = None
    for candidate in range(port, port + 20):
        try:
            httpd = ThreadingHTTPServer((host, candidate), Handler)
            actual = candidate
            break
        except OSError as exc:
            last_exc = exc
    else:
        raise RuntimeError(f"无法在 {host}:{port}~{port+19} 绑定端口：{last_exc}")

    print(f"TradeCheck 本地界面已启动： http://{host}:{actual}/")
    print("（仅本机可访问；关闭窗口或按 Ctrl+C 停止）")
    return httpd


if __name__ == "__main__":
    import sys

    p = int(sys.argv[1]) if len(sys.argv) > 1 else 8753
    srv = run(port=p)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        _clear_session()
