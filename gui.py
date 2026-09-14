# -*- coding: utf-8 -*-
"""
表格智能体 · 图形界面（聊天式）

一个本地网页版聊天界面：选择表格文件 → 输入中文查询 → 智能体处理，
聊天区实时显示处理过程（加载、每一步工具调用、表格预览、最终结果）。

仅依赖标准库（http.server），无需安装 Flask 等第三方库。

启动：
    python gui.py
    然后浏览器会自动打开 http://127.0.0.1:8000

接口：
    GET  /               返回聊天界面页面（index.html）
    GET  /api/files      返回目录下的 .xlsx 文件列表
    POST /api/chat       提交查询，以 NDJSON 流式返回处理过程（每行一个 JSON 事件）
    GET  /api/download?name=xxx.xlsx   下载结果文件
"""

import json
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs

from agent import TableAgent


def _resource_dir():
    """静态资源目录。PyInstaller 打包后资源会被解压到 sys._MEIPASS。"""
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def _data_dir():
    """用户数据目录（表格/上传/下载/结果）。打包后用可执行文件所在目录，否则用脚本目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


RESOURCE_DIR = _resource_dir()
DATA_DIR = _data_dir()
INDEX = RESOURCE_DIR / "index.html"
HOST = "127.0.0.1"
PORT = 8000


def list_xlsx():
    """列出数据目录下所有 .xlsx 文件（含智能体生成的结果文件）。"""
    return sorted(p.name for p in DATA_DIR.glob("*.xlsx"))


class Handler(BaseHTTPRequestHandler):
    def _json(self, obj, code=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _html(self):
        html = INDEX.read_text(encoding="utf-8").encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(html)))
        self.end_headers()
        self.wfile.write(html)

    def _handle_upload(self):
        qs = parse_qs(urlparse(self.path).query)
        name = qs.get("name", [""])[0]
        if not name.lower().endswith(".xlsx"):
            self._json({"error": "仅支持 .xlsx 文件"}, 400)
            return
        name = Path(name).name  # 去掉任何路径前缀，防止路径穿越
        length = int(self.headers.get("Content-Length", 0))
        data = self.rfile.read(length)
        if not data:
            self._json({"error": "文件内容为空"}, 400)
            return
        try:
            (DATA_DIR / name).write_bytes(data)
        except Exception as e:
            self._json({"error": f"保存失败：{e}"}, 500)
            return
        self._json({"ok": True, "name": name})

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path in ("/", "/index.html"):
            self._html()
        elif path == "/api/files":
            self._json({"files": list_xlsx()})
        elif path == "/api/download":
            qs = parse_qs(parsed.query)
            name = qs.get("name", [""])[0]
            fp = DATA_DIR / name
            if name.endswith(".xlsx") and fp.is_file() and fp.resolve().parent == DATA_DIR:
                data = fp.read_bytes()
                self.send_response(200)
                self.send_header("Content-Type",
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f'attachment; filename="{name}"')
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            else:
                self._json({"error": "文件不存在"}, 404)
        else:
            self._json({"error": "not found"}, 404)

    def do_POST(self):
        path = urlparse(self.path).path
        if path == "/api/upload":
            self._handle_upload()
            return
        if path != "/api/chat":
            self._json({"error": "not found"}, 404)
            return

        length = int(self.headers.get("Content-Length", 0))
        try:
            req = json.loads(self.rfile.read(length).decode("utf-8"))
        except (json.JSONDecodeError, ValueError):
            self._json({"error": "请求体不是合法 JSON"}, 400)
            return

        query = (req.get("query") or "").strip()
        files = req.get("files") or []
        if not query:
            self._json({"error": "查询不能为空"}, 400)
            return

        available = list_xlsx()
        files = [f for f in files if f.endswith(".xlsx") and (DATA_DIR / f).is_file()]
        if not files and available:
            files = [available[0]]
        if not files:
            self._json({"error": "目录下没有可处理的 .xlsx 表格"}, 400)
            return
        paths = [str(DATA_DIR / f) for f in files]

        # 流式响应：NDJSON，每行一个事件，逐条 flush
        self.send_response(200)
        self.send_header("Content-Type", "application/x-ndjson; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("X-Accel-Buffering", "no")
        self.end_headers()

        def emit(event):
            line = json.dumps(event, ensure_ascii=False) + "\n"
            try:
                self.wfile.write(line.encode("utf-8"))
                self.wfile.flush()
            except (BrokenPipeError, ConnectionResetError):
                pass

        try:
            agent = TableAgent(verbose=False)
            agent.run(paths, query, on_event=emit)
        except Exception as e:
            emit({"type": "error", "text": f"处理失败：{e}"})

    def log_message(self, *args):
        pass  # 静默访问日志，避免刷屏


def main():
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}"
    print("=" * 56)
    print("表格智能体 · 图形界面已启动")
    print(f"  地址：{url}")
    print("  提示：浏览器会自动打开，若无请在浏览器手动访问上述地址")
    print("  退出：按 Ctrl+C")
    print("=" * 56)
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已退出。")


if __name__ == "__main__":
    main()
