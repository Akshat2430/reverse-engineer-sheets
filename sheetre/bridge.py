"""Excel bridge for sandboxed sessions.

Claude Code's macOS sandbox lets Excel answer read-only Apple Events but rejects anything that
changes state (-10004 privilege violation), so a sandboxed /reverse-engineer run cannot drive
Excel itself. Start this bridge in a normal terminal first; the sandboxed session's engine
then talks to it over one Unix socket (allowed in scoring/sealed.settings.json).

    .venv/bin/python -m sheetre.bridge            # leave it running

The bridge only opens workbooks inside <repo>/runs, never runs macros and never executes
code sent to it: it is a remote control for set / calc / get on copies.
"""
import datetime as dt
import json
import os
import socket
import socketserver
import sys
import threading
from pathlib import Path

from .engine import Book, EngineError, ExcelEngine

SOCKET = os.environ.get("SHEETRE_BRIDGE", "/tmp/sheetre-bridge.sock")
ROOT = Path(__file__).resolve().parents[1]
ALLOWED = (ROOT / "runs").resolve()


def encode(v):
    if isinstance(v, (dt.datetime, dt.date)):
        if not isinstance(v, dt.datetime):
            v = dt.datetime(v.year, v.month, v.day)
        return {"date": v.isoformat()}
    if isinstance(v, (int, float, str, bool)) or v is None:
        return v
    return str(v)


def _inside(path):
    p = Path(path).resolve()
    return p == ALLOWED or ALLOWED in p.parents


class Handler(socketserver.StreamRequestHandler):
    engine = None
    books = {}
    lock = threading.Lock()

    def handle(self):
        for line in self.rfile:
            try:
                req = json.loads(line)
                with self.lock:
                    res = self.dispatch(req.get("cmd"), req.get("args", {}))
                out = {"ok": True, "result": res}
            except Exception as e:
                out = {"ok": False, "error": f"{type(e).__name__}: {e}", "blocked": "EngineBlocked" in type(e).__name__}
            self.wfile.write((json.dumps(out, default=str) + "\n").encode())
            self.wfile.flush()

    def dispatch(self, cmd, a):
        eng = Handler.engine
        if cmd == "hello":
            return dict(eng.info, bridge=True, allowed_root=str(ALLOWED))
        if cmd == "open":
            if not (_inside(a["path"]) and _inside(a["work_dir"])):
                raise PermissionError(f"bridge only opens workbooks under {ALLOWED}")
            if Path(a["path"]).suffix.lower() not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
                raise PermissionError("bridge only opens Excel workbooks")
            book = eng.open(a["path"], a["work_dir"])
            Handler.books[book.path] = book
            return dict(path=book.path, iteration=eng.info.get("iteration"))
        book = Handler.books.get(a.get("book"))
        if book is None:
            raise EngineError("unknown book (open it first)")
        if cmd == "set":
            eng.set(book, a["values"])
            return {}
        if cmd == "calc":
            eng.calc(book, full=bool(a.get("full")))
            return {}
        if cmd == "get":
            return {k: encode(v) for k, v in eng.get(book, a["cells"]).items()}
        if cmd == "close":
            eng.close(book)
            Handler.books.pop(book.path, None)
            return {}
        raise EngineError(f"command not allowed over the bridge: {cmd}")


class Server(socketserver.ThreadingMixIn, socketserver.UnixStreamServer):
    daemon_threads = True


def main():
    if os.path.exists(SOCKET):
        try:
            s = socket.socket(socket.AF_UNIX)
            s.connect(SOCKET)
            s.close()
            sys.exit(f"a bridge is already listening on {SOCKET}")
        except OSError:
            os.unlink(SOCKET)
    Handler.engine = ExcelEngine().start()
    with Server(SOCKET, Handler) as srv:
        os.chmod(SOCKET, 0o600)
        print(f"sheetre bridge: Excel {Handler.engine.info.get('version')} on {SOCKET}; workbooks under {ALLOWED}. Ctrl-C to stop.", flush=True)
        try:
            srv.serve_forever()
        except KeyboardInterrupt:
            pass
        finally:
            Handler.engine.stop()
            if os.path.exists(SOCKET):
                os.unlink(SOCKET)


if __name__ == "__main__":
    main()
