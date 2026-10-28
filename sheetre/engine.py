"""Recalculation adapter. One interface, two backends.

ExcelEngine drives installed Excel through xlwings (appscript on macOS, COM on Windows).
Excel runs inside a worker process so a modal dialog in Excel (which blocks every Apple
Event / COM call) becomes a clear EngineBlocked error after a timeout instead of a hang.

The engine only ever opens COPIES placed in the run's work folder and never saves them.
"""
import multiprocessing as mp
import os
import platform
import shutil
import subprocess
import traceback
from collections import defaultdict
from pathlib import Path

from . import refs as R
from .xlsxml import Package


class EngineError(Exception):
    pass


class EngineUnavailable(EngineError):
    pass


class EngineBlocked(EngineError):
    pass


class EngineBusy(EngineError):
    pass


EXCEL_LOCK = os.environ.get("SHEETRE_EXCEL_LOCK", "/tmp/sheetre-excel.lock")
# macOS Excel is sandboxed: every new file outside its container raises a "Grant File Access" dialog
# (folder grants don't stick), so Excel's working copies live inside the container instead
EXCEL_HOME = Path.home() / "Library/Containers/com.microsoft.Excel/Data/sheetre"


def _acquire_excel_lock():
    """One stage drives Excel at a time; a second gets EngineBusy at once instead of corrupting the first."""
    try:
        import fcntl
    except ImportError:                      # Windows: COM serialises calls; skip the advisory lock
        return None
    f = open(EXCEL_LOCK, "w")
    try:
        fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        f.close()
        raise EngineBusy("Excel is busy with another stage (probe, try or prove). Wait for it to finish, then run this stage.")
    return f


def safe_name(name):
    stem, ext = os.path.splitext(name)
    return "".join(ch if ch.isalnum() else "_" for ch in stem)[:60] + ext


def system_locale():
    try:
        if platform.system() == "Darwin":
            return subprocess.run(["defaults", "read", "-g", "AppleLocale"], capture_output=True, text=True).stdout.strip()
        import locale
        return locale.getlocale()[0] or ""
    except Exception:
        return ""


# ---------------------------------------------------------------- Excel worker (child process)

def _group_runs(cells):
    """{sheet: {col: [(r1, r2, [keys])]}} for contiguous vertical runs."""
    by = defaultdict(lambda: defaultdict(list))
    for k in cells:
        s, c, r = R.split_key(k)
        by[s][c].append((r, k))
    out = {}
    for s, cols in by.items():
        out[s] = {}
        for c, items in cols.items():
            items.sort()
            runs, cur = [], []
            for r, k in items:
                if cur and r != cur[-1][0] + 1:
                    runs.append(cur)
                    cur = []
                cur.append((r, k))
            if cur:
                runs.append(cur)
            out[s][c] = runs
    return out


CELL_ERRORS = {"#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NUM!", "#NAME?", "#NULL!", "#SPILL!", "#CALC!", "#GETTING_DATA"}
MAC = platform.system() == "Darwin"


def _flat(v):
    if not isinstance(v, list):
        return [v]
    return [row[0] if isinstance(row, list) else row for row in v]


def _clean(v):
    if v is None or type(v).__name__ in ("Keyword", "AEType") or v == "":
        return None                                 # appscript k.missing_value / empty
    return v


def _excel_worker(conn):
    import datetime as dt
    import xlwings as xw
    books, app, saved = {}, None, {}

    def to_excel(v):
        if isinstance(v, str):
            return "'" + v           # keep text as text: stops Excel turning "12,345" or "DEC1" into numbers/dates
        if isinstance(v, dict) and "date" in v:
            return dt.datetime.fromisoformat(v["date"])
        return v

    def handle(cmd, args):
        nonlocal app
        if cmd == "start":
            app = xw.apps.active if xw.apps.count else xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False
            stale = [b.name for b in app.books if b.name.startswith("engine_")]
            for b in list(app.books):                 # left open by a crashed earlier run
                if b.name.startswith("engine_"):
                    b.close()
            return dict(version=app.version, platform=platform.system(), closed_stale=stale)
        if cmd == "open":
            path, iterate = args
            b = app.books.open(path, update_links=False)
            books[path] = b
            if "calculation" not in saved:          # macOS can't report this until a book is open
                saved["calculation"] = app.calculation
            app.calculation = "manual"
            if iterate is not None:
                try:
                    app.api.iteration.set(iterate["on"]) if platform.system() == "Darwin" else setattr(app.api, "Iteration", iterate["on"])
                    if iterate["on"]:
                        if platform.system() == "Darwin":
                            app.api.max_iterations.set(iterate["count"])
                            app.api.max_change.set(iterate["delta"])
                        else:
                            app.api.MaxIterations, app.api.MaxChange = iterate["count"], iterate["delta"]
                except Exception as e:
                    return dict(warning=f"could not set iteration: {e}")
            return {}
        if cmd == "set":
            path, values = args
            b = books[path]
            for s, cols in _group_runs(values.keys()).items():
                sh = b.sheets[s]
                for c, runs in cols.items():
                    for run in runs:
                        vals = [to_excel(values[k]) for _, k in run]
                        rng = sh.range((run[0][0], c), (run[-1][0], c))
                        if MAC:      # raw Apple Event: ~60x faster than xlwings' converter on macOS
                            rng.api.value.set([["" if v is None else v] for v in vals])
                        else:
                            rng.options(transpose=True).value = vals
            return {}
        if cmd == "calc":
            full = args[0]
            if full:
                try:
                    app.api.calculate_full() if platform.system() == "Darwin" else app.api.CalculateFull()
                except Exception:
                    app.calculate()
            else:
                app.calculate()
            return {}
        if cmd == "get":
            path, cells = args
            b = books[path]
            out = {}
            for s, cols in _group_runs(cells).items():
                sh = b.sheets[s]
                for c, runs in cols.items():
                    for run in runs:
                        rng = sh.range((run[0][0], c), (run[-1][0], c))
                        if MAC:
                            # xlwings 0.37 err_to_str breaks on multi-cell ranges on macOS (ensure_2d returns
                            # None), so read raw values and display strings and take errors from the strings
                            raw = _flat(rng.api.value.get())
                            txt = _flat(rng.api.string_value.get())
                            if len(raw) != len(run) or len(txt) != len(run):
                                # Excel collapses some ranges (e.g. all errors) to one value: read singly
                                saved.setdefault("shape_fallbacks", 0)
                                saved["shape_fallbacks"] += 1
                                raw, txt = [], []
                                for r, _ in run:
                                    cell = sh.range((r, c)).api
                                    raw.append(cell.value.get())
                                    txt.append(cell.string_value.get())
                            for (_, k), v, t in zip(run, raw, txt):
                                out[k] = t if t in CELL_ERRORS else _clean(v)
                        else:
                            vals = rng.options(ndim=1, err_to_str=True).value
                            for (_, k), v in zip(run, vals):
                                out[k] = v
            return out
        if cmd == "macro":
            path, name, margs = args
            return dict(result=books[path].macro(name)(*margs))
        if cmd == "ping":
            return dict(version=str(app.version), books=[b.name for b in app.books])
        if cmd == "close":
            path = args[0]
            b = books.pop(path, None)
            if b is not None:
                b.close()
            if not books and "calculation" in saved:
                app.calculation = saved["calculation"]
                app.screen_updating = True
            return {}
        raise EngineError(f"unknown command {cmd}")

    while True:
        try:
            cmd, args = conn.recv()
        except EOFError:
            break
        if cmd == "quit":
            break
        try:
            conn.send(("ok", handle(cmd, args)))
        except Exception as e:
            conn.send(("err", f"{type(e).__name__}: {e}\n{traceback.format_exc(limit=3)}"))


# ---------------------------------------------------------------- public engines

class Book:
    def __init__(self, engine, work_path, original):
        self.engine, self.path, self.original = engine, str(work_path), str(original)


class ExcelEngine:
    name = "excel"

    def __init__(self, timeout=120, lock=True):
        self.timeout = timeout
        self.proc = self.conn = None
        self.info = {}
        self.use_lock = lock
        self._lockf = None

    @staticmethod
    def detect():
        try:
            import xlwings  # noqa: F401
        except ImportError:
            return False, "xlwings not installed"
        if platform.system() == "Darwin":
            # Presence only. Whether Excel will actually accept automation is settled by
            # start(), which needs a live handshake to find out.
            ok = Path("/Applications/Microsoft Excel.app").exists()
            return ok, None if ok else "Microsoft Excel not found in /Applications"
        if platform.system() == "Windows":
            return True, None
        return False, "Excel automation needs macOS or Windows"

    def _call(self, cmd, *args, timeout=None):
        if self.proc is None or not self.proc.is_alive():
            raise EngineError("Excel worker is not running")
        self.conn.send((cmd, args))
        limit, waited = timeout or self.timeout, 0.0
        while not self.conn.poll(0.5):
            waited += 0.5
            if not self.proc.is_alive():
                self.proc = None
                raise EngineError(f"Excel worker process exited during '{cmd}' (exit code crash, not a dialog)")
            if waited < limit:
                continue
            self.proc.terminate()
            self.proc = None
            raise EngineBlocked(f"Excel did not answer '{cmd}' within {timeout or self.timeout}s. Excel is probably "
                                "showing a dialog (VBA error, Save As, update links, repair). Close it in Excel, "
                                "then rerun this stage.")
        status, payload = self.conn.recv()
        if status == "err":
            raise EngineError(f"Excel '{cmd}' failed: {payload}")
        return payload

    def start(self):
        ok, why = self.detect()
        if not ok:
            raise EngineUnavailable(why)
        if self.use_lock:
            self._lockf = _acquire_excel_lock()
        ctx = mp.get_context("spawn")
        self.conn, child = ctx.Pipe()
        self.proc = ctx.Process(target=_excel_worker, args=(child,), daemon=True)
        self.proc.start()
        try:
            self.info = self._call("start", timeout=180)
        except EngineBlocked:
            raise
        except EngineError as e:
            # Excel is installed but cannot be driven from this session. The usual cause on
            # macOS is no window server (a headless / non-interactive shell), where Excel
            # answers 'get version' but refuses display_alerts and 'open workbook' with
            # OSERROR -10004. Report it as unavailable so callers fall through to the next
            # engine and the no-engine path runs, instead of a raw traceback.
            self.stop()
            # Keep only the operator-useful lines; the appscript traceback is noise here.
            keep = [ln.strip() for ln in str(e).splitlines()
                    if ln.strip().startswith(("OSERROR", "MESSAGE", "COMMAND"))]
            brief = " ".join(dict.fromkeys(keep)) or str(e).splitlines()[0]
            raise EngineUnavailable(
                f"Excel is installed but will not accept automation in this session ({brief}). "
                "If this is a headless or non-interactive shell, rerun this stage from an "
                "interactive macOS terminal that has a window server and Automation consent "
                "for Excel.") from e
        self.info.update(engine="excel", locale=system_locale())
        return self

    def open(self, path, work_dir):
        """Copy path into work_dir (ASCII name) and open the copy. Returns a Book."""
        work_dir = Path(work_dir)
        if platform.system() == "Darwin":
            work_dir = EXCEL_HOME / f"{work_dir.resolve().parent.name}_{work_dir.name}"
        work_dir.mkdir(parents=True, exist_ok=True)
        n = len(list(work_dir.glob("engine_*")))
        dst = work_dir / f"engine_{n}_{safe_name(Path(path).name)}"
        shutil.copy(path, dst)
        calc = Package(dst).workbook()["calc"]
        iterate = dict(on=calc.get("iterate") in ("1", "true"), count=int(calc.get("iterateCount", 100)),
                       delta=float(calc.get("iterateDelta", 0.001)))
        res = self._call("open", str(dst.resolve()), iterate, timeout=240)
        if res.get("warning"):
            self.info.setdefault("warnings", []).append(res["warning"])
        self.info["iteration"] = iterate
        return Book(self, dst.resolve(), path)

    def set(self, book, values):
        self._call("set", book.path, values)

    def calc(self, book, full=False):
        self._call("calc", full, timeout=self.timeout * (3 if full else 1))

    def get(self, book, cells):
        return self._call("get", book.path, list(cells))

    def run_macro(self, book, name, *args):
        return self._call("macro", book.path, name, args)["result"]

    def close(self, book):
        try:
            self._call("close", book.path)
        except EngineError:
            pass

    def ping(self):
        res = self._call("ping", timeout=30)
        return dict(self.info, **res)

    def stop(self):
        if self.proc is not None and self.proc.is_alive():
            try:
                self.conn.send(("quit", ()))
                self.proc.join(5)
            except Exception:
                pass
            if self.proc.is_alive():
                self.proc.terminate()
        self.proc = None
        if self._lockf is not None:
            self._lockf.close()
            self._lockf = None


class LibreOfficeEngine:
    """Fallback when Excel is missing. Starts a private headless LibreOffice that runs
    sheetre/lo_worker.py as a Python macro and serves set/calc/get over a Unix socket.
    Results are marked engine=libreoffice: no macros, external links never updated, and some
    functions differ from Excel."""
    name = "libreoffice"

    def __init__(self, timeout=180):
        self.timeout = timeout
        self.proc = self.sock = self.f = None
        self.info = {}

    @staticmethod
    def _soffice():
        for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice", shutil.which("soffice") or "",
                  r"C:\Program Files\LibreOffice\program\soffice.exe", "/usr/lib/libreoffice/program/soffice"):
            if p and Path(p).exists():
                return p
        return None

    @classmethod
    def detect(cls):
        p = cls._soffice()
        return (True, p) if p else (False, "LibreOffice (soffice) not found")

    def _call(self, cmd, timeout=None, **args):
        import json
        self.sock.settimeout(timeout or self.timeout)
        try:
            self.f.write((json.dumps({"cmd": cmd, "args": args}, default=str) + "\n").encode())
            self.f.flush()
            line = self.f.readline()
        except OSError as e:
            self.stop()
            raise EngineBlocked(f"LibreOffice did not answer '{cmd}': {e}")
        if not line:
            self.stop()
            raise EngineError(f"LibreOffice worker closed the connection during '{cmd}'")
        res = json.loads(line)
        if not res["ok"]:
            raise EngineError(f"LibreOffice '{cmd}' failed: {res['error']}")
        return res["result"]

    def start(self):
        import socket
        import tempfile
        import time
        soffice = self._soffice()
        if not soffice:
            raise EngineUnavailable("LibreOffice (soffice) not found")
        profile = Path(tempfile.mkdtemp(prefix="sheetre-lo-"))
        scripts = profile / "user" / "Scripts" / "python"
        scripts.mkdir(parents=True)
        shutil.copy(Path(__file__).with_name("lo_worker.py"), scripts / "sheetre_worker.py")
        sock_path = str(profile / "w.sock")
        env = dict(os.environ, SHEETRE_LO_SOCKET=sock_path)
        url = "vnd.sun.star.script:sheetre_worker.py$serve?language=Python&location=user"
        self.proc = subprocess.Popen([soffice, "--headless", "--invisible", "--norestore", "--nologo", "--nodefault",
                                      f"-env:UserInstallation={profile.as_uri()}", url],
                                     env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        for _ in range(240):
            if os.path.exists(sock_path):
                break
            if self.proc.poll() is not None:
                raise EngineUnavailable(f"LibreOffice exited before its worker started (code {self.proc.returncode})")
            time.sleep(0.25)
        else:
            self.stop()
            raise EngineUnavailable("LibreOffice started but its worker never opened a socket (macro blocked?)")
        self.sock = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        self.sock.connect(sock_path)
        self.f = self.sock.makefile("rwb")
        self._call("hello", timeout=30)
        ver = subprocess.run([soffice, "--version"], capture_output=True, text=True).stdout.strip()
        self.info = dict(engine="libreoffice", version=ver, locale=system_locale(),
                         caveats="no macros; external links not updated; some functions differ from Excel")
        return self

    def open(self, path, work_dir):
        work_dir = Path(work_dir)
        work_dir.mkdir(parents=True, exist_ok=True)
        n = len(list(work_dir.glob("engine_*")))
        dst = work_dir / f"engine_{n}_{safe_name(Path(path).name)}"
        shutil.copy(path, dst)
        calc = Package(dst).workbook()["calc"]
        iterate = dict(on=calc.get("iterate") in ("1", "true"), count=int(calc.get("iterateCount", 100)),
                       delta=float(calc.get("iterateDelta", 0.001)))
        self._call("open", timeout=240, path=str(dst.resolve()), iterate=iterate)
        self.info["iteration"] = iterate
        return Book(self, dst.resolve(), path)

    def set(self, book, values):
        import datetime as dt
        enc = {k: ({"date": (v if isinstance(v, dt.datetime) else dt.datetime(v.year, v.month, v.day)).isoformat()}
                   if isinstance(v, (dt.date, dt.datetime)) else v) for k, v in values.items()}
        self._call("set", book=book.path, values=enc)

    def calc(self, book, full=False):
        self._call("calc", book=book.path)

    def get(self, book, cells):
        return self._call("get", book=book.path, cells=list(cells))

    def run_macro(self, book, name, *args):
        raise EngineError("macros are not run on the LibreOffice engine")

    def close(self, book):
        try:
            self._call("close", book=book.path)
        except EngineError:
            pass

    def stop(self):
        try:
            if self.f is not None:
                self.f.write(b'{"cmd": "quit"}\n')
                self.f.flush()
        except Exception:
            pass
        for x in (self.f, self.sock):
            try:
                x and x.close()
            except Exception:
                pass
        self.f = self.sock = None
        if self.proc is not None:
            try:
                self.proc.wait(15)
            except Exception:
                self.proc.kill()
        self.proc = None


class BridgeEngine:
    """Client for sheetre.bridge: same interface as ExcelEngine, Excel driven by a process outside
    the sandbox. Used automatically when the bridge socket exists."""
    name = "excel-bridge"

    def __init__(self, timeout=600, path=None):
        self.path = path or os.environ.get("SHEETRE_BRIDGE", "/tmp/sheetre-bridge.sock")
        self.timeout = timeout
        self.sock = self.f = None
        self.info = {}

    @staticmethod
    def detect(path=None):
        p = path or os.environ.get("SHEETRE_BRIDGE", "/tmp/sheetre-bridge.sock")
        return os.path.exists(p), None if os.path.exists(p) else f"no bridge socket at {p}"

    def _call(self, cmd, **args):
        import json
        self.f.write((json.dumps({"cmd": cmd, "args": args}, default=str) + "\n").encode())
        self.f.flush()
        line = self.f.readline()
        if not line:
            raise EngineError("bridge closed the connection")
        res = json.loads(line)
        if not res["ok"]:
            if res.get("busy"):
                raise EngineBusy(res["error"])
            raise (EngineBlocked if res.get("blocked") else EngineError)(f"bridge '{cmd}': {res['error']}")
        return res["result"]

    def start(self):
        import socket
        ok, why = self.detect(self.path)
        if not ok:
            raise EngineUnavailable(why)
        try:
            self.sock = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
            self.sock.settimeout(self.timeout)
            self.sock.connect(self.path)
        except OSError as e:
            raise EngineUnavailable(f"bridge socket {self.path} not reachable ({e}); start it with "
                                    "'.venv/bin/python -m sheetre.bridge' in a normal terminal")
        self.f = self.sock.makefile("rwb")
        self.info = self._call("hello")
        self.info["engine"] = "excel (via bridge)"
        return self

    def open(self, path, work_dir):
        Path(work_dir).mkdir(parents=True, exist_ok=True)
        res = self._call("open", path=str(Path(path).resolve()), work_dir=str(Path(work_dir).resolve()))
        self.info["iteration"] = res.get("iteration")
        return Book(self, res["path"], path)

    def set(self, book, values):
        import datetime as dt
        enc = {k: ({"date": (v if isinstance(v, dt.datetime) else dt.datetime(v.year, v.month, v.day)).isoformat()}
                   if isinstance(v, (dt.date, dt.datetime)) else v) for k, v in values.items()}
        self._call("set", book=book.path, values=enc)

    def calc(self, book, full=False):
        self._call("calc", book=book.path, full=full)

    def get(self, book, cells):
        import datetime as dt
        out = self._call("get", book=book.path, cells=list(cells))
        return {k: dt.datetime.fromisoformat(v["date"]) if isinstance(v, dict) and "date" in v else v for k, v in out.items()}

    def run_macro(self, book, name, *args):
        raise EngineError("macros cannot be run through the bridge")

    def close(self, book):
        try:
            self._call("close", book=book.path)
        except EngineError:
            pass

    def stop(self):
        try:
            if self.sock:
                self.sock.close()
        finally:
            self.sock = self.f = None


def get_engine(prefer=None, timeout=120):
    """Bridge (if running), then Excel, then LibreOffice. Raises EngineUnavailable with every reason."""
    order = [BridgeEngine, ExcelEngine, LibreOfficeEngine]
    if prefer == "libreoffice":
        order = [LibreOfficeEngine, BridgeEngine, ExcelEngine]
    elif prefer == "excel":
        order = [ExcelEngine, BridgeEngine, LibreOfficeEngine]
    reasons = []
    for cls in order:
        if cls is BridgeEngine and not cls.detect()[0]:
            continue
        try:
            eng = cls(timeout=timeout) if cls in (ExcelEngine, LibreOfficeEngine) else cls()
            return eng.start()
        except EngineUnavailable as e:
            reasons.append(f"{cls.name}: {e}")
    raise EngineUnavailable("; ".join(reasons))
