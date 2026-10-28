"""LibreOffice worker: runs INSIDE soffice as a Python macro (location=user of a private profile),
because recent LibreOffice builds forbid launching their bundled Python from outside the app
(macOS launch constraint). It serves JSON lines on the Unix socket named in SHEETRE_LO_SOCKET.
Imports nothing from sheetre.

Commands: start, open, set, calc, get, close, quit.
Values: numbers, strings, None (blank), {"date": iso} (written as a serial number).
Errors come back as Excel-style strings (#N/A, #VALUE!, ...) where LibreOffice has an equivalent.
"""
import datetime as dt
import json
import os

import uno
from com.sun.star.beans import PropertyValue

LO_ERRORS = {503: "#NUM!", 502: "#VALUE!", 504: "#VALUE!", 519: "#VALUE!", 522: "#REF!", 523: "#NUM!", 524: "#REF!",
             525: "#NAME?", 532: "#DIV/0!", 539: "#VALUE!", 540: "#VALUE!", 32767: "#N/A"}
EPOCH = dt.datetime(1899, 12, 30)

state = dict(desktop=None, books={})


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def cell(doc, key):
    sheet, addr = key.rsplit("!", 1)
    return doc.Sheets.getByName(sheet).getCellRangeByName(addr)


def to_lo(c, v):
    if v is None:
        c.setFormula("")
    elif isinstance(v, bool):
        c.setValue(1 if v else 0)
    elif isinstance(v, (int, float)):
        c.setValue(v)
    elif isinstance(v, dict) and "date" in v:
        d = dt.datetime.fromisoformat(v["date"])
        c.setValue((d - EPOCH).total_seconds() / 86400)
    else:
        c.setString(str(v))


def from_lo(c):
    t = c.getType().value                     # EMPTY, VALUE, TEXT, FORMULA
    if t == "EMPTY":
        return None
    if t == "VALUE":
        return c.getValue()
    if t == "TEXT":
        return c.getString()
    err = c.getError()
    if err:
        return LO_ERRORS.get(err, f"Err:{err}")
    rt = c.FormulaResultType2               # com.sun.star.sheet.FormulaResult: VALUE=1, STRING=2, ERROR=4
    if rt == 2:
        s = c.getString()
        return s if s != "" else None
    return c.getValue()


def snapshot(doc):
    out = []
    for i in range(doc.Sheets.Count):
        sh = doc.Sheets.getByIndex(i)
        cur = sh.createCursor()
        cur.gotoEndOfUsedArea(False)
        a = cur.getRangeAddress()
        out.append(sh.getCellRangeByPosition(0, 0, a.EndColumn, a.EndRow).getDataArray())
    return out


def handle(cmd, a):
    if cmd == "hello":
        return dict(pid=os.getpid())
    if cmd == "open":
        url = uno.systemPathToFileUrl(a["path"])
        doc = state["desktop"].loadComponentFromURL(url, "_blank", 0, (
            prop("Hidden", True), prop("UpdateDocMode", 0), prop("MacroExecutionMode", 0), prop("ReadOnly", False)))
        if doc is None:
            raise RuntimeError("LibreOffice could not open the workbook")
        it = a.get("iterate")
        if it:
            doc.IsIterationEnabled = bool(it["on"])
            doc.IterationCount = int(it["count"])
            doc.IterationEpsilon = float(it["delta"])
        state["books"][a["path"]] = doc
        return {}
    if cmd == "set":
        doc = state["books"][a["book"]]
        for k, v in a["values"].items():
            to_lo(cell(doc, k), v)
        return {}
    if cmd == "calc":
        doc = state["books"][a["book"]]
        doc.calculateAll()
        passes = 1
        if doc.IsIterationEnabled:
            # LibreOffice runs only a few iterations per recalculation, while Excel converges a circular
            # chain in one go: recalculate until every sheet's values stop changing
            prev = snapshot(doc)
            while passes < 50:
                doc.calculateAll()
                passes += 1
                cur = snapshot(doc)
                if cur == prev:
                    break
                prev = cur
        return dict(passes=passes)
    if cmd == "get":
        doc = state["books"][a["book"]]
        return {k: from_lo(cell(doc, k)) for k in a["cells"]}
    if cmd == "close":
        doc = state["books"].pop(a["book"], None)
        if doc is not None:
            doc.close(True)
        return {}
    raise RuntimeError(f"unknown command {cmd}")


def serve(*args):
    """Macro entry point, run as vnd.sun.star.script:sheetre_worker.py$serve?language=Python&location=user."""
    import socket
    state["desktop"] = XSCRIPTCONTEXT.getDesktop()  # noqa: F821 (provided by LibreOffice's script provider)
    path = os.environ["SHEETRE_LO_SOCKET"]
    if os.path.exists(path):
        os.unlink(path)
    srv = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
    srv.bind(path)
    srv.listen(1)
    running = True
    while running:
        conn, _ = srv.accept()
        f = conn.makefile("rwb")
        for line in f:
            req = json.loads(line)
            if req["cmd"] == "quit":
                running = False
                break
            try:
                res = {"ok": True, "result": handle(req["cmd"], req.get("args", {}))}
            except Exception as e:
                res = {"ok": False, "error": f"{type(e).__name__}: {e}"}
            f.write((json.dumps(res, default=str) + "\n").encode())
            f.flush()
        conn.close()
    for doc in list(state["books"].values()):
        try:
            doc.close(True)
        except Exception:
            pass
    srv.close()
    os.unlink(path)
    state["desktop"].terminate()


g_exportedScripts = (serve,)
