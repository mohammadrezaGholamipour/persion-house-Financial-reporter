# -*- coding: utf-8 -*-
"""سرور محلی سبک (فقط کتابخانه استاندارد پایتون). صفحه‌بندی، جستجو و مرتب‌سازی در SQLite انجام می‌شود."""
import csv, io, json, mimetypes, sqlite3, webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

HERE = Path(__file__).parent
REASONS = ["عدم ثبت فروش در سامانه مودیان", "عدم ثبت فروش در نرم‌افزار", "مغایرت مبلغ فروش",
           "برگشت فقط در نرم‌افزار", "برگشت فقط در سامانه مودیان", "مغایرت مبلغ برگشت"]


def make_handler(db_path):
    def q(sql, args=()):
        con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
        con.row_factory = sqlite3.Row
        try:
            return [dict(r) for r in con.execute(sql, args)]
        finally:
            con.close()

    def where(p, table, cols):
        w, a = [], []
        if table == "mis":
            r = p.get("reason", [""])[0]
            if r == "UNMATCHED":  # رکورد متناظر پیدا نشده؛ مستقل از آستانه مبلغ
                w.append("(reasons LIKE '%عدم ثبت%' OR reasons LIKE '%فقط در%')")
            else:
                w.append("max_abs > ?"); a.append(int(p.get("th", ["10000"])[0]) * 10)
                if r:
                    w.append("reasons LIKE ?"); a.append("%" + r + "%")
        elif p.get("reason", [""])[0]:
            w.append("issue = ?" if table == "iss" else "status = ?"); a.append(p["reason"][0])
        if p.get("q", [""])[0]:
            w.append("(" + " || ' ' || ".join(f"COALESCE(CAST({c} AS TEXT),'')" for c in cols) + ") LIKE ?")
            a.append("%" + p["q"][0] + "%")
        return (" WHERE " + " AND ".join(w)) if w else "", a

    def table_query(p, limit=True):
        t = p.get("t", ["mis"])[0]
        if t not in ("mis", "can", "iss", "map"):
            raise ValueError("table")
        cols = [r["name"] for r in q(f"PRAGMA table_info({t})")]
        wh, a = where(p, t, cols)
        sort = p.get("sort", [""])[0]
        order = f" ORDER BY {sort} {'DESC' if p.get('dir', ['1'])[0] == '-1' else 'ASC'}" if sort in cols else ""
        if t == "mis" and not order:
            order = " ORDER BY max_abs DESC"
        sql = f"SELECT * FROM {t}{wh}{order}"
        if limit:
            size = min(int(p.get("size", ["100"])[0]), 500)
            sql += f" LIMIT {size} OFFSET {int(p.get('page', ['0'])[0]) * size}"
        return t, cols, wh, a, sql

    class H(BaseHTTPRequestHandler):
        def log_message(self, *a): pass

        def send(self, body, ctype="application/json; charset=utf-8", extra=None):
            b = body if isinstance(body, bytes) else body.encode("utf-8")
            self.send_response(200); self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(b)))
            for k, v in (extra or {}).items(): self.send_header(k, v)
            self.end_headers(); self.wfile.write(b)

        def do_GET(self):
            u = urlparse(self.path); p = parse_qs(u.query)
            try:
                if u.path == "/":
                    return self.send((HERE / "index.html").read_bytes(), "text/html; charset=utf-8")
                if u.path.startswith(("/css/", "/assets/")):  # فایل‌های ثابت: فونت، چارت، استایل
                    f = (HERE / unquote(u.path.lstrip("/"))).resolve()
                    if any((HERE / d).resolve() in f.parents for d in ("css", "assets")) and f.is_file():
                        mt = {".woff2": "font/woff2", ".woff": "font/woff", ".eot": "application/vnd.ms-fontobject"}
                        ct = mt.get(f.suffix.lower()) or mimetypes.guess_type(str(f))[0] or "application/octet-stream"
                        return self.send(f.read_bytes(), ct, {"Cache-Control": "max-age=3600"})
                    self.send_response(404); self.end_headers()
                    return
                if u.path == "/api/insights":
                    th = int(p.get("th", ["10000"])[0]) * 10
                    top = q("SELECT COALESCE(NULLIF(name,''),'(بدون اطلاعات طرف حساب - فقط در مودیان)') AS name, COUNT(*) n, "
                            "SUM(d_net) d FROM mis WHERE max_abs > ? GROUP BY 1 ORDER BY ABS(SUM(d_net)) DESC LIMIT 10", (th,))
                    b = q("SELECT SUM(max_abs<=1000000) a, SUM(max_abs>1000000 AND max_abs<=10000000) b, "
                          "SUM(max_abs>10000000 AND max_abs<=100000000) c, SUM(max_abs>100000000 AND max_abs<=1000000000) d, "
                          "SUM(max_abs>1000000000) e FROM mis WHERE max_abs > ?", (th,))[0]
                    return self.send(json.dumps({"top": top, "buckets": [b[k] or 0 for k in "abcde"]}, ensure_ascii=False))
                if u.path == "/api/summary":
                    kv = {r["k"]: json.loads(r["v"]) for r in q("SELECT * FROM kv")}
                    th = int(p.get("th", ["10000"])[0]) * 10
                    kv["visible"] = q("SELECT COUNT(*) n FROM mis WHERE max_abs > ?", (th,))[0]["n"]
                    kv["total_raw"] = q("SELECT COUNT(*) n FROM mis")[0]["n"]
                    kv["reasons"] = {r: q("SELECT COUNT(*) n FROM mis WHERE max_abs > ? AND reasons LIKE ?",
                                          (th, "%" + r + "%"))[0]["n"] for r in REASONS}
                    kv["issue_types"] = [r["issue"] for r in q("SELECT DISTINCT issue FROM iss")]
                    kv["cancel_types"] = [r["status"] for r in q("SELECT DISTINCT status FROM can")]
                    return self.send(json.dumps(kv, ensure_ascii=False))
                if u.path == "/api/table":
                    t, cols, wh, a, sql = table_query(p)
                    total = q(f"SELECT COUNT(*) n FROM {t}{wh}", a)[0]["n"]
                    return self.send(json.dumps({"total": total, "rows": q(sql, a)}, ensure_ascii=False))
                if u.path == "/api/csv":
                    t, cols, wh, a, sql = table_query(p, limit=False)
                    buf = io.StringIO(); w = csv.writer(buf); w.writerow(cols)
                    for r in q(sql, a): w.writerow([r[c] for c in cols])
                    return self.send(("\ufeff" + buf.getvalue()).encode("utf-8"), "text/csv; charset=utf-8",
                                     {"Content-Disposition": f"attachment; filename={t}.csv"})
            except Exception as e:
                self.send_response(400); self.end_headers(); self.wfile.write(str(e).encode())
                return
            self.send_response(404); self.end_headers()
    return H


def serve(db_path, port=8765):
    srv = ThreadingHTTPServer(("127.0.0.1", port), make_handler(str(db_path)))
    print(f"گزارش روی http://127.0.0.1:{port} در حال اجراست (برای توقف: Ctrl+C)")
    webbrowser.open(f"http://127.0.0.1:{port}")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass