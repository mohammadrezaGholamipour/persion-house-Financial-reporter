# -*- coding: utf-8 -*-
"""
سیستم مغایرت‌گیری ماژول (نرم‌افزار شرکت) و سامانه مودیان - نسخه ۵
اجرا:  py main.py ماژول.xlsx مودیان.xlsx      (ترتیب فایل‌ها مهم نیست)
خروجی: پوشه output  ->  گزارش_مغایرت.html  و  result.json
"""
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd

# ---------------- تنظیمات ----------------
THRESHOLD_TOMAN = 10_000            # اختلاف کمتر یا مساوی این مقدار (تومان) پنهان می‌شود
RIAL = 10                           # هر تومان = ۱۰ ریال؛ مبالغ فایل‌ها ریال فرض می‌شوند
DIG = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

LABELS = {"key": "شناسه صورتحساب مالیاتی", "ref": "شناسه صورتحساب مرجع", "type": "نوع فاکتور",
          "subject": "موضوع صورتحساب", "pattern": "الگوی صورتحساب", "gross": "جمع ناخالص", "discount": "جمع تخفیفات",
          "tax": "مالیات", "total": "مبلغ کل", "status": "وضعیت صورتحساب", "date": "تاریخ",
          "customer": "نام طرف حساب", "nid": "کد ملی / شناسه ملی", "eco": "کد اقتصادی",
          "ptype": "نوع شخص", "sub": "کد اشتراک", "modno": "شماره ماژول مودیان"}
ALIASES = {
    "key": ["شناسه صورتحساب مالیاتی", "شماره مالیاتی صورتحساب", "شناسه مالیاتی"],
    "ref": ["شناسه صورتحساب مرجع", "شماره مالیاتی صورتحساب مرجع"],
    "type": ["نوع فاکتور"],
    "subject": ["موضوع صورتحساب", "موضوع فاکتور"],
    "pattern": ["الگو صورتحساب", "الگوی صورتحساب"],
    "gross": ["جمع ناخالص فاکتور", "جم ناخالص فاکتور",
              "مجموع بهای کالا و خدمات صورت حساب بدون مالیات ها و عوارض (ریال)"],
    "discount": ["جمع تخفیفات", "مبلغ تخفیف"],
    "tax": ["مبلغ مالیات", "مالیات بر ارزش افزوده (ریال)"],
    "total": ["مجموع صورتحساب (ریال)", "مبلغ کل صورتحساب", "مبلغ کل شامل مالیات"],
    "status": ["وضعیت صورتحساب"],
    "date": ["تاریخ فاکتور", "تاریخ صدور صورتحساب"],
    "customer": ["نام مشتری", "نام خریدار"],
    "nid": ["کد ملی", "شناسه ملی"],
    "eco": ["کد اقتصادی"],
    "ptype": ["نوع شخص"],
    "sub": ["کد اشتراک"],
    "modno": ["شماره ماژول مودیان"],
}
META = ["customer", "nid", "eco", "ptype", "sub", "date"]


# ---------------- توابع کمکی ----------------
def fa(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).translate(DIG).replace("ي", "ی").replace("ك", "ک").replace("\u200c", " ").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip()


def hdr(v):
    return fa(v).replace(" ", "")


def key_of(v):
    return re.sub(r"[^0-9A-Z]", "", fa(v).upper())


def num(v):
    try:
        return int(round(float(fa(v).replace(",", "").replace("٬", ""))))
    except ValueError:
        return 0


def month_of(d):
    m = re.match(r"(\d{4})/(\d{1,2})", d)
    return f"{m[1]}/{int(m[2]):02d}" if m else "نامشخص"


def classify(*texts):
    s = " ".join(texts)
    for kind, words in (("return", ("برگشت", "مرجوع", "عودت")), ("cancel", ("ابطال", "باطل", "لغو")),
                        ("amend", ("اصلاح",)), ("sale", ("فروش", "اصلی"))):
        if any(w in s for w in words):
            return kind
    return "unknown"


AR = re.compile(r"[\u0600-\u06FF\uFB50-\uFDFF\uFE70-\uFEFF]")
LTR = r"[A-Za-z0-9.,:/\\%_+-]+"
BRACKETS = str.maketrans("()[]<>«»", ")(][><»«")


def visual(text):
    """ترتیب دیداری یک خط برای ترمینال چپ‌به‌راست: ترتیب کلمات و حروف هر کلمه فارسی برعکس می‌شود؛
    اعداد و متن لاتین (مسیر، نام فایل) دست‌نخورده می‌مانند. جایگزین ساده python-bidi."""
    lead = len(text) - len(text.lstrip(" "))
    toks = re.split(r"(\s+)", text[lead:])
    isl = lambda t: bool(t) and not t.isspace() and not AR.search(t) and re.search("[A-Za-z0-9]", t)
    units, i = [], 0
    while i < len(toks):  # کلمه‌های لاتین/عددی پشت‌سرهم یک واحد هستند و ترتیبشان حفظ می‌شود
        if isl(toks[i]):
            j = i
            while j + 2 < len(toks) and toks[j + 1].isspace() and isl(toks[j + 2]):
                j += 2
            units.append("".join(toks[i:j + 1]))
            i = j + 1
        else:
            units.append(toks[i])
            i += 1
    out = []
    for u in reversed(units):
        if AR.search(u):
            runs = re.split(f"({LTR})", u[::-1])
            u = "".join(r[::-1] if re.fullmatch(LTR, r or "") else r for r in runs).translate(BRACKETS)
        out.append(u)
    return " " * lead + "".join(out)


def setup_rtl():
    """ترمینال ویندوز حروف فارسی را جدا و برعکس نشان می‌دهد؛ متن قبل از چاپ به شکل دیداری تبدیل می‌شود."""
    try:
        import arabic_reshaper
    except ImportError:
        print("برای نمایش درست فارسی در ترمینال اجرا کنید:  py -m pip install arabic-reshaper")
        return

    class RTLOut:
        def __init__(self, base):
            self.base = base

        def write(self, text):
            parts = re.split(r"(\r|\n)", text)
            return self.base.write("".join(
                visual(arabic_reshaper.reshape(p.replace("\u200c", " "))) if AR.search(p) else p for p in parts))

        def flush(self):
            self.base.flush()

        def __getattr__(self, name):
            return getattr(self.base, name)

    sys.stdout = RTLOut(sys.stdout)


def ask(prompt=""):
    """input() در ترمینال، متن سؤال را مستقیم و بدون تبدیل فارسی چاپ می‌کند؛ پس سؤال را خودمان می‌نویسیم."""
    sys.stdout.write(prompt)
    sys.stdout.flush()
    return input()


class Progress:
    """نوار پیشرفت ترمینال با درصد و زمان باقی‌مانده."""
    def __init__(self):
        self.reset()

    def reset(self):
        self.t0, self.p, self.label = time.time(), 0.0, ""

    @staticmethod
    def fmt(sec):
        sec = int(sec)
        return f"{sec // 60}:{sec % 60:02d}"

    def to(self, p, label=None):
        self.p = max(self.p, min(p, 1.0))
        self.label = label or self.label
        el = time.time() - self.t0
        eta = self.fmt(el * (1 - self.p) / self.p) if self.p > 0.03 else "..."
        n = int(28 * self.p)
        sys.stdout.write(f"\r[{'=' * n}{'.' * (28 - n)}] {self.p * 100:5.1f}% | {self.label} | "
                         f"سپری‌شده {self.fmt(el)} | باقی‌مانده {eta}      ")
        sys.stdout.flush()

    def log(self, msg):
        sys.stdout.write("\r" + " " * 118 + "\r" + msg + "\n")
        self.to(self.p)


P = Progress()


def peek_headers(path):
    p = str(path)
    if p.lower().endswith(".csv"):
        return list(pd.read_csv(p, nrows=0).columns)
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    h = next(wb.active.iter_rows(values_only=True), ())
    wb.close()
    return [x for x in h if x is not None]


def detect_type(path):
    """نوع فایل از روی عنوان ستون‌ها تشخیص داده می‌شود (نه نام فایل و نه ترتیب): ستون «موضوع صورتحساب» = مودیان."""
    heads = {hdr(c) for c in peek_headers(path)}
    return "tax" if any(hdr(a) in heads for a in ALIASES["subject"]) else "module"


TYPE_FA = {"module": "ماژول (نرم‌افزار شرکت)", "tax": "مودیان (سامانه مالیاتی)"}


def load(path, p0=0.0, p1=0.0, label=""):
    p = str(path).strip().strip('"')
    if p.lower().endswith(".csv"):
        df = pd.read_csv(p, dtype=object)
    else:
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        ws = wb.active
        total = ws.max_row or 0
        it = ws.iter_rows(values_only=True)
        head, seen = next(it), {}
        names = []
        for j, h in enumerate(head):
            h = str(h) if h is not None else f"ستون{j}"
            seen[h] = seen.get(h, 0) + 1
            names.append(h if seen[h] == 1 else f"{h}.{seen[h] - 1}")
        n, rows = len(names), []
        for i, r in enumerate(it, 1):
            rows.append((tuple(r) + (None,) * n)[:n])
            if i % 5000 == 0 and total:
                P.to(p0 + (p1 - p0) * min(i / total, 1), f"خواندن {label}")
        wb.close()
        df = pd.DataFrame(rows, columns=names, dtype=object)
    df = df.dropna(how="all")
    cols = {hdr(c): c for c in df.columns}
    mp = {f: next((cols[hdr(a)] for a in al if hdr(a) in cols), None) for f, al in ALIASES.items()}
    P.to(p1, f"خواندن {label}")
    need = ["key", "ref", "subject"] if mp["subject"] else ["key", "ref", "customer", "nid", "ptype", "sub"]
    miss = [LABELS[f] for f in need if not mp[f]] + ([] if mp["total"] or mp["gross"] else ["مبلغ"])
    P.log(f"  ✔ {label}: {len(df):,} ردیف، {len(df.columns)} ستون | ستون‌های شناسایی‌شده: "
          f"{sum(1 for v in mp.values() if v)}/{len(mp)}" + (f" | یافت نشد: {'، '.join(miss)}" if miss else ""))
    return df, mp


def prepare(df, mp, src, discount_before=False, prefix=""):
    g = lambda f: df[mp[f]] if mp[f] else pd.Series("", index=df.index)
    o = pd.DataFrame({"row": [prefix + str(i + 2) for i in df.index], "key": g("key").map(key_of), "ref": g("ref").map(key_of)})
    o["gross"], o["tax"] = g("gross").map(num), g("tax").map(num)
    if mp["total"]:
        o["total"] = g("total").map(num)
    else:  # مبلغ کل = ناخالص + مالیات؛ اگر ناخالص قبل از تخفیف باشد، تخفیف هم کم می‌شود
        o["total"] = o.gross + o.tax - (g("discount").map(num) if discount_before else 0)
    for f in ["type", "subject", "pattern", "status", "date", "customer", "nid", "eco", "ptype", "sub", "modno"]:
        o[f] = g(f).map(fa)
    o["src"] = src
    return o


# ---------------- پردازش هر طرف ----------------
def clean_side(D, issues, tax_side):
    """حذف تکراری/بدون شناسه و برگرداندن فروش‌ها و برگشت‌های قابل استفاده."""
    def add(rows, issue, detail=""):
        for _, r in rows.iterrows():
            issues.append(dict(src=r.src, row=str(r.row), key=r.key, ref=r.ref, type=r.type or r.subject,
                               amount=int(r.total), issue=issue, detail=detail))

    S = D[D.k == "sale"]
    add(S[S.key == ""], "فروش بدون شناسه مالیاتی", "در محاسبه نیامد")
    S = S[S.key != ""]
    d = S.duplicated(["key", "total", "tax", "gross", "date", "customer"])
    add(S[d], "ردیف فروش کاملاً تکراری", "از محاسبه حذف شد")
    S = S[~d]
    c = S.key.duplicated(keep=False)
    add(S[c], "شناسه فروش تکراری/متعارض", "در محاسبه نیامد")
    S = S[~c]

    R = D[D.k == "return"]
    if tax_side:  # برگشت‌های مودیان شناسه یکتا دارند
        d = (R.key != "") & R.duplicated(["key", "total", "ref"])
        add(R[d], "ردیف برگشت کاملاً تکراری", "از محاسبه حذف شد")
        R = R[~d]
    else:         # برگشت‌های ماژول شناسه ندارند؛ ردیف‌های یکسان جمع می‌شوند ولی گزارش می‌شوند
        cols = ["ref", "total", "tax", "date", "modno", "customer"]
        g = R.duplicated(cols, keep=False)
        for _, grp in R[g].groupby(cols):
            r = grp.iloc[0].copy()
            r["row"] = grp.row.iloc[0]
            add(pd.DataFrame([r]), "ردیف‌های برگشت یکسان (همه جمع شدند)",
                f"{len(grp)} ردیف: " + "، ".join(map(str, grp.row)) + " - اگر سطرهای یک سند نیستند بررسی شود")
    add(R[R.ref == ""], "برگشت بدون شناسه مرجع", "کسر نشد")
    R = R[R.ref != ""]

    add(D[D.k == "unknown"], "نوع نامشخص", "در مبلغ نیامد")
    add(D[D.k == "amend"], "اصلاحی", "در مبلغ نیامد؛ بررسی دستی")
    add(D[D.status.str.contains("رد شده")], "صورتحساب رد شده", "")
    add(D[D.status.str.contains("انتظار")], "در انتظار واکنش", "")
    add(D[(D.k == "cancel") & (D.ref == "")], "ابطال بدون شناسه مرجع", "")
    return S, R


def ledger(S, R):
    r = R.groupby("ref").agg(ret=("total", "sum"), rrows=("row", lambda x: "، ".join(map(str, x))))
    meta = pd.concat([R.groupby("ref")[META].first(), S.set_index("key")[META]])
    meta = meta[~meta.index.duplicated(keep="last")]
    s = S.set_index("key")[["total", "row"]].rename(columns={"total": "sale", "row": "srow"})
    return s.join(r, how="outer").join(meta, how="outer"), s


def month_net(S, R):
    a = S.assign(m=S.date.map(month_of)).groupby("m").total.sum()
    b = R.assign(m=R.date.map(month_of)).groupby("m").total.sum()
    return a.sub(b, fill_value=0)


# ---------------- اجرای اصلی ----------------
def ask_discount(parts):
    """اگر فایل ماژول ستون تخفیف با مقدار غیرصفر داشته باشد، از ادمین معنی ستون ناخالص پرسیده می‌شود."""
    if not any((not mp["total"]) and mp["discount"] and df[mp["discount"]].map(num).any() for df, mp in parts):
        return False
    print("\nدر فایل ماژول ردیف‌هایی با «جمع تخفیفات» غیرصفر پیدا شد. ستون «جمع ناخالص» چیست؟")
    print("  1) بعد از کسر تخفیف (تخفیف قبلاً اعمال شده؛ مبلغ کل = ناخالص + مالیات)")
    print("  2) قبل از کسر تخفیف (مبلغ کل = ناخالص - تخفیف + مالیات)")
    while True:
        a = ask("شماره گزینه (1 یا 2): ").strip().translate(DIG)
        if a in ("1", "2"):
            return a == "2"

def run_files(items, threshold_toman=THRESHOLD_TOMAN, discount_before=None):
    """items = [(مسیر فایل, 'module' یا 'tax'), ...]؛ از هر نوع می‌توان چند فایل داد (ادغام می‌شوند)."""
    P.reset()
    sizes = [max(os.path.getsize(p), 1) for p, _ in items]
    cum, parts = 0, {"tax": [], "module": []}
    for (path, typ), sz in zip(items, sizes):
        p0, p1 = 0.02 + 0.38 * cum / sum(sizes), 0.02 + 0.38 * (cum + sz) / sum(sizes)
        cum += sz
        df, mp = load(path, p0, p1, f"{TYPE_FA[typ]} - {Path(path).name}")
        if typ == "tax" and not mp["subject"]:
            raise SystemExit(f"خطا: فایل «{Path(path).name}» ستون «موضوع صورتحساب» ندارد؛ نوع آن مودیان نیست.")
        if not mp["key"] or not (mp["total"] or mp["gross"]):
            raise SystemExit(f"خطا: ستون شناسه مالیاتی یا مبلغ در «{Path(path).name}» پیدا نشد.")
        parts[typ].append((df, mp))
    if not parts["tax"] or not parts["module"]:
        raise SystemExit("خطا: حداقل یک فایل ماژول و یک فایل مودیان لازم است.")
    tmp, cmp_ = parts["tax"][0][1], parts["module"][0][1]
    if discount_before is None:
        sys.stdout.write("\r" + " " * 118 + "\r")
        discount_before = ask_discount(parts["module"])
    P.to(0.42, "آماده‌سازی داده‌ها")

    def build(kind, src, **kw):
        many = len(parts[kind]) > 1
        return pd.concat([prepare(df, mp, src, prefix=(f"{n}:" if many else ""), **kw)
                          for n, (df, mp) in enumerate(parts[kind], 1)], ignore_index=True)

    T, C = build("tax", "سامانه مودیان"), build("module", "نرم‌افزار شرکت", discount_before=discount_before)
    T["k"] = [classify(a, b) for a, b in zip(T.subject, T.pattern)]
    C["k"] = C.type.map(classify)
    for nm, D in (("مودیان", T), ("ماژول", C)):
        c = D.k.value_counts()
        P.log(f"  ● استخراج {nm}: فروش {c.get('sale', 0):,} | برگشت {c.get('return', 0):,} | "
              f"ابطال {c.get('cancel', 0):,} | اصلاحی {c.get('amend', 0):,} | نامشخص {c.get('unknown', 0):,}")
    P.to(0.47, "تشخیص ابطال و برگشت")

    # ابطال: فاکتورهایی که در مودیان ردیف ابطال یا وضعیت ابطال دارند؛ از جمع و مغایرت خارج می‌شوند
    canc = T[T.k == "cancel"]
    cancelled = set(canc.ref[canc.ref != ""]) | set(T.key[(T.k == "sale") & T.status.str.contains("ابطال")])
    # ماژول، ابطال و اصلاحی را «برگشت از فروش» ثبت می‌کند؛ نوع واقعی از روی مودیان مشخص می‌شود
    tret = set(T.ref[T.k == "return"])
    amend = set(T.ref[(T.k == "amend") & (T.ref != "")])
    isret = C.k == "return"
    C.loc[isret & C.ref.isin(cancelled), "k"] = "cancel"
    C.loc[(C.k == "return") & C.ref.isin(amend - tret), "k"] = "amend"

    issues = []
    P.to(0.52, "تمیزسازی و حذف تکراری‌ها")
    Sc, Rc = clean_side(C, issues, False)
    St, Rt = clean_side(T, issues, True)

    P.to(0.60, "گزارش ابطال‌ها")
    # گزارش ابطال‌ها
    cm_rows = C[C.k == "cancel"].groupby("ref").row.agg(lambda x: "، ".join(map(str, x)))
    tc_rows = canc[canc.ref != ""].groupby("ref").row.agg(lambda x: "، ".join(map(str, x)))
    sale_c, sale_t = Sc.set_index("key").total, St.set_index("key").total
    cancels = []
    for k in sorted(cancelled):
        in_m = k in sale_c.index
        if k in cm_rows.index:
            st = "ابطال در هر دو سیستم"
        elif in_m:
            st = "اختلاف وضعیت: ابطال فقط در مودیان - در ماژول فعال است"
        else:
            st = "ابطال در مودیان - فاکتور اصلی در ماژول نیست"
        cancels.append(dict(key=k, status=st, m_sale=int(sale_c.get(k, 0)), t_sale=int(sale_t.get(k, 0)),
                            m_rows=cm_rows.get(k, ""), t_rows=tc_rows.get(k, ""),
                            diff=int(st.startswith("اختلاف"))))

    cs, cr = Sc[~Sc.key.isin(cancelled)], Rc[~Rc.ref.isin(cancelled)]
    ts, tr = St[~St.key.isin(cancelled)], Rt[~Rt.ref.isin(cancelled)]
    LC, _ = ledger(cs, cr)
    LT, _ = ledger(ts, tr)

    P.to(0.64, "محاسبه خالص هر شناسه")
    # مقایسه به‌ازای هر شناسه (فروش اصلی)
    idx = LC.index.union(LT.index)
    X = pd.DataFrame(index=idx)
    for p, L in (("m", LC), ("t", LT)):
        X[p + "_has"] = L.srow.reindex(idx).notna()
        X[p + "_sale"] = L.sale.reindex(idx).fillna(0).astype("int64")
        X[p + "_ret"] = L.ret.reindex(idx).fillna(0).astype("int64")
        X[p + "_net"] = X[p + "_sale"] - X[p + "_ret"]
        X[p + "_srow"] = L.srow.reindex(idx)
        X[p + "_rrows"] = L.rrows.reindex(idx).fillna("")
    X["d_sale"], X["d_ret"], X["d_net"] = X.m_sale - X.t_sale, X.m_ret - X.t_ret, X.m_net - X.t_net
    M = pd.DataFrame({
        "عدم ثبت فروش در سامانه مودیان": X.m_has & ~X.t_has,
        "عدم ثبت فروش در نرم‌افزار": X.t_has & ~X.m_has,
        "مغایرت مبلغ فروش": X.m_has & X.t_has & (X.d_sale != 0),
        "برگشت فقط در نرم‌افزار": (X.m_ret != 0) & (X.t_ret == 0),
        "برگشت فقط در سامانه مودیان": (X.t_ret != 0) & (X.m_ret == 0),
        "مغایرت مبلغ برگشت": (X.m_ret != 0) & (X.t_ret != 0) & (X.d_ret != 0)})
    bad = M.any(axis=1)
    X["max_abs"] = X[["d_sale", "d_ret", "d_net"]].abs().max(axis=1)
    mis = []
    meta = LC[META].reindex(idx)
    nb = int(bad.sum())
    P.log(f"  ● شناسه‌های مقایسه‌شده: {len(idx):,} | کاملاً مطابق: {int((~bad).sum()):,} | دارای اختلاف خام: {nb:,}")
    for i, (k, r) in enumerate(X[bad].iterrows()):
        if i % 2000 == 0:
            P.to(0.68 + 0.24 * i / max(nb, 1), "ساخت جدول مغایرت‌ها")
        notes = []
        if r.m_ret and not r.m_has:
            notes.append("فروش اصلی در نرم‌افزار نیست (احتمالاً دوره قبل)")
        if r.t_ret and not r.t_has:
            notes.append("فروش اصلی در مودیان نیست (احتمالاً دوره قبل)")
        mt = meta.loc[k]
        mis.append(dict(
            key=k, name=fa(mt.customer), nid=fa(mt.nid), eco=fa(mt.eco), ptype=fa(mt.ptype), sub=fa(mt["sub"]),
            reasons=[c for c in M.columns if M.at[k, c]], max_abs=int(r.max_abs),
            m_sale=int(r.m_sale), t_sale=int(r.t_sale), m_ret=int(r.m_ret), t_ret=int(r.t_ret),
            m_net=int(r.m_net), t_net=int(r.t_net), d_net=int(r.d_net), notes="؛ ".join(notes),
            m_rows=(f"فروش {r.m_srow}" if r.m_has else "") + (f" برگشت {r.m_rrows}" if r.m_rrows else ""),
            t_rows=(f"فروش {r.t_srow}" if r.t_has else "") + (f" برگشت {r.t_rrows}" if r.t_rrows else "")))

    P.to(0.92, "محاسبه خلاصه و نمودار")
    mm, mt_ = month_net(cs, cr), month_net(ts, tr)
    monthly = [dict(m=m, c=int(mm.get(m, 0)), t=int(mt_.get(m, 0))) for m in sorted(set(mm.index) | set(mt_.index))]

    def kpi(D, S0, S, R):
        return dict(rows=len(D), sales=len(S), sale=int(S.total.sum()), ret=int(R.total.sum()),
                    net=int(S.total.sum() - R.total.sum()),
                    cancel_amt=int(S0[S0.key.isin(cancelled)].total.sum()))

    result = dict(
        meta=dict(generated=datetime.now().strftime("%Y-%m-%d %H:%M"), threshold_toman=threshold_toman, rial=RIAL),
        kpi=dict(c=kpi(C, Sc, cs, cr), t=kpi(T, St, ts, tr), n_keys=len(idx), n_exact=int((~bad).sum()),
                 n_cancel=len(cancelled)),
        monthly=monthly, mis=mis, can=cancels, iss=issues,
        map=[dict(f=LABELS[f], c=cmp_[f] or "-", t=tmp[f] or "-") for f in ALIASES])
    return result


COLS = dict(
    mis="key name nid eco ptype sub reasons max_abs m_sale t_sale m_ret t_ret m_net t_net d_net notes m_rows t_rows",
    can="key status m_sale t_sale m_rows t_rows diff",
    iss="src row key ref type amount issue detail",
    map="f c t")


def write_db(result, db_path):
    import sqlite3
    db_path = Path(db_path)
    db_path.parent.mkdir(exist_ok=True)
    if db_path.exists():
        db_path.unlink()
    con = sqlite3.connect(db_path)
    for i, (name, cols) in enumerate(COLS.items()):
        P.to(0.93 + 0.07 * i / len(COLS), "ذخیره در دیتابیس")
        df = pd.DataFrame(result[name if name != "map" else "map"], columns=cols.split())
        if name == "mis":
            df["reasons"] = df.reasons.map(" | ".join)
        df.to_sql(name, con, index=False)
    con.execute("CREATE INDEX i_mis ON mis(max_abs)")
    con.execute("CREATE TABLE kv(k TEXT, v TEXT)")
    for k in ("meta", "kpi", "monthly"):
        con.execute("INSERT INTO kv VALUES(?,?)", (k, json.dumps(result[k], ensure_ascii=False)))
    con.commit()
    con.close()


def run(path_a, path_b, threshold_toman=THRESHOLD_TOMAN, discount_before=None):
    """سازگار با نسخه قبل: نوع هر فایل از روی ستون‌ها تشخیص داده می‌شود."""
    return run_files([(path_a, detect_type(path_a)), (path_b, detect_type(path_b))], threshold_toman, discount_before)


def show_items(items):
    print("\n  فایل‌های واردشده:")
    if not items:
        print("   (هیچ فایلی وارد نشده)")
    for n, (p, t) in enumerate(items, 1):
        mb = os.path.getsize(p) / 1e6 if os.path.exists(p) else 0
        print(f"   [{n}] {TYPE_FA[t]:<26} {mb:7.1f}MB  {p}")


def menu():
    items, th = [], THRESHOLD_TOMAN
    while True:
        show_items(items)
        print(f"\n  1) افزودن فایل   2) حذف فایل   3) شروع محاسبه   4) تغییر آستانه اختلاف (فعلی {th:,} تومان)   0) خروج")
        c = ask("  انتخاب شما: ").strip().translate(DIG)
        if c == "1":
            p = ask("  مسیر کامل فایل (اکسل یا CSV): ").strip().strip('"')
            if not os.path.isfile(p):
                print("  ✖ فایل پیدا نشد.")
                continue
            t = ask("  نوع فایل؟  1) ماژول (نرم‌افزار شرکت)   2) مودیان (سامانه مالیاتی): ").strip().translate(DIG)
            if t not in ("1", "2"):
                print("  ✖ نوع نامعتبر؛ فایل اضافه نشد.")
                continue
            t = "module" if t == "1" else "tax"
            auto = detect_type(p)
            if auto != t:
                print(f"  ⚠ عنوان ستون‌های این فایل شبیه «{TYPE_FA[auto]}» است، ولی شما «{TYPE_FA[t]}» انتخاب کردید.")
                if ask("  1) پذیرفتن تشخیص خودکار   2) نگه‌داشتن انتخاب من: ").strip().translate(DIG) != "2":
                    t = auto
            items.append((p, t))
        elif c == "2" and items:
            n = ask("  شماره فایل برای حذف: ").strip().translate(DIG)
            if n.isdigit() and 1 <= int(n) <= len(items):
                print("  حذف شد:", items.pop(int(n) - 1)[0])
        elif c == "3":
            if {"module", "tax"} <= {t for _, t in items}:
                return items, th
            print("  ✖ حداقل یک فایل ماژول و یک فایل مودیان لازم است.")
        elif c == "4":
            v = ask("  آستانه جدید (تومان): ").strip().translate(DIG).replace(",", "")
            th = int(v) if v.isdigit() else th
        elif c == "0":
            raise SystemExit(0)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    if "--no-rtl" not in sys.argv:
        setup_rtl()
    items, th = [], THRESHOLD_TOMAN
    for a in sys.argv[1:]:
        if a.startswith("--module="):
            items.append((a[9:], "module"))
        elif a.startswith("--tax="):
            items.append((a[6:], "tax"))
        elif a.startswith("--threshold="):
            th = int(a[12:])
        elif not a.startswith("--"):
            items.append((a, detect_type(a)))
    if not items:
        items, th = menu()
    flag = next((a for a in sys.argv if a.startswith("--discount=")), "")
    res = run_files(items, th, {"--discount=before": True, "--discount=after": False}.get(flag))
    db = Path(__file__).parent / "output" / "report.db"
    write_db(res, db)
    P.to(1.0, "پایان")
    k = res["kpi"]
    print(f"\n\nمغایرت‌های خام: {len(res['mis']):,} | کاملاً مطابق: {k['n_exact']:,} | ابطال کنار گذاشته‌شده: {k['n_cancel']:,}")
    if "--no-serve" not in sys.argv:
        from server import serve
        serve(db)


if __name__ == "__main__":
    main()