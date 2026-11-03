# -*- coding: utf-8 -*-
"""بيانات تاب الشركاء: كشوف المصروفات + نسبة زايا التجميعية، مع تشفير بكلمة سر."""
import re, json, base64, hashlib, hmac, os

NUM = (int, float)

def _txt(v):
    return re.sub(r"\s+", " ", v).strip() if isinstance(v, str) else None

# ---------------- تصنيف المصاريف (عدّل الكلمات هنا لو احتجت) ----------------
CATS = [
    ("شراء قسط الأرض",        ["الارض", "الأرض"]),
    ("استشاري ومكاتب هندسية", ["رينو", "صبور", "استشار"]),
    ("غفرة وأمن",             ["الغفر"]),
    ("مكتب المبيعات والتسويق", ["sales", "سيلز", "يفط", "اعلام"]),
    ("إقامة وتسكين",          ["الفندق", "التسكين", "السياحه"]),
    ("استرداد للعملاء",       ["استرداد"]),
    ("مصروفات الجهاز والمرافق", ["الجهاز", "المياه"]),
    ("مبالغ مسلّمة للأفراد (عهد وتحويلات)", ["عبدالله", "عبد الله", "ممدوح", "محمود اسماعيل", "الدبيحه", "مازن", "التسعير", "رحمي"]),
]
def categorize(name):
    n = (name or "").lower()
    for cat, keys in CATS:
        if any(k.lower() in n for k in keys): return cat
    return "أخرى"

def _date(s):
    m = re.search(r"(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})", s or "")
    return (int(m.group(3)), int(m.group(2))) if m else None

def parse_expenses(ws):
    stmts, cur = [], None
    for row in ws.iter_rows(values_only=True, max_col=5):
        a, b, c, d, e = (list(row) + [None] * 5)[:5]
        if isinstance(b, str) and "المصروفات كشف رقم" in b:
            n = int(re.search(r"\(\s*(\d+)\s*\)", b).group(1))
            ds = re.findall(r"\d{1,2}\s*/\s*\d{1,2}\s*/\s*\d{4}", b)
            if ds: ym = _date(ds[0]); m = f"{ym[0]}-{ym[1]:02d}"
            else:  m = "2026-03"        # الكشف 1 بدون تاريخ = حتى 31/3/2026
            cur = dict(n=n, m=m, title=_txt(b), items=[], total=None); continue
        if cur is None: continue
        if _txt(a) == "الاجمالي":
            cur["total"] = dict(total=c or 0, paid=d or 0, rem=e or 0); stmts.append(cur); cur = None; continue
        if isinstance(a, NUM) and _txt(b):
            if not any(isinstance(x, NUM) and x for x in (c, d, e)): continue
            name = _txt(b)
            cur["items"].append(dict(i=int(a), name=name, total=float(c or 0), paid=float(d or 0), rem=float(e or 0), cat=categorize(name)))
    for s in stmts:
        s["paid"] = sum(i["paid"] for i in s["items"])
        s["check"] = abs(s["paid"] - float(s["total"]["paid"] or 0)) < 1
    return stmts

def parse_zaya(ws):
    """تجميعي نسبة زايا من شيت تقفيل شركاء: لكل شهر (تعاقدات + نسبة زايا + بنودها)."""
    months, cur, mode = [], None, None
    for row in ws.iter_rows(values_only=True, max_col=6):
        a, b, c, d, e, f = (list(row) + [None] * 6)[:6]
        if isinstance(a, str) and "كشف تفصيلي مشروع" in a:
            ym = _date(a)
            if not ym: cur = None; continue
            cur = dict(m=f"{ym[0]}-{ym[1]:02d}", contracts=0.0, downs=0.0, units=0, zaya=0.0, items=[]); months.append(cur); mode = "sales"; continue
        if cur is None: continue
        if _txt(b) and "تسويه نسبه شركه زايا" in _txt(b): mode = "zaya"; continue
        if _txt(b) and "تسويه حساب المشروع في البنك" in _txt(b) or (_txt(a) and "تسويه حساب المشروع في البنك" in _txt(a)): mode = "done"; continue
        if mode == "sales" and _txt(c) == "الرصيد":
            cur["contracts"] = float(b or 0); cur["downs"] = float(d or 0); cur["units"] = int(e or 0)
        elif mode == "zaya":
            if isinstance(b, NUM) and isinstance(c, NUM) and _txt(d):
                cur["items"].append(dict(name=_txt(d), amount=float(c)))
            elif b is None and isinstance(c, NUM) and _txt(d) and _txt(d).startswith("نسبه شركه زايا"):
                cur["zaya"] = float(c); mode = "done"
    return [m for m in months if m["contracts"] or m["zaya"]]


# ---------------- تسوية الشركاء الشهرية (نفس ترتيب شيت تقفيل شركاء) ----------------
HDR = {"م", "المبلغ", "بيان", "التاريخ", "اجمالي تعاقدات", "اجمالي مقدمات", "عدد الوحدات"}
def _row(cells):
    """يرجّع (قبل, النص, بعد): الأرقام قبل البيان، البيان، والأرقام بعده."""
    txt, idx = None, None
    for i, c in enumerate(cells):
        t = _txt(c)
        if t and t not in HDR and not isinstance(c, NUM): txt, idx = t, i; break
    if txt is None: return None
    before = [float(c) for c in cells[:idx] if isinstance(c, NUM)]
    after = [float(c) for c in cells[idx + 1:] if isinstance(c, NUM)]
    return before, txt, after

def parse_settle(ws):
    months, cur, mode = [], None, None
    for row in ws.iter_rows(values_only=True, max_col=7):
        cells = list(row)
        r = _row(cells)
        if r is None: continue
        before, txt, after = r
        if "كشف تفصيلي مشروع" in txt:
            ym = _date(txt)
            if not ym: cur = None; continue
            cur = dict(m=f"{ym[0]}-{ym[1]:02d}", title=txt, sales=[], sales_total=None, proj=[], proj_total=None, zaya=[], zaya_total=None,
                       rev=[], rev_total=None, exp=[], exp_total=None, bal=None, actual=None, actual_label=None, net=None)
            months.append(cur); mode = "sales"; continue
        if cur is None: continue
        if txt == "تسويه رصيد المشروع": mode = "proj"; continue
        if txt == "تسويه نسبه شركه زايا": mode = "zaya"; continue
        if txt.startswith("تسويه حساب المشروع في البنك"): mode = "bank"; continue
        if txt == "الايرادات": mode = "rev"; continue
        if txt == "المصروفات": mode = "exp"; continue
        amt = before[-1] if before else None
        if mode == "sales":
            if txt == "الرصيد": cur["sales_total"] = dict(contracts=amt or 0, downs=after[0] if after else 0, units=after[1] if len(after) > 1 else 0)
            else: cur["sales"].append(dict(name=txt, contracts=amt or 0, downs=after[0] if after else 0, units=after[1] if len(after) > 1 else 0))
        elif mode == "proj":
            if txt == "الرصيد": cur["proj_total"] = amt
            else: cur["proj"].append(dict(name=txt, amount=amt, sub=after[0] if after else None))
        elif mode == "zaya":
            if txt.startswith("نسبه شركه زايا"): cur["zaya_total"] = amt; mode = "bank_wait"
            else: cur["zaya"].append(dict(name=txt, amount=amt))
        elif mode in ("rev", "exp"):
            tot = txt.startswith("اجمالي الايراد") or txt.startswith("اجمالي المصروفات")
            if tot:
                cur["rev_total" if mode == "rev" else "exp_total"] = amt
                if mode == "exp": mode = "bank_end"
            else: cur[mode].append(dict(name=txt, amount=amt))
        elif mode == "bank_end":
            if txt.startswith("رصيد البنك في"): cur["actual"] = amt; cur["actual_label"] = txt
            elif txt.startswith("رصيد البنك"): cur["bal"] = amt
            elif txt.startswith("الصافي"): cur["net"] = amt
    keep = []
    for m in months:
        if m["sales"] or m["sales_total"] and m["sales_total"]["contracts"] or m["zaya_total"] or any(x["amount"] for x in m["proj"]): keep.append(m)
    return keep


# ---------------- التقفيل الشهري (تسوية الشركاء) كما هو في شيت «تقفيل شركاء» ----------------
HDR = {"م", "المبلغ", "بيان", "التاريخ"}
def _split(row):
    cells = [(i, v) for i, v in enumerate(row) if v is not None and not (isinstance(v, str) and not v.strip())]
    strs = [(i, _txt(v)) for i, v in cells if isinstance(v, str)]
    nums = [(i, float(v)) for i, v in cells if isinstance(v, NUM) and not isinstance(v, bool)]
    return strs, nums

def group_proj(items):
    """تقسيم تسوية رصيد المشروع إلى مجموعة إيرادات (أزرق) ومجموعة مصروفات (برتقالي) + الرصيد."""
    its = [x for x in items if not x["total"]]
    if not its: return []
    k = next((i for i, x in enumerate(its) if "مصروفات" in x["label"]), None)
    if k is None:
        k = next((i for i, x in enumerate(its) if i > 0 and x.get("m") == 1), None)
    parts = [("rev", its)] if k is None else [("rev", its[:k]), ("exp", its[k:])]
    out = []
    for kind, part in parts:
        if not part: continue
        tot = sum((x["amount"] or 0) for x in part)
        # المجموع الفرعي = الخلية اللي قيمتها بتساوي مجموع المجموعة؛ وإلا نحسبه (يتجاهل أي رقم تاني على الشمال)
        sub = next((x["sub"] for x in part if x.get("sub") is not None and abs(x["sub"] - tot) < 1), tot)
        out.append(dict(kind=kind, sub=sub, items=[dict(label=x["label"], amount=x["amount"]) for x in part]))
    return out

def parse_closing(ws):
    out, cur, mode, sub = [], None, None, None
    for row in ws.iter_rows(values_only=True, max_col=7):
        strs, nums = _split(row)
        text = " ".join(s for _, s in strs)
        if "كشف تفصيلي مشروع" in text:
            ym = _date(text)
            if not ym: cur = None; continue
            cur = dict(m=f"{ym[0]}-{ym[1]:02d}", sales=[], proj=[], zaya=[], rev=[], exp=[], bank_items=[], bank={}); out.append(cur); mode = "sales"; sub = None; continue
        if cur is None: continue
        if "تسويه رصيد المشروع" in text: mode = "proj"; continue
        if "تسويه نسبه شركه زايا" in text: mode = "zaya"; continue
        if "تسويه حساب المشروع في البنك" in text: mode = "bank"; sub = None; continue
        if mode == "bank" and strs and all(s in ("الايرادات", "المصروفات") for _, s in strs):
            sub = "rev" if strs[0][1] == "الايرادات" else "exp"; continue
        lab = [(i, s) for i, s in strs if s not in HDR]
        if not lab: continue
        li, label = max(lab, key=lambda t: len(t[1]))
        before = [n for n in nums if n[0] < li]; after = [n for n in nums if n[0] > li]
        if not before: continue
        mnum = None
        if mode in ("proj", "zaya"):
            # المبلغ هو الرقم الملاصق للبيان؛ لو الخانة فاضية والرقم اللي قبلها هو (م) فقط، المبلغ فاضي
            adj = [n for n in before if n[0] == li - 1]
            if adj: amount = adj[0][1]
            elif len(before) == 1 and before[0][1] == int(before[0][1]) and before[0][1] <= 50: amount = None; mnum = before[0][1]
            else: amount = before[-1][1]
            if adj and len(before) > 1: mnum = before[0][1]
        else:
            amount = before[-1][1]
        if mode == "sales":
            cur["sales"].append(dict(label=label, contracts=amount, downs=after[0][1] if after else 0.0, units=int(after[1][1]) if len(after) > 1 else 0, total=label.startswith("الرصيد")))
        elif mode in ("proj", "zaya"):
            tot = (label == "الرصيد") if mode == "proj" else label.startswith("نسبه شركه زايا")
            it = dict(label=label, amount=amount, total=tot)
            if mode == "proj":
                it["m"] = mnum
                it["sub"] = after[0][1] if after else None   # المجموع الفرعي (الخلية المدموجة على اليسار)
            cur[mode].append(it)
        elif mode == "bank":
            b = cur["bank"]
            if label.startswith("اجمالي الايرداد"): b["rev_total"] = amount
            elif label.startswith("اجمالي المصروفات"): b["exp_total"] = amount
            elif label.startswith("رصيد البنك في") or label.startswith("اجمالي رصيد البنك"): b["actual"] = amount; b["actual_label"] = label
            elif label.startswith("رصيد البنك"): b["calc"] = amount
            elif label.startswith("الصافي"): b["net"] = amount
            elif sub in ("rev", "exp"): cur[sub].append(dict(label=label, amount=amount))
            else: cur["bank_items"].append(dict(label=label, amount=amount))
    for c in out: c["proj_g"] = group_proj(c["proj"])
    return [c for c in out if any(c[k] for k in ("sales", "proj", "zaya", "rev", "exp", "bank_items")) and any(abs(x.get("amount", x.get("contracts", 0)) or 0) > 0 for k in ("sales", "proj", "zaya", "rev", "exp", "bank_items") for x in c[k])]

def build_partners(wb):
    out = {"exp": parse_expenses(wb["كشوف مصروفات"]), "zaya": parse_zaya(wb["تقفيل شركاء"]), "settle": parse_settle(wb["تقفيل شركاء"]), "closing": parse_closing(wb["تقفيل شركاء"])}
    return out

