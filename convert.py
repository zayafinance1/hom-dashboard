"""يقرأ ملف الإكسل (شيت HOM) ويطلّع data.json للداشبورد.
الاستخدام:  python convert.py "مسار ملف الإكسل.xlsx" data.json
"""
import sys, os, json, datetime as dt
import openpyxl, pandas as pd

src = sys.argv[1]
out = sys.argv[2] if len(sys.argv) > 2 else "data.json"

wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
SHEETS = [("HOM", "هوم")]   # مشروع هوم فقط
frames = []
for sh, pname in SHEETS:
    rows = list(wb[sh].iter_rows(min_row=3, values_only=True, max_col=37))
    f = pd.DataFrame(rows[1:], columns=[str(c).strip() for c in rows[0]])
    f = f.loc[:, ~f.columns.duplicated()]
    f["project"] = pname
    frames.append(f)
df = pd.concat(frames, ignore_index=True)

for c in ["تاريخ السداد", "تاريخ التعاقد", "تاريخ الشيك"]:
    df[c] = pd.to_datetime(df[c], errors="coerce")
for c in ["المحصل", "المتبقى", "القيمة", "اجمالى الوحدة", "دفعه التعاقد", "مساحه الوحده", "مساحه الحديقه"]:
    df[c] = pd.to_numeric(df[c], errors="coerce")

def ym(s): return s.dt.strftime("%Y-%m")
def clean(v): return None if pd.isna(v) else v
def is_zero(v):   # نظام 0% (القيمة الرقمية صفر) وليس النص "صفر%"
    return (not isinstance(v, str)) and pd.notna(v) and v == 0
def system(v):
    if v is None or pd.isna(v): return "بدون"
    if is_zero(v): return "0%"
    if isinstance(v, str): return v.strip()
    return f"{v*100:g}%" if v < 1 else f"{v:g}%"

# لو القسط مسدّد وتاريخ السداد ناقص نستخدم تاريخ الشيك بدل ما نخسر المبلغ
miss = df["تاريخ السداد"].isna() & (df["المحصل"] > 0)
df["تاريخ السداد"] = df["تاريخ السداد"].fillna(df["تاريخ الشيك"].where(miss))
print("صفوف مسددة بدون تاريخ سداد:", int(miss.sum()))

# ---- 1) المتحصل من الأقساط (حسب تاريخ السداد)
inst_types = ["قسط", "قسط0", "صيانة", "جراج"]
a = df[df["نوع الدفعة"].isin(inst_types) & (df["المحصل"] > 0) & df["تاريخ السداد"].notna()].copy()
a["m"] = ym(a["تاريخ السداد"])
installments = []
for _, r in a.iterrows():
    installments.append(dict(p=r["project"], m=r["m"], date=r["تاريخ السداد"].strftime("%Y-%m-%d"), unit=clean(r["عمارة شقة"]),
        kind=r["نوع الدفعة"], method=clean(r["نوع السداد"]) or "-",
        bank=(clean(r["البنك المحصل به"]) or "غير محدد").strip(),
        zero=is_zero(r["نظام التعاقد"]), amount=float(r["المحصل"])))

# المستحق (المتبقي) لكل شهر استحقاق - للتوقع
d = df[df["نوع الدفعة"].isin(inst_types) & (df["المتبقى"] > 0) & df["تاريخ الشيك"].notna()].copy()
d["m"] = ym(d["تاريخ الشيك"])
rem = [dict(m=r["m"], d=r["تاريخ الشيك"].strftime("%Y-%m-%d"), u=clean(r["عمارة شقة"]), k=r["نوع الدفعة"],
            t=clean(r["نوع السداد"]) or "-", a=round(float(r["المتبقى"]), 2)) for _, r in d.iterrows()]

# ---- 2) المبيعات (حسب تاريخ التعاقد)
s = df[(df["نوع الدفعة"] == "اجمالي الوحده") & df["تاريخ التعاقد"].notna()].copy()
s["m"] = ym(s["تاريخ التعاقد"])
# الجراجات: (1) منفصلة = وحدات لها دفعات "جراج" -> القيمة = مجموع قيمة دفعات الجراج
#           (2) موزعة على الأقساط = مذكورة في الملاحظات على صف اجمالي الوحدة
gv = df[df["نوع الدفعة"] == "جراج"].groupby("عمارة شقة")["القيمة"].sum().to_dict()
def garage_kind(r):
    if r["عمارة شقة"] in gv: return "sep"
    n = str(r["الملاحظات"]) if pd.notna(r["الملاحظات"]) else ""
    if "الجراج موزع" in n or "شامل 100 الف جراج" in n: return "dist"
    return None
DIST_PRICE = 100000   # قيمة الجراج الموزع على الأقساط (افتراض: 100 ألف للوحدة)
sales = [dict(gk=garage_kind(r), gv=float(gv.get(r["عمارة شقة"], DIST_PRICE if garage_kind(r) == "dist" else 0)), p=r["project"], m=r["m"], date=r["تاريخ التعاقد"].strftime("%Y-%m-%d"), unit=clean(r["عمارة شقة"]),
              building=clean(r["عمارة"]), system=system(r["نظام التعاقد"]),
              area=clean(r["مساحه الوحده"]), garden=clean(r["مساحه الحديقه"]), value=float(r["اجمالى الوحدة"] or 0))
         for _, r in s.iterrows()]

# ---- 3) دفعات التعاقد (حسب تاريخ السداد)
c = df[(df["نوع الدفعة"] == "مقدم تعاقد") & (df["المحصل"] > 0) & df["تاريخ السداد"].notna()].copy()
c["m"] = ym(c["تاريخ السداد"])
downs = [dict(p=r["project"], m=r["m"], date=r["تاريخ السداد"].strftime("%Y-%m-%d"), unit=clean(r["عمارة شقة"]),
              bank=(clean(r["البنك المحصل به"]) or "غير محدد").strip(), amount=float(r["المحصل"]))
         for _, r in c.iterrows()]
pend = df[(df["نوع الدفعة"] == "مقدم تعاقد") & (df["المتبقى"] > 0)]["المتبقى"].sum()

json.dump(dict(updated=(open("modified.txt").read().strip() if os.path.exists("modified.txt") else dt.datetime.fromtimestamp(os.path.getmtime(src)).strftime("%Y-%m-%d %H:%M")), installments=installments, rem=rem,
               sales=sales, downs=downs, downs_pending=float(pend)),
          open(out, "w", encoding="utf-8"), ensure_ascii=False)
print("OK", len(installments), len(sales), len(downs), "->", out)

# ---- تاب الشركاء + تشفير الداشبورد كله بكلمة سر (DASH_PASSWORD)
d = json.load(open(out, encoding="utf-8"))
try:
    import partners
    _partners = partners.build_partners(wb)
    print("OK partners")
except Exception as ex:
    _partners = None
    print("تعذّر بناء تاب الشركاء:", ex)
pw = os.environ.get("DASH_PASSWORD", "")
ppw = os.environ.get("PARTNERS_PASSWORD", "")
if _partners is not None:
    if ppw:
        import dcrypt as _dc
        # بيانات الشركاء مشفّرة بكلمة سر منفصلة داخل الملف المشفّر الرئيسي
        d["partners_enc"] = _dc.encrypt(_partners, ppw, b"hom-partners-salt")
    else:
        print("تحذير: PARTNERS_PASSWORD غير مضبوط، تاب الشركاء لن يظهر (لا نكشف بياناته).")
if not pw:
    # أمان: من غير كلمة سر ما ننشرش بيانات مكشوفة
    os.remove(out)
    raise SystemExit("DASH_PASSWORD غير مضبوط: لم يتم إنشاء data.enc.json (البيانات القديمة المشفّرة كما هي).")
import dcrypt
enc_out = os.path.join(os.path.dirname(os.path.abspath(out)), "data.enc.json")
json.dump(dcrypt.encrypt(d, pw), open(enc_out, "w", encoding="utf-8"))
os.remove(out)            # لا نترك نسخة مكشوفة
print("OK encrypted ->", enc_out)
