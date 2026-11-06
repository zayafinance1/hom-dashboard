#!/usr/bin/env python3
"""convert.py - يحوّل ملف الإكسل إلى data.enc.json (مشفّر) للداشبورد.
الاستخدام (زي ما الـ workflow بيشغّله):  python convert.py book.xlsx data.json
- الباسورد من Secret اسمه DASH_PASSWORD
- المشروع من متغير PROJECT (الافتراضي: دريم). سيبه فاضي ('') لكل المشاريع.
الملف الناتج data.enc.json بيتكتب في نفس مكان الاسكريبت."""
import base64, gzip, json, os, re, sys
from datetime import datetime, timedelta, timezone
import pandas as pd
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes

DEFAULT_PROJECT = 'دريم'
XLSX = sys.argv[1] if len(sys.argv) > 1 else 'book.xlsx'
PW = os.environ.get('DASH_PASSWORD', '')
if len(PW) < 8:
    sys.exit('DASH_PASSWORD secret مش موجود أو قصير (8 حروف على الأقل).')
ONLY = os.environ['PROJECT'].strip() if 'PROJECT' in os.environ else DEFAULT_PROJECT
HERE = os.path.dirname(os.path.abspath(__file__))
# --- اقرا الشيت الأول: دوّر على صف العناوين (اللي فيه "اسم المشروع") بدل ما نفترض مكانه ---
raw_top = pd.read_excel(XLSX, sheet_name=0, header=None, nrows=15)
hdr_row = None
for i in range(len(raw_top)):
    if any(str(v).strip() == 'اسم المشروع' for v in raw_top.iloc[i].tolist()):
        hdr_row = i; break
if hdr_row is None:
    names = pd.ExcelFile(XLSX).sheet_names
    sys.exit('مش لاقي عمود "اسم المشروع" في أول 15 صف من الشيت الأول. الشيتات: %s | أول صف: %s'
             % (names, [str(v) for v in raw_top.iloc[0].tolist()[:12]]))
df = pd.read_excel(XLSX, sheet_name=0, header=hdr_row)
df.columns = [str(c).strip() for c in df.columns]
# لو عمود "المحصل" اتغير اسمه لـ "القيمة" (فبقى عندنا عمودين "القيمة")، التاني هو المحصل
_cols = list(df.columns)
if 'المحصل' not in _cols and _cols.count('القيمة') == 2:
    _i = len(_cols) - 1 - _cols[::-1].index('القيمة')
    _cols[_i] = 'المحصل'
    df.columns = _cols
need = ['اسم المشروع', 'نوع الدفعة', 'تاريخ الاستحقاق', 'اسم العميل', 'تاريخ التعاقد', 'فيلا- ارض', 'القطعة',
        'اجمالى الوحدة', 'القيمة', 'المحصل', 'المتبقى', 'تاريخ السداد', 'البنك المحصل به']
missing = [c for c in need if c not in df.columns]
if missing:
    sys.exit('الأعمدة دي ناقصة في الملف اللي اتحمّل من Drive: %s\nالأعمدة الموجودة: %s\nغالبًا GDRIVE_FILE_ID بيشاور على ملف غلط.'
             % (missing, [c for c in df.columns if not c.startswith('Unnamed')]))
df = df.loc[:, ~df.columns.duplicated()]
df = df[df['اسم المشروع'].notna()].copy()
df['اسم المشروع'] = df['اسم المشروع'].astype(str).str.strip()
if ONLY: df = df[df['اسم المشروع'] == ONLY].copy()
if df.empty:
    sys.exit('مفيش صفوف للمشروع "%s". المشاريع الموجودة في الملف: %s' % (ONLY, sorted(set(pd.read_excel(XLSX, sheet_name=0, header=hdr_row)['اسم المشروع'].dropna().astype(str).str.strip()))))
df['t'] = df['نوع الدفعة'].fillna('').astype(str).str.strip()
stats = {}

# تاريخ التحديث: أي تاريخ في الصف اللي فوق صف العناوين (كان الخلية Y2)، وإلا آخر تاريخ سداد، وإلا النهارده
upd = None
if hdr_row >= 1:
    import datetime as _dt
    for v in pd.read_excel(XLSX, sheet_name=0, header=None, nrows=hdr_row).iloc[hdr_row - 1].tolist():
        if isinstance(v, (_dt.datetime, pd.Timestamp)) and not pd.isna(v):
            upd = pd.Timestamp(v)
if upd is None:
    m = pd.to_datetime(df['تاريخ السداد'], errors='coerce').max()
    upd = m if not pd.isna(m) else pd.Timestamp.today()
updated = upd.strftime('%Y-%m-%d')

def norm_channel(v):
    if pd.isna(v) or str(v).strip() == '': return 'غير محدد'
    toks = [x.strip() for x in re.split(r'\s+-\s+|-(?=\s*[A-Za-zء-ي])', str(v)) if x.strip()]
    out = []
    for t in toks:
        l = t.lower().replace(' ', '')
        if l.startswith('eg'):
            m = re.search(r'\d+', l); n = 'EG ' + (m.group(0).zfill(3) if m else '')
        elif 'cib' in l: n = 'CIB'
        elif 'saib' in l: n = 'Saib'
        elif 'كاش' in l: n = 'كاش'
        elif 'الكويت' in l: n = 'الكويت الوطني'
        elif 'الاهلي' in l: n = 'الاهلي المصري'
        elif 'بنكمصر' in l: n = 'بنك مصر'
        elif l == 'mid': n = 'ميد بنك'
        elif 'كابتن' in l: n = 'الكابتن مازن'
        else: n = 'أخرى'
        if n not in out: out.append(n)
    return out[0] if len(out) == 1 else 'أكثر من جهة'

def is_inst(t): return 'قسط' in t and 'صيانة' not in t and 'جراج' not in t
def is_con(t): return t in ('تعاقد', 'مقدم', 'تعاقد + جراج')
def cat_of(t):
    """0 قسط | 1 تعاقد | 2 جراج | 3 أخرى | 4 صيانة"""
    if 'صيانة' in t: return 4
    if is_con(t): return 1
    if is_inst(t): return 0
    if 'جراج' in t: return 2
    return 3

projects = df['اسم المشروع'].value_counts().index.tolist()
pidx = {p: i for i, p in enumerate(projects)}
D = lambda x: x.strftime('%Y-%m-%d')
r2 = lambda x: round(float(x), 2)

pay = df[df['t'] != ''].copy()
# ---- collected
col = pay[pay['المحصل'].fillna(0) > 0].copy()
col['d'] = col['تاريخ السداد'].fillna(col['تاريخ الاستحقاق'])
stats['collected_rows_no_date'] = int(col['d'].isna().sum())
col = col[col['d'].notna()]
inst = [[D(r['d']), r2(r['المحصل']), norm_channel(r['البنك المحصل به']), pidx[r['اسم المشروع']]]
        for _, r in col[col['t'].map(is_inst)].iterrows()]
con = [[D(r['d']), r2(r['المحصل']), norm_channel(r['البنك المحصل به']), pidx[r['اسم المشروع']],
        str(r['القطعة']).strip() if pd.notna(r['القطعة']) else '', str(r['اسم العميل']).strip() if pd.notna(r['اسم العميل']) else '']
       for _, r in col[col['t'].map(is_con)].iterrows()]

# ---- كل المتحصل (للتابة "المتحصل من المشروع")
colall = [[D(r['d']), r2(r['المحصل']), norm_channel(r['البنك المحصل به']), pidx[r['اسم المشروع']], cat_of(r['t'])]
          for _, r in col.iterrows()]

# ---- remaining installments
rem_df = pay[(pay['t'].map(is_inst) | pay['t'].str.contains('صيانة')) & (pay['المتبقى'].fillna(0) > 0)]
stats['remaining_rows_no_date'] = int(rem_df['تاريخ الاستحقاق'].isna().sum())
rem = [[D(r['تاريخ الاستحقاق']) if pd.notna(r['تاريخ الاستحقاق']) else '', r2(r['المتبقى']),
        str(r['القطعة']).strip() if pd.notna(r['القطعة']) else '',
        str(r['اسم العميل']).strip() if pd.notna(r['اسم العميل']) else '', pidx[r['اسم المشروع']],
        1 if 'صيانة' in r['t'] else 0]
       for _, r in rem_df.iterrows()]

# ---- sales: every row carrying اجمالى الوحدة (the unit total may sit on the header row
# or on the first payment row). Unit total = sum of those rows per (project, unit, client);
# matches the reference sheet "Sheet1". Contract date = first contract date, else first payment date.
df['unit'] = df['القطعة'].astype(str).str.strip()
df['client'] = df['اسم العميل'].fillna('').astype(str).str.strip()
_is_g = df['t'].str.contains('جراج')
h = df[(df['اجمالى الوحدة'].fillna(0) > 0) & ~_is_g].copy()
def utype(v):
    v = str(v).strip()
    return v if v in ('فيلا', 'شقة', 'ارض') else 'غير محدد'
sales = []; nodate = 0
for (pj, unit, cl), g in h.groupby(['اسم المشروع', 'unit', 'client'], sort=False):
    d = g['تاريخ التعاقد'].dropna()
    if d.empty:
        allrows = df[(df['اسم المشروع'] == pj) & (df['unit'] == unit) & (df['client'] == cl)]
        d = allrows['تاريخ الاستحقاق'].dropna()
    if d.empty: nodate += 1; continue
    ty = g['فيلا- ارض'].map(utype); ty = ty[ty != 'غير محدد']
    sales.append([D(d.iloc[0]), unit, ty.iloc[0] if len(ty) else 'غير محدد', r2(g['اجمالى الوحدة'].sum()), cl, pidx[pj]])
stats['sales_no_date_skipped'] = nodate

# ---- الجراج: (أ) قيمة الجراج لكل وحدة (من صفوف إجمالى الوحدة اللي نوعها "جراج") - بتتعرض منفصلة عن المبيعات
unit_date = {}
for (pj, unit, cl), g in h.groupby(['اسم المشروع', 'unit', 'client'], sort=False):
    d = g['تاريخ التعاقد'].dropna()
    if not d.empty: unit_date.setdefault((pj, unit), D(d.iloc[0]))
gsales = []
gh = df[(df['اجمالى الوحدة'].fillna(0) > 0) & _is_g]
for (pj, unit, cl), g in gh.groupby(['اسم المشروع', 'unit', 'client'], sort=False):
    d = g['تاريخ التعاقد'].dropna()
    dd = D(d.iloc[0]) if not d.empty else unit_date.get((pj, unit))
    if not dd:
        ar = df[(df['اسم المشروع'] == pj) & (df['unit'] == unit)]['تاريخ الاستحقاق'].dropna()
        dd = D(ar.min()) if not ar.empty else ''
    gsales.append([dd, unit, r2(g['اجمالى الوحدة'].sum()), cl, pidx[pj]])
# (ب) أقساط الجراج المتبقية
remg = [[D(r['تاريخ الاستحقاق']) if pd.notna(r['تاريخ الاستحقاق']) else '', r2(r['المتبقى']),
         str(r['القطعة']).strip() if pd.notna(r['القطعة']) else '', str(r['اسم العميل']).strip() if pd.notna(r['اسم العميل']) else '', pidx[r['اسم المشروع']]]
        for _, r in pay[(pay['t'].map(cat_of) == 2) & (pay['المتبقى'].fillna(0) > 0)].iterrows()]


# ---- maintenance & infrastructure (pure types only)
MI = {'صيانة': 0, 'الصيانة': 0, 'البنية التحتية': 1, 'بنية تحتية': 1}
mi_df = pay[pay['t'].isin(MI.keys())]
mi = [[D(r['تاريخ الاستحقاق']) if pd.notna(r['تاريخ الاستحقاق']) else '',
       r2(r['القيمة'] if pd.notna(r['القيمة']) else (r['المحصل'] if pd.notna(r['المحصل']) else 0)),
       r2(r['المحصل']) if pd.notna(r['المحصل']) else 0,
       r2(r['المتبقى']) if pd.notna(r['المتبقى']) else 0,
       str(r['القطعة']).strip() if pd.notna(r['القطعة']) else '',
       str(r['اسم العميل']).strip() if pd.notna(r['اسم العميل']) else '',
       pidx[r['اسم المشروع']], MI[r['t']]] for _, r in mi_df.iterrows()]

try:
    from zoneinfo import ZoneInfo
    _now = datetime.now(ZoneInfo('Africa/Cairo'))
except Exception:
    _now = datetime.now(timezone(timedelta(hours=2)))
payload = dict(updated=updated, synced=_now.strftime('%Y-%m-%d %H:%M'), projects=projects, inst=inst, con=con, rem=rem, sales=sales, mi=mi, col=colall, gsales=gsales, remg=remg)
raw = gzip.compress(json.dumps(payload, ensure_ascii=False, separators=(',', ':')).encode('utf-8'), 9)
salt, iv = os.urandom(16), os.urandom(12)
key = PBKDF2HMAC(hashes.SHA256(), 32, salt, 200000).derive(PW.encode('utf-8'))
ct = AESGCM(key).encrypt(iv, raw, None)
blob = json.dumps({'s': base64.b64encode(salt).decode(), 'i': base64.b64encode(iv).decode(), 'c': base64.b64encode(ct).decode()})

open(os.path.join(HERE, 'data.enc.json'), 'w', encoding='utf-8').write(blob)
print('OK updated', updated, 'projects', projects, 'inst', len(inst), 'con', len(con), 'rem', len(rem), 'sales', len(sales), 'mi', len(mi), 'col', len(colall), 'gsales', len(gsales), 'remg', len(remg), stats)
