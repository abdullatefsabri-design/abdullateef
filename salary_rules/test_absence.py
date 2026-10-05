# اختبار قاعدة ABSENCE على بيانات مسوّية تغطي كل الحالات:
#   python3 salary_rules/test_absence.py
# إذا odoo موجود بالـ PYTHONPATH يشغّلها بـ safe_eval مال أودو نفسه، وإلا بـ exec بنفس builtins.
import os
import sys
from datetime import date, datetime, time, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'tools'))
from odoo_absence import SAFE_BUILTINS  # noqa: E402

try:
    from odoo.tools.safe_eval import safe_eval
except ImportError:
    def safe_eval(code, ld, mode='exec', nocopy=True):
        ld['__builtins__'] = SAFE_BUILTINS
        exec(compile(code, 'ABSENCE', mode), ld)

H = timedelta(hours=1)
def bg(y, m, d, hh=0, mm=0, ss=0):  # Baghdad local -> naive UTC
    return datetime(y, m, d, hh, mm, ss) - 3 * H

class Rec:
    _ids = iter(range(1, 10**6))
    def __init__(self, **kw):
        self.id = next(Rec._ids)
        self.__dict__.update(kw)
    def __or__(self, other):
        return RS([r for r in (self, other) if r])
    @property
    def ids(self):
        return [self.id]

class RS(list):
    def mapped(self, f):
        return [getattr(r, f) for r in self]
    @property
    def ids(self):
        return [r.id for r in self]
    def __or__(self, other):
        return RS(list(self) + [r for r in other if r not in self])

def getval(rec, path, lang):
    v = rec
    for p in path.split('.'):
        v = getattr(v, p, False)
    if isinstance(v, dict):  # translated field (jsonb)
        v = v.get(lang) or v['en_US']
    return v

def leaf(rec, l, lang):
    path, op, right = l
    v = getval(rec, path, lang)
    if isinstance(v, datetime):
        if isinstance(right, str):
            if len(right) == 10:
                right += ' 23:59:59' if op in ('>', '<=') else ' 00:00:00'
            right = datetime.strptime(right, '%Y-%m-%d %H:%M:%S')
        elif isinstance(right, date) and not isinstance(right, datetime):
            right = datetime.combine(right, time.max if op in ('>', '<=') else time.min)
    if isinstance(v, Rec):
        v = v.id
    if op == 'in':
        return (not v and False in right) or (v and v in right)
    if op == '=':
        return (not v) if right is False else v == right
    if op == '!=':
        return v != right
    if v is False or v is None:
        return False
    return {'>=': v >= right, '<=': v <= right, '>': v > right, '<': v < right}[op]

def match(rec, domain, lang):
    def parse(i):
        t = domain[i]
        if t in ('|', '&'):
            a, i = parse(i + 1); b, i = parse(i)
            return (a or b) if t == '|' else (a and b), i
        if t == '!':
            a, i = parse(i + 1); return (not a), i
        return leaf(rec, t, lang), i + 1
    i, res = 0, True
    while i < len(domain):
        r, i = parse(i); res = res and r
    return res

class Model:
    def __init__(self, recs, lang):
        self.recs, self.lang = recs, lang
    def sudo(self):
        return self
    def with_context(self, **kw):
        return Model(self.recs, kw.get('lang', self.lang))
    def search(self, domain):
        return RS([r for r in self.recs if match(r, domain, self.lang)])

class Env:
    def __init__(self, data, lang):
        self.data, self.lang = data, lang
    def __getitem__(self, name):
        return Model(self.data.get(name, []), self.lang)

def run(code, data, lang, emp, contract, structure=None, categories=None, inputs=None):  # noqa: E302
    if structure:
        code = code.replace("STRUCTURE = 'IQD'", "STRUCTURE = '%s'" % structure)
    payslip = Rec(env=Env(data, lang), employee_id=emp, date_from=date(2026, 10, 1), date_to=date(2026, 10, 31))
    ld = {'payslip': payslip, 'contract': contract, 'employee': emp,
          'inputs': inputs or {}, 'categories': categories or {},
          'result': None, 'result_qty': 1.0, 'result_rate': 100, 'result_name': False}
    safe_eval(code, ld, mode='exec', nocopy=True)
    return ld

# ---------------- data ----------------
ABS = Rec(name={'en_US': 'Absence', 'ar_001': 'غياب'})
cal = Rec(name='Day')
company = Rec(name='BRK')
pg = Rec(absence_policy_id=Rec(), work_days_per_month=26)
emp = Rec(category_ids=RS([Rec(name='موظف')]), resource_calendar_id=cal)
contract = Rec(date_start=date(2025, 1, 1), date_end=False, policy_group_id=pg, wage=900000.0,
               resource_calendar_id=cal, company_id=company)

we, att, zk, lv, hol = [], [], [], [], []
def block(a, b, state='draft'):
    we.append(Rec(employee_id=emp, work_entry_type_id=ABS, state=state, date_start=a, date_stop=b,
                  duration=(b - a).total_seconds() / 3600))
def day_abs(d):   # 08:00-16:00
    block(bg(2026, 10, d, 8), bg(2026, 10, d, 16))
def night_abs(d, split=True):  # 19:00 -> 07:00 next day
    if split:
        block(bg(2026, 10, d, 19), bg(2026, 10, d + 1, 0))
        block(bg(2026, 10, d + 1, 0), bg(2026, 10, d + 1, 7))
    else:
        block(bg(2026, 10, d, 19), bg(2026, 10, d + 1, 7))
def attendance(ci, co):
    att.append(Rec(employee_id=emp, check_in=ci, check_out=co,
                   worked_hours=(co - ci).total_seconds() / 3600 if co else False))
def punch(t):
    zk.append(Rec(employee_id=emp, date=t))
def leave(df, dt, rdf, rdt, hours=False, half=False):
    lv.append(Rec(employee_id=emp, state='validate', date_from=df, date_to=dt,
                  request_date_from=rdf, request_date_to=rdt,
                  request_unit_hours=hours, request_unit_half=half))

# 03 Sat: public holiday (National Day) with absence entries
day_abs(3)
hol.append(Rec(resource_id=False, time_type='leave', calendar_id=False, company_id=company,
               date_from=bg(2026, 10, 3, 0), date_to=bg(2026, 10, 3, 23, 59, 59)))
# 05 Mon: plain absence
day_abs(5)
# 06 Tue: shift change (old day + new night blocks), worked the night shift
day_abs(6); night_abs(6)
attendance(bg(2026, 10, 6, 19), bg(2026, 10, 7, 7))
# 07 Wed: punches misordered -> 0h attendance
day_abs(7); punch(bg(2026, 10, 7, 8)); punch(bg(2026, 10, 7, 8, 1)); punch(bg(2026, 10, 7, 16, 5))
attendance(bg(2026, 10, 7, 8), bg(2026, 10, 7, 8, 1))
# 08 Thu: single punch (forgot)
day_abs(8); punch(bg(2026, 10, 8, 8, 10))
# 09 Fri: absence on Friday
day_abs(9)
# 11 Sun: same punch imported twice
day_abs(11); punch(bg(2026, 10, 11, 8, 10)); punch(bg(2026, 10, 11, 8, 10))
# 12 Mon: full-day leave
day_abs(12); leave(bg(2026, 10, 12, 8), bg(2026, 10, 12, 16), date(2026, 10, 12), date(2026, 10, 12))
# 13 Tue: half-day leave
day_abs(13); leave(bg(2026, 10, 13, 8), bg(2026, 10, 13, 12), date(2026, 10, 13), date(2026, 10, 13), half=True)
# 14 Wed: wrong-shift blocks, attendance split by break (4h + 3.5h), no raw punches kept
night_abs(14)
attendance(bg(2026, 10, 14, 8), bg(2026, 10, 14, 12)); attendance(bg(2026, 10, 14, 12, 30), bg(2026, 10, 14, 16))
# 19 Mon: shift change, really absent; only prev-night checkout + next-morning check-in in window
day_abs(19); night_abs(19)
punch(bg(2026, 10, 19, 7, 5)); punch(bg(2026, 10, 20, 7, 55))
# 21 Wed: night shift absence split at midnight
night_abs(21)
# 22 Thu: no-show, hourly leave (zamaniya) at shift start
day_abs(22); leave(bg(2026, 10, 22, 8), bg(2026, 10, 22, 10), date(2026, 10, 22), date(2026, 10, 22), hours=True)
# 26 Mon: arrived 08:00, zamaniya 10:00-14:00 after arriving, never came back
block(bg(2026, 10, 26, 10), bg(2026, 10, 26, 16))
punch(bg(2026, 10, 26, 8)); punch(bg(2026, 10, 26, 10))
attendance(bg(2026, 10, 26, 8), bg(2026, 10, 26, 10))
leave(bg(2026, 10, 26, 10), bg(2026, 10, 26, 14), date(2026, 10, 26), date(2026, 10, 26), hours=True)
# 27 Tue: 2h late only
block(bg(2026, 10, 27, 8), bg(2026, 10, 27, 10))
# 28 Wed: cancelled absence entry
block(bg(2026, 10, 28, 8), bg(2026, 10, 28, 16), state='cancelled')
# 29 Thu: Thursday night shift, tail on Friday
night_abs(29, split=False)

data = {'hr.work.entry': we, 'hr.attendance': att, 'hr.attendance.zk.temp': zk,
        'hr.leave': lv, 'resource.calendar.leaves': hol}

code = open(os.path.join(HERE, 'absence.py'), encoding='utf-8').read()
EXPECTED = '(05/10، 19/10، 21/10، 22/10، 29/10)'

for lang in ('en_US', 'ar_001'):   # ar_001 = اللي يحسب الإيصال واجهته عربي
    r = run(code, data, lang, emp, contract)
    assert r['result_name'].endswith(EXPECTED), (lang, r['result_name'])
    assert abs(r['result'] - -(5 * 1.5 * 900000 / 30)) < 0.01, r['result']

# المدخلات: TRFOD تنضاف للأجر اليومي، و ABSENCE تنطرح
r = run(code, data, 'ar_001', emp, contract, inputs={'TRFOD': Rec(amount=60000.0), 'ABSENCE': Rec(amount=1000.0)})
assert abs(r['result'] - (-(5 * 1.5 * 960000 / 30) - 1000)) < 0.01, r['result']
# الدولار: اليومي = الأجر / work_days_per_month، والخصم ما يعبر المتبقي
r = run(code, data, 'ar_001', emp, contract, structure='USD',
        categories={'BASIC': 900000.0, 'ALW': 0.0, 'DED': -800000.0})
assert r['result'] == -100000.0, r['result']
r = run(code, data, 'ar_001', emp, contract, structure='USD', categories={'BASIC': 900000.0, 'ALW': 0.0, 'DED': 0.0})
assert abs(r['result'] - -(5 * 1.5 * 900000 / 26)) < 0.01, r['result']
# معفي من البصمة، أو بدون مجموعة سياسات ← صفر
emp2 = Rec(category_ids=RS([Rec(name='معفي من البصمة')]), resource_calendar_id=cal)
assert run(code, data, 'en_US', emp2, contract)['result'] == 0
c2 = Rec(**dict(contract.__dict__, policy_group_id=False))
assert run(code, data, 'en_US', emp, c2)['result'] == 0
# العقد ينتهي 20/10
c3 = Rec(**dict(contract.__dict__, date_end=date(2026, 10, 20)))
assert run(code, data, 'en_US', emp, c3)['result_name'].endswith('(05/10، 19/10)')
print('OK: كل الحالات طلعت صح')
