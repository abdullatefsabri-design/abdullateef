# اختبار قاعدة ABSENCE على بيانات مسوّية تغطي كل الحالات:
#   python3 salary_rules/test_absence.py
# إذا odoo موجود بالـ PYTHONPATH يشغّلها بـ safe_eval مال أودو نفسه، وإلا بـ exec بنفس builtins.
import os
import re
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

def run(code, data, lang, emp, contract, structure=None, categories=None, inputs=None, input_lines=None):  # noqa: E302
    """input_lines: أسطر إدخال بنفس الكود، يشغّلها مثل أودو: مرة لكل سطر ويجمع النتيجة."""
    if structure:
        code = re.sub(r"^STRUCTURE = 'IQD'", "STRUCTURE = '%s'" % structure, code, flags=re.M)
    inputs = dict(inputs or {})
    lines = list(input_lines or [])
    for c, line in inputs.items():
        line.code = c
        lines.append(line)
    payslip = Rec(env=Env(data, lang), employee_id=emp, date_from=date(2026, 10, 1), date_to=date(2026, 10, 31),
                  input_line_ids=RS(lines))
    ld = {'payslip': payslip, 'contract': contract, 'employee': emp,
          'inputs': inputs, 'categories': dict(categories or {}),
          'result': None, 'result_qty': 1.0, 'result_rate': 100, 'result_name': False}
    if not input_lines:
        safe_eval(code, ld, mode='exec', nocopy=True)
        return ld
    total, names = 0.0, []
    for line in input_lines:        # نفس حلقة same_type_input_lines بأودو (result_name ما يتصفّر)
        ld['inputs'][line.code] = line
        ld['result'] = None
        safe_eval(code, ld, mode='exec', nocopy=True)
        total += ld['result']
        names.append(ld['result_name'])
    ld['result'], ld['names'] = total, names
    return ld

# ---------------- data ----------------
ABS = Rec(name={'en_US': 'Absence', 'ar_001': 'غياب'})
# أنواع الإجازات مثل BRK: الزمنية نوع أودو القياسي (اسمه الإنكليزي غير)، والاستثناءات نوع مسوّى بالعربي
ZAM = Rec(name={'en_US': 'Extra Hours', 'ar_001': 'الزمنية'})
ZAM_EX = Rec(name={'en_US': 'الزمنية (استثناءات)', 'ar_001': 'الزمنية (استثناءات)'})
PAID = Rec(name={'en_US': 'Paid Time Off', 'ar_001': 'أيام الإجازة المدفوعة'})
WORK_H = Rec(name={'en_US': 'الإجازاة الساعية (عمل)', 'ar_001': 'الإجازاة الساعية (عمل)'})
LEAVE_TYPES = [ZAM, ZAM_EX, PAID, WORK_H]
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
def attendance(ci, co, in_mode='kiosk'):
    att.append(Rec(employee_id=emp, check_in=ci, check_out=co, in_mode=in_mode,
                   worked_hours=(co - ci).total_seconds() / 3600 if co else False))
def punch(t):
    zk.append(Rec(employee_id=emp, date=t))
def leave(df, dt, rdf, rdt, hours=False, half=False, kind=None):
    # بالساعات بدون نوع = زمنية، وغيرها = إجازة مدفوعة
    lv.append(Rec(employee_id=emp, state='validate', date_from=df, date_to=dt,
                  request_date_from=rdf, request_date_to=rdt,
                  request_unit_hours=hours, request_unit_half=half,
                  holiday_status_id=kind or (ZAM if hours else PAID)))

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
        'hr.leave': lv, 'resource.calendar.leaves': hol, 'hr.leave.type': LEAVE_TYPES}

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
# أكثر من سطر إدخال ABSENCE: أودو يشغّل القاعدة مرة لكل سطر، والأيام تنخصم مرة وحدة
r = run(code, data, 'en_US', emp, contract,
        input_lines=[Rec(code='ABSENCE', amount=1000.0), Rec(code='ABSENCE', amount=2000.0)])
assert abs(r['result'] - (-(5 * 1.5 * 900000 / 30) - 3000)) < 0.01, r['result']
assert r['names'][0].endswith(EXPECTED) and r['names'][1] is False, r['names']


# ---------------- حالات منفصلة (كل وحدة ببيانات لحالها) ----------------
def days_for(build):
    global we, att, zk, lv, hol
    we, att, zk, lv, hol = [], [], [], [], []
    build()
    d = {'hr.work.entry': we, 'hr.attendance': att, 'hr.attendance.zk.temp': zk,
         'hr.leave': lv, 'resource.calendar.leaves': hol, 'hr.leave.type': LEAVE_TYPES}
    m = re.search(r'\((.*)\)', run(code, d, 'en_US', emp, contract)['result_name'] or '')
    return m.group(1) if m else ''

def prev_night(d):   # داوم الشفت الليلي اللي قبل اليوم d (19:00 → 07:05)
    attendance(bg(2026, 10, d - 1, 19), bg(2026, 10, d, 7, 5))
    punch(bg(2026, 10, d - 1, 19)); punch(bg(2026, 10, d, 7, 5))

def zam(d, h1, h2, m1=0, m2=0, kind=None):
    leave(bg(2026, 10, d, h1, m1), bg(2026, 10, d, h2, m2), date(2026, 10, d), date(2026, 10, d), hours=True, kind=kind)

def parity(ts):
    """حضور مبني من البصمات بالتناوب دخول/خروج، مثل الجهاز: ضغطتين = سجل صفر، وبعدها كلشي ينقلب."""
    ts = sorted(ts)
    for t in ts:
        punch(t)
    for i in range(0, len(ts), 2):
        attendance(ts[i], ts[i + 1] if i + 1 < len(ts) else False)

def gaps(day_list, h1=8, h2=16):
    """سجلات غياب = ساعات الدوام اللي ما يغطيها حضور (مثل ما يولّدها الموديول)."""
    for d in day_list:
        segs = [(bg(2026, 10, d, h1), bg(2026, 10, d, h2))]
        for x in att:
            if not x.check_out or not x.worked_hours:
                continue
            new = []
            for a, b in segs:
                if x.check_out <= a or x.check_in >= b:
                    new.append((a, b))
                    continue
                if x.check_in > a:
                    new.append((a, x.check_in))
                if x.check_out < b:
                    new.append((x.check_out, b))
            segs = new
        for a, b in segs:
            if b > a:
                block(a, b)

def night_break(d):   # شفت ليلي 19:00–00:00 / 00:30–07:00
    block(bg(2026, 10, d, 19), bg(2026, 10, d + 1, 0))
    block(bg(2026, 10, d + 1, 0, 30), bg(2026, 10, d + 1, 7))

def punctual_week():
    ts = [bg(2026, 10, 4, 8), bg(2026, 10, 4, 8, 1), bg(2026, 10, 4, 16)]
    for d in (5, 6, 7, 8, 10, 11):
        ts += [bg(2026, 10, d, 7, 55), bg(2026, 10, d, 16)]
    parity(ts)
    gaps([4, 5, 6, 7, 8, 10, 11])

def forgot_out_7h():
    ts = []
    for d in (4, 5, 6, 7, 8):
        ts += [bg(2026, 10, d, 8), bg(2026, 10, d, 15)]
    ts.remove(bg(2026, 10, 4, 15))
    parity(ts)
    gaps([4, 5, 6, 7, 8], 8, 15)

CASES = [
    # (الحالة، البناء، الأيام المتوقعة)
    ('بدّل من ليلي لصباحي وغاب يومين', lambda: (
        prev_night(19), day_abs(19), day_abs(20)), '19/10، 20/10'),
    ('يوم التبديل غايب، ودخول اليوم الثاني 08:02', lambda: (
        prev_night(19), day_abs(19), night_abs(19),
        attendance(bg(2026, 10, 20, 8, 2), bg(2026, 10, 20, 16)),
        punch(bg(2026, 10, 20, 8, 2)), punch(bg(2026, 10, 20, 16))), '19/10'),
    ('يوم التبديل غايب، ودخول اليوم الثاني 07:20', lambda: (
        prev_night(19), day_abs(19), night_abs(19),
        attendance(bg(2026, 10, 20, 7, 20), bg(2026, 10, 20, 16)),
        punch(bg(2026, 10, 20, 7, 20)), punch(bg(2026, 10, 20, 16))), '19/10'),
    ('يوم التبديل ما جا، وعنده زمنية 08–10', lambda: (
        prev_night(19), day_abs(19), night_abs(19), punch(bg(2026, 10, 20, 7, 55)), zam(19, 8, 10)), '19/10'),
    ('بعد الليلي وصل 14:00 وطلع 16:00', lambda: (
        prev_night(19), block(bg(2026, 10, 19, 8), bg(2026, 10, 19, 14)),
        punch(bg(2026, 10, 19, 14)), punch(bg(2026, 10, 19, 16)),
        attendance(bg(2026, 10, 19, 14), bg(2026, 10, 19, 16))), '19/10'),
    ('بعد الليلي بصماته انرتبت غلط (حضور دقيقة)', lambda: (
        prev_night(19), day_abs(19), punch(bg(2026, 10, 19, 8)), punch(bg(2026, 10, 19, 8, 1)),
        punch(bg(2026, 10, 19, 16, 3)), attendance(bg(2026, 10, 19, 8), bg(2026, 10, 19, 8, 1))), ''),
    ('بعد الليلي دخل 07:50 ونسى الخروج', lambda: (
        prev_night(19), day_abs(19), punch(bg(2026, 10, 19, 7, 50)),
        attendance(bg(2026, 10, 19, 7, 50), False)), ''),
    ('ضغطتين بالجهاز بفرق ثواني ونسى الخروج', lambda: (
        day_abs(5), punch(bg(2026, 10, 5, 8, 0, 3)), punch(bg(2026, 10, 5, 8, 0, 9)),
        attendance(bg(2026, 10, 5, 8, 0, 3), bg(2026, 10, 5, 8, 0, 9))), ''),
    ('نسى الدخول، بصمة خروج وحدة 16:02', lambda: (
        day_abs(5), punch(bg(2026, 10, 5, 16, 2)), attendance(bg(2026, 10, 5, 16, 2), False)), ''),
    ('وصل متأخر 6 ساعات (14:00–16:00)', lambda: (
        block(bg(2026, 10, 5, 8), bg(2026, 10, 5, 14)), punch(bg(2026, 10, 5, 14)), punch(bg(2026, 10, 5, 16)),
        attendance(bg(2026, 10, 5, 14), bg(2026, 10, 5, 16))), '05/10'),
    ('ليلي بيه استراحة نص الليل: غاب 12 وداوم 13', lambda: (
        block(bg(2026, 10, 12, 19), bg(2026, 10, 13, 0)), block(bg(2026, 10, 13, 0, 30), bg(2026, 10, 13, 7)),
        punch(bg(2026, 10, 13, 19)), punch(bg(2026, 10, 14, 7)),
        attendance(bg(2026, 10, 13, 19), bg(2026, 10, 14, 7))), '12/10'),
    ('ليلي طلع 00:30 والباقي غياب', lambda: (
        block(bg(2026, 10, 13, 0, 30), bg(2026, 10, 13, 7)),
        attendance(bg(2026, 10, 12, 19), bg(2026, 10, 13, 0, 30)),
        punch(bg(2026, 10, 12, 19)), punch(bg(2026, 10, 13, 0, 30))), '12/10'),
    ('غاب ليلة 12', lambda: night_abs(12), '12/10'),
    ('غاب ليلة الجمعة', lambda: night_abs(16), ''),
    ('زمنية 08–12 ببداية الدوام وما جا (أودو أرشف ساعاتها)', lambda: (
        block(bg(2026, 10, 5, 13), bg(2026, 10, 5, 17)), zam(5, 8, 12)), '05/10'),
    ('وصل، طلع زمنية نص ساعة وما رجع', lambda: (
        day_abs(5), punch(bg(2026, 10, 5, 8)), punch(bg(2026, 10, 5, 8, 30)),
        attendance(bg(2026, 10, 5, 8), bg(2026, 10, 5, 8, 30)), zam(5, 8, 9, 30, 0)), '05/10'),
    ('زمنية بنص الدوام ورجع', lambda: (
        block(bg(2026, 10, 5, 10), bg(2026, 10, 5, 12)),
        attendance(bg(2026, 10, 5, 8), bg(2026, 10, 5, 10)), attendance(bg(2026, 10, 5, 12), bg(2026, 10, 5, 16)),
        zam(5, 10, 12)), ''),
    ('زمنية 08–10 وداوم 10–16', lambda: (
        attendance(bg(2026, 10, 5, 10), bg(2026, 10, 5, 16)), punch(bg(2026, 10, 5, 10)), punch(bg(2026, 10, 5, 16)),
        zam(5, 8, 10)), ''),
    ('غاب ليلتين، وحضور تقني 00:00 وزمنية 02–06', lambda: (
        night_abs(12), night_abs(13),
        attendance(bg(2026, 10, 13, 0), bg(2026, 10, 13, 0, 0, 1), in_mode='technical'),
        attendance(bg(2026, 10, 14, 0), bg(2026, 10, 14, 0, 0, 1), in_mode='technical'),
        leave(bg(2026, 10, 13, 2), bg(2026, 10, 13, 6), date(2026, 10, 13), date(2026, 10, 13), hours=True)),
     '12/10، 13/10'),
    # بصمات انرتبت غلط وصارت سجلات أطول من 16 ساعة (من خروج يوم لدخول اليوم اللي بعده)
    ('ضغطتين يوم الأربعاء وبعدها داوم 07 و 08 و 10', lambda: (
        parity([bg(2026, 10, 7, 8), bg(2026, 10, 7, 8, 1), bg(2026, 10, 7, 16, 5), bg(2026, 10, 8, 8, 10),
                bg(2026, 10, 8, 16, 3), bg(2026, 10, 10, 8, 5), bg(2026, 10, 10, 16)]),
        gaps([7, 8, 10])), ''),
    ('شفت 7 ساعات ونسى الخروج يوم 04 وداوم لحد 08', forgot_out_7h, ''),
    ('ضغطتين يوم الأحد وبعدها أسبوع كامل بوقته (الخميس للسبت 40 ساعة)', punctual_week, ''),
    # وصل قبل الزمنية، وأودو أرشف ساعاتها فأول غياب يبدي بعدها
    ('وصل 07:50، زمنية 09–12، غايب 13–16', lambda: (
        block(bg(2026, 10, 5, 13), bg(2026, 10, 5, 16)), punch(bg(2026, 10, 5, 7, 50)), punch(bg(2026, 10, 5, 9, 1)),
        attendance(bg(2026, 10, 5, 7, 50), bg(2026, 10, 5, 9, 1)), zam(5, 9, 12)), ''),
    ('وصل 08:00، طلع زمنية 10–14، غايب 14–16', lambda: (
        block(bg(2026, 10, 26, 14), bg(2026, 10, 26, 16)), punch(bg(2026, 10, 26, 8)), punch(bg(2026, 10, 26, 10)),
        attendance(bg(2026, 10, 26, 8), bg(2026, 10, 26, 10)), zam(26, 10, 14)), ''),
    ('ليلي داوم لنص الليل، زمنية 00–02، غايب 02:30–07', lambda: (
        attendance(bg(2026, 10, 12, 19), bg(2026, 10, 13, 0)), punch(bg(2026, 10, 12, 19)), punch(bg(2026, 10, 13, 0)),
        leave(bg(2026, 10, 13, 0), bg(2026, 10, 13, 2), date(2026, 10, 13), date(2026, 10, 13), hours=True),
        block(bg(2026, 10, 13, 2, 30), bg(2026, 10, 13, 7))), ''),
    ('صباحي 08–12 بس، والليلي ما جا وعنده زمنية 19–21', lambda: (
        attendance(bg(2026, 10, 12, 8), bg(2026, 10, 12, 12)), punch(bg(2026, 10, 12, 8)), punch(bg(2026, 10, 12, 12)),
        leave(bg(2026, 10, 12, 19), bg(2026, 10, 12, 21), date(2026, 10, 12), date(2026, 10, 12), hours=True),
        block(bg(2026, 10, 12, 21), bg(2026, 10, 13, 0)), block(bg(2026, 10, 13, 0), bg(2026, 10, 13, 7))), '12/10'),
    # الحضور بعد نص الليل يتبع يوم الشفت
    ('داوم ليلة 12 بالاستراحة، وما جا ليلة 13', lambda: (
        attendance(bg(2026, 10, 12, 19), bg(2026, 10, 13, 0)), attendance(bg(2026, 10, 13, 0, 30), bg(2026, 10, 13, 7)),
        punch(bg(2026, 10, 12, 19)), punch(bg(2026, 10, 13, 0)), punch(bg(2026, 10, 13, 0, 30)), punch(bg(2026, 10, 13, 7)),
        night_break(13)), '13/10'),
    ('نفسه بدون بصمات خام', lambda: (
        attendance(bg(2026, 10, 12, 19), bg(2026, 10, 13, 0)), attendance(bg(2026, 10, 13, 0, 30), bg(2026, 10, 13, 7)),
        night_break(13)), '13/10'),
    ('ليلة 12: غايب 19–01 وداوم 01–07', lambda: (
        block(bg(2026, 10, 12, 19), bg(2026, 10, 13, 0)), block(bg(2026, 10, 13, 0), bg(2026, 10, 13, 1)),
        attendance(bg(2026, 10, 13, 1), bg(2026, 10, 13, 7)), punch(bg(2026, 10, 13, 1)), punch(bg(2026, 10, 13, 7))), ''),
    ('نفسه وما جا ليلة 13', lambda: (
        block(bg(2026, 10, 12, 19), bg(2026, 10, 13, 0)), block(bg(2026, 10, 13, 0), bg(2026, 10, 13, 1)),
        attendance(bg(2026, 10, 13, 1), bg(2026, 10, 13, 7)), punch(bg(2026, 10, 13, 1)), punch(bg(2026, 10, 13, 7)),
        night_abs(13)), '13/10'),
    ('وصل 00:30 ليلة 12 وداوم، وما جا ليلة 13', lambda: (
        attendance(bg(2026, 10, 13, 0, 30), bg(2026, 10, 13, 7)), punch(bg(2026, 10, 13, 0, 30)), punch(bg(2026, 10, 13, 7)),
        block(bg(2026, 10, 12, 19), bg(2026, 10, 13, 0)), night_break(13)), '13/10'),
    ('شفت 07:00 غاب يوم 04، ويوم 05 دخل 05:50', lambda: (
        block(bg(2026, 10, 4, 7), bg(2026, 10, 4, 15)),
        attendance(bg(2026, 10, 5, 5, 50), bg(2026, 10, 5, 15)), punch(bg(2026, 10, 5, 5, 50)), punch(bg(2026, 10, 5, 15))),
     '04/10'),
    ('شفت 07:00: يوم 04 أربع ساعات، يوم 05 من 05:50 لـ 08:00', lambda: (
        attendance(bg(2026, 10, 4, 7), bg(2026, 10, 4, 11)), punch(bg(2026, 10, 4, 7)), punch(bg(2026, 10, 4, 11)),
        block(bg(2026, 10, 4, 11), bg(2026, 10, 4, 15)),
        attendance(bg(2026, 10, 5, 5, 50), bg(2026, 10, 5, 8)), punch(bg(2026, 10, 5, 5, 50)), punch(bg(2026, 10, 5, 8)),
        block(bg(2026, 10, 5, 8), bg(2026, 10, 5, 15))), '05/10'),
    # حالات طلعت من بيانات أيلول الحقيقية (154 و 303) ومن السجلات المكررة
    ('شفت 04:00: داوم الثلاثاء والخميس وغاب الأربعاء 21', lambda: (
        attendance(bg(2026, 10, 20, 3, 36), bg(2026, 10, 20, 12, 4)), punch(bg(2026, 10, 20, 3, 36)), punch(bg(2026, 10, 20, 12, 4)),
        block(bg(2026, 10, 21, 4), bg(2026, 10, 21, 12)),
        attendance(bg(2026, 10, 22, 3, 38), bg(2026, 10, 22, 12, 5)), punch(bg(2026, 10, 22, 3, 38)), punch(bg(2026, 10, 22, 12, 5))),
     '21/10'),
    ('تبديل من ليلي لصباحي الثلاثاء 13، وإجازة مدفوعة بالساعات 24 ساعة على ليلته', lambda: (
        attendance(bg(2026, 10, 12, 22, 52), bg(2026, 10, 13, 7, 8)), punch(bg(2026, 10, 12, 22, 52)), punch(bg(2026, 10, 13, 7, 8)),
        block(bg(2026, 10, 13, 9, 42), bg(2026, 10, 13, 15)), block(bg(2026, 10, 13, 23), bg(2026, 10, 13, 23, 59)),
        leave(bg(2026, 10, 13, 23), bg(2026, 10, 14, 23), date(2026, 10, 13), date(2026, 10, 14), hours=True, kind=PAID)), ''),
    ('ما جا، وعنده زمنية (استثناءات) 08–10', lambda: (
        day_abs(5), zam(5, 8, 10, kind=ZAM_EX)), '05/10'),
    ('غياب ساعتين مكرر 3 مرات (conflict)', lambda: (
        block(bg(2026, 10, 5, 8), bg(2026, 10, 5, 10)), block(bg(2026, 10, 5, 8), bg(2026, 10, 5, 10)),
        block(bg(2026, 10, 5, 8), bg(2026, 10, 5, 10))), ''),
    ('ليلي مكرر 3 مرات', lambda: (night_abs(12), night_abs(12), night_abs(12)), '12/10'),
    # الإجازة الساعية (عمل) تغطي يومها مثل أي إجازة (قرار الموارد معلّق)
    ('ما جا، وعنده إجازة ساعية (عمل) 08–10', lambda: (
        day_abs(5), zam(5, 8, 10, kind=WORK_H)), ''),
]
failed = []
for label, build, expected in CASES:
    got = days_for(build)
    if got != expected:
        failed.append('%s: المتوقع [%s] طلع [%s]' % (label, expected, got))
assert not failed, '\n'.join(failed)
print('OK: كل الحالات طلعت صح (%d حالة منفصلة)' % len(CASES))
