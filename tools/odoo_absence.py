#!/usr/bin/env python3
# أداة قاعدة ABSENCE على أودو اللايف (XML-RPC). ما تحتاج متصفح.
#
#   python3 tools/odoo_absence.py check   [--month 2026-09]   قراءة بس: يتأكد من افتراضات القاعدة
#   python3 tools/odoo_absence.py compare  --month 2026-09    قراءة بس: يشغّل القاعدة الجديدة على إيصالات
#                                                              الشهر ويقارنها ويه سطر ABSENCE الحالي
#   python3 tools/odoo_absence.py apply                       يعرض شنو راح يتغيّر بدون ما يكتب
#   python3 tools/odoo_absence.py apply --yes                 نسخة احتياطية من 37 و 57 ← يكتب الكود الجديد
#         [--clear-condition]                                 إذا القاعدة بيها شرط (Condition) يخلّيه Always True
#   python3 tools/odoo_absence.py restore FILE.json RULE_ID --yes  يرجّع القاعدة من نسخة احتياطية
#
# متغيرات البيئة:
#   ODOO_URL      (افتراضي https://brk.ejaferp.com)
#   ODOO_DB       (إذا السيرفر بيه قاعدة بيانات وحدة يلكاها بروحه)
#   ODOO_LOGIN    اسم الدخول
#   ODOO_API_KEY  مفتاح API (التفضيلات ← أمان الحساب ← New API Key)، وله نفس صلاحيات المستخدم
#
# المستخدم لازم يكون مدير حضور ومسؤول إجازات على الأقل: القاعدة بأودو تبحث بـ sudo()، وعبر XML-RPC
# صلاحيات المستخدم تنطبق، فإذا ناقصة يختفي حضور وإجازات بدون أي خطأ. compare يوكف إذا ناقصة.
#
# compare يشغّل نفس نص salary_rules/absence.py على بيانات اللايف (قراءة بس): كل payslip.env[...].search
# وكل حقل يقراه يروح طلب XML-RPC، فما أكو نسخة ثانية من المنطق تختلف عن القاعدة.

import argparse
import builtins
import csv
import datetime as dt
import functools
import json
import os
import re
import sys
import xmlrpc.client
from collections import defaultdict
from urllib.parse import urlsplit

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RULE_FILE = os.path.join(ROOT, 'salary_rules', 'absence.py')
BACKUP_DIR = os.path.join(ROOT, 'salary_rules', 'backup')
REPORT_DIR = os.path.join(ROOT, 'reports')
RULES = {37: 'IQD', 57: 'USD'}   # rule id ← قيمة STRUCTURE
RULE_CODE = 'ABSENCE'
COND_FIELDS = ('condition_select', 'condition_python', 'condition_range',
               'condition_range_min', 'condition_range_max')
# المجموعات اللي تخلّي بحث XML-RPC يشوف نفس اللي تشوفه القاعدة بـ sudo()
SUDO_GROUPS = ('hr_attendance.group_hr_attendance_manager', 'hr_holidays.group_hr_holidays_user')
RELATIONAL = ('many2one', 'one2many', 'many2many')

# نفس builtins اللي يسمح بيها safe_eval مال أودو
SAFE_BUILTINS = {name: getattr(builtins, name) for name in (
    'bytes', 'str', 'bool', 'int', 'float', 'enumerate', 'dict', 'list', 'tuple', 'map', 'abs',
    'min', 'max', 'sum', 'filter', 'sorted', 'round', 'len', 'repr', 'set', 'all', 'any', 'ord',
    'chr', 'divmod', 'isinstance', 'range', 'zip', 'Exception')}
SAFE_BUILTINS['reduce'] = functools.reduce


def _import(name, globals=None, locals=None, fromlist=None, level=-1):
    # مثل safe_eval: strftime وغيرها يستوردون time داخلياً، فلازم __import__ يرجّع الموديولات المسموحة بس
    if name not in ('_strptime', 'math', 'time'):
        raise ImportError(name)
    return __import__(name, globals, locals, fromlist or (), 0)


SAFE_BUILTINS['__import__'] = _import


def rule_source(structure):
    with open(RULE_FILE, encoding='utf-8') as f:
        code = f.read()
    # بس سطر الكود (أول السطر)، مو التعليق اللي بالرأس
    marker = re.compile(r"^STRUCTURE = 'IQD'", re.M)
    if len(marker.findall(code)) != 1:
        sys.exit("لازم يكون سطر STRUCTURE = 'IQD' مرة وحدة بالضبط بـ %s" % RULE_FILE)
    code = marker.sub("STRUCTURE = '%s'" % structure, code)
    compile(code, 'ABSENCE', 'exec')
    return code


# ---------------------------------------------------------------- XML-RPC

class Odoo:
    def __init__(self):
        url = os.environ.get('ODOO_URL', 'https://brk.ejaferp.com')
        parts = urlsplit(url if '://' in url else 'https://' + url)
        self.url = '%s://%s' % (parts.scheme, parts.netloc)   # https://x/odoo ← https://x
        self.login = os.environ.get('ODOO_LOGIN')
        self.key = os.environ.get('ODOO_API_KEY')
        if not self.login or not self.key:
            sys.exit('حدد ODOO_LOGIN و ODOO_API_KEY بمتغيرات البيئة')
        common = xmlrpc.client.ServerProxy(self.url + '/xmlrpc/2/common', allow_none=True)
        self.version = common.version()
        self.db = os.environ.get('ODOO_DB') or self._only_db()
        self.uid = common.authenticate(self.db, self.login, self.key, {})
        if not self.uid:
            sys.exit('فشل تسجيل الدخول (%s على %s / %s)' % (self.login, self.url, self.db))
        self.obj = xmlrpc.client.ServerProxy(self.url + '/xmlrpc/2/object', allow_none=True)
        self._fields = {}

    def _only_db(self):
        try:
            dbs = xmlrpc.client.ServerProxy(self.url + '/xmlrpc/2/db').list()
        except (xmlrpc.client.Fault, xmlrpc.client.ProtocolError):
            dbs = []
        if len(dbs) != 1:
            sys.exit('حدد ODOO_DB (قواعد البيانات اللي بانت: %s)' % (dbs or 'القائمة مسدودة'))
        return dbs[0]

    def call(self, model, method, *args, **kw):
        return self.obj.execute_kw(self.db, self.uid, self.key, model, method, list(args), kw)

    def fields(self, model):
        if model not in self._fields:
            self._fields[model] = self.call(model, 'fields_get', attributes=['type', 'relation', 'string'])
        return self._fields[model]

    def has_model(self, model):
        return bool(self.call('ir.model', 'search_count', [('model', '=', model)]))

    def missing_groups(self, xmlids=SUDO_GROUPS):
        mine = set(self.call('res.users', 'read', [self.uid], ['groups_id'])[0]['groups_id'])
        missing = []
        for xmlid in xmlids:
            module, name = xmlid.split('.')
            rows = self.call('ir.model.data', 'search_read', [
                ('module', '=', module), ('name', '=', name), ('model', '=', 'res.groups')], fields=['res_id'])
            if rows and rows[0]['res_id'] not in mine:
                missing.append(xmlid)
        return missing


# ------------------------------------------- ORM مصغّر حتى يشتغل نص القاعدة نفسه

def to_rpc(value):
    if isinstance(value, dt.datetime):
        return value.strftime('%Y-%m-%d %H:%M:%S')
    if isinstance(value, dt.date):
        return value.strftime('%Y-%m-%d')
    if isinstance(value, Records):
        return value.ids
    if isinstance(value, (list, tuple)):
        return [to_rpc(v) for v in value]
    return value


def from_rpc(env, field, value):
    kind = field['type']
    if kind == 'many2one':
        return Records(env, field['relation'], [value[0]] if value else [])
    if kind in ('one2many', 'many2many'):
        return Records(env, field['relation'], value or [])
    if value is False or value is None:
        return False
    if kind == 'datetime':
        return dt.datetime.strptime(value[:19], '%Y-%m-%d %H:%M:%S')
    if kind == 'date':
        return dt.datetime.strptime(value[:10], '%Y-%m-%d').date()
    return value


class Env:
    def __init__(self, odoo):
        self.odoo = odoo
        self.cache = defaultdict(dict)   # (model, id) ← {field: قيمة RPC}

    def __getitem__(self, model):
        return Model(self, model, {})


class Model:
    def __init__(self, env, name, context):
        self.env, self.name, self.context = env, name, context

    def sudo(self):
        return self

    def with_context(self, **kw):
        return Model(self.env, self.name, dict(self.context, **kw))

    def search(self, domain):
        ids = self.env.odoo.call(self.name, 'search', to_rpc(domain), context=self.context)
        return Records(self.env, self.name, ids)


class Records:
    def __init__(self, env, model, ids, prefetch=None):
        self._env, self._model, self._ids = env, model, tuple(ids)
        self._prefetch = prefetch if prefetch is not None else self._ids

    env = property(lambda self: self._env)
    ids = property(lambda self: list(self._ids))
    id = property(lambda self: self._ids[0] if len(self._ids) == 1 else False)

    def __len__(self):
        return len(self._ids)

    def __bool__(self):
        return bool(self._ids)

    def __iter__(self):
        for rid in self._ids:
            yield Records(self._env, self._model, [rid], self._prefetch)

    def __eq__(self, other):
        return isinstance(other, Records) and (self._model, self._ids) == (other._model, other._ids)

    def __hash__(self):
        return hash((self._model, self._ids))

    def __repr__(self):
        return '%s%r' % (self._model, self._ids)

    def __or__(self, other):
        return Records(self._env, self._model, list(dict.fromkeys(self._ids + other._ids)))

    def mapped(self, name):
        field = self._field(name)
        values = [getattr(rec, name) for rec in self]
        if field['type'] in RELATIONAL:
            ids = [i for v in values for i in v.ids]
            return Records(self._env, field['relation'], list(dict.fromkeys(ids)))
        return values

    def _field(self, name):
        fields = self._env.odoo.fields(self._model)
        if name not in fields:
            raise AttributeError('%s ما بيه حقل %s' % (self._model, name))
        return fields[name]

    def __getattr__(self, name):
        if name.startswith('_'):
            raise AttributeError(name)
        field = self._field(name)
        if len(self._ids) > 1:
            raise ValueError('Expected singleton: %r' % self)
        if not self._ids:
            return Records(self._env, field['relation'], []) if field['type'] in RELATIONAL else False
        rid = self._ids[0]
        cache = self._env.cache
        if name not in cache[(self._model, rid)]:
            todo = [i for i in self._prefetch if name not in cache[(self._model, i)]]
            if rid not in todo:
                todo.append(rid)
            for start in range(0, len(todo), 500):
                for row in self._env.odoo.call(self._model, 'read', todo[start:start + 500], [name]):
                    cache[(self._model, row['id'])][name] = row[name]
        return from_rpc(self._env, field, cache[(self._model, rid)][name])


def localdict(env, slip_id, rule):
    """نفس قاموس أودو 17+ للقاعدة: payslip و contract و inputs و categories (اللي قبل القاعدة)."""
    slip = Records(env, 'hr.payslip', [slip_id])
    categories = defaultdict(float)
    for line in slip.line_ids:
        if (line.sequence, line.salary_rule_id.id) >= (rule['sequence'], rule['id']):
            continue
        category = line.category_id
        while category:
            categories[category.code] += line.total
            category = category.parent_id
    return {
        'payslip': slip,
        'contract': slip.contract_id,
        'employee': slip.employee_id,
        'inputs': {line.code: line for line in slip.input_line_ids if line.code},
        'worked_days': {line.code: line for line in slip.worked_days_line_ids if line.code},
        'categories': categories,
        'result': None, 'result_qty': 1.0, 'result_rate': 100, 'result_name': False,
    }


def run_rule(code, ld):
    ld['__builtins__'] = SAFE_BUILTINS
    exec(compile(code, 'ABSENCE', 'exec'), ld)
    return float(ld['result'] or 0.0), ld['result_name'] or ''


def condition_ok(rule, ld):
    """مثل _satisfy_condition بأودو. يرجّع None إذا نوع الشرط ما نعرف نقيّمه هنا."""
    kind = rule.get('condition_select') or 'none'
    scope = dict(ld, result=None, __builtins__=SAFE_BUILTINS)
    if kind == 'none':
        return True
    if kind == 'python':
        exec(compile(rule.get('condition_python') or 'result = False', 'condition', 'exec'), scope)
        return bool(scope.get('result'))
    if kind == 'range':
        value = eval(compile(rule.get('condition_range') or '0', 'range', 'eval'), scope)
        return rule['condition_range_min'] <= value <= rule['condition_range_max']
    return None


def run_like_odoo(code, rule, ld, category_codes):
    """إذا الإيصال بيه أكثر من سطر إدخال بنفس كود القاعدة، أودو يشغّلها مرة لكل سطر
    (same_type_input_lines) ويضيف كل نتيجة للفئات قبل المرة الجاية، ويجمعها."""
    lines = [line for line in ld['payslip'].input_line_ids if line.code == rule['code']]
    if len(lines) < 2:
        return run_rule(code, ld)
    total, first_name = 0.0, ''
    for line in lines:
        ld['inputs'][rule['code']] = line
        amount, name = run_rule(code, ld)
        total += amount
        for c in category_codes:
            ld['categories'][c] += amount
        first_name = first_name or name
    return total, first_name


# ---------------------------------------------------------------- الأوامر

def month_range(month):
    year, mon = (int(x) for x in month.split('-'))
    first = dt.date(year, mon, 1)
    last = (first + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)
    return first, last


def read_rules(odoo):
    available = odoo.fields('hr.salary.rule')
    wanted = ['code', 'name', 'struct_id', 'sequence', 'category_id', 'amount_select', 'amount_python_compute']
    wanted += [f for f in COND_FIELDS if f in available]
    rows = odoo.call('hr.salary.rule', 'read', list(RULES), wanted)
    rules = {r['id']: r for r in rows}
    for rid in RULES:
        if rid not in rules:
            sys.exit('القاعدة %s مو موجودة' % rid)
        if rules[rid]['code'] != RULE_CODE:
            sys.exit('القاعدة %s كودها %r مو %r، وكفت' % (rid, rules[rid]['code'], RULE_CODE))
    return rules


def cmd_check(odoo, month):
    first, last = month_range(month)
    print('Odoo %s | DB %s | uid %s' % (odoo.version.get('server_version'), odoo.db, odoo.uid))
    missing = odoo.missing_groups()
    if missing:
        print('✗ المستخدم ناقصته صلاحيات %s: أرقام الحضور والإجازات تحت ممكن تطلع ناقصة' % missing)

    print('\n== 1) قواعد ABSENCE')
    for rid, rule in read_rules(odoo).items():
        code = rule['amount_python_compute'] or ''
        print('  %s %s | %s | sequence %s | %s | %d سطر' % (
            rid, rule['code'], rule['struct_id'] and rule['struct_id'][1], rule['sequence'],
            rule['amount_select'], len(code.splitlines())))
        if 'absence_deduction_amount' in code:
            print('     يستخدم الحقل القديم absence_deduction_amount')
        kind = rule.get('condition_select') or 'none'
        print('     الشرط: %s %s' % (kind, {'python': rule.get('condition_python'),
                                           'range': rule.get('condition_range')}.get(kind, '') or ''))
    as_dict = odoo.call('hr.salary.rule', 'search_count', [('amount_python_compute', 'ilike', "inputs['")])
    as_attr = odoo.call('hr.salary.rule', 'search_count', [('amount_python_compute', 'ilike', 'inputs.')])
    print('  قواعد تكتب inputs[...] = %d، و inputs.X = %d (القاعدة الجديدة تحتاج inputs[...])' % (as_dict, as_attr))

    print('\n== 2) الموديلات والحقول')
    need = {
        'hr.payslip': ['absence_deduction_amount', 'contract_id', 'input_line_ids', 'line_ids'],
        'hr.contract': ['policy_group_id', 'wage', 'resource_calendar_id', 'company_id', 'date_start', 'date_end'],
        'hr.work.entry': ['work_entry_type_id', 'duration', 'state', 'date_start', 'date_stop'],
        'hr.attendance': ['check_in', 'worked_hours'],
        'hr.attendance.zk.temp': ['employee_id', 'date'],
        'hr.leave': ['request_unit_hours', 'request_unit_half', 'request_date_from', 'request_date_to'],
        'resource.calendar.leaves': ['resource_id', 'time_type', 'calendar_id', 'company_id'],
    }
    for model, names in need.items():
        if not odoo.has_model(model):
            print('  ✗ %s مو موجود' % model)
            continue
        fields = odoo.fields(model)
        missing = [n for n in names if n not in fields]
        print('  %s %s%s' % ('✗' if missing else '✓', model, ' — ناقص: %s' % missing if missing else ''))
    zk = odoo.fields('hr.attendance.zk.temp') if odoo.has_model('hr.attendance.zk.temp') else {}
    for name in ('employee_id', 'date'):
        if name in zk:
            print('     zk.temp.%s: %s %s' % (name, zk[name]['type'], zk[name].get('relation') or ''))
    pg = odoo.fields('hr.contract').get('policy_group_id')
    if pg:
        pg_fields = odoo.fields(pg['relation'])
        print('  %s: absence_policy_id %s، work_days_per_month %s' % (
            pg['relation'], 'absence_policy_id' in pg_fields, 'work_days_per_month' in pg_fields))

    print('\n== 3) نوع سجل العمل Absence (الاسم مترجم)')
    ids = odoo.call('hr.work.entry.type', 'search', [('name', '=', 'Absence')], context={'lang': 'en_US'})
    print('  بالإنكليزي: %s' % (ids or '✗ ما لكيته'))
    for lang in ('ar_001', 'ar_SY', 'ar_IQ'):
        hits = odoo.call('hr.work.entry.type', 'search', [('name', '=', 'Absence')], context={'lang': lang})
        names = odoo.call('hr.work.entry.type', 'read', ids, ['name', 'code'], context={'lang': lang}) if ids else []
        print('  lang=%s: البحث القديم يلكى %s | الاسم: %s' % (
            lang, hits or 'ولا شي ✗', ['%s (%s)' % (n['name'], n['code']) for n in names]))
    langs = odoo.call('res.users', 'search_read', [('share', '=', False)], fields=['lang'])
    print('  لغات المستخدمين: %s' % dict(sorted(_count(u['lang'] or 'بدون' for u in langs).items())))

    if zk:
        print('\n== 4) البصمات الخام hr.attendance.zk.temp')
        oldest = odoo.call('hr.attendance.zk.temp', 'search_read', [], fields=['date'], order='date asc', limit=1)
        newest = odoo.call('hr.attendance.zk.temp', 'search_read', [], fields=['date'], order='date desc', limit=1)
        print('  أقدم بصمة %s | أحدث بصمة %s' % (oldest and oldest[0]['date'], newest and newest[0]['date']))
        lo, hi = str(first - dt.timedelta(days=2)), str(last + dt.timedelta(days=2))
        punches = odoo.call('hr.attendance.zk.temp', 'search_read',
                            [('date', '>=', lo), ('date', '<=', hi)], fields=['employee_id', 'date'])
        keys = [(p['employee_id'] and p['employee_id'][0], p['date']) for p in punches]
        print('  بصمات %s: %d | مكررة بنفس الثانية: %d' % (month, len(keys), len(keys) - len(set(keys))))
        attendance = odoo.call('hr.attendance', 'search_read',
                               [('check_in', '>=', lo), ('check_in', '<=', hi)], fields=['employee_id', 'check_in'])
        punch_set = set(keys)
        shifts = {'نفس الوقت (UTC ✓)': 0, '+3 ساعات': 0, '-3 ساعات': 0}
        for a in attendance:
            emp = a['employee_id'] and a['employee_id'][0]
            t = dt.datetime.strptime(a['check_in'][:19], '%Y-%m-%d %H:%M:%S')
            for label, delta in (('نفس الوقت (UTC ✓)', 0), ('+3 ساعات', 3), ('-3 ساعات', -3)):
                if (emp, (t + dt.timedelta(hours=delta)).strftime('%Y-%m-%d %H:%M:%S')) in punch_set:
                    shifts[label] += 1
        print('  دخول الحضور (%d) يطابق بصمة: %s' % (len(attendance), shifts))

    print('\n== 5) العطل الرسمية (Global Time Off) من %s' % (first - dt.timedelta(days=92)))
    holidays = odoo.call('resource.calendar.leaves', 'search_read', [
        ('resource_id', '=', False),
        ('date_to', '>=', str(first - dt.timedelta(days=92))),
        ('date_from', '<=', str(last + dt.timedelta(days=62)))],
        fields=['name', 'date_from', 'date_to', 'calendar_id', 'time_type'], order='date_from')
    for h in holidays:
        print('  %s → %s | %s | %s | %s' % (h['date_from'], h['date_to'], h['name'],
                                           h['calendar_id'] and h['calendar_id'][1] or 'كل التقويمات', h['time_type']))
    if not holidays:
        print('  ما أكو')

    print('\n== 6) قواعد تعتمد على الحقل القديم أو FRIDAY')
    for r in odoo.call('hr.salary.rule', 'search_read', [
            '|', '|', ('amount_python_compute', 'ilike', 'absence_deduction_amount'),
            ('condition_python', 'ilike', 'absence_deduction_amount'), ('code', 'ilike', 'FRIDAY')],
            fields=['code', 'name', 'struct_id', 'amount_python_compute', 'condition_python']):
        uses = 'absence_deduction_amount' in (r['amount_python_compute'] or '') + (r['condition_python'] or '')
        print('  %s %s | %s | %s' % (r['id'], r['code'], r['struct_id'] and r['struct_id'][1],
                                     'يستخدم الحقل القديم ✗' if uses else 'ما يستخدمه'))
        if r['code'] and 'FRIDAY' in r['code'].upper():
            for line in (r['amount_python_compute'] or '').splitlines()[:60]:
                print('      | ' + line)

    print('\n== 7) أنواع الإجازات الساعية (الزمنية)')
    leave_fields = odoo.fields('hr.leave.type')
    extra = [f for f in ('requires_allocation',) if f in leave_fields]
    for t in odoo.call('hr.leave.type', 'search_read', [('request_unit', '=', 'hour')],
                       fields=['name', 'request_unit'] + extra):
        count = odoo.call('hr.leave', 'search_count', [
            ('holiday_status_id', '=', t['id']), ('state', '=', 'validate'),
            ('request_date_from', '>=', str(first)), ('request_date_from', '<=', str(last))])
        print('  %s %s | %s | %d طلب معتمد بـ %s' % (t['id'], t['name'], t.get('requires_allocation', ''), count, month))


def _count(values):
    out = defaultdict(int)
    for v in values:
        out[v] += 1
    return out


def cmd_compare(odoo, month):
    first, last = month_range(month)
    missing = odoo.missing_groups()
    if missing:
        sys.exit('المستخدم ناقصته صلاحيات %s. القاعدة بأودو تبحث بـ sudo()، وهنا الحضور والإجازات راح '
                 'تختفي بدون خطأ والأرقام تطلع غلط. استخدم مفتاح API لمستخدم مدير.' % missing)
    rules = read_rules(odoo)
    by_struct = {r['struct_id'][0]: r for r in rules.values() if r['struct_id']}
    codes = {rid: rule_source(structure) for rid, structure in RULES.items()}
    has_field = 'absence_deduction_amount' in odoo.fields('hr.payslip')
    slips = odoo.call('hr.payslip', 'search_read', [
        ('date_from', '>=', str(first)), ('date_from', '<=', str(last)),
        ('struct_id', 'in', list(by_struct)), ('state', '!=', 'cancel')],
        fields=['number', 'employee_id', 'struct_id', 'state', 'line_ids'] +
               (['absence_deduction_amount'] if has_field else []),
        order='struct_id, id')
    print('%d إيصال بـ %s' % (len(slips), month))
    env = Env(odoo)
    chains = {}
    for rule in rules.values():   # فئة القاعدة وآباؤها: تنضاف لها النتيجة بين مرات الإدخال المتكررة
        codes_chain, category = [], Records(env, 'hr.salary.rule.category', [rule['category_id'][0]])
        while category:
            codes_chain.append(category.code)
            category = category.parent_id
        chains[rule['id']] = codes_chain
    rows, changed, errors, uncomputed, blocked = [], 0, 0, 0, 0
    for n, slip in enumerate(slips, 1):
        rule = by_struct[slip['struct_id'][0]]
        row = {'payslip': slip['number'] or slip['id'], 'employee': slip['employee_id'][1],
               'structure': RULES[rule['id']], 'state': slip['state'],
               'old_line': '', 'old_field': slip.get('absence_deduction_amount', ''),
               'new': '', 'diff': '', 'condition_now': '', 'new_days': ''}
        rows.append(row)
        if not slip['line_ids']:
            uncomputed += 1
            row['new_days'] = 'الإيصال مو محسوب (ما بيه أسطر)'
            continue
        lines = odoo.call('hr.payslip.line', 'search_read',
                          [('slip_id', '=', slip['id']), ('salary_rule_id', '=', rule['id'])], fields=['total'])
        old = sum(line['total'] for line in lines)
        row['old_line'] = round(old, 2)
        try:
            ld = localdict(env, slip['id'], rule)
            ok = condition_ok(rule, ld)
            row['condition_now'] = {True: 'يتحقق', False: 'ما يتحقق', None: '؟'}[ok]
            if ok is False:
                blocked += 1
            new, name = run_like_odoo(codes[rule['id']], rule, ld, chains[rule['id']])
        except Exception as e:   # noqa: BLE001 — نكمل باقي الإيصالات ونسجّل الخطأ
            new, name, errors = None, 'خطأ: %s: %s' % (type(e).__name__, e), errors + 1
        diff = None if new is None else round(new - old, 2)
        if diff:
            changed += 1
        row.update({'new': '' if new is None else round(new, 2), 'diff': '' if diff is None else diff,
                    'new_days': name})
        if diff or new is None:
            print('  %-14s %-35s قديم %12.2f ← جديد %12s | %s' % (
                row['payslip'], row['employee'][:35], old, row['new'], name))
        if n % 50 == 0:
            print('  … %d/%d' % (n, len(slips)))
    os.makedirs(REPORT_DIR, exist_ok=True)
    path = os.path.join(REPORT_DIR, 'absence_compare_%s.csv' % month)
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0]) if rows else ['payslip'])
        writer.writeheader()
        writer.writerows(rows)
    print('\nتغيّر %d من %d | أخطاء %d | مو محسوبة %d | التفاصيل: %s' % (
        changed, len(slips), errors, uncomputed, path))
    if blocked:
        print('✗ شرط القاعدة الحالي ما يتحقق بـ %d إيصال: بهذني القاعدة ما تشتغل أصلاً إلا إذا '
              'انشال الشرط (apply --clear-condition). عمود condition_now بالملف.' % blocked)


def save_backup(rule, stamp):
    """كل اللي يحتاجه الرجوع: نوع المبلغ والكود والشرط، ويه رقم القاعدة حتى restore يتأكد منه."""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    path = os.path.join(BACKUP_DIR, 'rule_%s_%s.json' % (rule['id'], stamp))
    values = {f: rule[f] for f in ('amount_select', 'amount_python_compute') + COND_FIELDS if f in rule}
    with open(path, 'w', encoding='utf-8') as f:
        json.dump({'rule_id': rule['id'], 'code': rule['code'], 'saved_at': stamp,
                   'structure': rule['struct_id'] and rule['struct_id'][1], 'values': values},
                  f, ensure_ascii=False, indent=1)
    return path


def write_and_verify(odoo, rid, values):
    odoo.call('hr.salary.rule', 'write', [rid], values)
    back = odoo.call('hr.salary.rule', 'read', [rid], list(values))[0]
    return all(back[k] == v for k, v in values.items())


def cmd_apply(odoo, yes, clear_condition):
    rules = read_rules(odoo)
    stamp = dt.datetime.now().strftime('%Y%m%d_%H%M%S')
    print('السيرفر: %s | قاعدة البيانات: %s' % (odoo.url, odoo.db))
    for rid, structure in RULES.items():
        rule, new = rules[rid], rule_source(structure)
        values = {'amount_select': 'code', 'amount_python_compute': new}
        kind = rule.get('condition_select') or 'none'
        if kind != 'none':
            print('%s: بيها شرط %s: %s' % (rid, kind, rule.get('condition_python') or rule.get('condition_range')))
            if not clear_condition:
                sys.exit('   القاعدة الجديدة تحسب كلشي بروحها وترجّع صفر إذا ماكو غياب، والشرط القديم ممكن '
                         'يمنعها تشتغل. شغّل apply --yes --clear-condition حتى يصير Always True.')
            values['condition_select'] = 'none'
        if all(rule.get(k) == v for k, v in values.items()):
            print('%s (%s): نفس الكود الجديد، ما يحتاج' % (rid, structure))
            continue
        print('%s %s (%s) | %s: %d سطر ← %d سطر' % (
            rid, rule['code'], structure, rule['struct_id'] and rule['struct_id'][1],
            len((rule['amount_python_compute'] or '').splitlines()), len(new.splitlines())))
        if not yes:
            continue
        backup = save_backup(rule, stamp)
        print('   نسخة احتياطية: %s' % backup)
        if not write_and_verify(odoo, rid, values):
            sys.exit('   ✗ اللي انكتب ما يطابق، رجّعه بـ restore %s %s --yes' % (backup, rid))
        print('   ✓ انكتب')
    if not yes:
        print('\nما انكتب شي. حتى تطبّق: apply --yes')
    else:
        print('\nالتغيير يأثر على الإيصالات اللي تنحسب بعد هسه بس. المؤكدة تبقى مثل ما هي، والمسودات '
              'المحسوبة قبل تبقى بالأرقام القديمة لحد ما ينضغط Compute Sheet عليها.')


def cmd_restore(odoo, path, rid, yes):
    with open(path, encoding='utf-8') as f:
        backup = json.load(f)
    if rid not in RULES or backup.get('rule_id') != rid:
        sys.exit('النسخة الاحتياطية للقاعدة %s، مو %s (المسموح %s)' % (backup.get('rule_id'), rid, list(RULES)))
    rule = read_rules(odoo)[rid]
    values = {k: v for k, v in backup['values'].items() if k in odoo.fields('hr.salary.rule')}
    print('يرجّع %s %s من %s (نسخة %s، %d سطر كود)' % (
        rid, rule['code'], path, backup.get('saved_at'), len((values.get('amount_python_compute') or '').splitlines())))
    if not yes:
        print('ما انكتب شي. أضف --yes')
        return
    current = save_backup(rule, dt.datetime.now().strftime('%Y%m%d_%H%M%S'))
    print('   نسخة من الحالي قبل الرجوع: %s' % current)
    if not write_and_verify(odoo, rid, values):
        sys.exit('   ✗ اللي انكتب ما يطابق النسخة الاحتياطية')
    print('✓ رجع')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='cmd', required=True)
    p = sub.add_parser('check')
    p.add_argument('--month', default='2026-09')
    p = sub.add_parser('compare')
    p.add_argument('--month', required=True)
    p = sub.add_parser('apply')
    p.add_argument('--yes', action='store_true')
    p.add_argument('--clear-condition', action='store_true')
    p = sub.add_parser('restore')
    p.add_argument('file')
    p.add_argument('rule_id', type=int)
    p.add_argument('--yes', action='store_true')
    args = parser.parse_args()

    odoo = Odoo()
    if args.cmd == 'check':
        cmd_check(odoo, args.month)
    elif args.cmd == 'compare':
        cmd_compare(odoo, args.month)
    elif args.cmd == 'apply':
        cmd_apply(odoo, args.yes, args.clear_condition)
    else:
        cmd_restore(odoo, args.file, args.rule_id, args.yes)


if __name__ == '__main__':
    main()
