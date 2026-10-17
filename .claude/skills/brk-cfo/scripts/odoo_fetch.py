#!/usr/bin/env python3
r"""قراءة البيانات الحية من Odoo (قراءة فقط) للمدير المالي.

الأوامر:
  snapshot        صورة مالية لحظية: النقد، الذمم المدينة والدائنة وأعمارها، الفواتير غير المسددة،
                  أوامر الشراء المفتوحة، طلبات الموافقة المعلقة، الإيرادات والمصاريف للشهر والسنة.
  costs           قيود المصاريف المرحّلة لفترة → CSV جاهز لـ build_cost_workbook.py.
  trial-balance   ميزان المراجعة لفترة (لمطابقة تقرير التكاليف مع الدفاتر).
  export          تصدير كامل صامت لكل ما يحتاجه المدير المالي إلى مجلد مؤرخ، مع فهرس مجاميع ضابطة وسجل.
                  يصلح للجدولة بـ pythonw (بدون نافذة). رمز الخروج 0 نجاح، 1 ملفات ناقصة، 2 خطأ اتصال.

الاتصال من متغيرات البيئة (لا تكتبها بالكود ولا تلصقها بالمحادثة):
  ODOO_URL      اختياري، والافتراضي رابط البركة https://brk.ejaferp.com
  ODOO_DB       اسم قاعدة البيانات
  ODOO_USER     اسم الدخول (login)
  ODOO_API_KEY  مفتاح API لمستخدم بصلاحيات قراءة المحاسبة فقط (Preferences → Account Security → New API Key)

أمثلة:
  python odoo_fetch.py snapshot
  python odoo_fetch.py costs --from 2026-07-01 --to 2026-09-30 --out costs_q3.csv
  python odoo_fetch.py costs --from 2026-07-01 --to 2026-09-30 --account-prefix 3 --code-digits 4 --out costs.csv
  python odoo_fetch.py trial-balance --from 2026-01-01 --to 2026-09-30 --out tb.csv
  pythonw odoo_fetch.py export --out-root "C:\Users\surface\Documents\الماليه\CFO\بيانات"

مكتوب لـ Odoo 16 فما فوق (account_type وanalytic_distribution) عبر واجهة JSON-RPC.
السكربت لا يستدعي أي دالة كتابة: المسموح فقط search_read وsearch_count وread وfields_get وread_group.
"""
import argparse
import csv
import datetime as dt
import json
import os
import sys
import urllib.request
from collections import defaultdict

DEFAULT_URL = 'https://brk.ejaferp.com'
READ_ONLY_METHODS = {'search_read', 'search_count', 'read', 'fields_get', 'read_group'}
PAGE = 2000
EXPENSE_TYPES = ['expense', 'expense_direct_cost', 'expense_depreciation']
PAYMENT_STATES = {'paid': 'مدفوع', 'in_payment': 'مدفوع', 'not_paid': 'غير مدفوع', 'partial': 'جزئي',
                  'reversed': 'معكوس'}


class Odoo:
    def __init__(self):
        missing = [k for k in ('ODOO_DB', 'ODOO_USER', 'ODOO_API_KEY') if not os.environ.get(k)]
        if missing:
            raise SystemExit('متغيرات البيئة غير معرّفة: %s، وتُضاف بإعدادات البيئة ثم تُبدأ جلسة جديدة.' % ', '.join(missing))
        self.url = os.environ.get('ODOO_URL', DEFAULT_URL).rstrip('/').removesuffix('/odoo') + '/jsonrpc'
        self.db, self.key = os.environ['ODOO_DB'], os.environ['ODOO_API_KEY']
        self.uid = self._call('common', 'login', [self.db, os.environ['ODOO_USER'], self.key])
        if not self.uid:
            raise SystemExit('فشل تسجيل الدخول إلى Odoo: تحقق من ODOO_DB وODOO_USER وODOO_API_KEY.')

    def _call(self, service, method, args):
        payload = json.dumps({'jsonrpc': '2.0', 'method': 'call', 'id': 1,
                              'params': {'service': service, 'method': method, 'args': args}}).encode()
        req = urllib.request.Request(self.url, payload, {'Content-Type': 'application/json'})
        try:
            with urllib.request.urlopen(req, timeout=120) as resp:
                data = json.load(resp)
        except OSError as e:
            raise SystemExit(f'تعذر الوصول إلى {self.url}: {e} — تحقق من الرابط ومن السماح بالمضيف بإعدادات الشبكة.')
        if data.get('error'):
            err = data['error'].get('data', {}).get('message') or data['error'].get('message')
            raise RuntimeError(err)
        return data['result']

    def call(self, model, method, args, **kwargs):
        if method not in READ_ONLY_METHODS:
            raise PermissionError(f'الدالة {method} غير مسموحة — هذا السكربت للقراءة فقط')
        return self._call('object', 'execute_kw', [self.db, self.uid, self.key, model, method, args, kwargs])

    def search_read(self, model, domain, fields, order=None):
        rows, offset = [], 0
        while True:
            page = self.call(model, 'search_read', [domain], fields=fields, offset=offset, limit=PAGE,
                             **({'order': order} if order else {}))
            rows += page
            if len(page) < PAGE:
                return rows
            offset += PAGE

    def try_count(self, model, domain):
        """للموديلات الاختيارية (مثل approvals): يرجع None إذا الموديل غير مثبت."""
        try:
            return self.call(model, 'search_count', [domain])
        except RuntimeError:
            return None


def name_of(m2o):
    return m2o[1] if isinstance(m2o, list) and len(m2o) > 1 else ''


def fmt(n):
    return f'({-n:,.0f})' if n < 0 else f'{n:,.0f}'


def company_currency(odoo):
    company = odoo.search_read('res.company', [], ['name', 'currency_id'])[0]
    return company['name'], name_of(company['currency_id'])


def aging(lines, today):
    buckets = {'غير مستحق': 0.0, '1–30': 0.0, '31–90': 0.0, '91–180': 0.0, 'أكثر من 180': 0.0}
    for ln in lines:
        due = dt.date.fromisoformat(ln['date_maturity'] or ln['date'])
        days = (today - due).days
        key = ('غير مستحق' if days <= 0 else '1–30' if days <= 30 else '31–90' if days <= 90
               else '91–180' if days <= 180 else 'أكثر من 180')
        buckets[key] += ln['amount_residual']
    return buckets


def cmd_snapshot(odoo, args):
    today = dt.date.today()
    company, cur = company_currency(odoo)
    month_start, year_start = today.replace(day=1).isoformat(), today.replace(month=1, day=1).isoformat()
    posted = [('parent_state', '=', 'posted')]
    print(f'# الصورة المالية: {company}\n\nبتاريخ {today} | العملة {cur} | المصدر: Odoo مباشرة (قيود مرحّلة فقط)\n')

    cash = odoo.search_read('account.move.line', posted + [('account_id.account_type', '=', 'asset_cash')],
                            ['account_id', 'balance'])
    by_acc = defaultdict(float)
    for ln in cash:
        by_acc[name_of(ln['account_id'])] += ln['balance']
    print('## النقد والمصارف\n\n| الحساب | الرصيد |\n|---|---:|')
    for acc, bal in sorted(by_acc.items()):
        print(f'| {acc} | {fmt(bal)} |')
    print(f'| **الإجمالي** | **{fmt(sum(by_acc.values()))}** |\n')

    for title, acc_type, sign in (('الذمم المدينة (الزبائن)', 'asset_receivable', 1),
                                  ('الذمم الدائنة (الموردون)', 'liability_payable', -1)):
        lines = odoo.search_read('account.move.line', posted + [('account_id.account_type', '=', acc_type),
                                                                ('reconciled', '=', False), ('amount_residual', '!=', 0)],
                                 ['partner_id', 'date', 'date_maturity', 'amount_residual'])
        for ln in lines:  # الدائنة رصيدها سالب بـ Odoo؛ نعرضها موجبة
            ln['amount_residual'] *= sign
        b = aging(lines, today)
        print(f'## {title}\n\n| العمر | المبلغ |\n|---|---:|')
        for k, v in b.items():
            print(f'| {k} | {fmt(v)} |')
        print(f'| **الإجمالي** | **{fmt(sum(b.values()))}** |\n')
        by_partner = defaultdict(float)
        for ln in lines:
            by_partner[name_of(ln['partner_id']) or '—'] += ln['amount_residual']
        top = sorted(by_partner.items(), key=lambda x: -abs(x[1]))[:5]
        print('أكبر 5: ' + '، '.join(f'{p} {fmt(v)}' for p, v in top) + '\n')

    bills = odoo.search_read('account.move', [('move_type', '=', 'in_invoice'), ('state', '=', 'posted'),
                                              ('payment_state', 'in', ['not_paid', 'partial'])],
                             ['amount_residual_signed', 'invoice_date_due'])
    overdue = [b for b in bills if b['invoice_date_due'] and b['invoice_date_due'] < today.isoformat()]
    print('## فواتير الموردين غير المسددة\n')
    print(f'- العدد {len(bills)} بمبلغ {fmt(-sum(b["amount_residual_signed"] for b in bills))}، '
          f'منها متأخرة {len(overdue)} بمبلغ {fmt(-sum(b["amount_residual_signed"] for b in overdue))}\n')

    pos = odoo.search_read('purchase.order', [('state', 'in', ['draft', 'sent', 'to approve', 'purchase'])],
                           ['state', 'invoice_status', 'amount_total', 'currency_id'])
    if pos:
        groups = defaultdict(lambda: [0, 0.0])
        for po in pos:
            label = {'draft': 'طلب عرض سعر (مسودة)', 'sent': 'طلب عرض سعر (مرسل)', 'to approve': 'بانتظار الموافقة',
                     'purchase': 'أمر شراء مؤكد لم يُفوتر بالكامل'}[po['state']]
            if po['state'] == 'purchase' and po['invoice_status'] == 'invoiced':
                continue
            key = (label, name_of(po['currency_id']))
            groups[key][0] += 1
            groups[key][1] += po['amount_total']
        print('## المشتريات المفتوحة\n\n| الحالة | العملة | العدد | المبلغ |\n|---|---|---:|---:|')
        for (label, c), (n, amt) in sorted(groups.items()):
            print(f'| {label} | {c} | {n} | {fmt(amt)} |')
        print()

    pending = odoo.try_count('approval.request', [('request_status', '=', 'pending')])
    if pending is not None:
        print(f'## طلبات الموافقة المعلقة\n\n- {pending} طلب بانتظار الموافقة\n')

    print('## الأداء\n\n| البند | الشهر الحالي | منذ بداية السنة |\n|---|---:|---:|')
    for label, types, sign in (('الإيرادات', ['income', 'income_other'], -1), ('المصاريف', EXPENSE_TYPES, 1)):
        vals = []
        for start in (month_start, year_start):
            lines = odoo.search_read('account.move.line', posted + [('account_id.account_type', 'in', types),
                                                                    ('date', '>=', start), ('date', '<=', today.isoformat())],
                                     ['balance'])
            vals.append(sign * sum(ln['balance'] for ln in lines))
        print(f'| {label} | {fmt(vals[0])} | {fmt(vals[1])} |')
    print('\nملاحظة: الأرقام من القيود المرحّلة فقط؛ المسودات والقيود غير المرحّلة غير مشمولة.')


def cmd_costs(odoo, args):
    _, cur = company_currency(odoo)
    domain = [('parent_state', '=', 'posted'), ('date', '>=', args.date_from), ('date', '<=', args.date_to),
              ('display_type', 'not in', ['line_section', 'line_note'])]
    if args.account_prefix:
        domain.append(('account_id.code', '=like', args.account_prefix + '%'))
    else:
        types = EXPENSE_TYPES + (['asset_fixed'] if args.include_capex else [])
        domain.append(('account_id.account_type', 'in', types))
    lines = odoo.search_read('account.move.line', domain,
                             ['date', 'move_id', 'move_name', 'ref', 'account_id', 'partner_id', 'name',
                              'balance', 'analytic_distribution'], order='date, id')
    move_ids = sorted({ln['move_id'][0] for ln in lines})
    moves = {m['id']: m for m in (odoo.call('account.move', 'read', [move_ids], fields=['ref', 'payment_state', 'move_type'])
                                  if move_ids else [])}
    analytic_ids = {int(i) for ln in lines for key in (ln['analytic_distribution'] or {}) for i in key.split(',')}
    analytic = {a['id']: a['name'] for a in (odoo.call('account.analytic.account', 'read', [sorted(analytic_ids)], fields=['name'])
                                             if analytic_ids else [])}

    out_rows = []
    for ln in lines:
        move = moves.get(ln['move_id'][0], {})
        code = name_of(ln['account_id']).split(' ')[0]
        if args.code_digits and code.isdigit():
            code = code[:args.code_digits]
        # رقم فاتورة المورد أفضل لكشف التكرار من الرقم الداخلي، لأن الفاتورة المكررة تأخذ رقماً داخلياً جديداً
        ref = (move.get('ref') or ln['ref'] or ln['move_name'] or '').strip()
        status = PAYMENT_STATES.get(move.get('payment_state'), '') if move.get('move_type', '').startswith('in_') else ''
        dist = ln['analytic_distribution'] or {'': 100}
        for key, pct in dist.items():
            center = ' / '.join(analytic.get(int(i), i) for i in key.split(',') if i) if key else ''
            out_rows.append({
                'date': ln['date'], 'ref': ref, 'cost_center': center, 'cost_code': code,
                'vendor': name_of(ln['partner_id']), 'description': ln['name'] or '',
                'qty': 1, 'unit': '', 'unit_price': round(ln['balance'] * pct / 100, 2), 'tax': '',
                'currency': cur, 'fx_rate': 1, 'payment_status': status,
                'notes': f'{ln["move_name"]}' + (f' — توزيع تحليلي {pct:g}%' if len(dist) > 1 else ''),
            })
    with open(args.out, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=list(out_rows[0].keys()) if out_rows else ['date'])
        writer.writeheader()
        writer.writerows(out_rows)
    total = sum(r['unit_price'] for r in out_rows)
    print(f'تم: {len(lines)} سطر قيد → {len(out_rows)} سطر بالملف {args.out} | الإجمالي {fmt(total)} {cur}')
    print('الخطوة الجاية: python build_cost_workbook.py %s --company "..." --base %s --out report.xlsx --check' % (args.out, cur))


def cmd_trial_balance(odoo, args):
    _, cur = company_currency(odoo)
    lines = odoo.search_read('account.move.line', [('parent_state', '=', 'posted'), ('date', '>=', args.date_from),
                                                   ('date', '<=', args.date_to)],
                             ['account_id', 'debit', 'credit'])
    tb = defaultdict(lambda: [0.0, 0.0])
    for ln in lines:
        tb[name_of(ln['account_id'])][0] += ln['debit']
        tb[name_of(ln['account_id'])][1] += ln['credit']
    with open(args.out, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['account', 'debit', 'credit', 'balance'])
        for acc in sorted(tb):
            d, c = tb[acc]
            writer.writerow([acc, round(d, 2), round(c, 2), round(d - c, 2)])
    td, tc = sum(v[0] for v in tb.values()), sum(v[1] for v in tb.values())
    print(f'تم: {len(tb)} حساب → {args.out} | مدين {fmt(td)} | دائن {fmt(tc)} {cur} | '
          + ('متوازن ✓' if round(td - tc, 2) == 0 else f'⚠ غير متوازن بفرق {fmt(td - tc)}'))


# ---------- التصدير الكامل (export) ----------

PL_PREFIXES = ('income', 'expense')


def fiscal_periods(start_month, today=None):
    today = today or dt.date.today()
    year = today.year if today.month >= start_month else today.year - 1
    start = dt.date(year, start_month, 1)
    prev_start = dt.date(year - 1, start_month, 1)
    return {'FROM': start, 'TO': today, 'PREV_FROM': prev_start, 'PREV_TO': start - dt.timedelta(days=1)}


def cell(v):
    if v is False or v is None:
        return ''
    if isinstance(v, list) and len(v) == 2 and isinstance(v[0], int) and isinstance(v[1], str):
        return v[1]
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    if isinstance(v, float):
        return f'{v:.6f}'.rstrip('0').rstrip('.')
    return v


def export_model(odoo, out_dir, fname, model, domain, wanted, control=(), order=None):
    """يصدّر موديلاً إلى CSV بالحقول المتوفرة فقط، ويرجع سطر الفهرس."""
    entry = {'file': fname + '.csv', 'model': model, 'rows': '', 'control': '', 'status': ''}
    try:
        meta = odoo.call(model, 'fields_get', [], attributes=['type', 'relation'])
    except RuntimeError:
        entry.update(file=fname, status='غير مثبت أو بدون صلاحية')
        return entry, []
    fields = [f for f in wanted if f in meta]
    rows = odoo.search_read(model, domain, fields, order=order)
    for f in fields:  # أسماء بدل أرقام لحقول القوائم المتعددة
        if meta[f]['type'] in ('many2many', 'one2many') and meta[f].get('relation'):
            ids = sorted({i for r in rows for i in (r[f] or [])})
            names = {}
            for k in range(0, len(ids), PAGE):
                for rec in odoo.call(meta[f]['relation'], 'read', [ids[k:k + PAGE]], fields=['display_name']):
                    names[rec['id']] = rec['display_name']
            for r in rows:
                r[f] = '، '.join(names.get(i, str(i)) for i in (r[f] or []))
    sums = {c: 0.0 for c in control if c in fields}
    with open(os.path.join(out_dir, entry['file']), 'w', newline='', encoding='utf-8-sig') as fh:
        w = csv.writer(fh)
        w.writerow(fields)
        for r in rows:
            w.writerow([cell(r.get(f)) for f in fields])
            for c in sums:
                sums[c] += r.get(c) or 0.0
    missing = [f for f in wanted if f not in meta]
    entry.update(rows=len(rows), control='، '.join(f'{c} = {v:,.2f}' for c, v in sums.items()),
                 status='تم' + (f' (حقول غير موجودة بهذه النسخة: {", ".join(missing)})' if missing else ''))
    return entry, rows


def grouped_by_account(odoo, domain):
    """مجاميع المدين والدائن لكل حساب: read_group، وإذا ما توفر فبالتجميع المحلي."""
    try:
        res = odoo.call('account.move.line', 'read_group', [domain, ['debit:sum', 'credit:sum'], ['account_id']], lazy=False)
        return {g['account_id'][0]: (g['debit'] or 0.0, g['credit'] or 0.0) for g in res if g.get('account_id')}
    except RuntimeError:
        agg = defaultdict(lambda: [0.0, 0.0])
        for ln in odoo.search_read('account.move.line', domain, ['account_id', 'debit', 'credit']):
            a = agg[ln['account_id'][0]]
            a[0] += ln['debit']
            a[1] += ln['credit']
        return {k: tuple(v) for k, v in agg.items()}


def export_trial_balance(odoo, out_dir, periods):
    """ميزان مراجعة: حركة الفترة لكل الحسابات، والرصيد التراكمي لحسابات الميزانية فقط،
    ونتائج السنوات السابقة سطراً مستقلاً حتى يتوازن الميزان."""
    posted = [('parent_state', '=', 'posted')]
    f, t = periods['FROM'].isoformat(), periods['TO'].isoformat()
    accounts = {a['id']: a for a in odoo.search_read('account.account', [], ['code', 'name', 'account_type'])}
    period = grouped_by_account(odoo, posted + [('date', '>=', f), ('date', '<=', t)])
    cumulative = grouped_by_account(odoo, posted + [('date', '<=', t)])
    before = grouped_by_account(odoo, posted + [('date', '<', f)])
    is_pl = lambda aid: str(accounts.get(aid, {}).get('account_type', '')).startswith(PL_PREFIXES)
    prior_pl = sum(d - c for aid, (d, c) in before.items() if is_pl(aid))
    fname = '13_ميزان_المراجعة.csv'
    tot = [0.0, 0.0, 0.0]
    by_type = defaultdict(float)
    with open(os.path.join(out_dir, fname), 'w', newline='', encoding='utf-8-sig') as fh:
        w = csv.writer(fh)
        w.writerow(['code', 'name', 'account_type', 'period_debit', 'period_credit', 'period_balance', 'closing_balance'])
        for aid in sorted(set(period) | set(cumulative), key=lambda a: str(accounts.get(a, {}).get('code', a))):
            acc = accounts.get(aid, {'code': aid, 'name': '', 'account_type': ''})
            d, c = period.get(aid, (0.0, 0.0))
            closing = d - c if is_pl(aid) else (lambda x: x[0] - x[1])(cumulative.get(aid, (0.0, 0.0)))
            w.writerow([acc['code'], acc['name'], acc['account_type'], cell(d), cell(c), cell(d - c), cell(closing)])
            tot[0] += d
            tot[1] += c
            tot[2] += closing
            by_type[acc['account_type']] += closing
        w.writerow(['', 'نتائج السنوات السابقة (محسوبة)', 'equity_unaffected', '', '', '', cell(prior_pl)])
    tot[2] += prior_pl
    ok = round(tot[0] - tot[1], 2) == 0 and round(tot[2], 2) == 0
    return {'file': fname, 'model': 'account.move.line (مجمّع)', 'rows': len(set(period) | set(cumulative)) + 1,
            'control': f'مدين الفترة = {tot[0]:,.2f}، دائن الفترة = {tot[1]:,.2f}، مجموع الأرصدة = {tot[2]:,.2f}',
            'status': 'تم، متوازن' if ok else 'تم، غير متوازن ويحتاج تحقق', '_by_type': by_type}


def export_aging(odoo, out_dir, fname, acc_type, sign, today):
    lines = odoo.search_read('account.move.line', [('parent_state', '=', 'posted'), ('account_id.account_type', '=', acc_type),
                                                   ('reconciled', '=', False), ('amount_residual', '!=', 0)],
                             ['partner_id', 'date', 'date_maturity', 'amount_residual'])
    labels = ['not_due', 'd1_30', 'd31_90', 'd91_180', 'd180_plus']
    per = defaultdict(lambda: [0.0] * 5)
    for ln in lines:
        days = (today - dt.date.fromisoformat(ln['date_maturity'] or ln['date'])).days
        k = 0 if days <= 0 else 1 if days <= 30 else 2 if days <= 90 else 3 if days <= 180 else 4
        per[name_of(ln['partner_id']) or 'بدون شريك'][k] += sign * ln['amount_residual']
    total = 0.0
    with open(os.path.join(out_dir, fname + '.csv'), 'w', newline='', encoding='utf-8-sig') as fh:
        w = csv.writer(fh)
        w.writerow(['partner'] + labels + ['total'])
        for partner, b in sorted(per.items(), key=lambda x: -sum(x[1])):
            w.writerow([partner] + [cell(v) for v in b] + [cell(sum(b))])
            total += sum(b)
    return {'file': fname + '.csv', 'model': 'account.move.line (غير مسوّاة)', 'rows': len(per),
            'control': f'الإجمالي = {total:,.2f}', 'status': 'تم', '_total': total}


def cmd_export(odoo, args):
    p = fiscal_periods(args.fy_start_month)
    f, t, pf, pt = (p[k].isoformat() for k in ('FROM', 'TO', 'PREV_FROM', 'PREV_TO'))
    out_dir = os.path.join(args.out_root, 'Odoo_تصدير_' + dt.datetime.now().strftime('%Y-%m-%d_%H%M'))
    os.makedirs(out_dir, exist_ok=True)
    log = open(os.path.join(out_dir, 'run.log'), 'a', encoding='utf-8')
    say = lambda m: (log.write(dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S  ') + m + '\n'), log.flush())
    started = dt.datetime.now()
    company, cur = company_currency(odoo)
    version = odoo._call('common', 'version', []).get('server_version', '?')
    say(f'اتصال: {company} | Odoo {version} | السنة الحالية {f} إلى {t}')

    posted = [('parent_state', '=', 'posted')]
    year = lambda a, b: [('date', '>=', a), ('date', '<=', b)]
    aml = ['date', 'move_name', 'journal_id', 'ref', 'account_id', 'partner_id', 'name', 'analytic_distribution',
           'debit', 'credit', 'balance', 'currency_id', 'amount_currency', 'matching_number', 'date_maturity']
    inv = ['name', 'partner_id', 'invoice_date', 'invoice_date_due', 'ref', 'amount_total_signed',
           'amount_residual_signed', 'payment_state', 'invoice_user_id', 'currency_id', 'move_type']
    specs = [
        ('10_دليل_الحسابات', 'account.account', [], ['code', 'name', 'account_type', 'reconcile', 'currency_id'], (), 'code'),
        ('11_قيود_اليومية_السنة_الحالية', 'account.move.line', posted + year(f, t), aml, ('debit', 'credit'), 'date, id'),
        ('12_قيود_اليومية_السنة_السابقة', 'account.move.line', posted + year(pf, pt), aml, ('debit', 'credit'), 'date, id'),
        ('16_فواتير_الزبائن', 'account.move', [('move_type', 'in', ['out_invoice', 'out_refund']), ('state', '=', 'posted')]
         + [('invoice_date', '>=', f), ('invoice_date', '<=', t)], inv, ('amount_total_signed', 'amount_residual_signed'), 'invoice_date'),
        ('17_فواتير_الموردين', 'account.move', [('move_type', 'in', ['in_invoice', 'in_refund']), ('state', '=', 'posted')]
         + [('invoice_date', '>=', f), ('invoice_date', '<=', t)], inv, ('amount_total_signed', 'amount_residual_signed'), 'invoice_date'),
        ('18_الدفعات', 'account.payment', [('date', '>=', f), ('date', '<=', t)],
         ['date', 'name', 'journal_id', 'partner_id', 'payment_type', 'partner_type', 'amount', 'currency_id', 'state'], ('amount',), 'date'),
        ('19_الوكلاء_والزبائن', 'res.partner', [('is_company', '=', True)],
         ['name', 'city', 'phone', 'user_id', 'property_payment_term_id', 'credit_limit', 'credit', 'debit', 'category_id'], ('credit',), 'name'),
        ('20_أسعار_العملات', 'res.currency.rate', [], ['name', 'currency_id', 'rate', 'inverse_company_rate', 'company_rate'], (), 'name'),
        ('21_المنتجات', 'product.product', [], ['default_code', 'name', 'categ_id', 'uom_id', 'standard_price', 'list_price',
                                               'detailed_type', 'type', 'qty_available', 'weight'], (), 'default_code'),
        ('22_تقييم_المخزون', 'stock.valuation.layer', [], ['product_id', 'create_date', 'quantity', 'value', 'description'], ('value',), 'create_date'),
        ('23_أوامر_الشراء', 'purchase.order', [('date_order', '>=', f)], ['name', 'partner_id', 'date_order', 'amount_total',
                                                                          'currency_id', 'state', 'invoice_status', 'receipt_status'], (), 'date_order'),
        ('24_أوامر_البيع', 'sale.order', [('date_order', '>=', f)], ['name', 'partner_id', 'date_order', 'amount_total',
                                                                     'state', 'invoice_status'], ('amount_total',), 'date_order'),
        ('25_أوامر_التصنيع', 'mrp.production', [('create_date', '>=', f)], ['name', 'product_id', 'product_qty', 'qty_produced',
                                                                             'date_start', 'date_planned_start', 'date_finished', 'bom_id', 'state'], (), 'create_date'),
        ('26_قوائم_المواد', 'mrp.bom.line', [], ['bom_id', 'product_id', 'product_qty', 'product_uom_id'], (), 'bom_id'),
        ('27_طلبات_الموافقة', 'approval.request', [], ['name', 'category_id', 'request_owner_id', 'date', 'request_status', 'amount'], (), 'id'),
        ('28_الدفاتر', 'account.journal', [], ['name', 'type', 'code', 'currency_id', 'default_account_id'], (), 'code'),
        ('29_الحسابات_التحليلية', 'account.analytic.account', [], ['name', 'plan_id', 'code', 'balance'], (), 'name'),
    ]
    entries, failures = [], 0
    for fname, model, domain, wanted, control, order in specs:
        try:
            entry, rows = export_model(odoo, out_dir, fname, model, domain, wanted, control, order)
            if fname.startswith(('11_', '12_')) and rows:
                d, c = sum(r['debit'] for r in rows), sum(r['credit'] for r in rows)
                entry['status'] += '، المدين = الدائن' if round(d - c, 2) == 0 else '، المدين لا يساوي الدائن ويحتاج تحقق'
        except Exception as e:  # ملف واحد فاشل ما يوقف الباقي
            entry, failures = {'file': fname, 'model': model, 'rows': '', 'control': '', 'status': f'فشل: {e}'}, failures + 1
        entries.append(entry)
        say(f'{entry["file"]}: {entry["status"]} ({entry["rows"]} سطر)')
    for fname, fn in (('13_ميزان_المراجعة', lambda: export_trial_balance(odoo, out_dir, p)),
                      ('14_أعمار_الذمم_المدينة', lambda: export_aging(odoo, out_dir, '14_أعمار_الذمم_المدينة', 'asset_receivable', 1, p['TO'])),
                      ('15_أعمار_الذمم_الدائنة', lambda: export_aging(odoo, out_dir, '15_أعمار_الذمم_الدائنة', 'liability_payable', -1, p['TO']))):
        try:
            entry = fn()
        except Exception as e:
            entry, failures = {'file': fname, 'model': 'account.move.line', 'rows': '', 'control': '', 'status': f'فشل: {e}'}, failures + 1
        entries.append(entry)
        say(f'{entry["file"]}: {entry["status"]}')

    # مطابقة الأعمار مع أرصدة الميزان: الفرق يعني قيوداً على حسابات الذمم بدون تسوية صحيحة
    found = {e['file'][:2]: e for e in entries}
    checks = []
    if '_by_type' in found.get('13', {}):
        for code, acc_type, sign, label in (('14', 'asset_receivable', 1, 'الذمم المدينة'), ('15', 'liability_payable', -1, 'الذمم الدائنة')):
            if '_total' in found.get(code, {}):
                diff = found[code]['_total'] - sign * found['13']['_by_type'].get(acc_type, 0.0)
                checks.append(f'- مطابقة {label}: الأعمار مقابل الميزان '
                              + ('متطابقة' if round(diff, 2) == 0 else f'بفرق {diff:,.2f} ويحتاج تحقق'))
                if round(diff, 2) != 0:
                    failures += 1
    entries.sort(key=lambda e: e['file'])
    lines = ['# فهرس تصدير Odoo (قراءة فقط)', '',
             f'- الشركة: {company}، العملة: {cur}، نسخة Odoo: {version}',
             f'- البداية: {started:%Y-%m-%d %H:%M:%S}، والنهاية: {dt.datetime.now():%Y-%m-%d %H:%M:%S}',
             f'- السنة الحالية: {f} إلى {t}، والسابقة: {pf} إلى {pt}. القيود المرحّلة فقط.',
             '- كل الاستدعاءات للقراءة فقط (search_read وread وfields_get وread_group). لم يُكتب شي على Odoo.',
             '- أسماء الأعمدة هي الأسماء التقنية للحقول بـ Odoo.', '',
             *checks, '',
             '| الملف | الموديل | الأسطر | المجاميع الضابطة | الحالة |', '|---|---|---:|---|---|']
    lines += [f"| {e['file']} | {e['model']} | {e['rows']} | {e['control']} | {e['status']} |" for e in entries]
    with open(os.path.join(out_dir, '00_فهرس.md'), 'w', encoding='utf-8-sig') as fh:
        fh.write('\n'.join(lines) + '\n')
    say(f'انتهى. فشل: {failures}')
    log.close()
    print(f'{out_dir} | فشل: {failures}')
    return 1 if failures else 0


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest='cmd', required=True)
    sub.add_parser('snapshot', help='صورة مالية لحظية')
    for name in ('costs', 'trial-balance'):
        p = sub.add_parser(name)
        p.add_argument('--from', dest='date_from', required=True, help='YYYY-MM-DD')
        p.add_argument('--to', dest='date_to', required=True, help='YYYY-MM-DD')
        p.add_argument('--out', required=True, help='مسار ملف CSV الناتج')
        if name == 'costs':
            p.add_argument('--account-prefix', help='بدل نوع الحساب: كل الحسابات التي يبدأ رمزها بهذا (مثل 3 للدليل الموحد)')
            p.add_argument('--code-digits', type=int, help='قص رمز الحساب لعدد أرقام (مثل 4) ليطابق دليل الترميز')
            p.add_argument('--include-capex', action='store_true', help='تضمين حسابات الأصول الثابتة')
    p = sub.add_parser('export', help='تصدير كامل صامت إلى مجلد مؤرخ')
    p.add_argument('--out-root', required=True, help='المجلد الأب للتصدير، مثل CFO\\بيانات')
    p.add_argument('--fy-start-month', type=int, default=1, help='شهر بداية السنة المالية (افتراضياً 1)')
    args = parser.parse_args()
    if args.cmd == 'export':
        # بالتشغيل الصامت ما توجد شاشة، فأي خطأ قبل إنشاء المجلد يُكتب بسجل بالمجلد الأب
        try:
            sys.exit(cmd_export(Odoo(), args))
        except (SystemExit, Exception) as e:
            if isinstance(e, SystemExit) and e.code in (0, 1, None):
                raise
            os.makedirs(args.out_root, exist_ok=True)
            with open(os.path.join(args.out_root, 'odoo_export_errors.log'), 'a', encoding='utf-8') as fh:
                fh.write(f'{dt.datetime.now():%Y-%m-%d %H:%M:%S}  {e}\n')
            sys.exit(2)
    odoo = Odoo()
    {'snapshot': cmd_snapshot, 'costs': cmd_costs, 'trial-balance': cmd_trial_balance}[args.cmd](odoo, args)


if __name__ == '__main__':
    try:
        main()
    except RuntimeError as e:
        sys.exit(f'خطأ من Odoo: {e}')
