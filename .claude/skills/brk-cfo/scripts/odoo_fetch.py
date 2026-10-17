#!/usr/bin/env python3
"""قراءة البيانات الحية من Odoo (قراءة فقط) للمدير المالي.

الأوامر:
  snapshot        صورة مالية لحظية: النقد، الذمم المدينة والدائنة وأعمارها، الفواتير غير المسددة،
                  أوامر الشراء المفتوحة، طلبات الموافقة المعلقة، الإيرادات والمصاريف للشهر والسنة.
  costs           قيود المصاريف المرحّلة لفترة → CSV جاهز لـ build_cost_workbook.py.
  trial-balance   ميزان المراجعة لفترة (لمطابقة تقرير التكاليف مع الدفاتر).

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

مكتوب لـ Odoo 16 فما فوق (account_type وanalytic_distribution) عبر واجهة JSON-RPC.
السكربت لا يستدعي أي دالة كتابة: المسموح فقط search_read وsearch_count وread.
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
READ_ONLY_METHODS = {'search_read', 'search_count', 'read'}
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
    args = parser.parse_args()
    odoo = Odoo()
    {'snapshot': cmd_snapshot, 'costs': cmd_costs, 'trial-balance': cmd_trial_balance}[args.cmd](odoo, args)


if __name__ == '__main__':
    try:
        main()
    except RuntimeError as e:
        sys.exit(f'خطأ من Odoo: {e}')
