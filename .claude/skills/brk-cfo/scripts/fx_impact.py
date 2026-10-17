#!/usr/bin/env python3
"""أثر تغيّر سعر الصرف على الأرصدة بالعملة الأجنبية (IAS 21).

البنود النقدية (نقد، ذمم مدينة ودائنة، قروض) يعاد تقييمها بسعر الإقفال ويظهر الفرق ربحاً أو خسارة.
البنود غير النقدية (مخزون، أصول ثابتة، دفعات مقدمة) تبقى بالسعر التاريخي محاسبياً،
لكن كلفة استبدالها ترتفع — وهذا مهم للتسعير وليس للقيد.

الاستخدام:
  python fx_impact.py balances.csv --new 1520 [--old 1320] [--base IQD]

أعمدة الملف: item, side (أصل/التزام أو asset/liability), currency, amount,
             booked_rate (اختياري، وإلا يُستخدم --old), monetary (نعم/لا، افتراضياً نعم)
"""
import argparse
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from validate_costs import parse_number  # noqa: E402

ASSET = {'أصل', 'موجود', 'asset', 'a'}
LIABILITY = {'التزام', 'مطلوب', 'liability', 'l'}
NO = {'لا', 'no', 'n', 'false', '0'}


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('balances', help='ملف الأرصدة CSV')
    parser.add_argument('--new', type=float, required=True, help='سعر الصرف الجديد (مثل 1520)')
    parser.add_argument('--old', type=float, help='السعر القديم إذا لم يُذكر booked_rate في الملف')
    parser.add_argument('--base', default='IQD')
    args = parser.parse_args()

    monetary, non_monetary = [], []
    with open(args.balances, encoding='utf-8-sig', newline='') as f:
        for rec in csv.DictReader(f):
            side = rec['side'].strip().lower()
            if side not in ASSET | LIABILITY:
                raise SystemExit(f'قيمة side غير مفهومة: {rec["side"]}')
            rate = parse_number(rec.get('booked_rate')) or args.old
            if not rate:
                raise SystemExit(f'لا يوجد سعر مسجل للبند "{rec["item"]}" — أضف booked_rate أو --old')
            amount = parse_number(rec['amount'])
            sign = 1 if side in ASSET else -1
            item = {'item': rec['item'], 'side': 'أصل' if sign == 1 else 'التزام', 'currency': rec['currency'],
                    'amount': amount, 'old': amount * rate, 'new': amount * args.new}
            # الربح/الخسارة: ارتفاع السعر يربح على الأصول ويخسر على الالتزامات
            item['effect'] = sign * (item['new'] - item['old'])
            (non_monetary if (rec.get('monetary') or '').strip().lower() in NO else monetary).append(item)

    b = args.base
    print(f'## أثر سعر الصرف الجديد {args.new:,.2f}\n')
    print(f'| البند | النوع | المبلغ | بالسعر المسجل ({b}) | بالسعر الجديد ({b}) | ربح / (خسارة) |')
    print('|---|---|---:|---:|---:|---:|')
    for x in monetary:
        eff = f'{x["effect"]:,.0f}' if x['effect'] >= 0 else f'({-x["effect"]:,.0f})'
        print(f'| {x["item"]} | {x["side"]} | {x["amount"]:,.2f} {x["currency"]} | {x["old"]:,.0f} | {x["new"]:,.0f} | {eff} |')
    net_fc = sum((1 if x['side'] == 'أصل' else -1) * x['amount'] for x in monetary)
    total = sum(x['effect'] for x in monetary)
    total_txt = f'{total:,.0f}' if total >= 0 else f'({-total:,.0f})'
    net_txt = f'{net_fc:,.2f}' if net_fc >= 0 else f'({-net_fc:,.2f})'
    print(f'| **صافي البنود النقدية** | | **{net_txt}** | | | **{total_txt}** |\n')

    position = 'صافي أصول' if net_fc > 0 else 'صافي التزامات' if net_fc < 0 else 'مركز متوازن'
    print(f'- المركز بالعملة الأجنبية: **{position}** بمبلغ {abs(net_fc):,.2f}.')
    print(f'- الأثر على الربح: **{"ربح" if total >= 0 else "خسارة"} فروقات صرف {abs(total):,.0f} {b}** '
          '(قيد: البنود النقدية بسعر الإقفال، والفرق إلى قائمة الدخل — IAS 21.23 و21.28).')
    if non_monetary:
        repl = sum(x['new'] - x['old'] for x in non_monetary)
        print(f'- بنود غير نقدية ({len(non_monetary)}): لا يعاد تقييمها محاسبياً، لكن كلفة استبدالها ارتفعت '
              f'{repl:,.0f} {b} — راجع أسعار البيع وصافي القيمة القابلة للتحقق (IAS 2).')
    if net_fc < 0:
        print('- مركزك مكشوف على الدولار: كل ارتفاع 10 دنانير بالسعر يكلفك '
              f'{abs(net_fc) * 10:,.0f} {b}. فكّر بتسوية الالتزامات أو مطابقة عملة الإيراد مع عملة الكلفة.')


if __name__ == '__main__':
    main()
