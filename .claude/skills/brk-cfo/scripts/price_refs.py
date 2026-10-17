#!/usr/bin/env python3
"""مراجع الأسعار من تصدير الخازن: تضخم كلف الشركة الفعلي، وآخر أسعار الشراء والبيع لكل مادة ومنتج.

  1. تضخم الشراء الخاص بالشركة: لكل مادة انشرت بالفترتين، تغير متوسط سعرها بعملتها بين فترة الأساس
     والفترة الأخيرة، مرجح بقيمة مشترياتها بالدينار. يطلع للمواد بالدينار والمواد بالدولار كل وحدة.
  2. مراجع_اسعار_الشراء.csv: آخر سعر ومتوسط الفترة الأخيرة لكل مادة، بعملتها وسعر صرف قيدها.
  3. مراجع_اسعار_البيع.csv: متوسط صافي سعر البيع للوحدة بالفترة الأخيرة لكل منتج (قبل خصم 14% النصف سنوي).
  4. materials_seed.csv: بصيغة ملف المواد لـ build_pricing.py (السعر الجديد = آخر سعر شراء).

الاستخدام:
  python price_refs.py "<مجلد الخازن_تصدير_التاريخ>" [--out-dir <مجلد>] [--recent-days 90] [--base-months 3]

الأرقام من فواتير الشراء والبيع فقط، وما تشمل الكمرك والتخليص إذا انقيدت بقيود منفصلة.
"""
import argparse
import csv
import datetime as dt
import statistics
import sys
from collections import defaultdict
from pathlib import Path

PURCHASE_FILE = '18_فواتير_الشراء_تفاصيل.csv'
SALES_FILE = '16_فواتير_البيع_تفاصيل.csv'
SALE_RETURN_KINDS = {'021101'}
USD_NAMES = {'دولار أميركي', 'دولار امريكي', 'دولار', 'USD'}


def num(v):
    try:
        return float(str(v).replace(',', ''))
    except (TypeError, ValueError):
        return 0.0


def load(path):
    with open(path, encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))


def day(s):
    return dt.date.fromisoformat(str(s)[:10])


def purchases(folder, recent_days, base_months):
    rows = [r for r in load(folder / PURCHASE_FILE) if num(r['QtyDefaultUnit']) > 0]
    if not rows:
        sys.exit('لا توجد أسطر شراء')
    end = max(day(r['BillDate']) for r in rows)
    start = min(day(r['BillDate']) for r in rows)
    recent_from = end - dt.timedelta(days=recent_days)
    base_to = start + dt.timedelta(days=base_months * 30)
    items = defaultdict(lambda: {'base_q': 0.0, 'base_v': 0.0, 'rec_q': 0.0, 'rec_v': 0.0, 'value_iqd': 0.0, 'last': None})
    for r in rows:
        d = day(r['BillDate'])
        cur = 'دولار' if r['Currency'].strip() in USD_NAMES else 'دينار'
        key = (r['ItemCode'], cur)
        it = items[key]
        q, v = num(r['QtyDefaultUnit']), num(r['NetTotal'])
        rate = num(r['ExchangeRate']) or 1
        it.update(name=r['ItemName'].strip(), unit=r['DefaultUnit'].strip(), cur=cur)
        it['value_iqd'] += v * (rate if cur == 'دولار' else 1)
        if d <= base_to:
            it['base_q'] += q
            it['base_v'] += v
        if d >= recent_from:
            it['rec_q'] += q
            it['rec_v'] += v
        if it['last'] is None or d >= it['last'][0]:
            it['last'] = (d, v / q, rate)
    return items, (start, base_to, recent_from, end)


def own_inflation(items):
    out = {}
    for cur in ('دينار', 'دولار'):
        both = [it for it in items.values() if it['cur'] == cur and it['base_q'] and it['rec_q']]
        total = sum(it['value_iqd'] for it in items.values() if it['cur'] == cur)
        w = sum(it['value_iqd'] for it in both)
        if not both or not w:
            out[cur] = None
            continue
        changes = sorted((it['rec_v'] / it['rec_q']) / (it['base_v'] / it['base_q']) - 1 for it in both)
        chg = sum(it['value_iqd'] * ((it['rec_v'] / it['rec_q']) / (it['base_v'] / it['base_q']) - 1) for it in both) / w
        out[cur] = (chg, statistics.median(changes), len(both), w / total if total else 0)
    return out


def sales(folder, recent_days, base_months):
    rows = [r for r in load(folder / SALES_FILE) if r['BillKind'] not in SALE_RETURN_KINDS and num(r['QtyDefaultUnit']) > 0]
    end = max(day(r['BillDate']) for r in rows)
    recent_from = end - dt.timedelta(days=recent_days)
    base_to = min(day(r['BillDate']) for r in rows) + dt.timedelta(days=base_months * 30)
    items = defaultdict(lambda: {'q': 0.0, 'v': 0.0, 'bq': 0.0, 'bv': 0.0, 'bills': set(), 'last': None})
    for r in rows:
        d = day(r['BillDate'])
        it = items[r['ItemCode']]
        it.update(name=r['ItemName'].strip(), unit=r['DefaultUnit'].strip(), cat=r['CategoryName'].strip())
        q, v = num(r['QtyDefaultUnit']), num(r['NetTotal'])
        if d >= recent_from:
            it['q'] += q
            it['v'] += v
            it['bills'].add(r['BillNo'])
        if d <= base_to:
            it['bq'] += q
            it['bv'] += v
        if it['last'] is None or d >= it['last'][0]:
            it['last'] = (d, v / q)
    return items, (recent_from, end)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('folder', type=Path)
    ap.add_argument('--out-dir', type=Path, default=Path('.'))
    ap.add_argument('--recent-days', type=int, default=90)
    ap.add_argument('--base-months', type=int, default=3)
    args = ap.parse_args()
    args.out_dir.mkdir(parents=True, exist_ok=True)

    items, (start, base_to, recent_from, end) = purchases(args.folder, args.recent_days, args.base_months)
    print(f'الشراء: {start} إلى {end}. الأساس لحد {base_to}، والفترة الأخيرة من {recent_from}')
    for cur, res in own_inflation(items).items():
        if res is None:
            print(f'  مواد {cur}: ما اكو مواد انشرت بالفترتين')
            continue
        chg, med, n, cover = res
        print(f'  تضخم أسعار شراء المواد بالـ{cur} (بعملتها): {chg:+.1%} مرجح بالقيمة، والوسيط {med:+.1%}، '
              f'على {n} مادة تغطي {cover:.0%} من قيمة مشتريات {cur}')
    print('  (تغير خليط الأصناف تحت نفس الرمز يرفع أو يخفض الرقم، فأكبر التغيرات تلزم مراجعتها بملف المراجع)')

    p_out = args.out_dir / 'مراجع_اسعار_الشراء.csv'
    seed = args.out_dir / 'materials_seed.csv'
    with open(p_out, 'w', encoding='utf-8-sig', newline='') as f, open(seed, 'w', encoding='utf-8-sig', newline='') as g:
        w = csv.writer(f)
        w.writerow(['code', 'name', 'unit', 'currency', 'last_date', 'last_price', 'booked_rate', 'avg_recent', 'avg_base', 'change', 'value_iqd'])
        s = csv.writer(g)
        s.writerow(['code', 'name', 'unit', 'currency', 'old_price', 'new_price', 'source'])
        for (code, _), it in sorted(items.items(), key=lambda kv: -kv[1]['value_iqd']):
            avg_r = it['rec_v'] / it['rec_q'] if it['rec_q'] else ''
            avg_b = it['base_v'] / it['base_q'] if it['base_q'] else ''
            chg = f'{avg_r / avg_b - 1:.4f}' if avg_r and avg_b else ''
            last_d, last_p, rate = it['last']
            w.writerow([code, it['name'], it['unit'], it['cur'], last_d, round(last_p, 4), rate,
                        round(avg_r, 4) if avg_r else '', round(avg_b, 4) if avg_b else '', chg, round(it['value_iqd'])])
            s.writerow([code, it['name'], 'كغ' if it['unit'] in ('كيلو', 'كغ') else it['unit'], it['cur'],
                        round(avg_b, 4) if avg_b else '', round(last_p, 4), f'آخر شراء بالخازن {last_d}'])
    s_items, (s_from, s_end) = sales(args.folder, args.recent_days, args.base_months)
    both = [it for it in s_items.values() if it['q'] and it['bq']]
    if both:
        w = sum(it['v'] for it in both)
        chg = sum(it['v'] * ((it['v'] / it['q']) / (it['bv'] / it['bq']) - 1) for it in both) / w
        print(f'  تغير أسعار البيع الفعلية (نفس الفترتين): {chg:+.1%} مرجح بمبيعات الفترة الأخيرة، على {len(both)} منتج')
    s_out = args.out_dir / 'مراجع_اسعار_البيع.csv'
    with open(s_out, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['code', 'name', 'category', 'unit', 'avg_net_price_recent', 'qty_recent', 'bills_recent', 'last_date', 'last_price'])
        for code, it in sorted(s_items.items(), key=lambda kv: -kv[1]['v']):
            w.writerow([code, it['name'], it['cat'], it['unit'], round(it['v'] / it['q']) if it['q'] else '',
                        round(it['q']), len(it['bills']), it['last'][0], round(it['last'][1])])
    print(f'البيع: الفترة الأخيرة {s_from} إلى {s_end}، {sum(1 for it in s_items.values() if it["q"])} منتج مباع')
    print(f'الملفات: {p_out.name}، {s_out.name}، {seed.name}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
