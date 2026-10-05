# =====================================================================
#  قاعدة الراتب ABSENCE — نسخة جديدة (BRK)
#  المكان: الرواتب ← الإعدادات ← قواعد الراتب ← ABSENCE
#          (Base Structure = rule id 37)        ← الصق الكود مثل ما هو (STRUCTURE = 'IQD')
#          (الهيكل الأساسي USD = rule id 57)    ← نفس الكود، بس STRUCTURE = 'USD'
#  Amount Type = Python Code
#  مكتوب لـ Odoo 17+ (payslip سجل، و inputs و categories قواميس dict)
#
#  ليش: الحقل payslip.absence_deduction_amount (موديول EJAF: ejaf_hr_absence)
#  يحسب غياب غلط بثلاث حالات طلعت بمطابقة أيلول 2026:
#    1) تبديل الشفت (Working Hours History): Odoo يولّد ساعات الشفت القديم
#       والجديد سوة ليوم أو يومين حول التبديل ← غياب على شفت مو مالته
#       (125، 303، 435، 84، 418)
#    2) خطأ ترتيب البصمات: «دخول» مرتين قبل «الخروج» ← سجل حضور 0 ساعة ← غياب
#       (389، 403)
#    3) أيام غياب بالإيصال بدون أي سجل غياب (543)
#
#  القاعدة الجديدة تحسب اليوم غياب بس إذا:
#    - بيه ساعات غياب ≥ 6 بسجلات العمل (بتاريخ بغداد، وذيل الشفت الليلي 00:00–07:00
#      يتبع اليوم اللي قبله)، والجمعة لها قاعدة FRIDAY
#    - وما عنده حضور مجموعه ≥ 6 ساعات بنفس اليوم (= بدّل شفت، مو غايب)
#    - وما عنده بصمتين بالجهاز بيناتهن 6–16 ساعة حول الشفت (= خطأ ترتيب بصمات)
#    - ومو بصمة وحدة داخل وقت الشفت (= نسيان بصمة، تروح لسياسة نسيان البصمة)
#    - ومو بإجازة (يوم كامل، أو نص يوم/زمنية تتقاطع ويه ساعات الغياب)
#      الزمنية (4 ساعات بالشهر) ما تنحسب ببداية الدوام: تغطي بس إذا وصل الشركة قبلها
#    - ومو عطلة رسمية (Global Time Off) يبدي بيها الشفت
#    - والموظف عنده سياسة غياب بمجموعة السياسات، ومو «معفي من البصمة»
#
#  تجربة على بيانات أيلول (330 إيصال): يتغيّر 8 إيصالات بس، وكلها مطابقة للبصمة:
#    84: 7 ← 5 أيام | 125، 303، 435، 543: ← صفر | 418: 3 ← 2 | 389، 403: ← صفر
#  وباقي 322 إيصال نفس المبلغ بالضبط.
#  العطلة الرسمية تنحسب بس إذا مدخلة بأودو كـ Global Time Off (اللي مو مدخلة = يوم دوام عادي).
#  الشفت نفس الطول طول السنة (حتى رمضان)، فحد الـ 6 ساعات ثابت.
#  جرّبها على Staging أول، وطابق أيلول وتشرين قبل الإنتاج (tools/odoo_absence.py compare).
# =====================================================================

STRUCTURE = 'IQD'     # 'IQD' لقاعدة 37 (الدينار) | 'USD' لقاعدة 57 (الدولار)
RATE = 1.5            # غياب دبل
MIN_HOURS = 6.0       # أقل ساعات غياب حتى ينحسب يوم
WORK_MIN = 6.0        # شفت حقيقي (حضور أو زوج بصمات) بين 6 و 16 ساعة
WORK_MAX = 16.0
EXEMPT_TAG = 'معفي من البصمة'
ABSENCE_TYPE = 'Absence'   # اسم نوع سجل العمل بالإنكليزي

env = payslip.env
emp = payslip.employee_id
to_date = payslip.date_from.fromordinal
one_hour = (to_date(2) - to_date(1)) / 24

start = max(payslip.date_from, contract.date_start or payslip.date_from).toordinal()
end = payslip.date_to.toordinal()
if contract.date_end:
    end = min(end, contract.date_end.toordinal())

days = []
pg = contract.policy_group_id
if pg and pg.absence_policy_id and EXEMPT_TAG not in emp.category_ids.mapped('name'):
    # 1) ساعات الغياب لكل يوم شفت (التواريخ مخزونة UTC؛ بغداد = UTC+3)
    #    lang=en_US: اسم النوع حقل مترجم، وبدونها البحث يصير بلغة اللي يحسب الإيصال
    hours = {}
    blocks = {}
    for w in env['hr.work.entry'].sudo().with_context(lang='en_US').search([
            ('employee_id', '=', emp.id),
            ('work_entry_type_id.name', '=', ABSENCE_TYPE),
            ('state', '!=', 'cancelled'),
            ('date_start', '>=', str(to_date(start - 2))),
            ('date_start', '<=', str(to_date(end + 1)))]):
        ds = w.date_start
        if ds.hour == 21 and ds.minute == 0:
            o = ds.toordinal()                               # يبدي 00:00 بغداد = ذيل شفت ليلي
        else:
            o = ds.toordinal() + (1 if ds.hour >= 21 else 0)  # تاريخ بغداد
        if start <= o <= end and to_date(o).weekday() != 4:  # الجمعة لها قاعدة FRIDAY
            hours[o] = hours.get(o, 0.0) + (w.duration or 0.0)
            blocks.setdefault(o, []).append(w)
    cand = sorted([o for o in hours if hours[o] >= MIN_HOURS])

    if cand:
        # 2) ساعات الحضور لكل يوم (مجموع السجلات، حتى الطلعة بالاستراحة ما تخرّب الحساب)
        worked = {}
        check_ins = []
        for a in env['hr.attendance'].sudo().search([
                ('employee_id', '=', emp.id),
                ('check_in', '>=', str(to_date(start - 2))),
                ('check_in', '<=', str(to_date(end + 2)))]):
            check_ins.append(a.check_in)
            if 0 < (a.worked_hours or 0.0) <= WORK_MAX:
                d = a.check_in.toordinal() + (1 if a.check_in.hour >= 21 else 0)
                worked[d] = worked.get(d, 0.0) + a.worked_hours

        leaves = env['hr.leave'].sudo().search([
            ('employee_id', '=', emp.id),
            ('state', '=', 'validate'),
            ('request_date_from', '<=', to_date(end + 1)),
            ('request_date_to', '>=', to_date(start - 1))])

        # العطل الرسمية: Global Time Off (بدون موظف)، للكل أو لتقويم الموظف
        holidays = env['resource.calendar.leaves'].sudo().search([
            ('resource_id', '=', False),
            ('time_type', '=', 'leave'),
            ('calendar_id', 'in', [False] + (contract.resource_calendar_id | emp.resource_calendar_id).ids),
            ('company_id', 'in', [False, contract.company_id.id]),
            ('date_from', '<=', str(to_date(end + 2))),
            ('date_to', '>=', str(to_date(start - 2)))])

        for o in cand:
            if worked.get(o, 0.0) >= WORK_MIN:
                continue                                     # داوم (بدّل شفت)، مو غايب
            b0 = min([w.date_start for w in blocks[o]])
            b1 = max([w.date_stop for w in blocks[o]])

            if [h for h in holidays if h.date_from <= b0 < h.date_to]:
                continue                                     # الشفت يبدي بعطلة رسمية

            # 3) البصمات الخام من الجهاز حول وقت الشفت (المكررة بنفس الثانية = بصمة وحدة)
            punches = sorted(set(env['hr.attendance.zk.temp'].sudo().search([
                ('employee_id', '=', emp.id),
                ('date', '>=', b0 - 5 * one_hour),
                ('date', '<=', b1 + one_hour)]).mapped('date')))
            # زوج بصمات بيناتهن 6–16 ساعة. الحد الأعلى ضروري: بيوم تبديل الشفت النافذة
            # تصير ~29 ساعة، وخروج الشفت اللي قبله ويه دخول اللي بعده يبينون «دوام»
            shift_done = False
            for i in range(len(punches)):
                for j in range(i + 1, len(punches)):
                    if WORK_MIN * one_hour <= punches[j] - punches[i] <= WORK_MAX * one_hour:
                        shift_done = True
            if shift_done:
                continue                                     # داوم، بس البصمات انرتبت غلط
            if len(punches) == 1 and b0 - one_hour <= punches[0] <= b1:
                continue                                     # نسيان بصمة، مو غياب

            # 4) الإجازات. الزمنية (إجازة ساعية) ما تنحسب ببداية الدوام: تغطي بس إذا
            #    الموظف وصل الشركة (بصمة أو تسجيل دخول) قبل ما تبدي، وإلا اليوم غياب
            arrivals = punches + [c for c in check_ins if b0 - 5 * one_hour <= c <= b1]
            on_leave = False
            for lv in leaves:
                if lv.request_unit_hours and not [p for p in arrivals if p < lv.date_from]:
                    continue
                if not lv.request_unit_hours and not lv.request_unit_half and \
                        lv.request_date_from.toordinal() <= o <= lv.request_date_to.toordinal():
                    on_leave = True
                for w in blocks[o]:
                    if lv.date_from < w.date_stop and lv.date_to > w.date_start:
                        on_leave = True
            if not on_leave:
                days.append(o)

# ---- المبلغ ----
if STRUCTURE == 'USD':
    daily = contract.wage / ((pg.work_days_per_month or 30) if pg else 30.0)
else:
    allw = (inputs['TRFOD'].amount if 'TRFOD' in inputs else 0.0) + \
           (inputs['ALWNT'].amount if 'ALWNT' in inputs else 0.0)
    daily = (contract.wage + allw) / 30.0
result = -(len(days) * RATE * daily)
if 'ABSENCE' in inputs:
    result -= inputs['ABSENCE'].amount
if STRUCTURE == 'USD':
    # الخصم ما يعبر المتبقي من الراتب (الأساسي + البدلات + الخصومات اللي قبل هالقاعدة)
    available = categories['BASIC'] + categories['ALW'] + categories['DED']
    result = max(result, -max(available, 0))
if days:
    result_name = 'خصم الغياب: %d يوم × 1.5 (%s)' % (
        len(days), '، '.join([to_date(o).strftime('%d/%m') for o in days]))
