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
#    - بيه ساعات غياب ≥ 6 بسجلات العمل. يوم الشفت = منتصف سجل العمل بتوقيت بغداد ناقص 7 ساعات:
#      ذيل الشفت الليلي (00:00–07:00) يتبع اليوم اللي قبله، والشفت اللي يبدي 04:00 أو 06:00 يتبع
#      يومه. السجلات المكررة (نفس البداية والنهاية) تنحسب مرة وحدة. والجمعة لها قاعدة FRIDAY
#    - وما عنده حضور مجموعه ≥ 6 ساعات بنفس يوم الشفت (= بدّل شفت، مو غايب). سجل الحضور
#      يتبع يوم الشفت من منتصفه
#    - وما عنده بصمتين بالجهاز بيناتهن 6–16 ساعة حول الشفت (= خطأ ترتيب بصمات).
#      البصمة اللي داخل سجل حضور كامل (6–16 ساعة) لشفت ثاني ما تنحسب، والبصمات اللي بيناتها
#      ≤ 5 دقايق = وحدة
#    - ومو بصمة وحدة من ساعة قبل الشفت لنص ساعة بعده وبدون حضور مسجّل (= نسيان بصمة،
#      تروح لسياسة نسيان البصمة)
#    - ومو بإجازة يوم كامل أو نص يوم تتقاطع ويه ساعات الغياب
#    - الزمنية (4 ساعات بالشهر) ما تنحسب ببداية الدوام: إذا وصل الشركة قبلها تغطي ساعاتها بس،
#      وإذا ما وصل تنحسب ساعاتها غياب. اليوم غياب إذا الباقي بعدها ≥ 6 ساعات. الزمنية = نوعي
#      الإجازة بـ ZAM_TYPES بس؛ باقي الإجازات بالساعات (مدفوعة، عمل) تغطي ساعاتها مثل أي إجازة
#    - ومو عطلة رسمية (Global Time Off) يبدي بيها الشفت
#    - والموظف عنده سياسة غياب بمجموعة السياسات، ومو «معفي من البصمة»
#  وإذا الإيصال بيه أكثر من سطر إدخال ABSENCE، أودو يشغّل القاعدة مرة لكل سطر:
#  الأيام تنخصم بأول مرة بس.
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
ZAM_TYPES = ('الزمنية', 'الزمنية (استثناءات)')   # أنواع إجازة الزمنية (الاسم بالعربي أو الإنكليزي)

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
        # يوم الشفت من منتصف السجل (UTC - 4 = بغداد - 7)، نفس الحضور: شفت 04:00 يبقى بيومه
        o = (w.date_start + (w.date_stop - w.date_start) / 2 - 4 * one_hour).toordinal()
        if start <= o <= end and to_date(o).weekday() != 4:  # الجمعة لها قاعدة FRIDAY
            if [x for x in blocks.get(o, []) if x.date_start == w.date_start and x.date_stop == w.date_stop]:
                continue                                     # سجل مكرر (conflict) = مرة وحدة
            hours[o] = hours.get(o, 0.0) + (w.duration or 0.0)
            blocks.setdefault(o, []).append(w)
    leaves = env['hr.leave'].sudo().search([
        ('employee_id', '=', emp.id),
        ('state', '=', 'validate'),
        ('request_date_from', '<=', to_date(end + 1)),
        ('request_date_to', '>=', to_date(start - 1))])
    # نوعي الزمنية بالرقم. الاسم ينقرا بلغة اللي يحسب الإيصال وبالعربي وبالإنكليزي، لأن «الزمنية» هو نوع
    # أودو «Extra Hours» بعد ما تسمّى بالعربي، فاسمه يختلف حسب اللغة
    lt = env['hr.leave.type'].sudo()
    zam_types = (lt.search([('name', 'in', list(ZAM_TYPES))]) |
                 lt.with_context(lang='ar_001').search([('name', 'in', list(ZAM_TYPES))]) |
                 lt.with_context(lang='en_US').search([('name', 'in', list(ZAM_TYPES))])).ids
    zam = {}
    for lv in leaves:
        if lv.request_unit_hours and lv.holiday_status_id.id in zam_types:
            zo = (lv.date_from - 4 * one_hour).toordinal()
            zam[zo] = zam.get(zo, 0.0) + (lv.date_to - lv.date_from) / one_hour
    cand = sorted([o for o in hours if hours[o] + zam.get(o, 0.0) >= MIN_HOURS])

    if cand:
        # 2) ساعات الحضور لكل يوم (مجموع السجلات، حتى الطلعة بالاستراحة ما تخرّب الحساب)
        worked = {}
        check_ins = []
        full = []
        for a in env['hr.attendance'].sudo().search([
                ('employee_id', '=', emp.id),
                ('check_in', '>=', str(to_date(start - 2))),
                ('check_in', '<=', str(to_date(end + 2)))]):
            if a.in_mode != 'technical':
                check_ins.append(a.check_in)
            # سجل كامل (6–16 ساعة) = شفت حقيقي. أطول من 16 = بصمات انرتبت غلط، مو شفت ثاني
            if a.check_out and WORK_MIN <= (a.worked_hours or 0.0) <= WORK_MAX:
                full.append((a.check_in, a.check_out))
            if 0 < (a.worked_hours or 0.0) <= WORK_MAX:
                # الحضور ينحسب ليوم الشفت من منتصف السجل (نفس قاعدة ناقص 7 ساعات)
                d = (a.check_in + (a.check_out - a.check_in) / 2 - 4 * one_hour).toordinal()
                worked[d] = worked.get(d, 0.0) + a.worked_hours

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
            # البصمة اللي داخل سجل حضور كامل (6–16 ساعة) تخص شفت ثاني، مو هذا اليوم
            own = []
            for p in punches:
                inside = False
                for f in full:
                    if f[0] <= p <= f[1]:
                        inside = True
                # بصمات بيناتها ≤ 5 دقايق (ضغطتين على الجهاز) = بصمة وحدة
                if not inside and not (own and p - own[-1] <= one_hour / 12):
                    own.append(p)
            punches = own
            # زوج بصمات بيناتهن 6–16 ساعة. الحد الأعلى ضروري: بيوم تبديل الشفت النافذة
            # تصير ~29 ساعة، وخروج الشفت اللي قبله ويه دخول اللي بعده يبينون «دوام»
            shift_done = False
            for i in range(len(punches)):
                for j in range(i + 1, len(punches)):
                    if WORK_MIN * one_hour <= punches[j] - punches[i] <= WORK_MAX * one_hour:
                        shift_done = True
            if shift_done:
                continue                                     # داوم، بس البصمات انرتبت غلط
            if len(punches) == 1 and worked.get(o, 0.0) < 0.1 and \
                    b0 - one_hour <= punches[0] <= b1 + one_hour / 2:
                continue                                     # نسيان بصمة، مو غياب

            # 4) الإجازات. الزمنية (إجازة ساعية) ما تنحسب ببداية الدوام: تغطي بس إذا
            #    الموظف وصل الشركة (بصمة أو تسجيل دخول) قبل ما تبدي، وإلا اليوم غياب
            arrivals = list(punches)
            for c in check_ins:
                inside = False
                for f in full:
                    if f[0] <= c <= f[1]:
                        inside = True
                # من 06:00 بغداد بيوم الشفت أو 5 ساعات قبل أول غياب: إذا أودو أرشف ساعات الزمنية،
                # أول غياب يبدي بعدها، والوصول يكون قبلها
                if min(b0 - 5 * one_hour, b0.fromordinal(o) + 3 * one_hour) <= c <= b1 and not inside:
                    arrivals.append(c)
            on_leave = False
            lost = hours[o]
            for lv in leaves:
                if lv.request_unit_hours and lv.holiday_status_id.id in zam_types:
                    ov = 0.0
                    for w in blocks[o]:
                        ov += max(min(lv.date_to, w.date_stop) - max(lv.date_from, w.date_start), 0 * one_hour) / one_hour
                    if [p for p in arrivals if p < lv.date_from]:
                        lost -= ov                           # زمنية بعد الوصول: تغطي ساعاتها بس
                    elif (lv.date_from - 4 * one_hour).toordinal() == o:
                        lost += (lv.date_to - lv.date_from) / one_hour - ov   # زمنية ببداية الدوام = غياب
                    continue
                if not lv.request_unit_hours and not lv.request_unit_half and \
                        lv.request_date_from.toordinal() <= o <= lv.request_date_to.toordinal():
                    on_leave = True
                for w in blocks[o]:
                    if lv.date_from < w.date_stop and lv.date_to > w.date_start:
                        on_leave = True
            if not on_leave and lost >= MIN_HOURS:
                days.append(o)

# أودو يشغّل القاعدة مرة لكل سطر إدخال ABSENCE إذا الإيصال بيه أكثر من سطر (نفس كود القاعدة):
# الأيام تنخصم بأول مرة بس، والمرات الباقية تخصم مبلغ سطرها
abs_inputs = [line for line in payslip.input_line_ids if line.code == 'ABSENCE']
if len(abs_inputs) > 1 and inputs['ABSENCE'].id != abs_inputs[0].id:
    days = []

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
result_name = False   # أودو ما يصفّره بين المرات، فبدونه السطر الثاني ياخذ اسم الأيام
if days:
    result_name = 'خصم الغياب: %d يوم × 1.5 (%s)' % (
        len(days), '، '.join([to_date(o).strftime('%d/%m') for o in days]))
