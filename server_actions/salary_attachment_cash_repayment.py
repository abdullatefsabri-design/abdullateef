# إجراء خادم (Server Action) على نموذج استقطاعات الراتب: hr.salary.attachment
# يسجّل تسديد سلفة نقداً (خارج الراتب) ويغلقها إذا تسددت بالكامل.
#
# الاستخدام الأساسي مع "سلفة الاسترجاع" (ADVRET): الموظف يجيب الفلوس قبل تاريخ الاستحقاق
# فنشغّل الإجراء حتى ما يبدأ الحجز من راتبه. يشتغل أيضاً على السلف المؤقتة إذا تسددت نقداً.
# السلف الدائمية ما عليها مبلغ متبقي فيرفضها.
#
# المبلغ: إذا أُضيف بالستوديو حقل رقمي باسم PAYMENT_FIELD يُسجَّل المبلغ المكتوب فيه
# (تسديد جزئي)، وإذا الحقل غير موجود أو فارغ يُسجَّل كامل المتبقي.
#
# الإجراء يحدّث رصيد السلفة فقط؛ استلام النقد نفسه يُسجَّل بالحسابات (سند قبض بصندوق
# نفس عملة السلفة، دينار أو دولار) على حساب ذمم سلف الموظفين.
# المبالغ تُكتب بدون اسم عملة لأن نوع الاستقطاع (دينار/دولار) هو اللي يحدد العملة،
# وعملة الاستقطاع بأودو هي عملة الشركة وقد لا تطابق عملة سلفة الدولار.
#
# بيئة إجراءات الخادم (safe_eval) المتاح فيها: env, record, records, datetime, UserError, log.
# مكتوب لـ Odoo 17 فما فوق.

PAYMENT_FIELD = 'x_studio_cash_payment'  # حقل الستوديو للتسديد الجزئي (اختياري)

for attachment in records:
    if attachment.state != 'open':
        raise UserError("الاستقطاع \"%s\" غير جارٍ، لا يمكن تسجيل تسديد عليه." % attachment.display_name)
    if attachment.remaining_amount <= 0:
        raise UserError(
            "الاستقطاع \"%s\" ما عليه مبلغ متبقي (الدائمي ما له مبلغ كلي)." % attachment.display_name
        )

    has_field = PAYMENT_FIELD in attachment._fields
    amount = (has_field and attachment[PAYMENT_FIELD]) or attachment.remaining_amount
    if amount <= 0 or attachment.currency_id.compare_amounts(amount, attachment.remaining_amount) > 0:
        raise UserError(
            "مبلغ التسديد لـ \"%s\" لازم يكون أكبر من صفر وما يتجاوز المتبقي (%s)."
            % (attachment.display_name, attachment.remaining_amount)
        )

    # record_payment هي نفس الدالة التي تستعملها القسائم عند التأكيد: تزيد المبلغ المسدد،
    # تكتب بالـ Chatter، وتغلق الاستقطاع تلقائياً إذا صار المتبقي صفراً
    attachment.record_payment(amount)
    attachment.message_post(
        body="التسديد أعلاه نقدي خارج الراتب (%s)، سجّله %s." % (amount, env.user.name)
    )
    if has_field:
        attachment.write({PAYMENT_FIELD: 0})
