# إجراء خادم (Server Action) على نموذج استقطاعات الراتب: hr.salary.attachment
# يسجّل تسديد سلفة نقداً (خارج الراتب) ويغلقها إذا تسددت بالكامل.
#
# الاستخدام الأساسي مع "سلفة الاسترجاع" (ADVRET): الموظف يجيب الفلوس قبل تاريخ الاستحقاق
# فنشغّل الإجراء حتى ما يبدأ الحجز من راتبه. يشتغل أيضاً على السلف المؤقتة إذا تسددت نقداً.
# السلف الدائمية ما عليها مبلغ متبقي فيرفضها.
#
# يدعم تسديد سلفة الدولار بالدينار: المبلغ المستلم يتحوّل لعملة السلفة بسعر الصرف.
#
# حقول اختيارية تُضاف بالستوديو على الاستقطاع (الحقل غير الموجود يُتجاهل):
#   PAYMENT_FIELD        رقم: المبلغ المستلم بعملة التسديد. فارغ = كامل المتبقي بعملة السلفة.
#   PAY_CURRENCY_FIELD   Many2one إلى res.currency: عملة التسديد. فارغ = نفس عملة السلفة.
#   RATE_FIELD           رقم: كم وحدة من عملة التسديد = وحدة واحدة من عملة السلفة
#                        (سلفة دولار تتسدد بالدينار = سعر الدولار بالدينار، مثلاً 1450).
#                        فارغ = سعر الصرف المسجل بأودو بتاريخ اليوم.
#   LOAN_CURRENCY_FIELD  Many2one إلى res.currency: عملة السلفة إذا تختلف عن عملة الشركة
#                        (رواتب الدولار مسجلة داخل شركة الدينار). فارغ = عملة الشركة.
#
# الإجراء يحدّث رصيد السلفة فقط؛ استلام النقد نفسه يُسجَّل بالحسابات (سند قبض بصندوق
# عملة التسديد) على حساب ذمم سلف الموظفين، وفرق سعر الصرف يُعالج محاسبياً.
#
# بيئة إجراءات الخادم (safe_eval) المتاح فيها: env, record, records, datetime, UserError, log.
# مكتوب لـ Odoo 17 فما فوق.

PAYMENT_FIELD = 'x_studio_cash_payment'
PAY_CURRENCY_FIELD = 'x_studio_cash_currency_id'
RATE_FIELD = 'x_studio_cash_rate'
LOAN_CURRENCY_FIELD = 'x_studio_loan_currency_id'


def optional(rec, field_name):
    return field_name in rec._fields and rec[field_name]


def fmt(amount, currency):
    return "{:,.2f} {}".format(amount, currency.name)


for attachment in records:
    if attachment.state != 'open':
        raise UserError("الاستقطاع \"%s\" غير جارٍ، لا يمكن تسجيل تسديد عليه." % attachment.display_name)
    remaining = attachment.remaining_amount
    if remaining <= 0:
        raise UserError(
            "الاستقطاع \"%s\" ما عليه مبلغ متبقي (الدائمي ما له مبلغ كلي)." % attachment.display_name
        )

    loan_currency = optional(attachment, LOAN_CURRENCY_FIELD) or attachment.currency_id
    pay_currency = optional(attachment, PAY_CURRENCY_FIELD) or loan_currency
    paid = optional(attachment, PAYMENT_FIELD)

    if pay_currency == loan_currency:
        rate = 1.0
        amount = paid or remaining
        note = "تسديد نقدي خارج الراتب: %s" % fmt(amount, loan_currency)
    else:
        if not paid:
            raise UserError(
                "اكتب المبلغ المستلم بـ %s للاستقطاع \"%s\"." % (pay_currency.name, attachment.display_name)
            )
        # سعر أودو بنفس صيغة الحقل: كم من عملة التسديد = 1 من عملة السلفة
        rate = optional(attachment, RATE_FIELD) or loan_currency._convert(
            1.0, pay_currency, attachment.company_id, datetime.date.today(), round=False
        )
        amount = loan_currency.round(paid / rate)
        note = "تسديد نقدي خارج الراتب: استُلم %s بسعر صرف %s = %s" % (
            fmt(paid, pay_currency), "{:,.2f}".format(rate), fmt(amount, loan_currency)
        )

    if amount <= 0 or loan_currency.compare_amounts(amount, remaining) > 0:
        raise UserError(
            "مبلغ التسديد لـ \"%s\" لازم يكون أكبر من صفر وما يتجاوز المتبقي: %s (= %s)."
            % (attachment.display_name, fmt(remaining, loan_currency), fmt(remaining * rate, pay_currency))
        )

    # record_payment هي نفس الدالة التي تستعملها القسائم عند التأكيد: تزيد المبلغ المسدد،
    # تكتب بالـ Chatter، وتغلق الاستقطاع تلقائياً إذا صار المتبقي صفراً
    attachment.record_payment(amount)
    attachment.message_post(body="%s، سجّله %s." % (note, env.user.name))

    # تصفير المبلغ والسعر حتى ما يتكرر التسديد بالخطأ (عملة التسديد تبقى للمرة الجاية)
    reset = {name: 0 for name in (PAYMENT_FIELD, RATE_FIELD) if name in attachment._fields}
    if reset:
        attachment.write(reset)
