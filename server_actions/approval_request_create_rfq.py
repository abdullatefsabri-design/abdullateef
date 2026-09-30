# إجراء خادم (Server Action) على نموذج طلب الموافقة: approval.request
# ينشئ طلب عرض سعر (RFQ) مسودة من أسطر المنتجات ويسنده إلى مجهز المشتريات.
#
# بيئة إجراءات الخادم (safe_eval) لا توفر `fields` ولا `Markup`،
# المتاح فيها: env, record, records, datetime, UserError, log.
# الكود مكتوب لـ Odoo 16 فما فوق (بسبب _get_html_link).

BUYER_NAME = 'محمود'  # الأفضل البحث بالـ login أو بمعرف ثابت بدل الاسم

# التحقق من وجود أسطر منتجات في طلب الموافقة
if record.product_line_ids:
    origin = "طلب موافقة رقم: %s" % record.name

    # منع التكرار: إذا شُغّل الإجراء مرة أخرى على نفس الطلب لا ننشئ أمر شراء ثانياً
    if not env['purchase.order'].search_count([('origin', '=', origin)]):

        # أسطر أمر الشراء تتطلب منتجاً، فلا نسمح بأسطر نصية فقط
        lines_without_product = record.product_line_ids.filtered(lambda l: not l.product_id)
        if lines_without_product:
            raise UserError(
                "لا يمكن إنشاء طلب الشراء، الأسطر التالية بدون منتج: %s"
                % "، ".join(lines_without_product.mapped('description'))
            )

        # المورد حقل إلزامي في أمر الشراء: جهة الاتصال في الطلب، وإلا أول مورد لأول منتج
        vendor = record.partner_id or record.product_line_ids.product_id.seller_ids[:1].partner_id
        if not vendor:
            raise UserError(
                "لا يمكن إنشاء طلب الشراء: حدّد جهة الاتصال (المورد) في طلب الموافقة "
                "أو أضف مورداً في تبويب المشتريات للمنتج."
            )

        # 1. جلب معرف المستخدم لمجهز المشتريات (محمود) - مستخدمون داخليون فقط
        buyer_user = env['res.users'].search(
            [('name', 'ilike', BUYER_NAME), ('share', '=', False)], order='id', limit=1
        )

        # 2. إنشاء أمر الشراء مسودة (RFQ)
        now = datetime.datetime.now()
        order_lines = []
        for line in record.product_line_ids:
            order_lines.append((0, 0, {
                'product_id': line.product_id.id,
                'name': line.description or line.product_id.display_name,
                'product_qty': line.quantity,
                'product_uom': line.product_uom_id.id or line.product_id.uom_po_id.id,
                'price_unit': 0.0,
                'date_planned': now,
            }))

        po = env['purchase.order'].create({
            'partner_id': vendor.id,
            'origin': origin,
            'user_id': buyer_user.id or env.user.id,
            'order_line': order_lines,
        })

        # 3. ربط الطلبين في الـ Chatter
        # message_post يهرّب النص العادي (Odoo 17+)، أما _get_html_link() فيعيد Markup
        # ودمج النص معه بـ + يبقي النتيجة Markup فيظهر الرابط قابلاً للنقر
        record.message_post(
            body="تم إنشاء طلب شراء جديد بنجاح برقم: " + po._get_html_link()
            + " وتم إسناده إلى المشتريات."
        )
        po.message_post(body="تم إنشاؤه من " + record._get_html_link())
