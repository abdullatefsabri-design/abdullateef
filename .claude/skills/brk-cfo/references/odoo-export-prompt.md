# برومبت التصدير من Odoo عبر Claude Cowork

يُستخدم عندما يتعذر الربط المباشر بـ API. المستخدم يلصق البرومبت أدناه بـ Claude Cowork على جهازه، فيصدّر Cowork البيانات من Odoo المفتوح بـ Chrome إلى مجلد واحد. بعدها:
- **على جهاز المستخدم:** السكل يقرأ المجلد مباشرة من `Documents\الماليه\CFO\بيانات\Odoo_تصدير_<التاريخ>\`.
- **بالكلاود:** المستخدم يضغط المجلد ويرفعه بالمحادثة.

**عند استلام المجلد:** اقرأ `00_فهرس.md` أولاً، وطابق المجاميع الضابطة المسجلة فيه مع مجاميع الملفات قبل أي تحليل، وحدّث `company.md` بالحقائق الجديدة مع تاريخها.

---

```text
أنت تساعد الإدارة المالية لشركة البركة للاستثمارات الصناعية. المطلوب تصدير بيانات من Odoo
المفتوح بمتصفح Chrome على الرابط https://brk.ejaferp.com/odoo/accounting وحفظها بمجلد واحد،
حتى يحللها المدير المالي.

== قواعد إلزامية (قراءة فقط) ==
1. لا تنشئ ولا تعدّل ولا تحذف ولا ترحّل أي شي. ممنوع الضغط على: Create/جديد، Edit/تعديل، Save/حفظ،
   Delete/حذف، Archive/أرشفة، Duplicate/نسخ، Confirm/تأكيد، Post/ترحيل، Validate/اعتماد،
   Register Payment/تسجيل دفعة، Reconcile/تسوية، Reset to Draft/إرجاع لمسودة، Send/إرسال، Cancel/إلغاء.
2. المسموح فقط: القوائم، والبحث، والفلاتر، والتجميع (Group By)، وفتح التقارير، و Export/تصدير،
   وزر تحميل XLSX أو PDF بالتقارير.
3. لا تغيّر أي إعداد، ولا الشركة المختارة، ولا لغة الواجهة. إذا ظهرت نافذة تطلب تأكيد عملية، اضغط إلغاء.
4. إذا الشاشة فيها أكثر من شركة، صدّر لشركة البركة فقط واكتب اسم الشركة بالفهرس.
5. إذا قائمة أو تطبيق غير موجود (مثل التصنيع أو الرواتب)، لا تحاول تثبيته. سجّل "غير مثبت" بالفهرس وكمّل.
6. لا تحفظ كلمات مرور ولا مفاتيح، ولا تصدّر بيانات شخصية للموظفين (رواتب فردية، هويات، حسابات بنكية شخصية).
7. إذا التصدير كبير جداً أو بطيء، قسّمه حسب الفترة (فصلياً) بنفس اسم الملف مع اللاحقة _Q1 وغيرها.

== المجلد والأسماء ==
- أنشئ المجلد: C:\Users\surface\Documents\الماليه\CFO\بيانات\Odoo_تصدير_<تاريخ اليوم YYYY-MM-DD>\
- صيغة الملفات XLSX، والتقارير اللي ما تدعم XLSX تكون PDF.
- الفترات:
  - "السنة الحالية" = من بداية السنة المالية الحالية لحد اليوم.
  - "السنة السابقة" = السنة المالية السابقة كاملة.
  - تأكد من بداية السنة المالية من إعدادات المحاسبة (قراءة فقط). إذا ما تبين، اعتبرها 1 كانون الثاني.

== المطلوب (بالترتيب) ==
أ) معلومات النظام
  01_معلومات_النظام.md: نسخة Odoo (من Settings ← About أو قائمة ؟)، والتطبيقات المثبتة، واسم الشركة،
  والعملة الأساسية، وبداية السنة المالية، واللغة.

ب) التقارير المحاسبية (Accounting ← Reporting / التقارير)
  02_ميزان_المراجعة_السنة_الحالية.xlsx      Trial Balance، قيود مرحّلة فقط
  03_ميزان_المراجعة_السنة_السابقة.xlsx
  04_الميزانية_اليوم.xlsx                   Balance Sheet بتاريخ اليوم، مع عمود مقارنة نهاية السنة السابقة
  05_الأرباح_والخسائر_شهري.xlsx            Profit and Loss للسنة الحالية، مقارنة شهرية (عمود لكل شهر)
  06_الأرباح_والخسائر_السنة_السابقة.xlsx
  07_أعمار_الذمم_المدينة.xlsx               Aged Receivable بتاريخ اليوم، لكل زبون/وكيل
  08_أعمار_الذمم_الدائنة.xlsx               Aged Payable بتاريخ اليوم
  09_دفتر_الشركاء_المدينين.xlsx             Partner Ledger للذمم المدينة، السنة الحالية
  10_دفتر_الأستاذ_النقد_والمصارف.xlsx       General Ledger لحسابات النقد والمصارف فقط، السنة الحالية
  11_التدفق_النقدي.xlsx                     Cash Flow Statement إذا موجود (وإلا سجّل "غير متوفر")

ج) القوائم (Export مع الحقول المذكورة؛ إذا حقل غير موجود تجاوزه واذكره بالفهرس)
  12_قيود_اليومية_السنة_الحالية.xlsx
      Accounting ← Journal Items، الفلتر Posted، والفترة السنة الحالية.
      الحقول: Date، Journal Entry، Journal، Reference، Account، Partner، Label،
      Analytic Distribution، Debit، Credit، Balance، Currency، Amount in Currency،
      Matching Number، Due Date.
  13_دليل_الحسابات.xlsx
      Chart of Accounts: Code، Account Name، Type، Allow Reconciliation، Currency.
  14_الدفاتر_والصناديق.xlsx
      Journals: Name، Type، Short Code، Currency، Default Account.
  15_الحسابات_التحليلية.xlsx
      Analytic Accounts: Name، Plan، Reference، Balance.
  16_فواتير_الزبائن.xlsx
      Customer Invoices + Credit Notes للسنة الحالية: Number، Customer، Invoice Date، Due Date،
      Total، Amount Due، Payment Status، Salesperson، Currency.
  17_فواتير_الموردين.xlsx
      Vendor Bills + Refunds للسنة الحالية: Number، Vendor، Bill Reference، Bill Date، Due Date،
      Total، Amount Due، Payment Status، Currency.
  18_الدفعات.xlsx
      Payments للسنة الحالية: Date، Number، Journal، Partner، Payment Type، Amount، Currency، Status.
  19_الوكلاء_والزبائن.xlsx
      Contacts/Customers (الشركات فقط): Name، City، Phone، Salesperson، Payment Terms،
      Credit Limit (إذا موجود)، Total Receivable، Tags.
  20_أسعار_العملات.xlsx
      Accounting ← Configuration ← Currencies ← USD ← Rates: كل الأسعار المسجلة مع تواريخها.

د) المخزون والمشتريات والمبيعات والتصنيع (إذا مثبتة)
  21_تقييم_المخزون.xlsx       Inventory ← Reporting ← Valuation/Stock: Product، Quantity، Unit Cost، Total Value، Location.
  22_المنتجات.xlsx            Products: Internal Reference، Name، Category، Unit، Cost، Sales Price، Type، On Hand، Weight.
  23_أوامر_الشراء.xlsx        Purchase Orders للسنة الحالية: Reference، Vendor، Order Date، Total، Currency، Status،
                              Billing Status، Receipt Status.
  24_أوامر_البيع.xlsx         Sales Orders للسنة الحالية: Reference، Customer، Order Date، Total، Status، Invoice Status.
  25_أوامر_التصنيع.xlsx       Manufacturing Orders للسنة الحالية: Reference، Product، Quantity Produced، Date،
                              Bill of Materials، State.
  26_قوائم_المواد.xlsx        Bills of Materials مع المكونات: Product، Component، Quantity، Unit.
  27_طلبات_الموافقة.xlsx      Approvals: Name، Category، Requested By، Date، Status، Amount.

هـ) الرواتب (إذا مثبتة) — مجاميع فقط
  28_مجاميع_الرواتب_الشهرية.xlsx  لكل شهر بالسنة الحالية: عدد الموظفين، وإجمالي الرواتب، والاستقطاعات،
                                    والصافي، وحصة صاحب العمل بالضمان. بدون أسماء أو رواتب فردية.

== الفهرس (إلزامي) ==
أنشئ 00_فهرس.md بالمجلد، وبيه جدول بالأعمدة:
اسم الملف | القائمة بـ Odoo | الفلاتر والفترة | عدد الأسطر | المجموع الضابط كما ظاهر بالشاشة | ملاحظات

المجموع الضابط أمثلة:
- ميزان المراجعة: إجمالي المدين وإجمالي الدائن.
- أعمار الذمم: الإجمالي الكلي.
- قيود اليومية: مجموع المدين والدائن.
- الفواتير: مجموع Amount Due.
- تقييم المخزون: إجمالي القيمة.

انسخ الرقم من شاشة Odoo نفسها، مو من الملف المصدّر، حتى نطابق الاثنين.

بآخر الفهرس اكتب:
- أي ملف ما تصدّر وليش.
- أي تحذير أو رسالة ظهرت.
- وقت بداية ونهاية التصدير.

لا تحلل البيانات ولا تلخصها. المطلوب تصدير نظيف وكامل فقط.
```
