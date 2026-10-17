سكل المدير المالي للبركة (brk-cfo)، نسخة 2026-10-17
==================================================

بالمجلد:
  brk-cfo\         السكل نفسه (SKILL.md والمراجع والسكربتات)
  install.cmd      التثبيت لـ Claude Code بدبل كلك
  install.ps1      سكربت التثبيت (يشغله install.cmd)
  brk-cfo.skill    نفس السكل بملف واحد، للرفع على تطبيق Claude أو claude.ai

الطريقة 1: Claude Code على ويندوز
  1. فك الضغط كامل (Extract All)، مو تشغيل من داخل الـ zip.
  2. دبل كلك على install.cmd.
  3. سكّر Claude Code وافتحه من جديد، واكتب: brk-cfo شنو تكدر تسوي؟

  شنو يسوي التثبيت:
  - ينسخ السكل إلى %USERPROFILE%\.claude\skills\brk-cfo
  - إذا اكو نسخة قديمة، ينقلها إلى %USERPROFILE%\.claude\skills-backup\ ويطبع أي ملف بيها ماكو بالجديدة.
  - إذا مجلد المستندات بهذا الجهاز غير C:\Users\surface\Documents، يصحح المسارات داخل السكل.
  - ينشئ مجلد العمل Documents\الماليه\CFO وبيه أدوات وبيانات ومخرجات إذا ما موجود.
  - ينسخ odoo_fetch.py وkhazen_export.ps1 إلى أدوات إذا ما موجودة. إذا موجودة ومختلفة، يحفظ الجديد باسم "(محدث)" وما يمس القديم.
  - يفحص Python وopenpyxl ويكلك شنو الناقص. ما يحذف ولا يكتب فوق أي ملف بالماليه.

الطريقة 2: تطبيق Claude أو claude.ai أو Cowork
  الإعدادات، ثم Capabilities (أو Customize)، ثم Skills، ثم Upload skill، واختار brk-cfo.skill.
  إذا ما قبل الامتداد، غيّر اسمه إلى brk-cfo.zip وارفعه.

بعد التثبيت على جهاز جديد
  - Python 3 من python.org (أشّر Add python.exe to PATH)، وبعدها: py -m pip install openpyxl
  - ربط Odoo المباشر اختياري ويحتاج متغيرات ODOO_DB وODOO_USER وODOO_API_KEY لكل جهاز:
    الخطوات بـ brk-cfo\references\odoo-live.md. المفتاح ينكتب بالجهاز فقط، لا بالمحادثة.
  - تصدير الخازن الصامت يحتاج khazen_config.json وkhazen_queries.sql بمجلد أدوات لكل جهاز
    (brk-cfo\references\khazen-export-prompt.md، المرحلة 3).

سرية
  references\company.md بيه أرقام الشركة المجمعة (أرصدة وكلاء ونقد ورواتب). الملف للاستخدام الداخلي فقط،
  لا يندز برا الشركة.

التحديث
  أي نسخة جديدة: نفس الخطوات. النسخة القديمة تنحفظ تلقائياً بـ skills-backup.
