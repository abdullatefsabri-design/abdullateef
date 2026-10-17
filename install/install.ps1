<#
  تثبيت سكل المدير المالي للبركة (brk-cfo) لـ Claude Code على هذا الجهاز.

  1. ينسخ مجلد brk-cfo إلى %USERPROFILE%\.claude\skills\brk-cfo
  2. إذا موجودة نسخة قديمة، ينقلها إلى %USERPROFILE%\.claude\skills-backup\brk-cfo_<التاريخ>
     (برا مجلد skills حتى ما تنحمل مرتين)، ويطبع الملفات اللي بالقديمة وماكو بالجديدة.
  3. إذا مجلد المستندات بهذا الجهاز غير C:\Users\surface\Documents، يصحح المسارات داخل السكل.
  4. ينشئ مجلد العمل CFO (أدوات، بيانات، مخرجات) إذا ما موجود، وما يحذف أو يكتب فوق أي ملف.
  5. ينسخ odoo_fetch.py وkhazen_export.ps1 إلى أدوات فقط إذا ما موجودة. إذا موجودة ومختلفة،
     يحفظ الجديدة باسم "(محدث)" ويخلي القديمة على حالها.
  6. يفحص Python وopenpyxl ويطبع الناقص.

  الاستخدام: دبل كلك على install.cmd، أو:
    powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 [-Workspace <مسار CFO>] [-SkillsRoot <مسار skills>]
#>
param(
    [string]$Workspace = '',
    [string]$SkillsRoot = ''
)
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$source = Join-Path $here 'brk-cfo'
$userHome = if ($env:USERPROFILE) { $env:USERPROFILE } else { $HOME }
if (-not $SkillsRoot) { $SkillsRoot = Join-Path (Join-Path $userHome '.claude') 'skills' }
if (-not $Workspace) {
    # المستندات ممكن تكون محولة لـ OneDrive، فنفضّل المكان اللي بيه مجلد الماليه فعلاً
    $docsList = @((Join-Path $userHome 'Documents'), [Environment]::GetFolderPath('MyDocuments')) | Select-Object -Unique
    $docs = $docsList | Where-Object { Test-Path (Join-Path $_ 'الماليه') } | Select-Object -First 1
    if (-not $docs) { $docs = $docsList | Where-Object { Test-Path $_ } | Select-Object -First 1 }
    $Workspace = Join-Path (Join-Path $docs 'الماليه') 'CFO'
}
$target = Join-Path $SkillsRoot 'brk-cfo'
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$utf8Bom = New-Object System.Text.UTF8Encoding($true)

function Say([string]$msg) { [Console]::WriteLine($msg) }

if (-not (Test-Path (Join-Path $source 'SKILL.md'))) {
    Say "خطأ: ما لگيت brk-cfo\SKILL.md جنب السكربت. فك الضغط كامل وشغّله من نفس المجلد."
    exit 1
}

# 1-2. النسخة القديمة
if (Test-Path $target) {
    $backupRoot = Join-Path (Split-Path -Parent $SkillsRoot) 'skills-backup'
    New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null
    $backup = Join-Path $backupRoot "brk-cfo_$stamp"
    $oldFiles = @(Get-ChildItem -Recurse -File $target | ForEach-Object { $_.FullName.Substring($target.Length + 1) })
    $newFiles = @(Get-ChildItem -Recurse -File $source | ForEach-Object { $_.FullName.Substring($source.Length + 1) })
    Move-Item -Path $target -Destination $backup
    Say "النسخة القديمة انحفظت بـ: $backup"
    $onlyOld = @($oldFiles | Where-Object { ($newFiles -notcontains $_) -and ($_ -notmatch '__pycache__') })
    if ($onlyOld.Count -gt 0) {
        Say "تنبيه: هذي الملفات چانت بالنسخة القديمة وماكو بالجديدة (موجودة بالنسخة الاحتياطية):"
        $onlyOld | ForEach-Object { Say "   $_" }
    }
}
New-Item -ItemType Directory -Force -Path $SkillsRoot | Out-Null
Copy-Item -Recurse -Path $source -Destination $target
Say "السكل انثبت بـ: $target"

# 3. تصحيح المسارات إذا المستندات بمكان ثاني
$oldDocs = 'C:\Users\surface\Documents'
$newDocs = Split-Path -Parent (Split-Path -Parent $Workspace)
if ($newDocs -and ($newDocs -ne $oldDocs)) {
    $changed = 0
    Get-ChildItem -Recurse -File $target -Include *.md, *.py, *.ps1 | ForEach-Object {
        $text = [System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::UTF8)
        if ($text.Contains($oldDocs)) {
            $enc = if ($_.Extension -eq '.ps1') { $utf8Bom } else { $utf8 }
            [System.IO.File]::WriteAllText($_.FullName, $text.Replace($oldDocs, $newDocs), $enc)
            $changed++
        }
    }
    Say "المسارات داخل السكل تصححت إلى $newDocs ($changed ملف)."
}

# 4. مجلد العمل
foreach ($sub in @('أدوات', 'بيانات', 'مخرجات')) {
    $p = Join-Path $Workspace $sub
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Force -Path $p | Out-Null; Say "انشأ: $p" }
}

# 5. الأدوات
$tools = Join-Path $Workspace 'أدوات'
foreach ($name in @('odoo_fetch.py', 'khazen_export.ps1')) {
    $src = Join-Path (Join-Path $target 'scripts') $name
    $dst = Join-Path $tools $name
    if (-not (Test-Path $dst)) {
        Copy-Item $src $dst
        Say "انسخ إلى أدوات: $name"
    } elseif ((Get-FileHash $src).Hash -ne (Get-FileHash $dst).Hash) {
        $upd = Join-Path $tools ([IO.Path]::GetFileNameWithoutExtension($name) + ' (محدث)' + [IO.Path]::GetExtension($name))
        Copy-Item -Force $src $upd
        Say "الموجود بـ أدوات\$name ما تغير، والجديد انحفظ باسم: $(Split-Path -Leaf $upd)"
    }
}

# 6. Python وopenpyxl
$py = $null
foreach ($c in @('py', 'python', 'python3')) {
    $cmd = Get-Command $c -ErrorAction SilentlyContinue
    if ($cmd -and ($cmd.Source -notmatch 'WindowsApps')) { $py = $cmd.Source; break }
}
if (-not $py) {
    Say "ناقص: Python 3 (من python.org، وأشّر Add python.exe to PATH). السكربتات تحتاجه."
} else {
    & $py -c "import openpyxl" 2>$null
    if ($LASTEXITCODE -eq 0) { Say "Python وopenpyxl موجودين: $py" }
    else { Say "ناقص: openpyxl. شغّل: `"$py`" -m pip install openpyxl" }
}

Say ""
Say "خلص التثبيت. افتح Claude Code من جديد حتى يحمّل السكل."
Say "ربط Odoo المباشر (اختياري) يحتاج متغيرات ODOO_DB وODOO_USER وODOO_API_KEY على هذا الجهاز: راجع references\odoo-live.md."
