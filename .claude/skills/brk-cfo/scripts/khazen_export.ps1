<#
.SYNOPSIS
    تصدير صامت للقراءة فقط من قاعدة بيانات الخازن إلى مجلد مؤرخ.

.DESCRIPTION
    يقرأ ملف الإعداد (khazen_config.json) وملف الاستعلامات المسماة (khazen_queries.sql)، وينفذ كل استعلام
    عبر حارس القراءة فقط داخل معاملة تُلغى دائماً (ROLLBACK)، ويكتب CSV لكل استعلام مع فهرس مجاميع وسجل.
    لا نوافذ ولا أسئلة، ولا كتابة إلا داخل مجلد التصدير. يعمل على Windows PowerShell 5.1 بدون أي تنصيب.

    ملف الاستعلامات: كل استعلام يبدأ بسطر اسم، واختيارياً سطر أعمدة المجاميع الضابطة:
        -- name: 13_ميزان_المراجعة
        -- control: Debit, Credit
        SELECT ... WHERE MoveDate BETWEEN '{FROM}' AND '{TO}'
    الرموز {FROM} و{TO} و{PREV_FROM} و{PREV_TO} تُستبدل بتواريخ السنة المالية الحالية والسابقة.
    لا تكتب تعليقات داخل نص الاستعلام: الحارس يفحص النص كله، وكلمة مثل update بتعليق ترفضه.

    رموز الخروج: 0 نجاح كامل، 1 استعلام أو أكثر رُفض أو فشل، 2 خطأ إعداد أو اتصال.

.EXAMPLE
    powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File khazen_export.ps1 -Config "C:\Users\surface\Documents\الماليه\CFO\أدوات\khazen_config.json"

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File khazen_export.ps1 -SelfTest
#>
param(
    [string]$Config,
    [switch]$SelfTest
)

$ErrorActionPreference = 'Stop'
$Invariant = [Globalization.CultureInfo]::InvariantCulture
# الثقافة الثابتة حتى ما تطلع الأرقام بفاصلة عشرية عربية أو التواريخ بصيغة الجهاز
[Threading.Thread]::CurrentThread.CurrentCulture = $Invariant
$Utf8Bom = New-Object System.Text.UTF8Encoding($true)

$AllowStart = '^\s*(SELECT|WITH)\b'
$Deny = '\b(INSERT|UPDATE|DELETE|MERGE|TRUNCATE|DROP|ALTER|CREATE|INTO|EXEC|EXECUTE|GRANT|REVOKE|DENY|BACKUP|RESTORE|DBCC|OPENROWSET|OPENQUERY|OPENDATASOURCE|SHUTDOWN|KILL|RECONFIGURE)\b|\bxp_|\bsp_|#'

function Test-ReadOnlySql([string]$Sql) {
    # يرجع سبب الرفض، أو نصاً فارغاً إذا الاستعلام قراءة فقط
    if ($Sql -notmatch $AllowStart) { return 'لا يبدأ بـ SELECT أو WITH' }
    if ($Sql -match $Deny) { return "كلمة ممنوعة: $($Matches[0])" }
    return ''
}

function Format-Cell($v) {
    if ($null -eq $v -or $v -is [DBNull]) { return '' }
    if ($v -is [datetime]) {
        if ($v.TimeOfDay.Ticks -eq 0) { return $v.ToString('yyyy-MM-dd', $Invariant) }
        return $v.ToString('yyyy-MM-dd HH:mm:ss', $Invariant)
    }
    if ($v -is [decimal] -or $v -is [double] -or $v -is [single]) { return $v.ToString('0.############', $Invariant) }
    if ($v -is [bool]) { if ($v) { return '1' } return '0' }
    if ($v -is [byte[]]) { return '' }
    return [string]$v
}

function ConvertTo-CsvField([string]$s) {
    if ($s -match '[",\r\n]') { return '"' + $s.Replace('"', '""') + '"' }
    return $s
}

function Write-TableCsv([System.Data.DataTable]$Table, [string]$Path, [string[]]$ControlColumns) {
    # يكتب CSV بترميز UTF-8 مع BOM ويرجع مجاميع الأعمدة الضابطة
    $sums = [ordered]@{}
    foreach ($c in $ControlColumns) {
        if (-not $Table.Columns.Contains($c)) { throw "عمود المجموع الضابط غير موجود بالنتيجة: $c" }
        $sums[$c] = [decimal]0
    }
    $n = $Table.Columns.Count
    $fields = New-Object string[] $n
    $w = New-Object System.IO.StreamWriter($Path, $false, $Utf8Bom)
    try {
        for ($i = 0; $i -lt $n; $i++) { $fields[$i] = ConvertTo-CsvField $Table.Columns[$i].ColumnName }
        $w.WriteLine(($fields -join ','))
        foreach ($row in $Table.Rows) {
            $items = $row.ItemArray
            for ($i = 0; $i -lt $n; $i++) { $fields[$i] = ConvertTo-CsvField (Format-Cell $items[$i]) }
            $w.WriteLine(($fields -join ','))
            foreach ($c in $ControlColumns) {
                if (-not ($row[$c] -is [DBNull])) { $sums[$c] += [decimal]$row[$c] }
            }
        }
    } finally { $w.Close() }
    return $sums
}

function Read-QueryFile([string]$Path) {
    $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    $queries = New-Object System.Collections.ArrayList
    $cur = $null
    foreach ($line in ($text -split "`r?`n")) {
        if ($line -match '^\s*--\s*name:\s*(.+?)\s*$') {
            $cur = [pscustomobject]@{ Name = ($Matches[1] -replace '[\\/:*?"<>|]', '_'); Control = @(); Sql = '' }
            [void]$queries.Add($cur)
        } elseif ($null -ne $cur -and $line -match '^\s*--\s*control:\s*(.*)$') {
            $cur.Control = @($Matches[1].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        } elseif ($null -ne $cur) {
            $cur.Sql += $line + "`n"
        }
    }
    foreach ($q in $queries) { $q.Sql = $q.Sql.Trim().TrimEnd(';').Trim() }
    return $queries
}

function Get-Periods([int]$StartMonth) {
    $today = (Get-Date).Date
    $year = if ($today.Month -ge $StartMonth) { $today.Year } else { $today.Year - 1 }
    $from = New-Object DateTime($year, $StartMonth, 1)
    return [ordered]@{ FROM = $from; TO = $today; PREV_FROM = $from.AddYears(-1); PREV_TO = $from.AddDays(-1) }
}

function Expand-Tokens([string]$Sql, $Periods) {
    foreach ($k in $Periods.Keys) { $Sql = $Sql.Replace('{' + $k + '}', $Periods[$k].ToString('yyyy-MM-dd', $Invariant)) }
    return $Sql
}

function Invoke-ReadOnlyQuery($Connection, [string]$Sql) {
    # القراءة داخل معاملة تُلغى دائماً: حتى لو تسرب أمر كتابة ما يُحفظ
    $tx = $Connection.BeginTransaction()
    try {
        $cmd = $Connection.CreateCommand()
        $cmd.Transaction = $tx
        $cmd.CommandText = $Sql
        $cmd.CommandTimeout = 900
        $ds = New-Object System.Data.DataSet
        $ds.EnforceConstraints = $false
        $dt = $ds.Tables.Add('result')
        $reader = $cmd.ExecuteReader()
        try { $dt.Load($reader) } finally { $reader.Close() }
        return , $dt
    } finally { $tx.Rollback() }
}

function Invoke-SelfTest {
    $fail = 0
    $cases = [ordered]@{
        'SELECT * FROM Accounts'                                        = $true
        'with t as (select AccNo, sum(Debit) d from Moves group by AccNo) select * from t' = $true
        'SELECT Update_Date, LastUpdate FROM Moves'                     = $true
        "SELECT N'حساب' AS [اسم] FROM Accounts"                         = $true
        'SELECT * INTO bak FROM Moves'                                  = $false
        'SELECT 1; DROP TABLE Moves'                                    = $false
        'UPDATE Moves SET Debit = 0'                                    = $false
        'EXEC sp_who'                                                   = $false
        'SELECT * FROM #tmp'                                            = $false
        'select * from Moves; delete from Moves'                        = $false
    }
    foreach ($sql in $cases.Keys) {
        $allowed = (Test-ReadOnlySql $sql) -eq ''
        if ($allowed -ne $cases[$sql]) { $fail++; [Console]::WriteLine("FAIL guard: $sql") }
    }
    $dt = New-Object System.Data.DataTable
    [void]$dt.Columns.Add('الحساب', [string]); [void]$dt.Columns.Add('Date', [datetime])
    [void]$dt.Columns.Add('Debit', [decimal]); [void]$dt.Columns.Add('Credit', [decimal])
    [void]$dt.Rows.Add('وكيل "النور", بغداد', (New-Object DateTime(2026, 9, 30)), [decimal]1500000.5, [DBNull]::Value)
    [void]$dt.Rows.Add('صندوق 1', (New-Object DateTime(2026, 10, 1, 14, 5, 0)), [decimal]0, [decimal]1500000.5)
    $tmp = [IO.Path]::GetTempFileName()
    try {
        $sums = Write-TableCsv $dt $tmp @('Debit', 'Credit')
        $lines = [IO.File]::ReadAllLines($tmp, [Text.Encoding]::UTF8)
        $expect = @('الحساب,Date,Debit,Credit',
                    '"وكيل ""النور"", بغداد",2026-09-30,1500000.5,',
                    'صندوق 1,2026-10-01 14:05:00,0,1500000.5')
        for ($i = 0; $i -lt 3; $i++) { if ($lines[$i] -ne $expect[$i]) { $fail++; [Console]::WriteLine("FAIL csv line $i : $($lines[$i])") } }
        if ($sums['Debit'] -ne [decimal]1500000.5 -or $sums['Credit'] -ne [decimal]1500000.5) { $fail++; [Console]::WriteLine('FAIL control sums') }
        $bytes = [IO.File]::ReadAllBytes($tmp)
        if (-not ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)) { $fail++; [Console]::WriteLine('FAIL BOM') }
    } finally { Remove-Item -LiteralPath $tmp -Force }
    $p = Get-Periods 1
    $sql = Expand-Tokens "SELECT * FROM M WHERE D BETWEEN '{FROM}' AND '{TO}' OR D BETWEEN '{PREV_FROM}' AND '{PREV_TO}'" $p
    if ($sql -match '\{') { $fail++; [Console]::WriteLine("FAIL tokens: $sql") }
    if ($p.PREV_TO.AddDays(1) -ne $p.FROM) { $fail++; [Console]::WriteLine('FAIL periods') }
    if ($fail) { [Console]::WriteLine("SELFTEST: FAIL ($fail)"); return 1 }
    [Console]::WriteLine('SELFTEST: PASS')
    return 0
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# ---------- التشغيل ----------
try {
    if (-not $Config) { throw 'المعامل -Config مطلوب' }
    $cfg = [IO.File]::ReadAllText($Config, [Text.Encoding]::UTF8) | ConvertFrom-Json
    if ($cfg.connection -match '(?i)(password|pwd)\s*=') {
        throw 'نص الاتصال فيه كلمة مرور: التشغيل الصامت يكون بحساب ويندوز فقط (Integrated Security=True)'
    }
    $provider = if ($cfg.provider) { [string]$cfg.provider } else { 'sqlserver' }
    $startMonth = if ($cfg.fiscal_year_start_month) { [int]$cfg.fiscal_year_start_month } else { 1 }
    $periods = Get-Periods $startMonth
    $queries = Read-QueryFile $cfg.queries
    if ($queries.Count -eq 0) { throw "ما لكيت استعلامات مسماة بالملف $($cfg.queries)" }
    $outDir = Join-Path $cfg.output_root ('الخازن_تصدير_' + (Get-Date).ToString('yyyy-MM-dd_HHmm', $Invariant))
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
} catch {
    [Console]::Error.WriteLine("خطأ إعداد: $($_.Exception.Message)")
    exit 2
}

$LogPath = Join-Path $outDir 'run.log'
function Write-Log([string]$Message) {
    $line = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $Invariant) + '  ' + $Message + [Environment]::NewLine
    [IO.File]::AppendAllText($LogPath, $line, $Utf8Bom)
}

$started = Get-Date
$copyDir = $null
$results = New-Object System.Collections.ArrayList
$exitCode = 0
try {
    if ($provider -eq 'access') {
        # Access: نسخة مؤقتة تُفتح بوضع القراءة، والأصل ما يُلمس
        $copyDir = Join-Path $outDir '_نسخة_القاعدة'
        New-Item -ItemType Directory -Path $copyDir -Force | Out-Null
        $copy = Join-Path $copyDir (Split-Path $cfg.access_file -Leaf)
        Copy-Item -LiteralPath $cfg.access_file -Destination $copy
        $conn = New-Object System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$copy;Mode=Read;")
        $source = "Access: $(Split-Path $cfg.access_file -Leaf) (نسخة مؤقتة)"
    } else {
        $conn = New-Object System.Data.SqlClient.SqlConnection([string]$cfg.connection)
        $b = New-Object System.Data.Common.DbConnectionStringBuilder
        $b.ConnectionString = [string]$cfg.connection
        $server = @('Server', 'Data Source') | Where-Object { $b.ContainsKey($_) } | ForEach-Object { $b[$_] } | Select-Object -First 1
        $db = @('Database', 'Initial Catalog') | Where-Object { $b.ContainsKey($_) } | ForEach-Object { $b[$_] } | Select-Object -First 1
        $source = "SQL Server: $server / $db"
    }
    $conn.Open()
    Write-Log "اتصال: $source | الاستعلامات: $($queries.Count)"
} catch {
    Write-Log "خطأ اتصال: $($_.Exception.Message)"
    [Console]::Error.WriteLine("خطأ اتصال: $($_.Exception.Message)")
    if ($copyDir) { Remove-Item -LiteralPath $copyDir -Recurse -Force -ErrorAction SilentlyContinue }
    exit 2
}

try {
    foreach ($q in $queries) {
        $t0 = Get-Date
        $sql = Expand-Tokens $q.Sql $periods
        $reason = Test-ReadOnlySql $sql
        if ($reason) {
            Write-Log "مرفوض $($q.Name): $reason"
            [void]$results.Add([pscustomobject]@{ Name = $q.Name; Rows = ''; Sums = ''; Status = "مرفوض: $reason"; Seconds = 0 })
            $exitCode = 1
            continue
        }
        try {
            $dt = Invoke-ReadOnlyQuery $conn $sql
            $sums = Write-TableCsv $dt (Join-Path $outDir ($q.Name + '.csv')) $q.Control
            $sumText = ($sums.Keys | ForEach-Object { '{0} = {1}' -f $_, $sums[$_].ToString('#,0.##', $Invariant) }) -join '، '
            $sec = [math]::Round(((Get-Date) - $t0).TotalSeconds, 1)
            Write-Log "تم $($q.Name): $($dt.Rows.Count) سطر ($sec ث)"
            [void]$results.Add([pscustomobject]@{ Name = $q.Name + '.csv'; Rows = $dt.Rows.Count; Sums = $sumText; Status = 'تم'; Seconds = $sec })
        } catch {
            Write-Log "فشل $($q.Name): $($_.Exception.Message)"
            [void]$results.Add([pscustomobject]@{ Name = $q.Name; Rows = ''; Sums = ''; Status = "فشل: $($_.Exception.Message)"; Seconds = 0 })
            $exitCode = 1
        }
    }
} finally {
    $conn.Close()
    if ($copyDir) { Remove-Item -LiteralPath $copyDir -Recurse -Force -ErrorAction SilentlyContinue }
}

$finished = Get-Date
$md = New-Object System.Text.StringBuilder
[void]$md.AppendLine('# فهرس تصدير الخازن (تشغيل صامت، قراءة فقط)')
[void]$md.AppendLine()
[void]$md.AppendLine("- البداية: $($started.ToString('yyyy-MM-dd HH:mm:ss', $Invariant))، والنهاية: $($finished.ToString('yyyy-MM-dd HH:mm:ss', $Invariant))")
[void]$md.AppendLine("- المصدر: $source")
[void]$md.AppendLine("- السنة الحالية: $($periods.FROM.ToString('yyyy-MM-dd')) إلى $($periods.TO.ToString('yyyy-MM-dd'))، والسابقة: $($periods.PREV_FROM.ToString('yyyy-MM-dd')) إلى $($periods.PREV_TO.ToString('yyyy-MM-dd'))")
[void]$md.AppendLine('- كل استعلام فُحص بحارس القراءة فقط، ونُفذ داخل معاملة أُلغيت بعد القراءة. لم يُكتب شي على القاعدة.')
[void]$md.AppendLine()
[void]$md.AppendLine('| الملف | الأسطر | المجاميع الضابطة | الحالة | المدة (ث) |')
[void]$md.AppendLine('|---|---:|---|---|---:|')
foreach ($r in $results) { [void]$md.AppendLine("| $($r.Name) | $($r.Rows) | $($r.Sums) | $($r.Status) | $($r.Seconds) |") }
[IO.File]::WriteAllText((Join-Path $outDir '00_فهرس.md'), $md.ToString(), $Utf8Bom)
Write-Log "انتهى برمز $exitCode"
Write-Output "$outDir | رمز الخروج $exitCode"
exit $exitCode
