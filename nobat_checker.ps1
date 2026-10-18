<#
=====================================================================
  nobat_checker.ps1  |  نسخه 1.0.0
  بررسی اولین نوبت آزاد و ظرفیت باقیمانده در سامانه نوبت‌دهی اینترنتی
  دانشگاه علوم پزشکی تهران (nobat.tums.ac.ir — موتور بایوویزیت)
---------------------------------------------------------------------
  بدون نیاز به نصب هیچ ماژول یا نرم‌افزار اضافه —
  فقط با PowerShell داخلی ویندوز (نسخه 5.1 به بالا) اجرا می‌شود.
---------------------------------------------------------------------
  نمونه‌های اجرا:
      .\nobat_checker.ps1 -Account 10460
      .\nobat_checker.ps1 -Account 10460 -Watch -Interval 300
      .\nobat_checker.ps1 -DoctorsFile doctors.txt -Watch
=====================================================================
#>
[CmdletBinding()]
param(
    [string[]]$Account,          # شناسه یا لینک کامل صفحه پزشک (مثال: 10460)
    [string]$DoctorsFile,        # مسیر فایل متنی حاوی فهرست پزشکان
    [switch]$Watch,              # حالت پایش مداوم
    [int]$Interval = 300,        # فاصله زمانی بررسی در حالت پایش (ثانیه)
    [int]$Days = 60,             # افق جستجوی نوبت (روز از امروز)
    [string]$BaseUrl = 'https://nobat.tums.ac.ir',   # نشانی سامانه
    [switch]$NoBrowser,          # باز نکردن مرورگر هنگام پیدا شدن نوبت
    [switch]$NoToast,            # غیرفعال کردن اعلان ویندوز
    [switch]$NoBeep,             # غیرفعال کردن صدای زنگ
    [string]$LogFile             # مسیر فایل گزارش (اختیاری)
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor `
        [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
} catch {}

$Script:UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36'
$Script:FoundAny = $false

# =======================  ابزارهای کمکی  =======================

function Write-Log {
    param([string]$Message)
    if ($LogFile) {
        try {
            $line = '{0} | {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
            Add-Content -Path $LogFile -Value $line -Encoding UTF8
        } catch {}
    }
}

function ConvertTo-Bool {
    param($Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return $Value }
    return ($Value.ToString().ToLower() -eq 'true')
}

function Get-Int {
    param($Value)
    if ($null -eq $Value) { return 0 }
    if ($Value -is [int]) { return $Value }
    $n = 0
    if ([int]::TryParse($Value.ToString(), [ref]$n)) { return $n }
    return 0
}

function Escape-Xml {
    param([string]$Text)
    if (-not $Text) { return '' }
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

# یک درخواست HTTP ساده؛ متن پاسخ را همیشه به‌صورت UTF-8 برمی‌گرداند
function Invoke-NobatRequest {
    param(
        [string]$Url,
        [object]$Session,
        [hashtable]$ExtraHeaders
    )
    $headers = @{ 'Accept-Language' = 'fa-IR,fa;q=0.9,en;q=0.8' }
    if ($ExtraHeaders) {
        foreach ($k in $ExtraHeaders.Keys) { $headers[$k] = $ExtraHeaders[$k] }
    }
    $p = @{
        Uri             = $Url
        UserAgent       = $Script:UA
        UseBasicParsing = $true
        TimeoutSec      = 45
        Headers         = $headers
    }
    if ($Session) { $p['WebSession'] = $Session }
    else          { $p['SessionVariable'] = 'nbSess' }

    $resp = Invoke-WebRequest @p
    $text = $null
    if ($resp.PSObject.Properties['RawContentStream'] -and $resp.RawContentStream) {
        $text = [System.Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
    } else {
        $text = [string]$resp.Content
    }
    $sessOut = $Session
    if (-not $sessOut) { $sessOut = $nbSess }
    return @{ Text = $text; Session = $sessOut; Status = [int]$resp.StatusCode }
}

# استخراج شناسه حساب از ورودی کاربر (شناسه خالص یا لینک کامل)
function Get-AccIdFromString {
    param([string]$InputText)
    $t = $InputText.Trim()
    if ($t -match '/account/(\d+)') { return $Matches[1] }
    if ($t -match '^(\d+)$')        { return $Matches[1] }
    throw "ورودی «$InputText» معتبر نیست. شناسه عددی یا لینک کامل صفحه پزشک را وارد کنید."
}

# =======================  منطق اصلی  =======================

# دریافت صفحه پزشک و استخراج پارامترهای لازم (uid، cid، wid، service_id)
function Get-AccountInfo {
    param([string]$AccId)

    $url = '{0}/book/account/{1}?lang=fa' -f $BaseUrl, $AccId
    $r = Invoke-NobatRequest -Url $url
    $html = $r.Text
    if (-not $html -or $html.Length -lt 200) { throw 'پاسخ خالی از سرور دریافت شد.' }
    if ($html -match 'Error 403|Access Denied') {
        throw 'سایت دسترسی را مسدود کرد (403). اگر VPN خارج از ایران دارید آن را خاموش کنید.'
    }

    # نام پزشک: اولین h2 که با «دکتر» شروع می‌شود
    $docName = $null
    if ($html -match '<h2[^>]*>\s*(دکتر[^<]{2,80})\s*<') { $docName = ($Matches[1].Trim()) }
    if (-not $docName -and $html -match '<h2[^>]*>\s*([^<]{3,80})\s*</h2>') { $docName = ($Matches[1].Trim()) }

    # استخراج همه عنصرهای data-info-item (هر مرکز/خدمت یک عنصر)
    $combos = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($m in [regex]::Matches($html, '<div[^>]{0,800}?data-info-item[^>]*>')) {
        $tag = $m.Value
        $uid = '0'; $cid = '0'; $wid = '0'; $sid = '0'
        if ($tag -match 'data-uid\s*=\s*["'']?(\d+)')   { $uid = $Matches[1] }
        if ($tag -match 'data-cid\s*=\s*["'']?(\d+)')   { $cid = $Matches[1] }
        if ($tag -match 'data-wid\s*=\s*["'']?(\d+)')   { $wid = $Matches[1] }
        if ($tag -match 'data-id\s*=\s*["'']?(\d+)')    { $sid = $Matches[1] }
        if ($uid -eq '0' -and $cid -eq '0') { continue }
        $key = "$uid|$cid|$wid|$sid"
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $combos.Add([pscustomobject]@{ uid = $uid; cid = $cid; wid = $wid; sid = $sid })
        }
    }

    # اگر عنصر data-info-item نبود، از data-info (مرحله انتخاب مرکز) استفاده کن
    if ($combos.Count -eq 0) {
        foreach ($m in [regex]::Matches($html, '<div[^>]{0,800}?[\s"'']data-info[\s"''=][^>]*>')) {
            $tag = $m.Value
            if ($tag -match 'data-info-item') { continue }
            $uid = '0'; $cid = '0'
            if ($tag -match 'data-uid\s*=\s*["'']?(\d+)') { $uid = $Matches[1] }
            if ($tag -match 'data-cid\s*=\s*["'']?(\d+)') { $cid = $Matches[1] }
            if ($uid -eq '0' -or $cid -eq '0') { continue }
            $key = "$uid|$cid|0|0"
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                $combos.Add([pscustomobject]@{ uid = $uid; cid = $cid; wid = '0'; sid = '0' })
            }
        }
    }

    return [pscustomobject]@{
        AccId   = $AccId
        Name    = $docName
        Url     = $url
        Session = $r.Session
        Combos  = $combos
    }
}

# فراخوانی اندپوینت time/source و گردآوری همه «اولین وقت‌های آزاد»
function Get-FirstTimes {
    param($Info, [int]$HorizonDays = 60)

    $start = (Get-Date).AddDays(-7).ToString('yyyy-MM-dd')
    $end   = (Get-Date).AddDays($HorizonDays).ToString('yyyy-MM-dd')

    $all = New-Object System.Collections.Generic.List[object]
    foreach ($c in $Info.Combos) {
        $api = '{0}/time/source?lang=fa&uid={1}&cid={2}&wid={3}&service_id={4}&start={5}&end={6}' -f `
            $BaseUrl, $c.uid, $c.cid, $c.wid, $c.sid, $start, $end
        try {
            $r = Invoke-NobatRequest -Url $api -Session $Info.Session -ExtraHeaders @{
                'X-Requested-With' = 'XMLHttpRequest'
                'Referer'          = $Info.Url
                'Accept'           = 'application/json, text/javascript, */*; q=0.01'
            }
            $data = $r.Text | ConvertFrom-Json
            if ($null -ne $data.first_time) {
                foreach ($ft in $data.first_time) { if ($ft) { $all.Add($ft) } }
            }
        } catch {
            # خطای یک ترکیب، بقیه را خراب نکند
        }
    }

    # حذف موارد تکراری
    $uniq = New-Object System.Collections.Generic.List[object]
    $seenKeys = @{}
    foreach ($it in $all) {
        $key = '{0}|{1}|{2}|{3}' -f $it.event_date, $it.cid, $it.schedule_id, $it.exact_type
        if (-not $seenKeys.ContainsKey($key)) {
            $seenKeys[$key] = $true
            $uniq.Add($it)
        }
    }
    return ,$uniq.ToArray()
}

# انتخاب نزدیک‌ترین نوبت قابل رزرو از میان موارد آزاد
function Select-BestSlot {
    param($Items)

    if ($null -eq $Items -or $Items.Count -eq 0) { return $null }

    $usable = @($Items | Where-Object {
        ($_ | Get-Member -Name 'type') -and ($_.type -ne 'unknown') -and
        (-not (ConvertTo-Bool $_.is_full)) -and ((Get-Int $_.exact_remain_count) -gt 0)
    })
    if ($usable.Count -eq 0) { return $null }

    $open = @($usable | Where-Object {
        ($_ | Get-Member -Name 'can_be_started') -and (ConvertTo-Bool $_.can_be_started) -and
        (-not (ConvertTo-Bool $_.is_past))
    })
    $list = $open
    if ($list.Count -eq 0) { $list = $usable }   # نوبت‌هایی که هنوز باز نشده‌اند (با توضیح)

    $sorted = @($list | Sort-Object -Property {
        $h = '99:99'
        if ($_ | Get-Member -Name 'freeHourServiceTime') {
            $arr = @($_.freeHourServiceTime) | Where-Object { $_ }
            if ($arr.Count -gt 0) { $h = [string]$arr[0] }
        }
        '{0} {1}' -f [string]$_.event_date, $h
    })
    return $sorted[0]
}

function Get-SlotHours {
    param($Slot)
    $hours = @()
    if ($Slot -and ($Slot | Get-Member -Name 'freeHourServiceTime')) {
        $hours = @($Slot.freeHourServiceTime) | Where-Object { $_ }
    }
    return $hours
}

# =======================  نمایش و هشدار  =======================

function Show-Slot {
    param($Info, $Slot)
    $hours = Get-SlotHours $Slot
    $remain = Get-Int $Slot.exact_remain_count

    Write-Host ''
    Write-Host ('  ' + ('=' * 62)) -ForegroundColor DarkGray
    Write-Host ("   پزشک                : {0}" -f $Info.Name) -ForegroundColor White
    Write-Host ("   شناسه صفحه          : {0}" -f $Info.AccId) -ForegroundColor White
    Write-Host ("   نشانی صفحه          : {0}" -f $Info.Url)
    Write-Host ("   تاریخ اولین نوبت    : {0}  ({1})" -f $Slot.event_persian_date_str, $Slot.event_date) -ForegroundColor Green
    Write-Host ("   ظرفیت باقیمانده     : {0}" -f $remain) -ForegroundColor Green
    if ($hours.Count -gt 0) {
        Write-Host ("   ساعت‌های آزاد        : {0}" -f ($hours -join '  ،  ')) -ForegroundColor Yellow
    }
    if ($Slot | Get-Member -Name 'title') {
        if ($Slot.title) { Write-Host ("   نوع نوبت            : {0}" -f $Slot.title) }
    }
    if ($Slot | Get-Member -Name 'age_type') {
        $at = [string]$Slot.age_type
        if ($at -eq 'children') { Write-Host '   بازه سنی            : اطفال' }
        elseif ($at -eq 'adult') { Write-Host '   بازه سنی            : بزرگسال' }
    }
    $tags = @()
    if (($Slot | Get-Member -Name 'remain_new_patient') -and (ConvertTo-Bool $Slot.remain_new_patient)) { $tags += 'بیمار جدید' }
    if (($Slot | Get-Member -Name 'remain_old_patient') -and (ConvertTo-Bool $Slot.remain_old_patient)) { $tags += 'بیمار قدیم' }
    if ($tags.Count -gt 0) { Write-Host ("   پذیرش               : {0}" -f ($tags -join ' + ')) }
    if ($Slot | Get-Member -Name 'waitingTime') {
        if ($Slot.waitingTime) { Write-Host ("   زمان انتظار اعلامی  : {0}" -f $Slot.waitingTime) }
    }
    if (-not (ConvertTo-Bool $Slot.can_be_started)) {
        $openAt = ''
        if ($Slot | Get-Member -Name 'open_date_time_str') { $openAt = [string]$Slot.open_date_time_str }
        Write-Host ("   توجه: رزرو این نوبت هنوز باز نشده است {0}" -f $openAt) -ForegroundColor DarkYellow
    }
    Write-Host ('  ' + ('=' * 62)) -ForegroundColor DarkGray
    Write-Host ''
}

function Show-Toast {
    param([string]$Title, [string]$Message)
    if ($NoToast) { return }
    try {
        [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
        $t = Escape-Xml $Title
        $m = Escape-Xml $Message
        $xml = "<toast duration=""long""><visual><binding template=""ToastGeneric""><text>$t</text><text>$m</text></binding></visual><audio src=""ms-winsoundevent:Notification.Looping.Alarm2"" loop=""true""/></toast>"
        $doc = New-Object Windows.Data.Xml.Dom.XmlDocument
        $doc.LoadXml($xml)
        $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        $toast = New-Object Windows.UI.Notifications.ToastNotification($doc)
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show($toast)
    } catch {
        Write-Host '   (نمایش اعلان ویندوز ممکن نشد — نتیجه در همین پنجره قابل مشاهده است)' -ForegroundColor DarkGray
    }
}

function Invoke-Alert {
    param($Info, $Slot)
    if (-not $NoBeep) {
        for ($i = 0; $i -lt 3; $i++) {
            try { [console]::beep(880, 450) } catch {}
            Start-Sleep -Milliseconds 250
        }
    }
    $msg = 'نوبت آزاد پیدا شد: {0} — ظرفیت باقیمانده: {1}' -f $Slot.event_persian_date_str, (Get-Int $Slot.exact_remain_count)
    Show-Toast 'nobat_checker — نوبت آزاد!' $msg
    if (-not $NoBrowser) {
        try { Start-Process $Info.Url } catch {}
    }
}

# =======================  بررسی یک حساب  =======================

function Check-Account {
    param($Info)

    if (-not $Info.Combos -or $Info.Combos.Count -eq 0) {
        Write-Host ('  [{0}] ظرفیت نوبت‌دهی این پزشک در حال حاضر تکمیل است یا صفحه فاقد برنامه فعال است.' -f $Info.AccId) -ForegroundColor DarkYellow
        Write-Log ('{0} | NO-SCHEDULE' -f $Info.AccId)
        return $null
    }

    $items = Get-FirstTimes -Info $Info -HorizonDays $Days
    $best = Select-BestSlot -Items $items
    if (-not $best) {
        Write-Host ('  [{0}] {1} : در {2} روز آینده نوبت آزادی یافت نشد. ({3} نوبت بررسی شد)' -f `
            $Info.AccId, $Info.Name, $Days, $items.Count) -ForegroundColor DarkGray
        Write-Log ('{0} | NO-FREE-SLOT' -f $Info.AccId)
        return $null
    }

    Show-Slot -Info $Info -Slot $best
    Write-Log ('{0} | FOUND | {1} | remain={2} | hours={3}' -f `
        $Info.AccId, $best.event_persian_date_str, (Get-Int $best.exact_remain_count), ((Get-SlotHours $best) -join '/'))
    $Script:FoundAny = $true
    return $best
}

# =======================  ورودی‌ها  =======================

Write-Host ''
Write-Host '=============================================================' -ForegroundColor Cyan
Write-Host '   nobat_checker  —  بررسی اولین نوبت آزاد (نوبت‌دهی TUMS)'      -ForegroundColor Cyan
Write-Host '=============================================================' -ForegroundColor Cyan

$targets = New-Object System.Collections.Generic.List[string]

if ($Account -and @($Account).Count -gt 0) {
    foreach ($a in $Account) { $targets.Add((Get-AccIdFromString $a)) }
}
elseif ($DoctorsFile) {
    if (-not (Test-Path $DoctorsFile)) { throw "فایل «$DoctorsFile» پیدا نشد." }
    foreach ($line in Get-Content -Path $DoctorsFile -Encoding UTF8) {
        $t = ($line -split '#')[0].Trim()
        if ($t) { $targets.Add((Get-AccIdFromString $t)) }
    }
}
else {
    Write-Host ''
    Write-Host 'شناسه یا لینک صفحه پزشک را وارد کنید. مثال:' -ForegroundColor Yellow
    Write-Host '   10460'
    Write-Host '   https://nobat.tums.ac.ir/book/account/10460?lang=fa'
    Write-Host '(چند مقدار را با ویرگول جدا کنید)' -ForegroundColor DarkGray
    $inp = Read-Host 'ورودی'
    foreach ($part in ($inp -split '[,،]')) {
        $t = $part.Trim()
        if ($t) { $targets.Add((Get-AccIdFromString $t)) }
    }
}

if ($targets.Count -eq 0) { throw 'هیچ پزشکی برای بررسی مشخص نشده است.' }

# دریافت اطلاعات و ساخت نشست برای هر حساب
$infos = New-Object System.Collections.Generic.List[object]
foreach ($acc in $targets) {
    try {
        $info = Get-AccountInfo -AccId $acc
        $infos.Add($info)
        Write-Host ("آماده شد: {0} (شناسه {1}) — {2} مرکز/خدمت فعال" -f $info.Name, $acc, $info.Combos.Count) -ForegroundColor Green
    } catch {
        Write-Host ("خطا در دریافت صفحه پزشک {0}: {1}" -f $acc, $_.Exception.Message) -ForegroundColor Red
        Write-Log ('{0} | PAGE-ERROR | {1}' -f $acc, $_.Exception.Message)
    }
}
if ($infos.Count -eq 0) { throw 'دریافت صفحه هیچ‌کدام از پزشکان ممکن نشد.' }

# =======================  اجرای اصلی  =======================

if ($Watch) {
    Write-Host ''
    Write-Host ('حالت پایش فعال شد — هر {0} ثانیه یک‌بار بررسی می‌شود. برای توقف Ctrl+C را فشار دهید.' -f $Interval) -ForegroundColor Cyan
    $seen = @{}
    while ($true) {
        Write-Host ''
        Write-Host ('----- بررسی {0} -----' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -ForegroundColor DarkCyan
        foreach ($info in $infos) {
            $best = $null
            try { $best = Check-Account -Info $info } catch {
                Write-Host ('  خطا در بررسی {0}: {1}' -f $info.AccId, $_.Exception.Message) -ForegroundColor Red
                Write-Host ($_.ScriptStackTrace) -ForegroundColor DarkGray
                continue
            }
            if ($best) {
                $sig = '{0}|{1}|{2}' -f $best.event_date, ($best.exact_remain_count), ((Get-SlotHours $best) -join ',')
                if (-not $seen.ContainsKey($info.AccId) -or $seen[$info.AccId] -ne $sig) {
                    $seen[$info.AccId] = $sig
                    Invoke-Alert -Info $info -Slot $best
                    Write-Host '   >> هشدار ارسال شد (زنگ + اعلان ویندوز + باز شدن مرورگر)' -ForegroundColor Magenta
                }
            }
        }
        $next = (Get-Date).AddSeconds($Interval).ToString('HH:mm:ss')
        Write-Host ('----- بررسی بعدی در ساعت {0} ... -----' -f $next) -ForegroundColor DarkCyan
        Start-Sleep -Seconds $Interval
    }
}
else {
    Write-Host ''
    foreach ($info in $infos) {
        try { $null = Check-Account -Info $info } catch {
            Write-Host ('  خطا در بررسی {0}: {1}' -f $info.AccId, $_.Exception.Message) -ForegroundColor Red
            Write-Host ($_.ScriptStackTrace) -ForegroundColor DarkGray
        }
    }
    if (-not $Script:FoundAny) {
        Write-Host 'نتیجه: در این بررسی نوبت آزادی پیدا نشد. برای بررسی خودکار مداوم از فلگ Watch- استفاده کنید:' -ForegroundColor Yellow
        Write-Host ('   .\nobat_checker.ps1 -Account {0} -Watch -Interval 180' -f $infos[0].AccId) -ForegroundColor Yellow
        exit 1
    }
}
