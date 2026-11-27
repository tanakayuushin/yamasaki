<#
.SYNOPSIS
  請求書メール 対応チェック（直近1週間版）

.DESCRIPTION
  直近の受信メール（既定は7日間）から請求書メール（PDF・画像・Excel・ZIP の添付、または請求書取得用のリンクがあるもの）を見つけ、
  件名・本文・添付ファイル名に「請求書」などがなければ、添付の中身（文字認識など）も確認する。
  社内の対応ルールどおりに処理されているかをチェックして、結果の画面（HTML）と CSV を出力する。

    対応段階 : 分類項目の確認者の数 → 0人=未対応 / 1人=1人目済 / 2人以上=完了
    コピー   : コピー先フォルダの「○○月分」（特定案件は専用フォルダ）にコピーがあるか（Message-ID で突き合わせ）
    格納     : 格納先\<年>\<月>（特定案件はその中の専用フォルダ）に、請求書のファイル（添付と中身が同じもの）があるか
               どの添付が請求書か分かれば（ファイル名・中身に「請求書」など）そのファイル、
               分からなければ添付のどれか1つが格納されていれば「済」
    返信     : コピー先（またはコピー元）が返信済みか。自動メールは不要
    フラグ   : 1人目済 → コピー先がフラグあり、完了 → コピー先が完了
    注意     : 「振込先変更」などの記載、危険な種類の添付（exe など）

  読み取りのみ。メール・フラグ・分類項目・フォルダは一切変更しない。
  Windows PowerShell 5.1 と、デスクトップ版（従来の）Outlook が必要。
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = '',
    [int]$DaysBack = 0
)

$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSVersion.Major -lt 3) {
    Write-Host ('[NG] PowerShell のバージョンが古すぎます（' + $PSVersionTable.PSVersion + '）。') -ForegroundColor Red
    exit 1
}

# スクリプトのフォルダ（$PSScriptRoot が空になる環境にも対応）
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir -and $MyInvocation.MyCommand.Definition) { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }
if (-not $ConfigPath) { $ConfigPath = Join-Path $ScriptDir 'config.json' }

# ---------------------------------------------------------------------------
# 設定
# ---------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "設定ファイルが見つかりません: $ConfigPath（config.sample.json をコピーして config.json を作ってください）"
}
$Config  = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($DaysBack -le 0) { $DaysBack = [int]$Config.Outlook.DaysBack }
$Det     = $Config.Detection
$Rules   = $Config.Rules
$Storage = $Config.Storage
$Caution = $Config.Caution
$Insp    = $Config.Inspection
$spLabel = if ($Rules.SpecialProject.Label) { [string]$Rules.SpecialProject.Label } else { '特定案件' }

$PR_INTERNET_MESSAGE_ID = 'http://schemas.microsoft.com/mapi/proptag/0x1035001F'
$PR_LAST_VERB_EXECUTED  = 'http://schemas.microsoft.com/mapi/proptag/0x10810003'
$PR_ATTACH_CONTENT_ID   = 'http://schemas.microsoft.com/mapi/proptag/0x3712001F'
$PR_ATTACHMENT_HIDDEN   = 'http://schemas.microsoft.com/mapi/proptag/0x7FFE000B'

$ImageExts  = @('.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.gif')
$OfficeExts = @('.xlsx', '.docx', '.pptx')   # マクロ付き（.xlsm など）は「危険な種類の添付」として扱う

$TempDir = Join-Path ([IO.Path]::GetTempPath()) ('invoice-check-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

# ---------------------------------------------------------------------------
# 汎用
# ---------------------------------------------------------------------------
function Test-ContainsAny {
    param([string]$Text, [object[]]$Words)
    if ([string]::IsNullOrEmpty($Text) -or -not $Words) { return $false }
    foreach ($w in $Words) {
        if ($w -and $Text.IndexOf([string]$w, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    }
    return $false
}

function Find-FirstWord {
    param([string]$Text, [object[]]$Words)
    if ([string]::IsNullOrEmpty($Text) -or -not $Words) { return '' }
    foreach ($w in $Words) {
        if ($w -and $Text.IndexOf([string]$w, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return [string]$w }
    }
    return ''
}

function Test-MatchAny {
    param([string]$Text, [object[]]$Patterns)
    if ([string]::IsNullOrEmpty($Text) -or -not $Patterns) { return $false }
    foreach ($p in $Patterns) { if ($p -and $Text -match [string]$p) { return $true } }
    return $false
}

# 照合用: 全角英数→半角・㈱→(株) などに揃え、空白を除く
function ConvertTo-MatchText {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return '' }
    return ($s.Normalize([Text.NormalizationForm]::FormKC) -replace '\s+', '')
}

# 「１０」「10」どちらでも数値にする
function ConvertTo-Int {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return 0 }
    $t = ($s.Normalize([Text.NormalizationForm]::FormKC)) -replace '[^0-9]', ''
    if (-not $t) { return 0 }
    return [int]$t
}

function Get-SafeFileName {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    return (-join ($Name.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '_' } else { $_ } }))
}

# 「株式会社〇〇」→「〇〇」（ファイル名の照合用）
function Get-CompanyKeyword {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    $k = $Name -replace '株式会社|（株）|\(株\)|㈱|有限会社|（有）|\(有\)|合同会社|一般社団法人|一般財団法人', ''
    $k = ($k -replace '[<>"].*$', '') -replace '\s+', ''
    $k = $k -replace '(様|御中)$', ''
    if ($k.Length -lt 2) { return $null }
    return $k
}

function Join-Notes {
    param([object[]]$Items)
    return ((@($Items) | Where-Object { $_ }) -join ' / ')
}

function ConvertTo-Defanged { param([string]$s) return ($s -replace '\.', '[.]') }   # クリックしても開かない表記

function Get-UrlHost {
    param([string]$Url)
    try { return ([Uri]$Url).Host.ToLower() } catch { return '' }
}

# ---------------------------------------------------------------------------
# 請求書取得用のリンク
# ---------------------------------------------------------------------------
function Get-InvoiceLinks {
    param([string]$Body, [string]$HtmlBody)
    $found = New-Object System.Collections.Generic.List[string]
    if ($Body) {
        foreach ($m in [regex]::Matches($Body, 'https?://[^\s"''<>()（）「」、。]+')) {
            $url = $m.Value.TrimEnd('.', ',', ')', '>')
            if (Test-MatchAny $url $Det.IgnoreUrlPatterns) { continue }
            $start = [Math]::Max(0, $m.Index - [int]$Det.LinkContextChars)
            $len   = [Math]::Min($Body.Length - $start, $m.Length + 2 * [int]$Det.LinkContextChars)
            $ctx   = $Body.Substring($start, $len)
            if ((Test-MatchAny $url $Det.LinkPatterns) -or (Test-ContainsAny $ctx $Det.LinkContextKeywords)) {
                if (-not $found.Contains($url)) { $found.Add($url) }
            }
        }
    }
    if ($HtmlBody) {
        foreach ($m in [regex]::Matches($HtmlBody, '<a\s[^>]*href\s*=\s*["'']([^"'']+)["''][^>]*>(.*?)</a>', 'IgnoreCase, Singleline')) {
            $url  = [System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value)
            $text = ($m.Groups[2].Value -replace '<[^>]+>', '')
            if ($url -notmatch '^https?://') { continue }
            if (Test-MatchAny $url $Det.IgnoreUrlPatterns) { continue }
            if ((Test-MatchAny $url $Det.LinkPatterns) -or (Test-ContainsAny $text $Det.LinkContextKeywords) -or (Test-ContainsAny $text $Det.Keywords)) {
                if (-not $found.Contains($url)) { $found.Add($url) }
            }
        }
    }
    return ,$found.ToArray()
}

# ---------------------------------------------------------------------------
# 請求月の候補（件名・添付名の「11月分」など → なければ受信月と前月）
# ---------------------------------------------------------------------------
function Get-BillingMonths {
    param([datetime]$Received, [string]$Text)
    $list = @()
    $m = [regex]::Match([string]$Text, '(?:(20\d{2}|２０[０-９]{2})年)?\s*(\d{1,2})\s*月\s*(?:分|度|ご請求|請求|\d{1,2}日締)')
    if ($m.Success) {
        $mo = ConvertTo-Int $m.Groups[2].Value
        if ($mo -ge 1 -and $mo -le 12) {
            if ($m.Groups[1].Success) { $y = ConvertTo-Int $m.Groups[1].Value }
            elseif ($mo -gt $Received.Month) { $y = $Received.Year - 1 }
            else { $y = $Received.Year }
            $list += (Get-Date -Year $y -Month $mo -Day 1).Date
        }
    }
    foreach ($o in @($Storage.SearchMonthOffsets)) {
        $d = (Get-Date -Year $Received.Year -Month $Received.Month -Day 1).Date.AddMonths([int]$o)
        if (-not ($list | Where-Object { $_ -eq $d })) { $list += $d }
    }
    return $list
}

# ---------------------------------------------------------------------------
# 格納先（<RootPath>\<年>\<月>）の検索
#   年・月フォルダの名前は「2026」「2026年」「11」「11月」「2026.11」などに対応
# ---------------------------------------------------------------------------
function Find-MonthFolders {
    param([datetime]$Month)
    $root = [string]$Storage.RootPath
    if (-not (Test-Path -LiteralPath $root -PathType Container)) { return }
    $years = @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue |
               Where-Object { $_.Name -match ('^' + $Month.Year + '(?!\d)') })
    foreach ($y in $years) {
        Get-ChildItem -LiteralPath $y.FullName -Directory -ErrorAction SilentlyContinue | Where-Object {
            ($_.Name -match '^(?:\d{4}[年.\-_/]?)?0?(\d{1,2})(?!\d)') -and ((ConvertTo-Int $Matches[1]) -eq $Month.Month)
        } | ForEach-Object { $_.FullName }
    }
}

# 特定案件: 月フォルダの中の専用フォルダ（設定がなければ月フォルダそのもの）
function Find-SpecialFolders {
    param([string[]]$MonthFolders)
    $pat = [string]$Rules.SpecialProject.StorageSubfolderPattern
    if (-not $pat) { return $MonthFolders }
    foreach ($mf in $MonthFolders) {
        Get-ChildItem -LiteralPath $mf -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match $pat } | ForEach-Object { $_.FullName }
    }
}

# 推測した月以外も探すための範囲：受信した年と前の年の、すべての月フォルダ（特定案件はその中の専用フォルダ）
$script:WideCache = @{}
function Get-WideFolders {
    param([datetime]$Received, [bool]$Special)
    $key = "$($Received.Year)|$Special"
    if ($script:WideCache.ContainsKey($key)) { return $script:WideCache[$key] }
    $list = @()
    $root = [string]$Storage.RootPath
    if (Test-Path -LiteralPath $root -PathType Container) {
        $years = @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -match ('^(' + $Received.Year + '|' + ($Received.Year - 1) + ')(?!\d)') })
        $months = @(foreach ($y in $years) { Get-ChildItem -LiteralPath $y.FullName -Directory -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName } })
        if ($Special) { $list = @(Find-SpecialFolders $months) } else { $list = $months }
    }
    $script:WideCache[$key] = $list
    return $list
}

$script:FolderCache = @{}
function Get-FolderIndex {
    param([string]$Folder)
    if ($script:FolderCache.ContainsKey($Folder)) { return $script:FolderCache[$Folder] }
    $index = [pscustomobject]@{ Files = @(); Hashes = @{} }
    if (Test-Path -LiteralPath $Folder -PathType Container) {
        $gci = @{ LiteralPath = $Folder; File = $true; ErrorAction = 'SilentlyContinue' }
        if ($Storage.Recurse) { $gci.Recurse = $true }
        $index.Files = @(Get-ChildItem @gci)
    }
    $script:FolderCache[$Folder] = $index
    return $index
}

function Get-FileSha256 {
    param([string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

# 添付と中身が同じファイルがあるか → なければ同じ名前のファイルがあるか
function Find-StoredFile {
    param([string]$TempFile, [string]$FileName, [string[]]$Folders, [switch]$HashOnly)
    $hash = Get-FileSha256 $TempFile
    $size = (Get-Item -LiteralPath $TempFile).Length
    foreach ($f in $Folders) {
        $idx = Get-FolderIndex $f
        foreach ($file in $idx.Files) {
            if ($file.Length -ne $size) { continue }
            if (-not $idx.Hashes.ContainsKey($file.FullName)) {
                try { $idx.Hashes[$file.FullName] = Get-FileSha256 $file.FullName } catch { $idx.Hashes[$file.FullName] = '' }   # 開けないファイルは飛ばす
            }
            if ($idx.Hashes[$file.FullName] -eq $hash) { return [pscustomobject]@{ Result = '済'; How = '中身が一致'; Path = $file.FullName } }
        }
    }
    if ($HashOnly) { return [pscustomobject]@{ Result = '未'; How = ''; Path = '' } }
    $want = ConvertTo-MatchText $FileName
    foreach ($f in $Folders) {
        $hit = (Get-FolderIndex $f).Files | Where-Object { (ConvertTo-MatchText $_.Name) -ieq $want } | Select-Object -First 1
        if ($hit) { return [pscustomobject]@{ Result = '要確認'; How = '同じ名前のファイルはあるが中身が違う'; Path = $hit.FullName } }
    }
    return [pscustomobject]@{ Result = '未'; How = ''; Path = '' }
}

# 取引先名を含むファイルがあるか（リンクで届いた請求書・画像の請求書用）
function Find-StoredByKeyword {
    param([string[]]$Keywords, [string[]]$Folders)
    $kws = @($Keywords | ForEach-Object { ConvertTo-MatchText $_ } | Where-Object { $_ })
    if ($kws.Count -eq 0) { return [pscustomobject]@{ Result = '要確認'; How = '取引先名が分からない（config.json の Vendors に登録すると判定できます）'; Path = '' } }
    $hits = @()
    foreach ($f in $Folders) { $hits += @((Get-FolderIndex $f).Files | Where-Object { Test-ContainsAny (ConvertTo-MatchText $_.Name) $kws }) }
    if ($hits.Count -gt 0) {
        return [pscustomobject]@{ Result = '済(推定)'; How = '取引先名を含むファイルあり'; Path = (($hits | Select-Object -ExpandProperty FullName -Unique) -join ' | ') }
    }
    return [pscustomobject]@{ Result = '未'; How = ''; Path = '' }
}

function Get-Vendor {
    param([string]$SenderAddress)
    foreach ($v in @($Config.Vendors)) { if (Test-MatchAny $SenderAddress $v.SenderPatterns) { return $v } }
    return $null
}

function Get-VendorKeywords {
    param($Vendor, [string]$SenderName)
    if ($Vendor -and $Vendor.FileKeywords) { return @($Vendor.FileKeywords) }
    $k = Get-CompanyKeyword $SenderName
    if ($k) { return @($k) }
    return @()
}

# ---------------------------------------------------------------------------
# 添付の中身の確認
#   PDF・画像 … Windows 標準の文字認識（OCR）/ Excel・Word・PowerPoint … 中の文字 / ZIP … 中のファイル名
#   添付を「開く」「実行する」ことはしない。中身は文字として読むだけ
# ---------------------------------------------------------------------------
function Save-AttTemp {
    param($A)
    if (-not $A.Temp) {
        $p = Join-Path $TempDir ([Guid]::NewGuid().ToString('N') + '_' + (Get-SafeFileName $A.Name))
        $A.Att.SaveAsFile($p)
        $A.Temp = $p
    }
    return $A.Temp
}

# 中身の文字に「請求書らしさ」があるか → 見つかった根拠を返す（なければ空）
function Find-InvoiceEvidence {
    param([string]$Text)
    $t = ConvertTo-MatchText $Text
    if (-not $t) { return '' }
    foreach ($w in @($Det.Keywords)) {
        $k = ConvertTo-MatchText ([string]$w)
        if ($k -and $t.IndexOf($k, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return "「$w」" }
    }
    # 「請求書」の文字がなくても、インボイスの登録番号（T＋13桁）と金額・支払いの記載があれば請求書とみなす
    if ($Insp.UseRegistrationNumber -and $t -match 'T\d{13}' -and (Test-ContainsAny $t @('合計', '金額', '税込', '消費税', 'お支払', '支払期限', '振込'))) {
        return '登録番号と金額の記載'
    }
    return ''
}

# Excel / Word / PowerPoint（新形式）の文字
function Get-OfficeText {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $sb = New-Object System.Text.StringBuilder
        foreach ($e in $zip.Entries) {
            if ($e.FullName -notmatch '^(xl/sharedStrings\.xml|xl/worksheets/sheet\d+\.xml|word/document\.xml|word/header\d*\.xml|ppt/slides/slide\d+\.xml)$') { continue }
            if ($e.Length -gt 20MB) { continue }
            $reader = New-Object System.IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
            try { [void]$sb.Append(($reader.ReadToEnd() -replace '<[^>]+>', ' ')) } finally { $reader.Dispose() }
            if ($sb.Length -gt 2000000) { break }
        }
        return $sb.ToString()
    } finally { $zip.Dispose() }
}

# ZIP の中のファイル名（展開はしない）
function Get-ZipNames {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Read, [Text.Encoding]::GetEncoding(932))
    try { return (@($zip.Entries | ForEach-Object { $_.FullName }) -join ' ') } finally { $zip.Dispose() }
}

# --- 文字認識（最初に使うときだけ準備） ---
$script:OcrReady   = $null
$script:OcrNote    = ''
$script:OcrSkipped = 0
function Initialize-Ocr {
    if ($null -ne $script:OcrReady) { return $script:OcrReady }
    $script:OcrReady = $false
    if ($PSVersionTable.PSEdition -eq 'Core') { $script:OcrNote = 'PowerShell 7 では文字認識を使えません（bat から起動してください）'; return $false }
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime
        $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
        $null = [Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
        $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
        $null = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.BitmapTransform, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
        $null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
        $null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime]
        $null = [Windows.Media.Ocr.OcrResult, Windows.Media.Ocr, ContentType = WindowsRuntime]
        $null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]
        $methods = [System.WindowsRuntimeSystemExtensions].GetMethods()
        $script:AsTaskOp = $methods | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
        $script:AsTaskAction = $methods | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' } | Select-Object -First 1
        $lang = New-Object Windows.Globalization.Language ([string]$Insp.OcrLanguage)
        $script:OcrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
        if (-not $script:OcrEngine) { $script:OcrNote = '日本語の文字認識機能がインストールされていません'; return $false }
        $script:OcrMaxDim = [int][Windows.Media.Ocr.OcrEngine]::MaxImageDimension
        $script:OcrReady = $true
    } catch {
        $script:OcrNote = "文字認識の準備に失敗しました（$($_.Exception.Message)）"
    }
    return $script:OcrReady
}
function Wait-WinRt {
    param($Operation, [Type]$ResultType)
    $task = $script:AsTaskOp.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    $null = $task.Wait(-1)
    return $task.Result
}
function Wait-WinRtAction {
    param($Operation)
    $task = $script:AsTaskAction.Invoke($null, @($Operation))
    $null = $task.Wait(-1)
}

# 画像のストリームを文字認識
function Read-OcrStream {
    param($Stream)
    $Stream.Seek(0)
    $dec = Wait-WinRt ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($Stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $pw = [double]$dec.PixelWidth; $ph = [double]$dec.PixelHeight
    $scale = [Math]::Min(1.0, $script:OcrMaxDim / [Math]::Max($pw, $ph))
    $tf = New-Object Windows.Graphics.Imaging.BitmapTransform
    $tf.ScaledWidth  = [uint32][Math]::Max(1, [Math]::Floor($pw * $scale))
    $tf.ScaledHeight = [uint32][Math]::Max(1, [Math]::Floor($ph * $scale))
    $bmp = Wait-WinRt ($dec.GetSoftwareBitmapAsync(
            [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
            [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied,
            $tf,
            [Windows.Graphics.Imaging.ExifOrientationMode]::RespectExifOrientation,
            [Windows.Graphics.Imaging.ColorManagementMode]::DoNotColorManage)) ([Windows.Graphics.Imaging.SoftwareBitmap])
    try {
        $res = Wait-WinRt ($script:OcrEngine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
        return [string]$res.Text
    } finally { $bmp.Dispose() }
}

# PDF（先頭から Inspection.MaxPdfPages ページ）または画像を文字認識
function Get-OcrText {
    param([string]$Path, [string]$Kind)
    $sf = Wait-WinRt ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) ([Windows.Storage.StorageFile])
    if ($Kind -ne 'PDF') {
        $stream = Wait-WinRt ($sf.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
        try { return (Read-OcrStream $stream) } finally { $stream.Dispose() }
    }
    $pdf = Wait-WinRt ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($sf)) ([Windows.Data.Pdf.PdfDocument])
    $pages = [Math]::Min([int]$pdf.PageCount, [Math]::Max(1, [int]$Insp.MaxPdfPages))
    $text = ''
    for ($i = 0; $i -lt $pages; $i++) {
        $page = $pdf.GetPage([uint32]$i)
        $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
        try {
            $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions
            $w = [Math]::Min([double]$script:OcrMaxDim, [double]$page.Size.Width * 2.5)
            $h = [double]$page.Size.Height * ($w / [double]$page.Size.Width)
            if ($h -gt $script:OcrMaxDim) { $w = $w * ($script:OcrMaxDim / $h) }
            $opt.DestinationWidth = [uint32][Math]::Floor($w)
            Wait-WinRtAction ($page.RenderToStreamAsync($stream, $opt))
            $text += ' ' + (Read-OcrStream $stream)
        } finally {
            $stream.Dispose()
            $page.Dispose()
        }
        if (Find-InvoiceEvidence $text) { break }   # 見つかったら残りのページは読まない
    }
    return $text
}

# 確認結果の記録（ファイルのハッシュ値 → 見つかった根拠）。中身の文字そのものは保存しない
$script:InspectCachePath = Join-Path $ScriptDir 'inspect_cache.json'
$script:InspectCache = @{}
if (Test-Path -LiteralPath $script:InspectCachePath) {
    try {
        (Get-Content -LiteralPath $script:InspectCachePath -Raw -Encoding UTF8 | ConvertFrom-Json).PSObject.Properties |
            ForEach-Object { $script:InspectCache[$_.Name] = [string]$_.Value }
    } catch { }
}

# 請求書に書かれた日付の月（請求月）を読み取る
#   「請求日」「発行日」などの見出しのすぐ後の日付 → なければ最初の日付。支払期限などの後ろの日付は使わない
$DatePatterns = @(
    '(?<![\d\-/.])(?<y>20\d{2})年(?<m>\d{1,2})月(?<d>\d{1,2})日?',
    '令和(?<r>\d{1,2}|元)年(?<m>\d{1,2})月(?<d>\d{1,2})日?',
    '(?<![\d\-/.])(?<y>20\d{2})[/.\-](?<m>\d{1,2})[/.\-](?<d>\d{1,2})(?!\d)',
    '(?<![A-Za-z\d])R(?<r>\d{1,2})[/.\-](?<m>\d{1,2})[/.\-](?<d>\d{1,2})(?!\d)'
)
function Find-InvoiceMonth {
    param([string]$Text, [datetime]$Received)
    $t = ConvertTo-MatchText $Text
    if (-not $t) { return $null }
    $labels   = @($Config.InvoiceDate.Labels | ForEach-Object { ConvertTo-MatchText $_ } | Where-Object { $_ })
    $excludes = @($Config.InvoiceDate.ExcludeLabels | ForEach-Object { ConvertTo-MatchText $_ } | Where-Object { $_ })
    $cands = @()
    foreach ($p in $DatePatterns) {
        foreach ($m in [regex]::Matches($t, $p)) {
            if ($m.Groups['y'].Success) { $y = ConvertTo-Int $m.Groups['y'].Value }
            else {
                if ($m.Groups['r'].Value -eq '元') { $rr = 1 } else { $rr = ConvertTo-Int $m.Groups['r'].Value }
                $y = 2018 + $rr
            }
            $mo = ConvertTo-Int $m.Groups['m'].Value
            $d  = ConvertTo-Int $m.Groups['d'].Value
            if ($y -lt 2000 -or $y -gt 2100 -or $mo -lt 1 -or $mo -gt 12 -or $d -lt 1 -or $d -gt 31) { continue }
            $dt = (Get-Date -Year $y -Month $mo -Day 1).Date
            if ([Math]::Abs(($dt - $Received.Date).TotalDays) -gt 400) { continue }   # 受信日から離れすぎた日付は使わない
            $from = [Math]::Max(0, $m.Index - 12)
            $ctx = $t.Substring($from, $m.Index - $from)
            if (Test-ContainsAny $ctx $excludes) { continue }
            $cands += [pscustomobject]@{ Month = $dt; Pos = $m.Index; Labeled = (Test-ContainsAny $ctx $labels) }
        }
    }
    if ($cands.Count -eq 0) { return $null }
    $pick = @($cands | Where-Object { $_.Labeled } | Sort-Object Pos) | Select-Object -First 1
    if (-not $pick) { $pick = @($cands | Sort-Object Pos) | Select-Object -First 1 }
    return $pick.Month
}

# 添付1件の中身に請求書らしさがあるか ＋ 請求書の日付の月
function Get-AttachmentFinding {
    param($A, [datetime]$Received = (Get-Date))
    $how = ''
    if ($A.Kind -eq 'PDF') { $how = 'PDFの文字認識' }
    elseif ($A.Kind -eq '画像') { $how = '画像の文字認識' }
    elseif ($A.Kind -eq 'Excel等') { $how = 'Excel・Wordの文字' }
    elseif ($A.Kind -eq 'ZIP') { $how = 'ZIPの中のファイル名' }
    $path = Save-AttTemp $A
    $hash = Get-FileSha256 $path
    $dateKey = "$hash#month"
    if ($script:InspectCache.ContainsKey($hash) -and $script:InspectCache.ContainsKey($dateKey)) {
        $cm = $null
        if ($script:InspectCache[$dateKey]) { $cm = [datetime]::ParseExact([string]$script:InspectCache[$dateKey], 'yyyy-MM', $null) }
        return [pscustomobject]@{ Found = [string]$script:InspectCache[$hash]; How = $how; Note = ''; Month = $cm }
    }
    $text = ''; $note = ''; $cache = $true
    try {
        if ($A.Kind -eq 'PDF' -or $A.Kind -eq '画像') {
            if (Initialize-Ocr) { $text = Get-OcrText $path $A.Kind }
            else { $script:OcrSkipped++; $cache = $false }
        }
        elseif ($A.Kind -eq 'Excel等') { $text = Get-OfficeText $path }
        elseif ($A.Kind -eq 'ZIP') { $text = Get-ZipNames $path }
    } catch {
        $note = '中身を読めませんでした（パスワード付き・壊れたファイルなど）'; $cache = $false
    }
    $found = Find-InvoiceEvidence $text
    $month = $null
    if ($A.Kind -ne 'ZIP') { $month = Find-InvoiceMonth $text $Received }
    if ($cache) {
        $script:InspectCache[$hash] = $found
        if ($month) { $script:InspectCache[$dateKey] = $month.ToString('yyyy-MM') } else { $script:InspectCache[$dateKey] = '' }
    }
    return [pscustomobject]@{ Found = $found; How = $how; Note = $note; Month = $month }
}

# ---------------------------------------------------------------------------
# Outlook
# ---------------------------------------------------------------------------
function Get-MapiProp {
    param($Item, [string]$Prop)
    try { return $Item.PropertyAccessor.GetProperty($Prop) } catch { return $null }
}

function Get-SenderSmtp {
    param($Mail)
    try {
        if ($Mail.SenderEmailType -eq 'EX') {
            $ex = $Mail.Sender.GetExchangeUser()
            if ($ex) { return [string]$ex.PrimarySmtpAddress }
        }
    } catch { }
    return [string]$Mail.SenderEmailAddress
}

function Get-MailKey {
    param($Mail)
    $id = [string](Get-MapiProp $Mail $PR_INTERNET_MESSAGE_ID)
    if ($id) { return $id.Trim() }
    return ('{0}|{1}|{2}' -f $Mail.Subject, ([datetime]$Mail.ReceivedTime).ToString('yyyyMMddHHmmss'), $Mail.SenderEmailAddress)
}

# 102=返信 103=全員に返信
function Test-Replied {
    param($Mail)
    $v = Get-MapiProp $Mail $PR_LAST_VERB_EXECUTED
    return ($v -eq 102 -or $v -eq 103)
}

function Get-FlagText {
    param($Mail)
    $st = 0
    try { $st = [int]$Mail.FlagStatus } catch { }
    if ($st -eq 1) { return '完了' }
    if ($st -eq 2) { return '赤旗' }
    return 'なし'
}

# 本文に埋め込まれた画像（署名のロゴなど）か
function Test-InlineAttachment {
    param($Att, [string]$HtmlBody)
    $cid = [string](Get-MapiProp $Att $PR_ATTACH_CONTENT_ID)
    if ($cid -and $HtmlBody -and $HtmlBody.IndexOf("cid:$cid", [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    if (Get-MapiProp $Att $PR_ATTACHMENT_HIDDEN) { return $true }
    return $false
}

function Resolve-OutlookFolder {
    param($Namespace, [string]$Mailbox, [string]$FolderPath)
    if ([string]::IsNullOrWhiteSpace($Mailbox)) { $current = $Namespace.GetDefaultFolder(6).Parent }
    else { $current = $Namespace.Folders.Item($Mailbox) }
    foreach ($part in ($FolderPath -split '[\\/]')) {
        if (-not $part) { continue }
        $next = $null
        try { $next = $current.Folders.Item($part) } catch { }
        if (-not $next) { throw "Outlook のフォルダが見つかりません: 「$part」（設定: $FolderPath）" }
        $current = $next
    }
    return $current
}

function Get-AllFolders {
    param($Folder, [bool]$IncludeSubfolders)
    $list = @($Folder)
    if ($IncludeSubfolders) { foreach ($sub in $Folder.Folders) { $list += Get-AllFolders $sub $true } }
    return $list
}

# 検索フォルダー（お気に入りにある、受信メールをまとめたもの）を名前で探す
function Find-SearchFolder {
    param($Namespace, [string]$Mailbox, [string]$Name)
    $stores = @()
    if ($Mailbox) { try { $stores += $Namespace.Folders.Item($Mailbox).Store } catch { } }
    else { try { $stores += $Namespace.DefaultStore } catch { } }
    try { foreach ($st in $Namespace.Stores) { $stores += $st } } catch { }   # 見つからなければ、ほかのメールボックスも探す
    foreach ($st in $stores) {
        try {
            foreach ($f in $st.GetSearchFolders()) { if ([string]$f.Name -eq $Name) { return $f } }
        } catch { }
    }
    return $null
}

# 分類項目 → 確認者名の一覧
function Get-Confirmers {
    param([string]$Categories)
    $cats  = @(([string]$Categories) -split '\s*[,;、]\s*' | Where-Object { $_ })
    $staff = @($Rules.Staff | Where-Object { $_ })
    if ($staff.Count -gt 0) { return ($cats | Where-Object { $staff -contains $_ } | Select-Object -Unique) }
    $ignore = @($Rules.DefectCategory) + @($Rules.IgnoreCategories) + @($Rules.ExcludeCategories)
    return ($cats | Where-Object { $ignore -notcontains $_ } | Select-Object -Unique)
}

# Outlook との接続が切れたことを示すエラーか
function Test-OutlookDisconnected {
    param($ErrorRecord)
    $e = $ErrorRecord.Exception
    while ($e) {
        if (@(-2147417848, -2147023174, -2147023170) -contains $e.HResult) { return $true }
        $e = $e.InnerException
    }
    return $false
}

# ---------------------------------------------------------------------------
# 一覧の列
# ---------------------------------------------------------------------------
$Columns = @('結果', 'やること', '受信日時', '請求月', '取引先', '件名', '差出人', '差出人アドレス', '対応段階', '確認者', '不備', '特定案件',
             '種別', '検出根拠', '注意', 'コピー', 'コピー先', '格納', '格納ファイル', '探したフォルダ', '返信',
             'コピー先フラグ', 'コピー元フラグ', '添付ファイル', 'リンク先', '備考', 'Outlookフォルダ')
function New-Row {
    $r = [ordered]@{}
    foreach ($c in $Columns) { $r[$c] = '' }
    return $r
}

# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------
# Outlook が開いていなければ、画面ありで起動してから接続
if (-not (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)) {
    Write-Host 'Outlook が開いていないため起動します。少しお待ちください...'
    try { Start-Process 'outlook.exe' } catch { throw 'Outlook を起動できませんでした。Outlook を開いてから、もう一度実行してください。' }
    $waited = 0
    while (-not (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) -and $waited -lt 60) { Start-Sleep -Seconds 2; $waited += 2 }
    Start-Sleep -Seconds 15
}
Write-Host 'Outlook に接続しています...（終わるまで Outlook を閉じないでください）'
$outlook = New-Object -ComObject Outlook.Application
$ns      = $outlook.GetNamespace('MAPI')
$mailbox = [string]$Config.Outlook.Mailbox

$since      = (Get-Date).Date.AddDays(-$DaysBack)
$filter     = "[ReceivedTime] >= '" + $since.ToString('g') + "'"
$copyFilter = "[ReceivedTime] >= '" + $since.AddDays(-7).ToString('g') + "'"   # コピーは元の受信日時を引き継ぐ

$results       = New-Object System.Collections.Generic.List[object]
$scopeText     = ''
$mailCount     = 0
$errorCount    = 0
$fatal         = ''

try {
    # --- コピー先フォルダの索引（必要な値だけを控える） ---
    Write-Host ('コピー先フォルダを読み込み中: ' + $Config.Outlook.CopyRootFolder)
    $copyRoot     = Resolve-OutlookFolder $ns $mailbox ([string]$Config.Outlook.CopyRootFolder)
    $copyRootPath = [string]$copyRoot.FolderPath
    $copyIndex    = @{}
    foreach ($cf in (Get-AllFolders $copyRoot $true)) {
        if ($cf.FolderPath -eq $copyRootPath) { continue }
        foreach ($ci in $cf.Items.Restrict($copyFilter)) {
            try {
                if ($ci.Class -ne 43) { continue }
                $k = Get-MailKey $ci
                if (-not $copyIndex.ContainsKey($k)) { $copyIndex[$k] = @() }
                $copyIndex[$k] += [pscustomobject]@{
                    Categories = [string]$ci.Categories
                    FlagText   = (Get-FlagText $ci)
                    Replied    = (Test-Replied $ci)
                    FolderName = [string]$cf.Name
                    FolderPath = [string]$cf.FolderPath
                }
            } finally {
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ci) } catch { }
            }
        }
    }

    # --- 調べる範囲 ---
    #   人が確認しているのと同じ検索フォルダー（SearchFolderName）を調べる。見つからなければ設定のフォルダ（受信トレイ）
    $sources = @()
    $sfName = [string]$Config.Outlook.SearchFolderName
    if ($sfName) {
        $sf = Find-SearchFolder $ns $mailbox $sfName
        if ($sf) { $sources += [pscustomobject]@{ Label = "「$sfName」"; Folder = $sf } }
        else { Write-Host "※ 検索フォルダー「$sfName」が見つからないため、設定のフォルダ（$(@($Config.Outlook.Folders) -join '、')）を調べます。" -ForegroundColor Yellow }
    }
    if ($sources.Count -eq 0) {
        foreach ($folderPath in @($Config.Outlook.Folders)) {
            $top = Resolve-OutlookFolder $ns $mailbox ([string]$folderPath)
            foreach ($fd in (Get-AllFolders $top ([bool]$Config.Outlook.IncludeSubfolders))) {
                if (([string]$fd.FolderPath).StartsWith($copyRootPath)) { continue }
                $sources += [pscustomobject]@{ Label = [string]$fd.FolderPath; Folder = $fd }
            }
        }
    }
    $scopeText = (@($sources | ForEach-Object { $_.Label }) -join '、')

    foreach ($src in $sources) {
            Write-Host ('メールを確認中: ' + $src.Label)
            try { $items = $src.Folder.Items.Restrict($filter) } catch { $items = $src.Folder.Items }
            try { $items.Sort('[ReceivedTime]', $true) } catch { }
            $total = [Math]::Max(1, [int]$items.Count)

            foreach ($mail in $items) {
                $atts = @()
                $mailFolder = ''
                try {
                    if ($mail.Class -ne 43) { continue }   # 43 = メール（会議通知などは除く）
                    # メールが入っているフォルダ（コピー先フォルダの中のコピーは対象外）
                    try {
                        $parent = $mail.Parent
                        $mailFolder = [string]$parent.FolderPath
                        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($parent)
                    } catch { }
                    if ($mailFolder -and $mailFolder.StartsWith($copyRootPath)) { continue }
                    if ([datetime]$mail.ReceivedTime -lt $since) { continue }
                    $mailCount++
                    if (($mailCount % 25) -eq 0) {
                        Write-Progress -Activity 'メールを確認中' -Status "$mailCount 件目" -PercentComplete ([Math]::Min(100, 100 * $mailCount / $total))
                    }

                    $subject  = [string]$mail.Subject
                    $body     = [string]$mail.Body
                    $htmlBody = ''
                    try { $htmlBody = [string]$mail.HTMLBody } catch { }

                    if ((Test-ContainsAny $subject $Det.ExcludeSubjectKeywords) -and -not (Test-ContainsAny $subject @('請求'))) { continue }

                    # --- 添付の仕分け（PDF・画像。危険な種類は注意に） ---
                    $warn = @()
                    foreach ($att in $mail.Attachments) {
                        if ($att.Type -ne 1) { continue }   # 1 = 実ファイル
                        $name = [string]$att.FileName
                        $ext  = [IO.Path]::GetExtension($name).ToLower()
                        if (@($Caution.DangerExtensions) -contains $ext) { $warn += "危険な種類の添付: $name"; continue }
                        if ($ext -eq '.pdf') { $kind = 'PDF' }
                        elseif ($ImageExts -contains $ext) {
                            if ([double]$att.Size -lt [double]$Det.MinImageBytes) { continue }   # 署名のロゴなど
                            if (Test-InlineAttachment $att $htmlBody) { continue }               # 本文に埋め込まれた画像
                            $kind = '画像'
                        }
                        elseif ($OfficeExts -contains $ext) { $kind = 'Excel等' }
                        elseif ($ext -eq '.zip') { $kind = 'ZIP' }
                        else { continue }
                        $atts += [pscustomobject]@{ Att = $att; Name = $name; Kind = $kind; Temp = ''; IsInvoice = (Test-ContainsAny $name $Det.Keywords) }
                    }
                    $attNames = (@($atts | ForEach-Object { $_.Name })) -join ' '

                    # --- 請求書メールかどうか ---
                    $reasons = @()
                    if (Test-ContainsAny $subject $Det.Keywords) { $reasons += '件名' }
                    if (Test-ContainsAny $attNames $Det.Keywords) { $reasons += '添付ファイル名' }
                    if (Test-ContainsAny $body $Det.Keywords) { $reasons += '本文' }
                    $notes = @()
                    if ($reasons.Count -eq 0 -and $Insp.Enabled -and $atts.Count -gt 0) {
                        # 件名・本文・添付ファイル名に言葉がなければ、添付の中身を確認
                        $checked = 0
                        foreach ($a in $atts) {
                            if ($checked -ge [int]$Insp.MaxAttachmentsPerMail) { break }
                            if ([double]$a.Att.Size -gt [double]$Insp.MaxFileMB * 1MB) { continue }
                            $checked++
                            $f = Get-AttachmentFinding $a ([datetime]$mail.ReceivedTime)
                            if ($f.Note) { $notes += "$($a.Name)：$($f.Note)" }
                            if ($f.Found) { $a.IsInvoice = $true; $reasons += "添付の中身（$($f.How)で$($f.Found)）"; break }
                        }
                    }
                    if ($reasons.Count -eq 0) { continue }

                    # 添付が複数あり、ファイル名ではどれが請求書か分からなければ、中身で見分ける
                    if ($Insp.Enabled -and $atts.Count -gt 1 -and @($atts | Where-Object { $_.IsInvoice }).Count -eq 0) {
                        $checked = 0
                        foreach ($a in $atts) {
                            if ($checked -ge [int]$Insp.MaxAttachmentsPerMail) { break }
                            if ([double]$a.Att.Size -gt [double]$Insp.MaxFileMB * 1MB) { continue }
                            $checked++
                            $f = Get-AttachmentFinding $a ([datetime]$mail.ReceivedTime)
                            if ($f.Found) { $a.IsInvoice = $true }
                        }
                    }
                    $links = @()
                    $links = @(Get-InvoiceLinks $body $htmlBody)   # 添付があってもリンクも見る（請求書がリンクで届き、添付は別の書類のこともあるため）
                    if ($atts.Count -eq 0 -and $links.Count -eq 0 -and $warn.Count -eq 0) { continue }

                    # --- 基本情報 ---
                    $received  = [datetime]$mail.ReceivedTime
                    $senderAdr = Get-SenderSmtp $mail
                    $senderNm  = [string]$mail.SenderName
                    $vendor    = Get-Vendor $senderAdr
                    $spKeys    = @($Rules.SpecialProject.Keywords | Where-Object { $_ })
                    $isSpecial = ($spKeys.Count -gt 0) -and ((Test-ContainsAny "$subject $attNames $senderNm" $spKeys) -or (Test-ContainsAny $body $spKeys))
                    $isAuto    = Test-MatchAny $senderAdr $Rules.AutoMailSenderPatterns
                    $bankWord  = Find-FirstWord (ConvertTo-MatchText "$subject $body") @($Caution.BankChangeKeywords | ForEach-Object { ConvertTo-MatchText $_ })
                    if ($bankWord) { $warn += "振込先・口座の変更に関する記載（「$bankWord」）" }

                    # --- 対応段階（分類項目の確認者の数） ---
                    $copies = @()
                    $key = Get-MailKey $mail
                    if ($copyIndex.ContainsKey($key)) { $copies = @($copyIndex[$key]) }
                    $copy = $null
                    if ($copies.Count -gt 0) { $copy = $copies[0] }

                    $catText = [string]$mail.Categories
                    if ($copy -and $copy.Categories) { $catText = "$catText, $($copy.Categories)" }
                    $catList    = @($catText -split '\s*[,;、]\s*' | Where-Object { $_ })
                    $confirmers = @(Get-Confirmers $catText)
                    $hasDefect  = @($catList | Where-Object { $_ -eq [string]$Rules.DefectCategory }).Count -gt 0
                    $isExcluded = @($catList | Where-Object { @($Rules.ExcludeCategories) -contains $_ }).Count -gt 0
                    if ($confirmers.Count -ge 2) { $stage = '完了' } elseif ($confirmers.Count -eq 1) { $stage = '1人目済' } else { $stage = '未対応' }

                    # --- コピー ---
                    $copyJudge = '未'; $copyNote = ''; $copyMonth = 0
                    if ($copy) {
                        $spPat = [string]$Rules.SpecialProject.CopyFolderPattern
                        $isSpecialFolder = [bool]$spPat -and ($copy.FolderName -match $spPat)
                        if ($isSpecial -and $spPat) { $pat = $spPat } else { $pat = [string]$Rules.NormalCopyFolderPattern }
                        $m = [regex]::Match($copy.FolderName, $pat)
                        if ($m.Success -and (-not $spPat -or ($isSpecial -eq $isSpecialFolder))) {
                            $copyJudge = '済'
                            if ($m.Groups.Count -gt 1) { $copyMonth = ConvertTo-Int $m.Groups[1].Value }
                        } else {
                            $copyJudge = '要確認'
                            if ($isSpecial -and -not $isSpecialFolder) { $copyNote = "$spLabel の案件なのに通常の月分フォルダにコピー" }
                            elseif (-not $isSpecial -and $isSpecialFolder) { $copyNote = "$spLabel 以外なのに $spLabel 用フォルダにコピー" }
                            else { $copyNote = 'コピー先のフォルダ名が想定と違う' }
                        }
                        if ($copies.Count -gt 1) { $copyNote = Join-Notes @($copyNote, "コピーが $($copies.Count) 件ある") }
                    }

                    # --- 格納 ---
                    # 請求書がどう届いたか
                    #   attach … ファイル名か中身で請求書と分かった添付がある → そのファイルがすべて格納されているか
                    #   link   … 分からず、取得用のリンクがある → リンクの請求書（取引先名で推定）。添付のどれかが格納されていても済
                    #   any    … 分からず、リンクもない → 添付（PDF、なければ全部）のどれか1つが格納されているか
                    $invoiceAtts = @($atts | Where-Object { $_.IsInvoice })
                    if ($invoiceAtts.Count -gt 0) { $mode = 'attach'; $files = $invoiceAtts }
                    elseif ($links.Count -gt 0) { $mode = 'link'; $files = @($atts | Where-Object { $_.Kind -eq 'PDF' }) }
                    else {
                        $mode = 'any'
                        $files = @($atts | Where-Object { $_.Kind -eq 'PDF' })
                        if ($files.Count -eq 0) { $files = $atts }
                    }

                    # 請求月 = 請求書に書かれた日付の月（請求書のファイルを文字認識などで読む）
                    $invMonth = $null
                    if ($mode -ne 'link' -and $Insp.Enabled) {
                        foreach ($a in $files) {
                            if ($a.Kind -eq 'ZIP') { continue }
                            if ([double]$a.Att.Size -gt [double]$Insp.MaxFileMB * 1MB) { continue }
                            $fi = Get-AttachmentFinding $a $received
                            if ($fi.Month) { $invMonth = $fi.Month; break }
                        }
                    }
                    if ($invMonth) {
                        $months = @($invMonth)
                        $monthText = $invMonth.ToString('yyyy/MM')
                        $monthSource = '請求書の日付'
                        if ($copyMonth -ge 1 -and $copyMonth -ne $invMonth.Month -and -not $isSpecial) {
                            $copyNote = Join-Notes @($copyNote, "コピー先の「$copyMonth 月分」と請求書の日付の月（$($invMonth.Month)月）が違う")
                        }
                    } else {
                        # 日付を読み取れなければ推測：コピー先フォルダの月 → 件名の「○月分」 → 受信月と前月
                        $months = @(Get-BillingMonths $received "$subject $attNames")
                        if ($copyMonth -ge 1 -and $copyMonth -le 12) {
                            if ($copyMonth -gt $received.Month) { $cy = $received.Year - 1 } else { $cy = $received.Year }
                            $months = @((Get-Date -Year $cy -Month $copyMonth -Day 1).Date) + @($months | Where-Object { $_.Month -ne $copyMonth })
                        }
                        $monthText = (@($months | ForEach-Object { $_.ToString('yyyy/MM') }) -join '・')
                        $monthSource = '推測（請求書の日付を読み取れず）'
                    }
                    $monthFolders = @(foreach ($mo in $months) { Find-MonthFolders $mo })
                    if ($isSpecial) { $searchFolders = @(Find-SpecialFolders $monthFolders) } else { $searchFolders = $monthFolders }

                    $storeJudge = '未'; $storeNote = ''; $storePaths = @(); $attDetail = @()
                    if ($searchFolders.Count -eq 0) {
                        if ($isSpecial -and $monthFolders.Count -gt 0) { $storeNote = "月フォルダの中に $spLabel 用フォルダがない" } else { $storeNote = "請求月（$monthText）のフォルダがない" }
                    }

                    $judges = @()
                    foreach ($a in $files) {
                        $tmp = Save-AttTemp $a
                        $r = [pscustomobject]@{ Result = '未'; How = ''; Path = '' }
                        if ($searchFolders.Count -gt 0) { $r = Find-StoredFile $tmp $a.Name $searchFolders }
                        if ($r.Result -ne '済') {
                            # 推測した月のフォルダになければ、受信した年と前の年のほかの月も探す（中身が同じものだけ）
                            $rw = Find-StoredFile $tmp $a.Name (Get-WideFolders $received $isSpecial) -HashOnly
                            if ($rw.Result -eq '済') {
                                if ($invMonth) { $r = [pscustomobject]@{ Result = '要確認'; How = "請求書の日付の月（$monthText）とは別の月のフォルダにありました"; Path = $rw.Path } }
                                else { $r = [pscustomobject]@{ Result = '済'; How = "推測した月（$monthText）とは別の月のフォルダにありました"; Path = $rw.Path } }
                            }
                        }
                        if ($r.Result -eq '未' -and $a.Kind -ne 'PDF' -and $searchFolders.Count -gt 0) {
                            # 画像・Excel などは PDF にして格納されることがある → 取引先名で推定
                            $r2 = Find-StoredByKeyword (Get-VendorKeywords $vendor $senderNm) $searchFolders
                            if ($r2.Result -like '済*') { $r = $r2 }
                        }
                        $judges += $r.Result
                        if ($a.IsInvoice) { $mark = '（請求書）' } else { $mark = '' }
                        $attDetail += ('{0}{1}：{2}' -f $a.Name, $mark, $r.Result)
                        if ($r.Path) { $storePaths += $r.Path }
                        if ($r.How) { $storeNote = Join-Notes @($storeNote, "$($a.Name)：$($r.How)") }
                    }

                    if ($mode -eq 'attach') {
                        if ($judges -contains '未') { $storeJudge = '未' }
                        elseif ($judges -contains '要確認') { $storeJudge = '要確認' }
                        elseif ($judges -contains '済(推定)') { $storeJudge = '済(推定)' }
                        else { $storeJudge = '済' }
                    }
                    elseif ($mode -eq 'link') {
                        $lr = [pscustomobject]@{ Result = '未'; How = ''; Path = '' }
                        if ($searchFolders.Count -gt 0) { $lr = Find-StoredByKeyword (Get-VendorKeywords $vendor $senderNm) $searchFolders }
                        if ($lr.Path) { $storePaths += $lr.Path }
                        if ($judges -contains '済') { $storeJudge = '済' }
                        elseif ($lr.Result -like '済*') { $storeJudge = $lr.Result; $storeNote = Join-Notes @($storeNote, $lr.How) }
                        elseif ($judges -contains '要確認' -or $lr.Result -eq '要確認') { $storeJudge = '要確認'; $storeNote = Join-Notes @($storeNote, $lr.How) }
                        else { $storeJudge = '未'; $storeNote = Join-Notes @($storeNote, 'リンク先からダウンロードして格納が必要') }
                    }
                    else {
                        if ($judges -contains '済') { $storeJudge = '済' }
                        elseif ($judges -contains '済(推定)') { $storeJudge = '済(推定)' }
                        elseif ($judges -contains '要確認') { $storeJudge = '要確認' }
                        else { $storeJudge = '未' }
                        if ($files.Count -gt 1) { $storeNote = Join-Notes @($storeNote, 'どの添付が請求書か分からないため、どれか1つが格納されていれば「済」としています') }
                    }
                    if ($storeJudge -eq '未' -and $hasDefect) { $storeJudge = '不備(再発行待ち)' }

                    # --- 返信・フラグ ---
                    if ($isAuto) { $replyJudge = '不要(自動メール)' }
                    elseif ($copy -and $copy.Replied) { $replyJudge = '済' }
                    elseif (Test-Replied $mail) { $replyJudge = '済(コピー元で返信)' }
                    else { $replyJudge = '未' }
                    if ($copy) { $copyFlag = $copy.FlagText } else { $copyFlag = '-' }
                    $srcFlag = Get-FlagText $mail

                    # --- やること・結果 ---
                    $todo = @()
                    if ($stage -eq '未対応') {
                        $todo += '1人目の確認待ち'
                    } else {
                        if ($copyJudge -ne '済') { $todo += "コピー($copyJudge)" }
                        if ($storeJudge -notlike '済*' -and $storeJudge -notlike '不備*') { $todo += "格納($storeJudge)" }
                        if ($replyJudge -eq '未') { $todo += '受領返信' }
                        if ($copy) {
                            if ($stage -eq '1人目済' -and $copyFlag -eq 'なし') { $todo += 'コピー先を赤旗に' }
                            if ($stage -eq '完了' -and $copyFlag -ne '完了') { $todo += 'コピー先のフラグを完了に' }
                        }
                        if ($stage -eq '1人目済' -and $todo.Count -eq 0) { $todo += '2人目の確認待ち' }
                    }
                    if ($warn.Count -gt 0) {
                        if ($bankWord) { $first = '振込先の変更は、メールではなく登録済みの電話番号で取引先に必ず確認' }
                        else { $first = '添付を開かずに、差出人へ電話などで確認' }
                        $todo = @($first) + $todo
                    }

                    if ($isExcluded) { $result = '対象外'; $todo = @() }
                    elseif ($warn.Count -gt 0) { $result = '要注意' }
                    elseif ($stage -eq '完了' -and $todo.Count -eq 0) { $result = 'OK' }
                    elseif ($stage -ne '未対応' -and @($todo | Where-Object { $_ -ne '2人目の確認待ち' }).Count -gt 0) { $result = '漏れあり' }
                    else { $result = '対応待ち' }

                    # --- 一覧に追加 ---
                    $row = New-Row
                    $row['結果']           = $result
                    $row['やること']       = ($todo -join ' / ')
                    $row['受信日時']       = $received.ToString('yyyy/MM/dd HH:mm')
                    $row['請求月']         = "$monthText（$monthSource）"
                    if ($vendor) { $row['取引先'] = [string]$vendor.Name } else { $row['取引先'] = $senderNm }
                    $row['件名']           = $subject
                    $row['差出人']         = $senderNm
                    $row['差出人アドレス'] = $senderAdr
                    $row['対応段階']       = $stage
                    $row['確認者']         = ($confirmers -join ',')
                    if ($hasDefect) { $row['不備'] = '○' }
                    if ($isSpecial) { $row['特定案件'] = '○' }
                    $kinds = @(@($atts | ForEach-Object { $_.Kind }) | Select-Object -Unique)
                    if ($links.Count -gt 0) { $kinds += 'リンク' }
                    if ($kinds.Count -gt 0) { $row['種別'] = ($kinds -join '+') } else { $row['種別'] = '危険な添付のみ' }
                    $row['検出根拠']       = ($reasons -join ' / ')
                    $row['注意']           = ($warn -join ' / ')
                    $row['コピー']         = $copyJudge
                    if ($copy) { $row['コピー先'] = $copy.FolderPath }
                    $row['格納']           = $storeJudge
                    $row['格納ファイル']   = ((@($storePaths) | Select-Object -Unique) -join ' | ')
                    $row['探したフォルダ'] = ($searchFolders -join ' | ')
                    $row['返信']           = $replyJudge
                    $row['コピー先フラグ'] = $copyFlag
                    $row['コピー元フラグ'] = $srcFlag
                    $row['添付ファイル']   = ($attDetail -join ' | ')
                    $row['リンク先']       = ((@($links | ForEach-Object { Get-UrlHost $_ } | Where-Object { $_ } | ForEach-Object { ConvertTo-Defanged $_ }) | Select-Object -Unique) -join ' ')
                    $row['備考']           = (Join-Notes (@($copyNote, $storeNote) + $notes))
                    $row['Outlookフォルダ'] = $mailFolder
                    $results.Add([pscustomobject]$row)
                }
                catch {
                    if (Test-OutlookDisconnected $_) { throw }   # Outlook との接続切れ → 全体を止めて、ここまでの結果を保存
                    # このメールは確認できなかった → 一覧に「エラー」として残して次へ
                    $errorCount++
                    $row = New-Row
                    try { $row['受信日時'] = ([datetime]$mail.ReceivedTime).ToString('yyyy/MM/dd HH:mm') } catch { }
                    try { $row['差出人'] = [string]$mail.SenderName; $row['取引先'] = [string]$mail.SenderName } catch { }
                    try { $row['件名'] = [string]$mail.Subject } catch { }
                    $row['結果'] = 'エラー'
                    $row['やること'] = 'このメールは直接確認してください'
                    $row['備考'] = [string]$_.Exception.Message
                    $row['Outlookフォルダ'] = $mailFolder
                    $results.Add([pscustomobject]$row)
                }
                finally {
                    # Outlook の負担を減らすため、読み終わったメールはすぐ手放す
                    foreach ($a in @($atts)) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($a.Att) } catch { } }
                    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) } catch { }
                    if (($mailCount % 50) -eq 0) { [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
                }
            }
    }
}
catch {
    if (Test-OutlookDisconnected $_) { $fatal = 'Outlook との接続が途中で切れました（Outlook が閉じた、または固まって再起動した可能性があります）。' }
    else { $fatal = "途中で止まりました: $($_.Exception.Message)" }
}
finally {
    Write-Progress -Activity 'メールを確認中' -Completed
    try { $script:InspectCache | ConvertTo-Json -Compress | Set-Content -LiteralPath $script:InspectCachePath -Encoding UTF8 } catch { }
    Remove-Item -LiteralPath $TempDir -Recurse -Force -ErrorAction SilentlyContinue
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns) } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) } catch { }
}

# ---------------------------------------------------------------------------
# 結果の保存（既定はひとつ上の「結果」フォルダ）
# ---------------------------------------------------------------------------
$outDir = [string]$Config.Report.OutputFolder
if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path (Split-Path $ScriptDir -Parent) '結果' }
elseif (-not [IO.Path]::IsPathRooted($outDir)) { $outDir = [IO.Path]::GetFullPath((Join-Path $ScriptDir $outDir)) }
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$stamp    = (Get-Date).ToString('yyyyMMdd_HHmmss')
$csvPath  = Join-Path $outDir ("請求書メールチェック_$stamp.csv")
$htmlPath = Join-Path $outDir ("請求書メールチェック_$stamp.html")

# 古い結果を削除（このツールが作ったファイルだけ）
$keep = [int]$Config.Report.KeepDays
if ($keep -gt 0) {
    $limit = (Get-Date).AddDays(-$keep)
    Get-ChildItem -LiteralPath $outDir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^(請求書メールチェック_\d{8}_\d{6}\.(csv|html)|環境チェック結果_\d{8}_\d{6}\.txt)$' -and $_.LastWriteTime -lt $limit } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

# 結果のグループ（並び順・色・説明）
$Groups = [ordered]@{
    '要注意'   = @{ Rank = 0; Css = 'g-alert'; Desc = '振込先変更の記載や危険な添付がある。返信・格納・振込の前に、電話などで取引先に確認' }
    '漏れあり' = @{ Rank = 1; Css = 'g-miss';  Desc = '確認済みなのに、コピー・格納・返信・フラグのどれかが足りない' }
    'エラー'   = @{ Rank = 2; Css = 'g-error'; Desc = 'ツールで確認できなかった。メールを直接確認' }
    '対応待ち' = @{ Rank = 3; Css = 'g-wait';  Desc = '1人目・2人目の確認待ち（いつもどおり対応）' }
    'OK'       = @{ Rank = 4; Css = 'g-ok';    Desc = '2人の確認が済み、すべて揃っている' }
    '対象外'   = @{ Rank = 5; Css = 'g-skip';  Desc = '分類項目「対象外」が付いている' }
}
$sorted = @($results | Sort-Object @{ Expression = { $Groups[[string]$_.結果].Rank } }, @{ Expression = { $_.受信日時 }; Descending = $true })

$enc = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
$sorted | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding $enc

# --- 結果の画面（HTML）。メールの文字はすべて無害化して埋め込む ---
function ConvertTo-HtmlText { param($s) return [System.Net.WebUtility]::HtmlEncode([string]$s) }
function Get-ChipState {
    param([string]$Kind, [string]$Value, [string]$Stage)
    $pending = ($Stage -eq '未対応')
    if ($Kind -eq 'copy')  { if ($Value -eq '済') { return 'ok' }; if ($Value -eq '要確認') { return 'warn' }; if ($pending) { return 'none' }; return 'bad' }
    if ($Kind -eq 'store') { if ($Value -like '済*') { return 'ok' }; if ($Value -like '不備*' -or $Value -eq '要確認') { return 'warn' }; if ($pending) { return 'none' }; return 'bad' }
    if ($Kind -eq 'reply') { if ($Value -like '済*') { return 'ok' }; if ($Value -like '不要*' -or $pending) { return 'none' }; return 'bad' }
    if ($Stage -eq '1人目済') { if ($Value -eq 'なし' -or $Value -eq '-') { return 'bad' }; return 'ok' }
    if ($Stage -eq '完了') { if ($Value -eq '完了') { return 'ok' }; return 'bad' }
    return 'none'
}
function New-Chip {
    param([string]$Label, [string]$Value, [string]$State)
    return "<span class=""chip $State""><b>$(ConvertTo-HtmlText $Label)</b>$(ConvertTo-HtmlText $Value)</span>"
}

$counts = @{}
foreach ($g in $Groups.Keys) { $counts[$g] = 0 }
foreach ($r in $sorted) { if ($counts.ContainsKey([string]$r.結果)) { $counts[[string]$r.結果]++ } }
$todoCount = $counts['要注意'] + $counts['漏れあり'] + $counts['エラー'] + $counts['対応待ち']
# 保管が未のもの：確認の段階に関係なく、請求書のファイルが格納されていない（済・不備以外）もの
function Test-StoreMissing {
    param($Row)
    if ([string]$Row.結果 -eq '対象外' -or [string]$Row.結果 -eq 'エラー') { return $false }
    $v = [string]$Row.格納
    return -not ($v -like '済*' -or $v -like '不備*')
}
$storeMissingCount = @($sorted | Where-Object { Test-StoreMissing $_ }).Count

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append(@'
<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>請求書メールチェック結果</title>
<style>
:root{--bg:#f5f6f8;--card:#fff;--ink:#1d2330;--sub:#5d6473;--line:#dde1e7;
--alert:#c62828;--miss:#e65100;--warn:#b88a00;--error:#6a1b9a;--wait:#546e7a;--ok:#2e7d32;--skip:#8a8f98}
@media (prefers-color-scheme:dark){:root{--bg:#15181e;--card:#1f242c;--ink:#e8eaee;--sub:#a3a9b5;--line:#343a45}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.6 "Yu Gothic UI","Meiryo",sans-serif}
header{background:var(--card);border-bottom:1px solid var(--line);padding:16px 20px}
h1{margin:0 0 4px;font-size:20px}.meta{color:var(--sub);font-size:13px}
main{max-width:1100px;margin:0 auto;padding:16px}
.notice{border-left:4px solid var(--alert);background:var(--card);padding:10px 14px;margin:0 0 12px;border-radius:6px}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px;margin-bottom:12px}
.card{border:2px solid transparent;background:var(--card);border-radius:10px;padding:10px 12px;text-align:left;cursor:pointer;color:var(--ink);font:inherit}
.card .n{display:block;font-size:26px;font-weight:700;line-height:1.2;color:var(--c)}.card .l{font-weight:700}
.card .d{display:block;color:var(--sub);font-size:12px;line-height:1.4;margin-top:2px}.card.active{border-color:var(--ink)}
.toolbar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin:6px 0 12px}
.toolbar input{flex:1;min-width:200px;padding:8px 10px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);font:inherit}
.count{color:var(--sub);font-size:13px}
.item{background:var(--card);border-radius:10px;border-left:6px solid var(--c);padding:12px 14px;margin-bottom:10px}
.top{display:flex;gap:10px;align-items:center;flex-wrap:wrap;font-size:13px;color:var(--sub)}
.badge{color:#fff;background:var(--c);border-radius:999px;padding:1px 10px;font-weight:700;font-size:13px}
.item h3{margin:6px 0 0;font-size:16px}.subject{margin:2px 0 6px;color:var(--sub);word-break:break-all}
.todo{margin:6px 0;padding:6px 10px;border-radius:6px;background:rgba(127,127,127,.12);font-weight:600}
.caution{color:var(--alert);font-weight:700;margin:6px 0}
.chips{display:flex;gap:6px;flex-wrap:wrap;margin-top:6px}
.chip{border:1px solid var(--line);border-radius:6px;padding:1px 8px;font-size:13px}.chip b{font-weight:400;color:var(--sub);margin-right:6px}
.chip.ok{border-color:var(--ok);color:var(--ok)}.chip.bad{border-color:var(--miss);color:var(--miss);font-weight:700}
.chip.warn{border-color:var(--warn);color:var(--warn)}.chip.none{color:var(--sub)}
details{margin-top:8px}summary{cursor:pointer;color:var(--sub);font-size:13px}
dl{display:grid;grid-template-columns:130px 1fr;gap:4px 12px;margin:8px 0 0;font-size:13px}dt{color:var(--sub)}dd{margin:0;word-break:break-all}
.g-alert{--c:var(--alert)}.g-miss{--c:var(--miss)}.g-error{--c:var(--error)}.g-wait{--c:var(--wait)}.g-ok{--c:var(--ok)}.g-skip{--c:var(--skip)}
.legend{background:var(--card);border-radius:10px;padding:12px 16px;margin-top:20px;font-size:14px}
.legend table{border-collapse:collapse;width:100%}.legend td{border-top:1px solid var(--line);padding:6px 8px;vertical-align:top}
#empty{color:var(--sub);text-align:center;padding:30px}a{color:inherit}
@media print{.toolbar,.cards,summary{display:none}.item{break-inside:avoid}}
@media (max-width:600px){dl{grid-template-columns:1fr}}
</style></head><body>
'@)

$runAt  = (Get-Date).ToString('yyyy/MM/dd HH:mm')
$period = "$scopeText の直近 $DaysBack 日（$($since.ToString('yyyy/MM/dd')) ～）のメール"
[void]$sb.Append("<header><h1>請求書メールチェック結果</h1><div class=""meta"">実行：$(ConvertTo-HtmlText $runAt)　／　対象：$(ConvertTo-HtmlText $period)　／　確認したメール：$mailCount 件　／　請求書メール：$($sorted.Count) 件　／　<a href=""../使い方ガイド.html"">使い方ガイド</a></div></header><main>")
if ($fatal) { [void]$sb.Append("<div class=""notice""><b>途中で止まりました。</b>$(ConvertTo-HtmlText $fatal) この結果は、止まるまでに確認できた $mailCount 件分です。Outlook を開いたまま、もう一度実行してください。</div>") }
if ($script:OcrSkipped -gt 0) { [void]$sb.Append("<div class=""notice"">文字認識を使えなかったため、添付 $($script:OcrSkipped) 件（PDF・画像）の中身を確認できませんでした（$(ConvertTo-HtmlText $script:OcrNote)）。</div>") }
if ($counts['要注意'] -gt 0) { [void]$sb.Append("<div class=""notice""><b>要注意のメールが $($counts['要注意']) 件あります。</b>添付やリンクを開かず、メール以外の方法（登録済みの電話番号など）で取引先に確認してください。</div>") }

[void]$sb.Append('<div class="cards">')
[void]$sb.Append("<button class=""card g-wait"" data-filter=""todo""><span class=""n"">$todoCount</span><span class=""l"">対応が必要なもの</span><span class=""d"">OK・対象外以外のすべて</span></button>")
[void]$sb.Append("<button class=""card g-miss"" data-filter=""store""><span class=""n"">$storeMissingCount</span><span class=""l"">保管が未のもの</span><span class=""d"">確認の段階に関係なく、請求書のファイルが既定の場所に見つからない</span></button>")
foreach ($g in $Groups.Keys) {
    [void]$sb.Append("<button class=""card $($Groups[$g].Css)"" data-filter=""$(ConvertTo-HtmlText $g)""><span class=""n"">$($counts[$g])</span><span class=""l"">$(ConvertTo-HtmlText $g)</span><span class=""d"">$(ConvertTo-HtmlText $Groups[$g].Desc)</span></button>")
}
[void]$sb.Append("<button class=""card g-skip"" data-filter=""all""><span class=""n"">$($sorted.Count)</span><span class=""l"">すべて</span><span class=""d"">見つかった請求書メール全部</span></button>")
[void]$sb.Append('</div><div class="toolbar"><input id="q" type="search" placeholder="取引先名・件名などで絞り込み"><span class="count" id="count"></span></div>')

foreach ($r in $sorted) {
    $g = [string]$r.結果
    $css = if ($Groups.Contains($g)) { $Groups[$g].Css } else { 'g-skip' }
    if ($r.確認者) { $stageText = "$($r.対応段階)（$($r.確認者)）" } else { $stageText = $r.対応段階 }
    if (Test-StoreMissing $r) { $sm = '1' } else { $sm = '0' }
    [void]$sb.Append("<article class=""item $css"" data-group=""$(ConvertTo-HtmlText $g)"" data-store=""$sm"">")
    [void]$sb.Append("<div class=""top""><span class=""badge"">$(ConvertTo-HtmlText $g)</span><span>$(ConvertTo-HtmlText $r.受信日時)</span><span>$(ConvertTo-HtmlText $stageText)</span>")
    if ($r.特定案件) { [void]$sb.Append("<span>$(ConvertTo-HtmlText $spLabel)</span>") }
    if ($r.不備) { [void]$sb.Append('<span>不備</span>') }
    if ($r.種別) { [void]$sb.Append("<span>$(ConvertTo-HtmlText $r.種別)</span>") }
    [void]$sb.Append("</div><h3>$(ConvertTo-HtmlText $r.取引先)</h3><p class=""subject"">$(ConvertTo-HtmlText $r.件名)</p>")
    if ($r.やること) { [void]$sb.Append("<p class=""todo"">やること：$(ConvertTo-HtmlText $r.やること)</p>") }
    if ($r.注意) { [void]$sb.Append("<p class=""caution"">注意：$(ConvertTo-HtmlText $r.注意)</p>") }
    if ($g -ne '対象外' -and $g -ne 'エラー') {
        [void]$sb.Append('<div class="chips">')
        [void]$sb.Append((New-Chip 'コピー' $r.コピー (Get-ChipState 'copy' $r.コピー $r.対応段階)))
        [void]$sb.Append((New-Chip '格納' $r.格納 (Get-ChipState 'store' $r.格納 $r.対応段階)))
        [void]$sb.Append((New-Chip '受領返信' $r.返信 (Get-ChipState 'reply' $r.返信 $r.対応段階)))
        [void]$sb.Append((New-Chip 'コピー先フラグ' $r.コピー先フラグ (Get-ChipState 'flag' $r.コピー先フラグ $r.対応段階)))
        [void]$sb.Append('</div>')
    }
    [void]$sb.Append('<details><summary>詳しく見る</summary><dl>')
    foreach ($c in @('請求月', '差出人', '差出人アドレス', '検出根拠', 'コピー先', '格納ファイル', '探したフォルダ', 'コピー元フラグ', '添付ファイル', 'リンク先', '備考', 'Outlookフォルダ')) {
        $v = [string]$r.$c
        if ($v) { [void]$sb.Append("<dt>$(ConvertTo-HtmlText $c)</dt><dd>$(ConvertTo-HtmlText $v)</dd>") }
    }
    [void]$sb.Append('</dl></details></article>')
}
[void]$sb.Append('<p id="empty" hidden>表示するメールはありません。</p>')
[void]$sb.Append('<section class="legend"><h2 style="font-size:16px;margin:0 0 6px">結果の意味とやること</h2><table>')
foreach ($g in $Groups.Keys) { [void]$sb.Append("<tr><td class=""$($Groups[$g].Css)""><span class=""badge"">$(ConvertTo-HtmlText $g)</span></td><td>$(ConvertTo-HtmlText $Groups[$g].Desc)</td></tr>") }
[void]$sb.Append('</table><p style="color:var(--sub);font-size:13px;margin:8px 0 0">各項目の色：緑＝できている／橙＝足りない／黄＝要確認／灰＝まだ不要。リンクのURLは安全のため載せていません（リンク先はドメインのみ）。</p></section>')
[void]$sb.Append(@'
</main>
<script>
(function(){
  var items=[].slice.call(document.querySelectorAll('.item'));
  var btns=[].slice.call(document.querySelectorAll('[data-filter]'));
  var q=document.getElementById('q'), cur='todo';
  var todo=['要注意','漏れあり','エラー','対応待ち'];
  function apply(){
    var t=q.value.trim().toLowerCase(), n=0;
    items.forEach(function(it){
      var g=it.getAttribute('data-group');
      var okF = cur==='all' || (cur==='todo' ? todo.indexOf(g)>=0 : (cur==='store' ? it.getAttribute('data-store')==='1' : g===cur));
      var okT = !t || it.textContent.toLowerCase().indexOf(t)>=0;
      it.hidden = !(okF && okT); if(!it.hidden) n++;
    });
    document.getElementById('count').textContent = n + ' 件を表示中';
    btns.forEach(function(b){ b.classList.toggle('active', b.getAttribute('data-filter')===cur); });
    document.getElementById('empty').hidden = n>0;
  }
  btns.forEach(function(b){ b.addEventListener('click', function(){ cur=b.getAttribute('data-filter'); apply(); }); });
  q.addEventListener('input', apply);
  apply();
})();
</script>
</body></html>
'@)
[System.IO.File]::WriteAllText($htmlPath, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))

# --- 画面のまとめ ---
Write-Host ''
Write-Host '================ 結果 ================'
Write-Host ('確認したメール : {0} 件' -f $mailCount)
Write-Host ('請求書メール   : {0} 件' -f $sorted.Count)
foreach ($g in $Groups.Keys) {
    if ($counts[$g] -eq 0) { continue }
    $color = 'Gray'
    if ($g -eq '要注意') { $color = 'Red' } elseif ($g -eq '漏れあり') { $color = 'Yellow' } elseif ($g -eq 'エラー') { $color = 'Magenta' } elseif ($g -eq 'OK') { $color = 'Green' }
    Write-Host ('  {0} : {1} 件' -f $g, $counts[$g]) -ForegroundColor $color
}
Write-Host ('  （保管が未のもの : {0} 件）' -f $storeMissingCount) -ForegroundColor $(if ($storeMissingCount) { 'Yellow' } else { 'Gray' })
if ($fatal) {
    Write-Host ''
    Write-Host "※ $fatal" -ForegroundColor Red
    Write-Host "  この結果は、止まるまでに確認できた $mailCount 件分です。Outlook を開いたまま、もう一度実行してください。" -ForegroundColor Red
}
if ($script:OcrSkipped -gt 0) { Write-Host ("※ 文字認識を使えなかったため、添付 $($script:OcrSkipped) 件（PDF・画像）の中身を確認できませんでした: " + $script:OcrNote) -ForegroundColor Yellow }
Write-Host ''
Write-Host "結果の画面 : $htmlPath"
Write-Host "Excel 用   : $csvPath"
if ($Config.Report.OpenReport) {
    Write-Host '結果の画面をブラウザで開きます...'
    try { Start-Process -FilePath $htmlPath } catch { Write-Host '自動で開けませんでした。上の「結果の画面」のファイルをダブルクリックしてください。' -ForegroundColor Yellow }
}
