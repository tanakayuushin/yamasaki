<#
.SYNOPSIS
  スキャンした請求書・注文請書などのPDF/画像を文字認識(OCR)し、
  「年月_相手の会社名_書類種別.pdf」の形にファイル名を付け直す。

.DESCRIPTION
  - Windows 標準の OCR（Windows.Media.Ocr）と PDF 描画機能（Windows.Data.Pdf）を使用。追加インストール不要。
  - 書類種別 : 本文の先頭付近に出てくる「請求書」「注文請書」などのキーワードで判定
  - 相手の会社名:
      1) config.json の Vendors（取引先マスタ）に登録された名前・別名が本文にあればそれを採用（確度: 高）
      2) なければ「株式会社〇〇」「〇〇(株)」などの表記を抜き出し、自社名と「御中/様」の宛名を除いて推定（確度: 推定）
  - 年月     : 「11月分」→「請求日/発行日」などのラベルの日付 →本文の最初の日付 の順に探す
               （支払期限・振込期日などの日付は除外。和暦「令和」にも対応）
  - 既定は「プレビュー」。新しいファイル名の一覧(CSV)を出すだけで、ファイルは変更しない。
    -Apply を付けたときだけ実際にリネーム／移動する。

  ※ Windows PowerShell 5.1（powershell.exe）専用。PowerShell 7（pwsh）では動きません。

.PARAMETER ConfigPath
  設定ファイル。省略時はスクリプトと同じフォルダの config.json

.PARAMETER InputFolder
  処理するフォルダ。省略時は config.json の Input.Folder

.PARAMETER Apply
  実際にリネーム／移動する。付けなければプレビューのみ。

.EXAMPLE
  .\Rename-ScannedDocs.ps1
  .\Rename-ScannedDocs.ps1 -InputFolder "C:\Scan\20261127"
  .\Rename-ScannedDocs.ps1 -Apply
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = '',
    [string]$InputFolder = '',
    [switch]$Apply
)

Set-StrictMode -Version 1.0
if ($PSVersionTable.PSVersion.Major -lt 3) {
    Write-Host ("[NG] PowerShell のバージョンが古すぎます（" + $PSVersionTable.PSVersion + "）。3.0 以上（Windows 10/11 は標準で 5.1）が必要です。") -ForegroundColor Red
    exit 1
}
# スクリプトのフォルダ（$PSScriptRoot が空になる環境・起動方法にも対応）
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir -and $MyInvocation.MyCommand.Definition) { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }
if (-not $ConfigPath) { $ConfigPath = Join-Path $ScriptDir 'config.json' }
$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSEdition -eq 'Core') {
    throw 'このスクリプトは Windows PowerShell 5.1 で実行してください（Run-ScanRename.bat から起動するか、powershell.exe で実行）。'
}

# ---------------------------------------------------------------------------
# 設定
# ---------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $ConfigPath)) { throw "設定ファイルが見つかりません: $ConfigPath" }
$Config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $InputFolder) { $InputFolder = [string]$Config.Input.Folder }
if (-not (Test-Path -LiteralPath $InputFolder -PathType Container)) { throw "入力フォルダが見つかりません: $InputFolder" }

# ---------------------------------------------------------------------------
# WinRT（OCR / PDF）の準備
# ---------------------------------------------------------------------------
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

$script:AsTaskOp = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
} | Select-Object -First 1
$script:AsTaskAction = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'
} | Select-Object -First 1

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

$lang = New-Object Windows.Globalization.Language ([string]$Config.Ocr.Language)
$script:Ocr = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
if (-not $script:Ocr) {
    throw ("日本語のOCR機能が見つかりません。`n" +
           "設定 → 時刻と言語 → 言語と地域 → 日本語 の「…」→ 言語のオプション → 「光学式文字認識」をインストールしてください。`n" +
           "（管理者 PowerShell で: Add-WindowsCapability -Online -Name 'Language.OCR~~~ja-JP~0.0.1.0'）")
}
$script:MaxDim = [int][Windows.Media.Ocr.OcrEngine]::MaxImageDimension

# ---------------------------------------------------------------------------
# OCR
# ---------------------------------------------------------------------------

# OcrLine → 文字列。単語の間隔が大きいところだけ空白を入れる（日本語は文字ごとに区切られるため）
function Convert-OcrLine {
    param($Line)
    $words = @($Line.Words | Sort-Object { $_.BoundingRect.X })
    $sb = New-Object System.Text.StringBuilder
    $prev = $null
    foreach ($w in $words) {
        if ($prev) {
            $gap = $w.BoundingRect.X - ($prev.BoundingRect.X + $prev.BoundingRect.Width)
            $h = [Math]::Max($prev.BoundingRect.Height, $w.BoundingRect.Height)
            if ($gap -gt ($h * 1.5)) { [void]$sb.Append(' ') }
        }
        [void]$sb.Append($w.Text)
        $prev = $w
    }
    return $sb.ToString()
}

# ストリーム（画像 or PDF描画結果）を OCR して行の配列を返す
function Invoke-OcrStream {
    param($Stream, [string]$Rotation = 'None')
    $Stream.Seek(0)
    $decoder = Wait-WinRt ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($Stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $w = [double]$decoder.PixelWidth; $h = [double]$decoder.PixelHeight
    $scale = [Math]::Min(1.0, $script:MaxDim / [Math]::Max($w, $h))

    $transform = New-Object Windows.Graphics.Imaging.BitmapTransform
    $transform.ScaledWidth  = [uint32][Math]::Floor($w * $scale)
    $transform.ScaledHeight = [uint32][Math]::Floor($h * $scale)
    $transform.InterpolationMode = [Windows.Graphics.Imaging.BitmapInterpolationMode]::Fant
    $transform.Rotation = [Windows.Graphics.Imaging.BitmapRotation]::$Rotation

    $bmp = Wait-WinRt ($decoder.GetSoftwareBitmapAsync(
            [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
            [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied,
            $transform,
            [Windows.Graphics.Imaging.ExifOrientationMode]::RespectExifOrientation,
            [Windows.Graphics.Imaging.ColorManagementMode]::DoNotColorManage)) ([Windows.Graphics.Imaging.SoftwareBitmap])
    try {
        $result = Wait-WinRt ($script:Ocr.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
        return ,@($result.Lines | ForEach-Object { Convert-OcrLine $_ })
    } finally {
        $bmp.Dispose()
    }
}

# 1ページ（ストリーム）を OCR。書類種別が読み取れなければ回転させて再試行
function Invoke-OcrPage {
    param($Stream)
    $rotations = @('None')
    if ($Config.Ocr.TryRotation) { $rotations += @('Clockwise180Degrees', 'Clockwise90Degrees', 'Clockwise270Degrees') }
    $first = $null
    foreach ($r in $rotations) {
        $lines = Invoke-OcrStream $Stream $r
        if ($null -eq $first) { $first = $lines }
        if (Get-DocType (Get-NormalizedLines $lines)) { return ,$lines }
    }
    return ,$first
}

# ファイルを OCR して行の配列を返す（元ファイルをロックしないよう一時コピーを読む）
function Get-OcrLines {
    param([string]$Path, [string]$TempDir)
    $ext = [IO.Path]::GetExtension($Path).ToLower()
    $tmp = Join-Path $TempDir ([Guid]::NewGuid().ToString('N') + $ext)
    Copy-Item -LiteralPath $Path -Destination $tmp
    $sf = Wait-WinRt ([Windows.Storage.StorageFile]::GetFileFromPathAsync($tmp)) ([Windows.Storage.StorageFile])
    $all = @()

    if ($ext -eq '.pdf') {
        $pdf = Wait-WinRt ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($sf)) ([Windows.Data.Pdf.PdfDocument])
        $pages = [Math]::Min([int]$pdf.PageCount, [Math]::Max(1, [int]$Config.Ocr.MaxPages))
        for ($i = 0; $i -lt $pages; $i++) {
            $page = $pdf.GetPage([uint32]$i)
            $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
            try {
                $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions
                $width = [Math]::Min($script:MaxDim, [double]$page.Size.Width * [double]$Config.Ocr.PdfRenderScale)
                $height = [double]$page.Size.Height * ($width / [double]$page.Size.Width)
                if ($height -gt $script:MaxDim) { $width = $width * ($script:MaxDim / $height) }
                $opt.DestinationWidth = [uint32][Math]::Floor($width)
                Wait-WinRtAction ($page.RenderToStreamAsync($stream, $opt))
                $all += Invoke-OcrPage $stream
            } finally {
                $stream.Dispose()
                $page.Dispose()
            }
        }
    } else {
        $stream = Wait-WinRt ($sf.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
        try { $all += Invoke-OcrPage $stream } finally { $stream.Dispose() }
    }
    return ,$all
}

# ---------------------------------------------------------------------------
# テキスト解析
# ---------------------------------------------------------------------------

# 全角英数→半角、㈱→(株)、半角カナ→全角 など
function ConvertTo-Nfkc {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return '' }
    return $s.Normalize([Text.NormalizationForm]::FormKC)
}
function ConvertTo-Compact {
    param([string]$s)
    return ((ConvertTo-Nfkc $s) -replace '\s+', '')
}

# 行ごとに「空白あり版」と「空白なし版」を持つ
function Get-NormalizedLines {
    param([string[]]$Lines)
    $out = @()
    foreach ($l in $Lines) {
        $n = (ConvertTo-Nfkc $l) -replace '\s+', ' '
        $out += [pscustomobject]@{ Spaced = $n.Trim(); Compact = ($n -replace '\s', '') }
    }
    return ,$out
}

# 書類種別: キーワードが本文中で一番早く出てくる種別（＝タイトル）
function Get-DocType {
    param($NLines)
    $text = (@($NLines | ForEach-Object { $_.Compact }) -join "`n")
    if (-not $text) { return $null }
    $best = $null; $bestPos = [int]::MaxValue
    foreach ($dt in @($Config.DocTypes)) {
        foreach ($kw in @($dt.Keywords)) {
            $k = ConvertTo-Compact $kw
            $pos = $text.IndexOf($k, [StringComparison]::OrdinalIgnoreCase)
            if ($pos -ge 0 -and $pos -lt $bestPos) { $best = $dt; $bestPos = $pos }
        }
    }
    return $best
}

$script:OwnNames = @(@($Config.OwnCompanyNames) | Where-Object { $_ } | ForEach-Object { ConvertTo-Compact $_ })
function Test-OwnCompany {
    param([string]$Name)
    $c = ConvertTo-Compact $Name
    if (-not $c) { return $false }
    foreach ($own in $script:OwnNames) {
        $ownCore = Get-CompanyCore $own
        $cCore = Get-CompanyCore $c
        if ($c -eq $own -or ($ownCore -and $cCore -and ($cCore.Contains($ownCore) -or $ownCore.Contains($cCore)))) { return $true }
    }
    return $false
}

$CorpForm = '株式会社|有限会社|合同会社|合資会社|合名会社|一般社団法人|一般財団法人|公益社団法人|公益財団法人|社会福祉法人|医療法人(?:社団|財団)?|特定非営利活動法人|NPO法人|\(株\)|\(有\)|\(合\)'
function Get-CompanyCore {
    param([string]$Name)
    $c = ($Name -replace $CorpForm, '') -replace '\s', ''
    if ($c.Length -lt 2) { return $null }
    return $c
}

$Prefectures = '北海道|青森県|岩手県|宮城県|秋田県|山形県|福島県|茨城県|栃木県|群馬県|埼玉県|千葉県|東京都|神奈川県|新潟県|富山県|石川県|福井県|山梨県|長野県|岐阜県|静岡県|愛知県|三重県|滋賀県|京都府|大阪府|兵庫県|奈良県|和歌山県|鳥取県|島根県|岡山県|広島県|山口県|徳島県|香川県|愛媛県|高知県|福岡県|佐賀県|長崎県|熊本県|大分県|宮崎県|鹿児島県|沖縄県'
$NameStop  = "御中|様|殿|〒|TEL|Tel|tel|FAX|Fax|電話|代表|登録番号|T\d{13}|住所|担当|$Prefectures|\s|$"
$NameChars = '[^\s、。,:;()\[\]「」『』【】<>]'
$RePrefixCorp = "(?<!$NameChars)(?:$CorpForm)\s?(?!$NameStop)$NameChars{1,20}?(?=$NameStop)"
$ReSuffixCorp = "$NameChars{1,20}?\s?(?:$CorpForm)"
$LeadingLabels = '^(?:請求元|発行元|発行者|会社名|社名|貴社名|取引先|宛先|請求先|受注者|発注者|売主|買主|商号|名称)'

# 相手の会社名を決める
function Get-Counterparty {
    param($NLines, $DocType)
    $compactText = (@($NLines | ForEach-Object { $_.Compact }) -join "`n")

    # 1) 取引先マスタ
    $hits = @()
    foreach ($v in @($Config.Vendors)) {
        $names = @($v.Name) + @($v.Aliases) | Where-Object { $_ }
        foreach ($n in $names) {
            $k = ConvertTo-Compact $n
            if (-not $k) { continue }
            $pos = $compactText.IndexOf($k, [StringComparison]::OrdinalIgnoreCase)
            if ($pos -ge 0) {
                # その行が「御中」「様」の宛名なら優先度を下げる
                $lineIdx = ($compactText.Substring(0, $pos) -split "`n").Count - 1
                $isAddr = $NLines[$lineIdx].Compact -match '御中|様$|様\s|殿'
                $hits += [pscustomobject]@{ Name = [string]$v.Name; Pos = $pos; Addr = $isAddr }
                break
            }
        }
    }
    $hits = @($hits | Where-Object { -not (Test-OwnCompany $_.Name) })
    if ($hits.Count -gt 0) {
        $wantAddr = ($DocType -and $DocType.Party -eq 'addressee')
        $pick = $hits | Sort-Object @{ Expression = { $_.Addr -ne $wantAddr } }, Pos | Select-Object -First 1
        return [pscustomobject]@{ Name = $pick.Name; Source = '取引先マスタ'; Confident = $true }
    }

    # 2) 「株式会社〇〇」表記から推定
    $cands = @()
    for ($i = 0; $i -lt $NLines.Count; $i++) {
        $line = $NLines[$i].Spaced
        $ms = @([regex]::Matches($line, $RePrefixCorp)) + @([regex]::Matches($line, $ReSuffixCorp))
        foreach ($m in $ms) {
            $name = ($m.Value -replace '\s', '') -replace $LeadingLabels, ''
            if (-not (Get-CompanyCore $name)) { continue }
            if (Test-OwnCompany $name) { continue }
            $after = $line.Substring($m.Index + $m.Length)
            $isAddr = $after -match '^\s?(御中|様|殿)'
            # 発行者らしさ: 近くに 〒 / TEL / 登録番号(T+13桁) / 代表 がある
            $block = (@($NLines[$i..([Math]::Min($NLines.Count - 1, $i + 4))] | ForEach-Object { $_.Compact }) -join ' ')
            $score = 0
            if ($block -match '〒|TEL|電話|FAX|登録番号|T\d{13}|代表取締役') { $score += 3 }
            if ($isAddr) { $score -= 5 }
            $cands += [pscustomobject]@{ Name = $name; Line = $i; Score = $score; Addr = $isAddr }
        }
    }
    if ($cands.Count -eq 0) { return $null }
    if ($DocType -and $DocType.Party -eq 'addressee') {
        $pick = $cands | Sort-Object @{ Expression = { -not $_.Addr } }, Line | Select-Object -First 1
    } else {
        $pick = $cands | Sort-Object @{ Expression = { $_.Score }; Descending = $true }, Line | Select-Object -First 1
    }
    return [pscustomobject]@{ Name = $pick.Name; Source = '本文から推定'; Confident = $false }
}

# 日付の抽出 --------------------------------------------------------------
$DatePatterns = @(
    '(?<![\d\-/.])(?<y>20\d{2})年(?<m>\d{1,2})月(?:(?<d>\d{1,2})日)?',
    '令和(?<r>\d{1,2}|元)年(?<m>\d{1,2})月(?:(?<d>\d{1,2})日)?',
    '(?<![\d\-/.])(?<y>20\d{2})[/.\-](?<m>\d{1,2})(?:[/.\-](?<d>\d{1,2}))?(?![\d\-/])',
    '(?<![A-Za-z\d])R(?<r>\d{1,2})[/.\-](?<m>\d{1,2})(?:[/.\-](?<d>\d{1,2}))?(?![\d\-/])'
)
$ReMonthBun = '(?:(?<y>20\d{2})年|令和(?<r>\d{1,2}|元)年)?(?<m>\d{1,2})月分'

function Convert-DateMatch {
    param($m)
    $y = $null
    if ($m.Groups['y'].Success) { $y = [int]$m.Groups['y'].Value }
    elseif ($m.Groups['r'].Success) {
        $r = if ($m.Groups['r'].Value -eq '元') { 1 } else { [int]$m.Groups['r'].Value }
        $y = 2018 + $r
    }
    $mo = [int]$m.Groups['m'].Value
    if ($mo -lt 1 -or $mo -gt 12) { return $null }
    if ($y -and ($y -lt 2000 -or $y -gt 2100)) { return $null }
    return [pscustomobject]@{ Year = $y; Month = $mo }
}

function Find-DatesInLine {
    param([string]$Compact)
    $found = @()
    foreach ($p in $DatePatterns) {
        foreach ($m in [regex]::Matches($Compact, $p)) {
            $d = Convert-DateMatch $m
            if ($d -and $d.Year) { $found += [pscustomobject]@{ Year = $d.Year; Month = $d.Month; Pos = $m.Index } }
        }
    }
    return ,@($found | Sort-Object Pos)
}

function Get-YearMonth {
    param($NLines)
    $labels   = @($Config.Date.Labels | ForEach-Object { ConvertTo-Compact $_ })
    $excludes = @($Config.Date.ExcludeLabels | ForEach-Object { ConvertTo-Compact $_ })

    $labelDate = $null; $firstDate = $null; $bun = $null
    for ($i = 0; $i -lt $NLines.Count; $i++) {
        $c = $NLines[$i].Compact
        $excluded = @($excludes | Where-Object { $c.Contains($_) }).Count -gt 0

        if (-not $bun) {
            $m = [regex]::Match($c, $ReMonthBun)
            if ($m.Success) { $bun = Convert-DateMatch $m }
        }
        if ($excluded) { continue }

        $dates = Find-DatesInLine $c
        if (-not $firstDate -and $dates.Count -gt 0) { $firstDate = $dates[0] }

        if (-not $labelDate -and (@($labels | Where-Object { $c.Contains($_) }).Count -gt 0)) {
            if ($dates.Count -gt 0) { $labelDate = $dates[0] }
            elseif ($i + 1 -lt $NLines.Count) {
                $next = Find-DatesInLine $NLines[$i + 1].Compact
                if ($next.Count -gt 0) { $labelDate = $next[0] }
            }
        }
    }

    $ref = if ($labelDate) { $labelDate } else { $firstDate }
    foreach ($p in @($Config.Date.Priority)) {
        switch ($p) {
            '月分' {
                if ($bun) {
                    $y = $bun.Year
                    if (-not $y) {
                        # 年が書かれていなければ、発行日などから補う（12月分を1月に発行 → 前年）
                        $baseY = if ($ref) { $ref.Year } else { (Get-Date).Year }
                        $baseM = if ($ref) { $ref.Month } else { (Get-Date).Month }
                        $y = if ($bun.Month -gt $baseM) { $baseY - 1 } else { $baseY }
                    }
                    return [pscustomobject]@{ Date = (Get-Date -Year $y -Month $bun.Month -Day 1); Source = "「$($bun.Month)月分」の記載"; ApplyOffset = $false }
                }
            }
            'ラベル' {
                if ($labelDate) { return [pscustomobject]@{ Date = (Get-Date -Year $labelDate.Year -Month $labelDate.Month -Day 1); Source = '請求日・発行日などの日付'; ApplyOffset = $true } }
            }
            '最初の日付' {
                if ($firstDate) { return [pscustomobject]@{ Date = (Get-Date -Year $firstDate.Year -Month $firstDate.Month -Day 1); Source = '本文の最初の日付'; ApplyOffset = $true } }
            }
        }
    }
    return $null
}

# ---------------------------------------------------------------------------
# ファイル名
# ---------------------------------------------------------------------------
function Get-SafeName {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $s = -join ($Name.ToCharArray() | ForEach-Object { if ($invalid -contains $_) { '_' } else { $_ } })
    return $s.Trim().TrimEnd('.')
}

function Expand-Pattern {
    param([string]$Pattern, $Date, [string]$Company, [string]$DocType, [string]$Original)
    $s = $Pattern
    if ($Date) {
        $s = $s.Replace('{yyyy}', $Date.ToString('yyyy')).Replace('{yy}', $Date.ToString('yy'))
        $s = $s.Replace('{MM}', $Date.ToString('MM')).Replace('{M}', [string]$Date.Month)
    }
    $s = $s.Replace('{company}', $Company).Replace('{doctype}', $DocType).Replace('{original}', $Original)
    return $s
}

$script:Reserved = @{}
function Get-UniquePath {
    param([string]$Folder, [string]$BaseName, [string]$Ext, [string]$SourcePath)
    $n = 1
    while ($true) {
        $name = if ($n -eq 1) { "$BaseName$Ext" } else { "${BaseName}_$n$Ext" }
        $path = Join-Path $Folder $name
        $key = $path.ToLower()
        $isSelf = ($path -ieq $SourcePath)
        if ($isSelf -or (-not (Test-Path -LiteralPath $path) -and -not $script:Reserved.ContainsKey($key))) {
            $script:Reserved[$key] = $true
            return $path
        }
        $n++
    }
}

# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------
$tempDir = Join-Path ([IO.Path]::GetTempPath()) ('scan-rename-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$logDir = [string]$Config.Log.Folder
if (-not $logDir) { $logDir = $ScriptDir }
if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$ocrDir = Join-Path $logDir "ocr_text_$stamp"
if ($Config.Log.SaveOcrText) { New-Item -ItemType Directory -Path $ocrDir -Force | Out-Null }

$exts = @($Config.Input.Extensions | ForEach-Object { $_.ToLower() })
$files = @(Get-ChildItem -LiteralPath $InputFolder -File | Where-Object { $exts -contains $_.Extension.ToLower() } | Sort-Object Name)
$skipRe = [string]$Config.Input.SkipIfNameMatches
$targets = @($Config.TargetDocTypes)

$results = New-Object System.Collections.Generic.List[object]
$mode = if ($Apply) { '実行' } else { 'プレビュー（ファイルは変更しません）' }
Write-Host "モード: $mode"
Write-Host "対象  : $InputFolder（$($files.Count) ファイル）"

$idx = 0
try {
    foreach ($f in $files) {
        $idx++
        $row = [pscustomobject][ordered]@{
            元ファイル名 = $f.Name; 新ファイル名 = ''; 判定 = ''; 書類種別 = ''; 相手先 = ''; 相手先の根拠 = ''
            年月 = ''; 年月の根拠 = ''; 移動先 = ''; 備考 = ''; OCR先頭 = ''
        }
        Write-Progress -Activity 'OCR 処理中' -Status "$idx / $($files.Count): $($f.Name)" -PercentComplete (100 * $idx / [Math]::Max(1, $files.Count))

        if ($skipRe -and $f.BaseName -match $skipRe) {
            $row.判定 = 'スキップ'; $row.備考 = 'リネーム済みの名前のため'
            $results.Add($row); continue
        }

        try {
            $lines = Get-OcrLines $f.FullName $tempDir
            if ($Config.Log.SaveOcrText) {
                Set-Content -LiteralPath (Join-Path $ocrDir ($f.Name + '.txt')) -Value $lines -Encoding UTF8
            }
            $nl = Get-NormalizedLines $lines
            $row.OCR先頭 = ((@($nl | ForEach-Object { $_.Compact }) -join ' ') -replace '^(.{0,120}).*$', '$1')

            $dt = Get-DocType $nl
            if (-not $dt) { $row.判定 = '要確認'; $row.備考 = '書類種別を読み取れませんでした'; $results.Add($row); continue }
            $row.書類種別 = [string]$dt.Name
            if ($targets.Count -gt 0 -and $targets -notcontains $dt.Name) {
                $row.判定 = '対象外'; $row.備考 = "対象の書類種別ではありません（$($dt.Name)）"
                $results.Add($row); continue
            }

            $cp = Get-Counterparty $nl $dt
            $ym = Get-YearMonth $nl
            if ($cp) { $row.相手先 = $cp.Name; $row.相手先の根拠 = $cp.Source }
            if ($ym) {
                $d = $ym.Date
                if ($ym.ApplyOffset -and [int]$Config.Date.MonthOffset -ne 0) { $d = $d.AddMonths([int]$Config.Date.MonthOffset) }
                $row.年月 = $d.ToString('yyyy/MM'); $row.年月の根拠 = $ym.Source
            }

            $missing = @()
            if (-not $cp) { $missing += '相手先' }
            if (-not $ym) { $missing += '年月' }
            if ($missing.Count -gt 0) {
                $row.判定 = '要確認'; $row.備考 = (($missing -join '・') + 'を読み取れませんでした')
            } else {
                $company = $cp.Name
                if ($Config.Naming.StripCorporateForm) { $c2 = Get-CompanyCore $company; if ($c2) { $company = $c2 } }
                $base = Get-SafeName (Expand-Pattern ([string]$Config.Naming.Template) $d $company $dt.Name $f.BaseName)
                $row.判定 = if ($cp.Confident) { 'OK' } else { '推定' }

                $outRoot = if ($Config.Output.Folder) { [string]$Config.Output.Folder } else { $f.DirectoryName }
                # SubfolderPattern の「\」は階層として扱う
                $sub = ''
                if ($Config.Output.SubfolderPattern) {
                    $sub = (@(([string](Expand-Pattern ([string]$Config.Output.SubfolderPattern) $d $company $dt.Name $f.BaseName)) -split '[\\/]' | ForEach-Object { Get-SafeName $_ }) -join '\')
                }
                $destDir = if ($sub) { Join-Path $outRoot $sub } else { $outRoot }
                $destPath = Get-UniquePath $destDir $base $f.Extension.ToLower() $f.FullName
                $row.新ファイル名 = [IO.Path]::GetFileName($destPath)
                $row.移動先 = $destDir
            }

            # 実行
            if ($Apply) {
                $doRename = ($row.判定 -eq 'OK') -or ($row.判定 -eq '推定' -and $Config.Output.RenameEstimated)
                if ($doRename -and $row.新ファイル名) {
                    $destPath = Join-Path $row.移動先 $row.新ファイル名
                    if ($destPath -ine $f.FullName) {
                        if (-not (Test-Path -LiteralPath $row.移動先)) { New-Item -ItemType Directory -Path $row.移動先 -Force | Out-Null }
                        Move-Item -LiteralPath $f.FullName -Destination $destPath
                    }
                    $row.備考 = (($row.備考, '変更しました') | Where-Object { $_ }) -join ' / '
                } elseif ($row.判定 -eq '要確認' -and $Config.Output.ReviewFolder) {
                    $rv = [string]$Config.Output.ReviewFolder
                    if (-not (Test-Path -LiteralPath $rv)) { New-Item -ItemType Directory -Path $rv -Force | Out-Null }
                    $rp = Get-UniquePath $rv $f.BaseName $f.Extension $f.FullName
                    Move-Item -LiteralPath $f.FullName -Destination $rp
                    $row.移動先 = $rv
                    $row.備考 = (($row.備考, '要確認フォルダへ移動しました') | Where-Object { $_ }) -join ' / '
                }
            }
        } catch {
            $row.判定 = 'エラー'; $row.備考 = $_.Exception.Message
        }
        $results.Add($row)
    }
} finally {
    Write-Progress -Activity 'OCR 処理中' -Completed
    Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------------
# 結果
# ---------------------------------------------------------------------------
$csv = Join-Path $logDir ("リネーム結果_$stamp.csv")
$results | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8   # PS5.1 は BOM 付き → Excel で文字化けしない

Write-Host ''
Write-Host '==== 結果 ===='
$results | Group-Object 判定 | ForEach-Object { Write-Host ('  {0,-6}: {1}' -f $_.Name, $_.Count) }
Write-Host ''
$results | Where-Object { $_.判定 -ne 'スキップ' } |
    Format-Table 判定, 元ファイル名, 新ファイル名, 備考 -AutoSize | Out-String -Width 220 | Write-Host
Write-Host "一覧: $csv"
if (-not $Apply) {
    Write-Host '※ プレビューです。内容がよければ -Apply を付けて（または Run-ScanRename-Apply.bat で）実行してください。' -ForegroundColor Cyan
}
