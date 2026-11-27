<#
.SYNOPSIS
  請求書メールチェックツールがこのPCで動くかを確認する（読み取りのみ。何も変更しません）。
  結果は画面と「環境チェック結果_日時.txt」に出力します。
#>
[CmdletBinding()]
param([string]$ConfigPath = '')

$ErrorActionPreference = 'Continue'
if ($PSVersionTable.PSVersion.Major -lt 3) {
    Write-Host ("[NG] PowerShell のバージョンが古すぎます（" + $PSVersionTable.PSVersion + "）。3.0 以上（Windows 10/11 は標準で 5.1）が必要です。") -ForegroundColor Red
    exit 1
}
# スクリプトのフォルダ（$PSScriptRoot が空になる環境・起動方法にも対応）
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir -and $MyInvocation.MyCommand.Definition) { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }
if (-not $ConfigPath) { $ConfigPath = Join-Path $ScriptDir 'config.json' }
$lines = New-Object System.Collections.Generic.List[string]
$script:ng = 0
function Out-Result {
    param([string]$Status, [string]$Item, [string]$Detail = '')
    $mark = switch ($Status) { 'OK' { '[OK]  ' } 'NG' { '[NG]  ' } 'WARN' { '[注意]' } default { '[情報]' } }
    if ($Status -eq 'NG') { $script:ng++ }
    $text = "$mark $Item" + $(if ($Detail) { " : $Detail" } else { '' })
    $color = switch ($Status) { 'OK' { 'Green' } 'NG' { 'Red' } 'WARN' { 'Yellow' } default { 'Gray' } }
    Write-Host $text -ForegroundColor $color
    $lines.Add($text)
}
function Out-Section { param([string]$Title) $t = "`n==== $Title ===="; Write-Host $t -ForegroundColor Cyan; $lines.Add($t) }

# ---------------------------------------------------------------------------
Out-Section 'Windows / PowerShell'
try {
    $os = Get-CimInstance Win32_OperatingSystem
    Out-Result 'INFO' 'Windows' "$($os.Caption) ($($os.Version))"
} catch { Out-Result 'INFO' 'Windows' ([Environment]::OSVersion.VersionString) }
Out-Result 'INFO' 'PowerShell' "$($PSVersionTable.PSVersion) / $($PSVersionTable.PSEdition)"

$lm = $ExecutionContext.SessionState.LanguageMode
if ($lm -eq 'FullLanguage') { Out-Result 'OK' 'PowerShell の言語モード' $lm }
else { Out-Result 'NG' 'PowerShell の言語モード' "$lm（会社のセキュリティ設定で制限されています。Outlook を操作できないため情報システム担当に相談が必要）" }

try {
    $pol = Get-ExecutionPolicy -List
    $mp = ($pol | Where-Object Scope -eq 'MachinePolicy').ExecutionPolicy
    $up = ($pol | Where-Object Scope -eq 'UserPolicy').ExecutionPolicy
    $detail = ($pol | ForEach-Object { "$($_.Scope)=$($_.ExecutionPolicy)" }) -join ', '
    if (@('AllSigned', 'Restricted') -contains [string]$mp -or @('AllSigned', 'Restricted') -contains [string]$up) {
        Out-Result 'NG' '実行ポリシー' "$detail（グループポリシーで制限。bat からの起動もブロックされる可能性）"
    } else { Out-Result 'OK' '実行ポリシー' $detail }
} catch { Out-Result 'WARN' '実行ポリシー' $_.Exception.Message }

# ---------------------------------------------------------------------------
Out-Section 'Outlook'
$progId = $null
try { $progId = (Get-ItemProperty 'Registry::HKEY_CLASSES_ROOT\Outlook.Application\CurVer' -ErrorAction Stop).'(default)' } catch { }
if ($progId) { Out-Result 'OK' 'Outlook（デスクトップ版）のインストール' $progId }
else { Out-Result 'NG' 'Outlook（デスクトップ版）のインストール' '見つかりません。新しい Outlook / Web版 Outlook では動きません' }

$classic = Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue
$newOl   = Get-Process -Name olk -ErrorAction SilentlyContinue
if ($classic) { Out-Result 'OK' '起動中の Outlook' '従来の Outlook（OUTLOOK.EXE）' }
elseif ($newOl) { Out-Result 'NG' '起動中の Outlook' '新しい Outlook（olk.exe）です。右上の「新しい Outlook」をオフにして従来版で起動してください' }
else { Out-Result 'WARN' '起動中の Outlook' '起動していません（チェック時は自動で起動を試みます）' }

$Config = $null
if (Test-Path -LiteralPath $ConfigPath) {
    try { $Config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json; Out-Result 'OK' 'config.json の読み込み' }
    catch { Out-Result 'NG' 'config.json の読み込み' "書式エラー: $($_.Exception.Message)" }
} else { Out-Result 'NG' 'config.json' "見つかりません: $ConfigPath（config.sample.json をコピーして作ってください）" }

$ns = $null
if ($progId -and $lm -eq 'FullLanguage') {
    try {
        Write-Host '  Outlook に接続中...（「プログラムによるアクセス」の警告が出たら「許可」を押してください）'
        $ol = New-Object -ComObject Outlook.Application
        $ns = $ol.GetNamespace('MAPI')
        $inbox = $ns.GetDefaultFolder(6)
        Out-Result 'OK' 'Outlook への接続' "既定のメールボックス: $($inbox.Parent.Name) / 受信トレイ名: $($inbox.Name)"
    } catch { Out-Result 'NG' 'Outlook への接続' $_.Exception.Message }
}

function Resolve-Folder {
    param($Namespace, [string]$Mailbox, [string]$Path)
    $cur = if ($Mailbox) { $Namespace.Folders.Item($Mailbox) } else { $Namespace.GetDefaultFolder(6).Parent }
    foreach ($p in ($Path -split '[\\/]')) { if ($p) { $cur = $cur.Folders.Item($p) } }
    return $cur
}

if ($ns -and $Config) {
    $mb = [string]$Config.Outlook.Mailbox
    foreach ($fp in @($Config.Outlook.Folders)) {
        try {
            $f = Resolve-Folder $ns $mb $fp
            $since = (Get-Date).Date.AddDays(-[int]$Config.Outlook.DaysBack)
            $cnt = $f.Items.Restrict("[ReceivedTime] >= '" + $since.ToString('g') + "'").Count
            Out-Result 'OK' "対象フォルダ「$fp」" "過去$($Config.Outlook.DaysBack)日のメール $cnt 件"

            # 1通目で Message-ID / 返信マークが取れるか
            $sample = $null
            foreach ($it in $f.Items) { if ($it.Class -eq 43) { $sample = $it; break } }
            if ($sample) {
                try {
                    $mid = $sample.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001F')
                    if ($mid) { Out-Result 'OK' 'メールの識別番号（Message-ID）の取得' } else { Out-Result 'WARN' 'メールの識別番号（Message-ID）の取得' '空でした（代わりに件名＋受信日時で突き合わせます）' }
                } catch { Out-Result 'WARN' 'メールの識別番号（Message-ID）の取得' $_.Exception.Message }
            }
        } catch { Out-Result 'NG' "対象フォルダ「$fp」" "見つかりません（$($_.Exception.Message)）" }
    }

    try {
        $cr = Resolve-Folder $ns $mb ([string]$Config.Outlook.CopyRootFolder)
        $subs = @($cr.Folders | ForEach-Object { $_.Name })
        Out-Result 'OK' "コピー先「$($Config.Outlook.CopyRootFolder)」" "サブフォルダ $($subs.Count) 個"
        Out-Result 'INFO' '  サブフォルダ名（最大20件）' (($subs | Select-Object -First 20) -join ' / ')
        $spPat = [string]$Config.Rules.SpecialProject.CopyFolderPattern
        $spLabel = if ($Config.Rules.SpecialProject.Label) { [string]$Config.Rules.SpecialProject.Label } else { '特定案件' }
        $normal = @($subs | Where-Object { $_ -match [string]$Config.Rules.NormalCopyFolderPattern -and -not ($spPat -and $_ -match $spPat) })
        if ($normal.Count -gt 0) { Out-Result 'OK' '  「○○月分」フォルダの名前の形式' "$($normal.Count) 個が一致" } else { Out-Result 'WARN' '  「○○月分」フォルダの名前の形式' '一致するフォルダがありません（設定の調整が必要）' }
        if ($spPat) {
            $sp = @($subs | Where-Object { $_ -match $spPat })
            if ($sp.Count -gt 0) { Out-Result 'OK' "  $spLabel 用フォルダの名前の形式" "$($sp.Count) 個が一致" } else { Out-Result 'WARN' "  $spLabel 用フォルダの名前の形式" '一致するフォルダがありません（今は該当がないだけなら問題なし）' }
        }
    } catch { Out-Result 'NG' "コピー先「$($Config.Outlook.CopyRootFolder)」" "見つかりません（$($_.Exception.Message)）" }

    # 検索フォルダー（SearchFolderName）
    $sfName = [string]$Config.Outlook.SearchFolderName
    if ($sfName) {
        $sf = $null
        $stores = @()
        if ($mb) { try { $stores += $ns.Folders.Item($mb).Store } catch { } } else { try { $stores += $ns.DefaultStore } catch { } }
        try { foreach ($st in $ns.Stores) { $stores += $st } } catch { }
        foreach ($st in $stores) {
            if ($sf) { break }
            try { foreach ($f in $st.GetSearchFolders()) { if ([string]$f.Name -eq $sfName) { $sf = $f; break } } } catch { }
        }
        if ($sf) {
            $since = (Get-Date).Date.AddDays(-[int]$Config.Outlook.DaysBack)
            $names = @{}; $n = 0
            # 直近の日数に絞ってから数える（全部を1通ずつ読むと、とても時間がかかるため）
            try { $sfItems = $sf.Items.Restrict("[ReceivedTime] >= '" + $since.ToString('g') + "'") } catch { $sfItems = $sf.Items }
            $sortedNewFirst = $false
            try { $sfItems.Sort('[ReceivedTime]', $true); $sortedNewFirst = $true } catch { }
            foreach ($it in $sfItems) {
                $stop = $false
                try {
                    if ($it.Class -eq 43) {
                        if ([datetime]$it.ReceivedTime -lt $since) { $stop = $sortedNewFirst }
                        else {
                            $n++
                            $fn = [string]$it.Parent.Name
                            if ($names.ContainsKey($fn)) { $names[$fn]++ } else { $names[$fn] = 1 }
                        }
                    }
                } catch { }
                if ($stop -or $n -ge 2000) { break }
            }
            Out-Result 'OK' "検索フォルダー「$sfName」" "直近 $($Config.Outlook.DaysBack) 日のメール $n 件"
            Out-Result 'INFO' '  入っているフォルダ（件数）' ((@($names.Keys | Sort-Object | ForEach-Object { "$_($($names[$_]))" })) -join ' / ')
        } else {
            Out-Result 'WARN' "検索フォルダー「$sfName」" '見つかりません。名前が違う場合は config.json の SearchFolderName を直してください（見つからないときは受信トレイを調べます）'
        }
    }

    try {
        $cats = @($ns.Categories | ForEach-Object { $_.Name })
        Out-Result 'INFO' '登録されている分類項目' ($cats -join ' / ')
        $staff = @($Config.Rules.Staff | Where-Object { $_ })
        if ($staff.Count -eq 0) { Out-Result 'WARN' '確認者の名前（Rules.Staff）' '未設定です。上の分類項目から確認者の名前を config.json に入れてください' }
        else {
            $missing = @($staff | Where-Object { $cats -notcontains $_ })
            if ($missing.Count) { Out-Result 'WARN' '確認者の名前（Rules.Staff）' "分類項目にない名前: $($missing -join ', ')" } else { Out-Result 'OK' '確認者の名前（Rules.Staff）' ($staff -join ', ') }
        }
    } catch { Out-Result 'WARN' '分類項目の一覧' $_.Exception.Message }
}

# ---------------------------------------------------------------------------
Out-Section '請求書PDFの格納場所'
if ($Config) {
    $root = [string]$Config.Storage.RootPath
    if (Test-Path -LiteralPath $root -PathType Container) {
        Out-Result 'OK' '格納場所にアクセス' $root
        $years = @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
        Out-Result 'INFO' '  年フォルダ（新しい順・最大5件）' (($years | Select-Object -First 5 | ForEach-Object Name) -join ' / ')
        $ty = $years | Where-Object { $_.Name -match ('^' + (Get-Date).Year + '(?!\d)') } | Select-Object -First 1
        if ($ty) {
            $months = @(Get-ChildItem -LiteralPath $ty.FullName -Directory -ErrorAction SilentlyContinue | ForEach-Object Name)
            Out-Result 'OK' "  今年のフォルダ「$($ty.Name)」" "中のフォルダ: $($months -join ' / ')"
            $mf = Get-ChildItem -LiteralPath $ty.FullName -Directory -ErrorAction SilentlyContinue | Select-Object -Last 1
            if ($mf) {
                $subs = @(Get-ChildItem -LiteralPath $mf.FullName -Directory -ErrorAction SilentlyContinue | ForEach-Object Name)
                $pdfs = @(Get-ChildItem -LiteralPath $mf.FullName -Filter *.pdf -File -Recurse -ErrorAction SilentlyContinue)
                Out-Result 'INFO' "  「$($mf.Name)」の中" "PDF $($pdfs.Count) 件 / サブフォルダ: $($subs -join ' / ')"
                Out-Result 'INFO' '  PDFファイル名の例（最大3件）' (($pdfs | Select-Object -First 3 | ForEach-Object Name) -join ' / ')
            }
        } else { Out-Result 'WARN' '  今年の年フォルダ' '見つかりません（フォルダ名の形式が想定と違う可能性）' }
    } else {
        Out-Result 'NG' '格納場所にアクセス' "$root にアクセスできません（ドライブ未接続・権限なし・パス違い）"
    }
}

# ---------------------------------------------------------------------------
Out-Section '添付の中身の確認'
if ($PSVersionTable.PSEdition -eq 'Core') {
    Out-Result 'WARN' '日本語の文字認識（OCR）' 'PowerShell 7 では使えません（bat から起動すれば使えます）'
} else {
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime
        $null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime]
        $null = [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime]
        $eng = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage((New-Object Windows.Globalization.Language 'ja'))
        if ($eng) { Out-Result 'OK' '日本語の文字認識（OCR）' 'PDF・画像の添付の中身を確認できます' }
        else { Out-Result 'WARN' '日本語の文字認識（OCR）' '未インストール。設定 → 時刻と言語 → 言語と地域 → 日本語の「…」→ 言語のオプション →「光学式文字認識」を追加すると、PDF・画像の中身も確認できます' }
    } catch { Out-Result 'WARN' '日本語の文字認識（OCR）' "使えません（$($_.Exception.Message)）" }
}
try {
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    Out-Result 'OK' 'Excel・Word・ZIP の中身の読み取り'
} catch { Out-Result 'WARN' 'Excel・Word・ZIP の中身の読み取り' $_.Exception.Message }

Out-Section 'その他'
# 結果の保存先（本体と同じ決め方。既定はひとつ上の「結果」フォルダ）
$outDir = if ($Config) { [string]$Config.Report.OutputFolder } else { '' }
if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path (Split-Path $ScriptDir -Parent) '結果' }
elseif (-not [IO.Path]::IsPathRooted($outDir)) { $outDir = [IO.Path]::GetFullPath((Join-Path $ScriptDir $outDir)) }
try { if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null } } catch { }
try {
    $t = Join-Path ([IO.Path]::GetTempPath()) ('invoice-check-test-' + [Guid]::NewGuid().ToString('N') + '.tmp')
    Set-Content -LiteralPath $t -Value 'test'; Remove-Item -LiteralPath $t
    Out-Result 'OK' '一時フォルダへの書き込み' ([IO.Path]::GetTempPath())
} catch { Out-Result 'NG' '一時フォルダへの書き込み' $_.Exception.Message }
try {
    $t = Join-Path $outDir ('write-test-' + [Guid]::NewGuid().ToString('N') + '.tmp')
    Set-Content -LiteralPath $t -Value 'test'; Remove-Item -LiteralPath $t
    Out-Result 'OK' '結果の保存先への書き込み' $outDir
} catch { Out-Result 'NG' '結果の保存先への書き込み' $_.Exception.Message }

if ($ns) {
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns)
}

# ---------------------------------------------------------------------------
$summary = if ($script:ng -eq 0) { "`n結果: 問題なし（[注意] の項目は確認してください）" } else { "`n結果: [NG] が $($script:ng) 件あります" }
Write-Host $summary -ForegroundColor $(if ($script:ng) { 'Red' } else { 'Green' })
$lines.Add($summary)
$out = Join-Path $outDir ('環境チェック結果_' + (Get-Date).ToString('yyyyMMdd_HHmmss') + '.txt')
try { Set-Content -LiteralPath $out -Value $lines -Encoding UTF8; Write-Host "保存しました: $out" } catch { }
