<#
.SYNOPSIS
  結果の画面の「Outlookで開く」から呼ばれ、そのメールを Outlook で開く（表示する）だけの補助スクリプト。
  リンクの形：invoicecheck-mail:<EntryID>?store=<StoreID>
  形が正しくないものは何もしない。メールの変更・送信・削除はしない。
#>
param([string]$Uri = '')

$ErrorActionPreference = 'Stop'
try {
    $u = [Uri]::UnescapeDataString([string]$Uri).Trim()
    $m = [regex]::Match($u, '^invoicecheck-mail:(?://)?([0-9A-Fa-f]{20,})(?:\?store=([0-9A-Fa-f]{20,}))?/?$')
    if (-not $m.Success) { throw 'リンクの形が正しくありません。' }
    $entryId = $m.Groups[1].Value
    $storeId = $m.Groups[2].Value

    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace('MAPI')
    if ($storeId) { $item = $ns.GetItemFromID($entryId, $storeId) } else { $item = $ns.GetItemFromID($entryId) }
    $item.Display()
    try { $item.GetInspector.Activate() } catch { }
}
catch {
    Add-Type -AssemblyName PresentationFramework
    [void][System.Windows.MessageBox]::Show(
        "メールを開けませんでした。`n（メールが削除・移動された、または Outlook が開いていない可能性があります）`n`n$($_.Exception.Message)",
        '請求書メールチェック')
}
