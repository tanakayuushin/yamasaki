<#
.SYNOPSIS
  結果の画面の「Outlookで開く」を使えるようにする準備（最初に1回だけ）。
  このユーザーの設定（HKEY_CURRENT_USER）に、「invoicecheck-mail:」で始まるリンクを Open-Mail.ps1 で開く登録を追加する。
  管理者の権限は不要。-Remove を付けると登録を消す。ツールのフォルダを移動したら、もう一度実行する。
#>
param([switch]$Remove)

$ErrorActionPreference = 'Stop'
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir -and $MyInvocation.MyCommand.Definition) { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $ScriptDir) { $ScriptDir = (Get-Location).Path }

$key = 'HKCU:\Software\Classes\invoicecheck-mail'

if ($Remove) {
    Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host '「Outlookで開く」の登録を消しました。' -ForegroundColor Green
    exit 0
}

$openScript = Join-Path $ScriptDir 'Open-Mail.ps1'
if (-not (Test-Path -LiteralPath $openScript)) { throw "Open-Mail.ps1 が見つかりません: $openScript" }

New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name '(default)' -Value 'URL:請求書メールチェック（メールを開く）'
New-ItemProperty -Path $key -Name 'URL Protocol' -Value '' -PropertyType String -Force | Out-Null
New-Item -Path "$key\shell\open\command" -Force | Out-Null
$command = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $openScript + '" "%1"'
Set-ItemProperty -Path "$key\shell\open\command" -Name '(default)' -Value $command

Write-Host '「Outlookで開く」を使えるようにしました。' -ForegroundColor Green
Write-Host '結果の画面で「Outlookで開く」を押すと、そのメールが Outlook で開きます。'
Write-Host '初めて押したときにブラウザで「アプリを開こうとしています」と出たら、「開く」を押してください。'
Write-Host '（ツールのフォルダを別の場所に移したときは、もう一度これを実行してください）'
