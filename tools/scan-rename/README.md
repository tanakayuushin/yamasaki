# スキャン書類 自動リネームツール（OCR）

スキャンした請求書・注文請書などのPDFや画像を文字認識(OCR)し、
**「年月_相手の会社名_書類種別.pdf」** の形にファイル名を付け直します。

```
scan_0001.pdf  →  202611_株式会社サンプル商事_請求書.pdf
scan_0002.pdf  →  202611_山口工業株式会社_注文請書.pdf
```

- 必要なもの: Windows 10/11 ＋ 日本語OCR（日本語版Windowsなら通常入っています）
- 追加インストール不要（Windows 標準の OCR と PDF 表示機能を使用）
- **既定はプレビューのみ。** ファイルは変更せず、新しいファイル名の一覧(CSV)を出します

## ファイル

| ファイル | 内容 |
|---|---|
| `Rename-ScannedDocs.ps1` | 本体 |
| `config.json` | 設定（入力フォルダ、自社名、取引先、命名ルールなど） |
| `Run-ScanRename.bat` | ダブルクリックで **プレビュー** |
| `Run-ScanRename-Apply.bat` | ダブルクリックで **実際にリネーム** |

## 使い方

1. `config.json` を設定する（最低限、次の2つ）
   - `Input.Folder` … スキャンしたファイルが入るフォルダ（`\` は `\\` と書く）
   - `OwnCompanyNames` … **自社名**（宛名の自社名を相手先と間違えないため。表記ゆれがあれば複数書く）
2. `Run-ScanRename.bat` を実行 → `リネーム結果_YYYYMMDD_HHMMSS.csv` を Excel で確認
3. よければ `Run-ScanRename-Apply.bat` を実行

フォルダを指定して実行したいとき:

```powershell
powershell -ExecutionPolicy Bypass -File .\Rename-ScannedDocs.ps1 -InputFolder "C:\Scan\20261127"
powershell -ExecutionPolicy Bypass -File .\Rename-ScannedDocs.ps1 -InputFolder "C:\Scan\20261127" -Apply
```

## 読み取りのしくみ

| 項目 | 方法 |
|---|---|
| 書類種別 | `DocTypes` のキーワード（請求書・注文請書など）のうち、**本文で一番先に出てくるもの**（＝タイトル）。`TargetDocTypes` 以外の書類は「対象外」として触りません |
| 相手の会社名 | ① `Vendors`（取引先マスタ）の名前・別名が本文にあればそれ（判定「OK」）<br>② なければ「株式会社〇〇」「〇〇(株)」などの表記を抜き出し、自社名と「御中・様」の宛名を除いて、〒・TEL・登録番号(T+13桁)が近くにあるもの＝発行元を採用（判定「推定」） |
| 年月 | `Date.Priority` の順に探す：<br>「11月分」の記載 → 「請求日」「発行日」などのラベルの日付 → 本文で最初に出てくる日付<br>支払期限・振込期日・納期などの行の日付は除外。和暦（令和8年、R8.11.30）にも対応 |

### 判定

| 判定 | 意味 | -Apply 時 |
|---|---|---|
| OK | 取引先マスタで会社名を確定 | リネーム |
| 推定 | 会社名を本文から推定 | リネーム（`Output.RenameEstimated: false` なら触らない） |
| 要確認 | 会社名・年月・書類種別のどれかが読み取れない | 触らない（`Output.ReviewFolder` を指定すればそこへ移動） |
| 対象外 | 対象外の書類種別 | 触らない |
| スキップ | すでにリネーム済み（`Input.SkipIfNameMatches`、既定は先頭が6桁の数字＋`_`） | 触らない |

同じ名前のファイルがすでにある場合は `_2`, `_3` … を付けます。

## 精度を上げるには

- **取引先マスタ（`Vendors`）を登録するのが一番効きます。** OCR の誤読（「ロ」と「口」など）や略称も `Aliases` に書いておけば拾えます
  ```json
  "Vendors": [
    { "Name": "株式会社サンプル商事", "Aliases": ["サンプル商事", "サンブル商事"] },
    { "Name": "山口工業株式会社",     "Aliases": ["山口工業"] }
  ]
  ```
- うまく読めないファイルは `Log.SaveOcrText: true` にすると、OCR結果がテキストで保存されるので、どう読まれているか確認できます
- スキャンは 300dpi 程度・白黒またはグレーがおすすめ。上下逆・横向きのスキャンは自動で回転して読み直します（`Ocr.TryRotation`）

## 設定項目

| 設定 | 説明 |
|---|---|
| `Ocr.MaxPages` | OCR するページ数（既定1＝1ページ目だけ。表紙だけで足りる書類が多いため） |
| `Date.MonthOffset` | 請求日・発行日から決めた年月をずらす（例: `-1` で前月扱い）。「〇月分」の記載には適用しません |
| `Naming.Template` | ファイル名の形。`{yyyy}` `{yy}` `{MM}` `{M}` `{company}` `{doctype}` `{original}`（元の名前）が使えます |
| `Naming.StripCorporateForm` | `true` で「株式会社」などを外して名前を付ける |
| `Output.Folder` | 移動先。空なら元のフォルダでリネームのみ |
| `Output.SubfolderPattern` | 例: `"{yyyy}\\{MM}"` で年・月ごとのフォルダに振り分け |

## 注意

- **Windows PowerShell 5.1（powershell.exe）専用**です。PowerShell 7（pwsh）では動きません。bat から起動すれば自動で 5.1 が使われます。
- 「日本語のOCR機能が見つかりません」と出たら：設定 → 時刻と言語 → 言語と地域 → 日本語 →「言語のオプション」→「光学式文字認識」を追加してください。
- `config.json` は **UTF-8** で保存してください。
