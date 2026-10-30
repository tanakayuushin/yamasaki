Attribute VB_Name = "ConsolidateCompanies"
Option Explicit

'==============================================================================
' 同じ企業を1行に集約するマクロ
'
'   工種ごとに企業を入力した一覧で、同じ企業が複数行に点在しているとき、
'   企業ごとに1行へまとめ、工種を「、」などの区切りで結合して出力します。
'
'   例）
'     企業名      工種            →   企業名      工種
'     山口建設    土木                 山口建設    土木、舗装、電気
'     周南電業    電気                 周南電業    電気
'     山口建設    舗装
'     山口建設    電気
'
'   使い方
'     1. Alt + F11 で VBE を開き、「ファイル」→「ファイルのインポート」で
'        この .bas を取り込む（または標準モジュールを追加して本文を貼り付け）
'     2. 下の「設定」を自分の表に合わせて書き換える
'     3. 集約したいシートを開いた状態で Alt + F8 →「企業を1行に集約」を実行
'
'   元のシートは変更しません。結果は別シートに書き出します。
'==============================================================================

'---------------------------- 設定（ここを書き換える） ------------------------

' 見出し行の行番号
Private Const HEADER_ROW As Long = 1

' 企業名の列。見出しの文字列（例 "企業名"）か、列記号（例 "B"）のどちらでも可
Private Const COMPANY_COL As String = "企業名"

' 工種の列。見出しの文字列（例 "工種"）か、列記号（例 "C"）のどちらでも可
Private Const WORKTYPE_COL As String = "工種"

' 工種をつなぐ区切り文字
Private Const DELIMITER As String = "、"

' 結果を書き出すシート名（同名シートがあれば中身を消して上書き）
Private Const OUTPUT_SHEET As String = "集約結果"

' 同じ工種が重複して出てきたとき 1 回だけにする（True）／そのまま並べる（False）
Private Const DEDUPE_WORKTYPES As Boolean = True

' 企業名の表記ゆれ吸収：全角・半角スペースを無視し、英数カナを半角に揃えて比較する
Private Const NORMALIZE_COMPANY_NAME As Boolean = True

' 結果の末尾に「元の行数」列を追加する
Private Const ADD_COUNT_COLUMN As Boolean = True

'------------------------------------------------------------------------------

Public Sub 企業を1行に集約()
    Dim wsSrc As Worksheet
    Set wsSrc = ActiveSheet

    If wsSrc.Name = OUTPUT_SHEET Then
        MsgBox "集約元のシートを開いた状態で実行してください。" & vbCrLf & _
               "（「" & OUTPUT_SHEET & "」シートは結果用です）", vbExclamation
        Exit Sub
    End If

    Dim colCompany As Long, colWork As Long
    colCompany = ResolveColumn(wsSrc, COMPANY_COL)
    colWork = ResolveColumn(wsSrc, WORKTYPE_COL)
    If colCompany = 0 Or colWork = 0 Then
        MsgBox "企業名または工種の列が見つかりません。" & vbCrLf & _
               "モジュール先頭の COMPANY_COL / WORKTYPE_COL を確認してください。", vbCritical
        Exit Sub
    End If

    Dim lastRow As Long, lastCol As Long
    lastRow = LastUsedRow(wsSrc)
    lastCol = LastUsedCol(wsSrc, HEADER_ROW)
    If lastRow <= HEADER_ROW Then
        MsgBox "データ行がありません。", vbExclamation
        Exit Sub
    End If

    Dim src As Variant
    src = wsSrc.Range(wsSrc.Cells(HEADER_ROW, 1), wsSrc.Cells(lastRow, lastCol)).Value

    ' key(正規化した企業名) → 出力配列上の行番号
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' 出力用の配列（最大でも元の行数と同じ）
    Dim out() As Variant
    ReDim out(1 To UBound(src, 1), 1 To lastCol + 1)
    Dim outCount As Long
    outCount = 0

    Dim r As Long, c As Long, key As String, idx As Long
    Dim company As String, work As String

    For r = 2 To UBound(src, 1)               ' src の 1 行目は見出し
        company = Trim$(CStr(src(r, colCompany)))
        If Len(company) > 0 Then
            work = Trim$(CStr(src(r, colWork)))
            key = NormalizeKey(company)

            If Not dict.Exists(key) Then
                ' 初出：その行を丸ごとコピー
                outCount = outCount + 1
                dict.Add key, outCount
                For c = 1 To lastCol
                    out(outCount, c) = src(r, c)
                Next c
                out(outCount, colCompany) = company
                out(outCount, colWork) = work
                out(outCount, lastCol + 1) = 1
            Else
                ' 既出：工種を結合、その他の列は空欄なら補完
                idx = dict(key)
                out(idx, colWork) = AppendWorkType(CStr(out(idx, colWork)), work)
                For c = 1 To lastCol
                    If c <> colCompany And c <> colWork Then
                        If Len(Trim$(CStr(out(idx, c)))) = 0 Then out(idx, c) = src(r, c)
                    End If
                Next c
                out(idx, lastCol + 1) = out(idx, lastCol + 1) + 1
            End If
        End If
    Next r

    If outCount = 0 Then
        MsgBox "企業名が入力された行がありません。", vbExclamation
        Exit Sub
    End If

    '---- 書き出し ----
    Dim wsOut As Worksheet
    Set wsOut = GetOrCreateSheet(wsSrc.Parent, OUTPUT_SHEET)
    wsOut.Cells.Clear

    Dim outCols As Long
    outCols = IIf(ADD_COUNT_COLUMN, lastCol + 1, lastCol)

    ' 見出し
    For c = 1 To lastCol
        wsOut.Cells(1, c).Value = src(1, c)
    Next c
    If ADD_COUNT_COLUMN Then wsOut.Cells(1, lastCol + 1).Value = "元の行数"

    ' 本体（配列を一括で貼り付け）
    Dim body() As Variant
    ReDim body(1 To outCount, 1 To outCols)
    For r = 1 To outCount
        For c = 1 To outCols
            body(r, c) = out(r, c)
        Next c
    Next r
    wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(outCount + 1, outCols)).Value = body

    ' 体裁
    With wsOut
        .Rows(1).Font.Bold = True
        .Columns(colWork).WrapText = False
        .Range(.Cells(1, 1), .Cells(outCount + 1, outCols)).Columns.AutoFit
        .Activate
        .Range("A1").Select
    End With

    MsgBox "集約が完了しました。" & vbCrLf & _
           "元データ " & (UBound(src, 1) - 1) & " 行 → " & outCount & " 社", vbInformation
End Sub

'------------------------------------------------------------------------------
' 補助関数
'------------------------------------------------------------------------------

' 既存の工種リストに新しい工種を追加（重複排除はオプション）
Private Function AppendWorkType(ByVal current As String, ByVal newWork As String) As String
    If Len(newWork) = 0 Then
        AppendWorkType = current
        Exit Function
    End If
    If Len(current) = 0 Then
        AppendWorkType = newWork
        Exit Function
    End If

    If DEDUPE_WORKTYPES Then
        Dim parts() As String, i As Long
        parts = Split(current, DELIMITER)
        For i = LBound(parts) To UBound(parts)
            If NormalizeKey(parts(i)) = NormalizeKey(newWork) Then
                AppendWorkType = current
                Exit Function
            End If
        Next i
    End If

    AppendWorkType = current & DELIMITER & newWork
End Function

' 比較用に企業名を正規化する
Private Function NormalizeKey(ByVal s As String) As String
    s = Trim$(s)
    If NORMALIZE_COMPANY_NAME Then
        s = Replace(s, " ", "")          ' 半角スペース
        s = Replace(s, "　", "")         ' 全角スペース
        s = Replace(s, vbTab, "")
        s = StrConv(s, vbNarrow)         ' 英数カナを半角に
        s = StrConv(s, vbUpperCase)      ' 英字の大文字小文字を無視
    End If
    NormalizeKey = s
End Function

' 見出し文字列または列記号から列番号を得る（見つからなければ 0）
Private Function ResolveColumn(ByVal ws As Worksheet, ByVal spec As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = LastUsedCol(ws, HEADER_ROW)

    ' まず見出し行から完全一致で探す
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(HEADER_ROW, c).Value)) = spec Then
            ResolveColumn = c
            Exit Function
        End If
    Next c

    ' 見つからなければ列記号（A, B, AA …）として解釈
    On Error Resume Next
    ResolveColumn = ws.Columns(spec).Column
    On Error GoTo 0
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim r As Range
    Set r = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If r Is Nothing Then LastUsedRow = 0 Else LastUsedRow = r.Row
End Function

Private Function LastUsedCol(ByVal ws As Worksheet, ByVal headerRow As Long) As Long
    LastUsedCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
End Function

Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = name
    End If
    Set GetOrCreateSheet = ws
End Function
