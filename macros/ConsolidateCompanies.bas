Attribute VB_Name = "ConsolidateCompanies"
Option Explicit

'==============================================================================
' 同じ業者を1行に集約するマクロ
'
'   工種ごとに業者を入力した一覧で、同じ業者が複数行に点在しているとき、
'   業者ごとに1行へまとめ、工種を「、」などの区切りで結合して出力します。
'
'   想定している表（見出しは 2 行目、A 列は余白）
'     B:工種  C:業者  D:担当者  E:メールアドレス  F:依頼  G:返答  H:見積書
'     I:お礼  J:見積入力  K:経過連絡  L:備考  M:NET金額（合計）
'
'   例）
'     工種    業者        →   工種              業者
'     土木    山口建設         土木、舗装、電気  山口建設
'     電気    周南電業         電気              周南電業
'     舗装    山口建設
'     電気    山口建設
'
'   使い方
'     1. Alt + F11 で VBE を開き、「ファイル」→「ファイルのインポート」で
'        この .bas を取り込む（または標準モジュールを追加して本文を貼り付け）
'     2. 必要なら下の「設定」を書き換える
'     3. 集約したいシートを開いた状態で Alt + F8 →「業者を1行に集約」を実行
'
'   元のシートは変更しません。元シートを複製した「集約結果」シートに、
'   罫線・塗り・列幅などの体裁をそのまま保って書き出します。
'==============================================================================

'---------------------------- 設定（必要なら書き換える） ----------------------

' 見出し行の行番号
Private Const HEADER_ROW As Long = 2

' 業者名の列。見出しの文字列（例 "業者"）か、列記号（例 "C"）のどちらでも可
Private Const COMPANY_COL As String = "業者"

' 工種の列。見出しの文字列（例 "工種"）か、列記号（例 "B"）のどちらでも可
Private Const WORKTYPE_COL As String = "工種"

' 工種をつなぐ区切り文字（セル内改行にしたいときは vbLf）
Private Const DELIMITER As String = "、"

' 数値として合計する列（見出し名をカンマ区切りで。不要なら ""）
Private Const SUM_COLS As String = "NET金額（合計）"

' 文字列を結合する列（見出し名をカンマ区切りで。不要なら ""）
'   ※ ここに無い列は「最初に出てきた行の値」を使い、空欄なら後の行で補う
Private Const JOIN_COLS As String = "備考"

' 結合する列（JOIN_COLS）の区切り文字
Private Const JOIN_DELIMITER As String = " ／ "

' 結果を書き出すシート名（同名シートがあれば中身を消して上書き）
Private Const OUTPUT_SHEET As String = "集約結果"

' 同じ工種が重複して出てきたとき 1 回だけにする（True）／そのまま並べる（False）
Private Const DEDUPE_WORKTYPES As Boolean = True

' 業者名の表記ゆれ吸収：全角・半角スペースを無視し、英数カナを半角に揃えて比較する
Private Const NORMALIZE_COMPANY_NAME As Boolean = True

' 結果の末尾に「元の行数」列を追加する（表の形を一切変えたくない場合は False）
Private Const ADD_COUNT_COLUMN As Boolean = False

' 結合した工種が列幅に収まらないとき、セル内で折り返す（列幅は変えない）
Private Const WRAP_WORKTYPE As Boolean = True

'------------------------------------------------------------------------------

Public Sub 業者を1行に集約()
    Dim wsSrc As Worksheet
    Set wsSrc = ActiveSheet

    If wsSrc.Name = OUTPUT_SHEET Then
        MsgBox "集約元のシートを開いた状態で実行してください。" & vbCrLf & _
               "（「" & OUTPUT_SHEET & "」シートは結果用です）", vbExclamation
        Exit Sub
    End If

    '---- 表の範囲を決める（見出し行の最初の文字～最後の文字がある列） ----
    Dim firstCol As Long, lastCol As Long, lastRow As Long
    firstCol = FirstUsedCol(wsSrc, HEADER_ROW)
    lastCol = LastUsedCol(wsSrc, HEADER_ROW)
    lastRow = LastUsedRow(wsSrc)
    If firstCol = 0 Or lastRow <= HEADER_ROW Then
        MsgBox "見出し行（" & HEADER_ROW & " 行目）またはデータ行が見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim nCols As Long
    nCols = lastCol - firstCol + 1

    Dim src As Variant
    src = wsSrc.Range(wsSrc.Cells(HEADER_ROW, firstCol), wsSrc.Cells(lastRow, lastCol)).Value
    ' src(1, *) が見出し、src(2〜, *) がデータ。列番号は 1 始まり（= シート列 - firstCol + 1）

    Dim colCompany As Long, colWork As Long
    colCompany = ResolveColumn(wsSrc, src, firstCol, COMPANY_COL)
    colWork = ResolveColumn(wsSrc, src, firstCol, WORKTYPE_COL)
    If colCompany = 0 Or colWork = 0 Then
        MsgBox "業者または工種の列が見つかりません。" & vbCrLf & _
               "モジュール先頭の COMPANY_COL / WORKTYPE_COL を確認してください。", vbCritical
        Exit Sub
    End If

    ' 列ごとの扱い: 0=初出優先(空欄なら補完) 1=数値合計 2=文字列結合
    Dim mode() As Long
    ReDim mode(1 To nCols)
    MarkColumns mode, src, SUM_COLS, 1
    MarkColumns mode, src, JOIN_COLS, 2

    ' key(正規化した業者名) → 出力配列上の行番号（Mac 版 Excel でも動くよう Collection を使用）
    Dim dict As Collection
    Set dict = New Collection

    Dim out() As Variant
    ReDim out(1 To UBound(src, 1), 1 To nCols + 1)
    Dim outCount As Long

    Dim r As Long, c As Long, key As String, idx As Long
    Dim company As String, work As String

    For r = 2 To UBound(src, 1)
        company = Trim$(CStr(src(r, colCompany)))
        If Len(company) > 0 Then
            work = Trim$(CStr(src(r, colWork)))
            key = NormalizeKey(company)

            If Not CollHas(dict, key) Then
                '---- 初出：その行を丸ごとコピー ----
                outCount = outCount + 1
                dict.Add outCount, key
                For c = 1 To nCols
                    out(outCount, c) = src(r, c)
                Next c
                out(outCount, colCompany) = company
                out(outCount, colWork) = work
                out(outCount, nCols + 1) = 1
            Else
                '---- 既出：工種を結合、他の列は設定どおりに統合 ----
                idx = dict(key)
                out(idx, colWork) = AppendWorkType(CStr(out(idx, colWork)), work)
                For c = 1 To nCols
                    If c <> colCompany And c <> colWork Then
                        Select Case mode(c)
                            Case 1   ' 数値合計
                                If IsNumeric(src(r, c)) And Len(Trim$(CStr(src(r, c)))) > 0 Then
                                    If IsNumeric(out(idx, c)) And Len(Trim$(CStr(out(idx, c)))) > 0 Then
                                        out(idx, c) = CDbl(out(idx, c)) + CDbl(src(r, c))
                                    Else
                                        out(idx, c) = CDbl(src(r, c))
                                    End If
                                End If
                            Case 2   ' 文字列結合（同じ文言は 1 回だけ）
                                out(idx, c) = JoinText(CStr(out(idx, c)), CStr(src(r, c)))
                            Case Else ' 初出優先、空欄なら補完
                                If Len(Trim$(CStr(out(idx, c)))) = 0 Then out(idx, c) = src(r, c)
                        End Select
                    End If
                Next c
                out(idx, nCols + 1) = out(idx, nCols + 1) + 1
            End If
        End If
    Next r

    If outCount = 0 Then
        MsgBox "業者名が入力された行がありません。", vbExclamation
        Exit Sub
    End If

    '---- 書き出し：元シートをまるごと複製し、データ行だけ差し替える（罫線・塗り・列幅・印刷設定はそのまま） ----
    Dim wb As Workbook
    Set wb = wsSrc.Parent

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(OUTPUT_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    wsSrc.Copy After:=wb.Worksheets(wb.Worksheets.Count)
    Dim wsOut As Worksheet
    Set wsOut = wb.Worksheets(wb.Worksheets.Count)
    wsOut.Name = OUTPUT_SHEET

    ' 複製されたボタンなどの図形は不要なので削除
    Dim shp As Shape
    For Each shp In wsOut.Shapes
        shp.Delete
    Next shp

    ' データ行の値だけ消す（書式は残す）
    wsOut.Range(wsOut.Cells(HEADER_ROW + 1, firstCol), wsOut.Cells(lastRow, lastCol)).ClearContents

    ' 集約結果を書き込む
    Dim body() As Variant
    ReDim body(1 To outCount, 1 To nCols)
    For r = 1 To outCount
        For c = 1 To nCols
            body(r, c) = out(r, c)
        Next c
    Next r
    wsOut.Range(wsOut.Cells(HEADER_ROW + 1, firstCol), wsOut.Cells(HEADER_ROW + outCount, lastCol)).Value = body

    ' 「元の行数」列（任意）：見出しと各行の書式は隣の列をコピー
    If ADD_COUNT_COLUMN Then
        Dim cntCol As Long
        cntCol = lastCol + 1
        wsOut.Range(wsOut.Cells(HEADER_ROW, lastCol), wsOut.Cells(HEADER_ROW + outCount, lastCol)).Copy
        wsOut.Cells(HEADER_ROW, cntCol).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        wsOut.Cells(HEADER_ROW, cntCol).Value = "元の行数"
        For r = 1 To outCount
            wsOut.Cells(HEADER_ROW + r, cntCol).Value = out(r, nCols + 1)
        Next r
        wsOut.Columns(cntCol).ColumnWidth = 9
    End If

    ' 工種が長くなった行はセル内で折り返す（列幅は変えない）
    If WRAP_WORKTYPE Then
        wsOut.Range(wsOut.Cells(HEADER_ROW + 1, firstCol + colWork - 1), wsOut.Cells(HEADER_ROW + outCount, firstCol + colWork - 1)).WrapText = True
    End If

    wsOut.Activate
    wsOut.Cells(HEADER_ROW, firstCol).Select
    Application.ScreenUpdating = True

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

' 備考などの文字列列を結合（同じ文言は追加しない）
Private Function JoinText(ByVal current As String, ByVal newText As String) As String
    current = Trim$(current): newText = Trim$(newText)
    If Len(newText) = 0 Then
        JoinText = current
    ElseIf Len(current) = 0 Then
        JoinText = newText
    ElseIf InStr(1, JOIN_DELIMITER & current & JOIN_DELIMITER, JOIN_DELIMITER & newText & JOIN_DELIMITER) > 0 Then
        JoinText = current
    Else
        JoinText = current & JOIN_DELIMITER & newText
    End If
End Function

' 比較用に業者名を正規化する
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

' 見出し文字列または列記号から、src 配列上の列番号を得る（見つからなければ 0）
Private Function ResolveColumn(ByVal ws As Worksheet, ByRef src As Variant, _
                               ByVal firstCol As Long, ByVal spec As String) As Long
    Dim c As Long
    For c = 1 To UBound(src, 2)
        If Trim$(CStr(src(1, c))) = spec Then
            ResolveColumn = c
            Exit Function
        End If
    Next c

    Dim sheetCol As Long
    On Error Resume Next
    sheetCol = ws.Columns(spec).Column
    On Error GoTo 0
    If sheetCol >= firstCol And sheetCol - firstCol + 1 <= UBound(src, 2) Then
        ResolveColumn = sheetCol - firstCol + 1
    End If
End Function

' カンマ区切りの見出し名リストに該当する列へ mode を設定する
Private Sub MarkColumns(ByRef mode() As Long, ByRef src As Variant, ByVal names As String, ByVal value As Long)
    If Len(Trim$(names)) = 0 Then Exit Sub
    Dim parts() As String, i As Long, c As Long
    parts = Split(Replace(names, "，", ","), ",")
    For i = LBound(parts) To UBound(parts)
        For c = 1 To UBound(src, 2)
            If Trim$(CStr(src(1, c))) = Trim$(parts(i)) Then mode(c) = value
        Next c
    Next i
End Sub

' Collection に指定キーがあるか
Private Function CollHas(ByVal c As Collection, ByVal k As String) As Boolean
    On Error Resume Next
    c.Item k
    CollHas = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim r As Range
    Set r = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If r Is Nothing Then LastUsedRow = 0 Else LastUsedRow = r.Row
End Function

Private Function FirstUsedCol(ByVal ws As Worksheet, ByVal headerRow As Long) As Long
    Dim c As Long
    For c = 1 To ws.Columns.Count
        If Len(Trim$(CStr(ws.Cells(headerRow, c).Value))) > 0 Then
            FirstUsedCol = c
            Exit Function
        End If
        If c > 50 Then Exit For   ' 見出しが 50 列以上右にあることは想定しない
    Next c
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
