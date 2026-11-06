Attribute VB_Name = "ConsolidateCompanies"
Option Explicit

'==============================================================================
' 同じ業者を1行に集約するマクロ
'
'   工種ごとに業者を入力した一覧で、同じ業者が複数行に点在しているとき、
'   業者ごとに1行へまとめ、工種を「、」などの区切りで結合して出力します。
'
'   想定している表（見出しは 2 行目、A 列は余白）
'     B:工種  C:会社名（業者）  D:担当者  E:メールアドレス  F:依頼  G:返答  H:見積書
'     I:お礼  J:見積入力  K:経過連絡  L:備考  M:NET金額（合計）
'
'   例）
'     工種    業者        →   工種              業者
'     土木    山口建設         土木、舗装、電気  山口建設
'     電気    周南電業         電気              周南電業
'     舗装    山口建設
'     電気    山口建設
'
'   シート構成
'     「依頼リスト」シート … 依頼する業者・工種などを 1 行ずつ入力する（入力用）
'     「Sheet1」シート     … 業者ごとに 1 行へまとめた結果（出力用。無ければ自動で作る）
'
'   マクロは 2 つ
'     業者を1行に集約       … 業者ごとに 1 行（担当者やメールアドレスが違っても同じ業者なら 1 行。最初の行の値を採用）
'     業者と担当者ごとに集約 … 業者と担当者の組み合わせごとに 1 行（担当者が違えば別の行）
'
'   使い方
'     1. Alt + F11 で VBE を開き、「ファイル」→「ファイルのインポート」で
'        この .bas を取り込む（または標準モジュールを追加して本文を貼り付け）
'     2. 必要なら下の「設定」を書き換える
'     3. 「依頼リスト」に入力し、Alt + F8 → どちらかのマクロを実行
'        （シート上のボタンでも可。どのシートを開いていても、何回でも実行できる）
'
'   「依頼リスト」は変更しません。「Sheet1」の見出し行より下を消して結果を書き直します
'   （Sheet1 の罫線・塗り・列幅などの書式はそのまま）。Sheet1 が無ければ「依頼リスト」を
'   複製して作ります。Sheet1 の列は見出しの文字で対応づけるので、列の順番が違っても構いません。
'==============================================================================

'---------------------------- 設定（必要なら書き換える） ----------------------

' 入力用シート名（依頼する業者・工種を 1 行ずつ書くシート）
Private Const INPUT_SHEET As String = "依頼リスト"

' 出力用シート名（業者ごとに 1 行にまとめた結果を書くシート。無ければ依頼リストを複製して作る）
Private Const OUTPUT_SHEET As String = "Sheet1"

' 見出し行の行番号（依頼リスト・Sheet1 とも同じ行に見出しがある前提）
Private Const HEADER_ROW As Long = 2

' 業者名の列。見出しの文字列か列記号（例 "C"）。候補をカンマ区切りで複数書ける（先に見つかったものを使う）
Private Const COMPANY_COL As String = "会社名,業者"

' 工種の列。見出しの文字列か列記号（例 "B"）。候補をカンマ区切りで複数書ける
Private Const WORKTYPE_COL As String = "工種,工事種"

' 工種をつなぐ区切り文字（セル内改行にしたいときは vbLf）
Private Const DELIMITER As String = "、"

' 項目ごとの内訳として並べる列（見出し名をカンマ区切りで。不要なら ""）
'   業者の金額が 2 件以上になるとき（複数の行に分かれている／セルの中に項目が複数書かれている）、
'   セル内に 1 行ずつ「項目名：金額」（項目名がなければ金額だけ）を並べ、最後に「合計：金額」を付ける。
'   業者ごとの金額なので工種名は付けない
Private Const BREAKDOWN_COLS As String = "NET金額（合計）"

' 内訳の金額の表示形式
Private Const BREAKDOWN_NUMFMT As String = "#,##0"

' 金額が未入力の行も内訳に「未入力」として載せる
Private Const BREAKDOWN_SHOW_EMPTY As Boolean = False

' 単純に数値を合計する列（見出し名をカンマ区切りで。不要なら ""）
Private Const SUM_COLS As String = ""

' 文字列を結合する列（見出し名をカンマ区切りで。不要なら ""）
'   ※ ここに無い列は「最初に出てきた行の値」を使い、空欄なら後の行で補う
Private Const JOIN_COLS As String = "備考"

' 結合する列（JOIN_COLS）の区切り文字
Private Const JOIN_DELIMITER As String = " ／ "

' 同じ工種が重複して出てきたとき 1 回だけにする（True）／そのまま並べる（False）
Private Const DEDUPE_WORKTYPES As Boolean = True

' 業者名の表記ゆれ吸収：全角・半角スペースを無視し、英数カナを半角に揃えて比較する
Private Const NORMALIZE_COMPANY_NAME As Boolean = True

' 結果の末尾に「元の行数」列を追加する（表の形を一切変えたくない場合は False）
Private Const ADD_COUNT_COLUMN As Boolean = False

' 結合した工種が列幅に収まらないとき、セル内で折り返す（列幅は変えない）
Private Const WRAP_WORKTYPE As Boolean = True

' 「業者と担当者ごとに集約」で、業者に加えて行を分ける基準にする列（見出し名をカンマ区切りで）
'   例 "担当者" … 同じ業者でも担当者が違えば別の行にする（担当者が空欄の行は「担当者なし」のグループ）
Private Const SPLIT_KEY_COLS As String = "担当者"

'------------------------------------------------------------------------------

' 業者ごとに 1 行にまとめる（担当者などが違っても同じ業者なら 1 行）
Public Sub 業者を1行に集約()
    Consolidate ""
End Sub

' 業者と担当者（SPLIT_KEY_COLS）の組み合わせごとに 1 行にまとめる
Public Sub 業者と担当者ごとに集約()
    Consolidate SPLIT_KEY_COLS
End Sub

Private Sub Consolidate(ByVal extraKeyCols As String)
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    ' 入力は常に「依頼リスト」シート
    Dim wsSrc As Worksheet
    If Not SheetExists(wb, INPUT_SHEET) Then
        MsgBox "「" & INPUT_SHEET & "」シートが見つかりません。" & vbCrLf & _
               "依頼する業者・工種を入力するシートの名前を「" & INPUT_SHEET & "」にしてください。", vbExclamation
        Exit Sub
    End If
    Set wsSrc = wb.Worksheets(INPUT_SHEET)

    '---- 表の範囲を決める（見出し行の最初の文字～最後の文字がある列） ----
    Dim firstCol As Long, lastCol As Long, lastRow As Long
    firstCol = FirstUsedCol(wsSrc, HEADER_ROW)
    lastCol = LastUsedCol(wsSrc, HEADER_ROW)
    If firstCol = 0 Then
        MsgBox "見出し行（" & HEADER_ROW & " 行目）が見つかりません。", vbExclamation
        Exit Sub
    End If

    Dim nCols As Long
    nCols = lastCol - firstCol + 1

    ' 見出し行から業者・工種の列を決める
    Dim hdr As Variant
    hdr = To2D(wsSrc.Range(wsSrc.Cells(HEADER_ROW, firstCol), wsSrc.Cells(HEADER_ROW, lastCol)).Value)
    Dim colCompany As Long, colWork As Long
    colCompany = ResolveColumn(wsSrc, hdr, firstCol, COMPANY_COL)
    colWork = ResolveColumn(wsSrc, hdr, firstCol, WORKTYPE_COL)
    If colCompany = 0 Or colWork = 0 Then
        MsgBox "「" & INPUT_SHEET & "」に業者または工種の列が見つかりません。" & vbCrLf & _
               "モジュール先頭の COMPANY_COL / WORKTYPE_COL を確認してください。", vbCritical
        Exit Sub
    End If

    ' 行を分ける基準に加える列（担当者など）。見つからない列は無視
    Dim keyCols() As Long, nKeyCols As Long, kc As Long, kcol As Long, kparts() As String
    ReDim keyCols(1 To nCols)
    If Len(Trim$(extraKeyCols)) > 0 Then
        kparts = Split(Replace(extraKeyCols, "，", ","), ",")
        For kc = LBound(kparts) To UBound(kparts)
            kcol = ResolveColumn(wsSrc, hdr, firstCol, Trim$(kparts(kc)))
            If kcol > 0 And kcol <> colCompany Then
                nKeyCols = nKeyCols + 1
                keyCols(nKeyCols) = kcol
            End If
        Next kc
    End If

    ' データの最終行は業者列・工種列のどちらか遅い方（表の下にあるメモなどは対象外）
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, firstCol + colCompany - 1).End(xlUp).Row
    Dim lastRow2 As Long
    lastRow2 = wsSrc.Cells(wsSrc.Rows.Count, firstCol + colWork - 1).End(xlUp).Row
    If lastRow2 > lastRow Then lastRow = lastRow2
    If lastRow <= HEADER_ROW Then
        MsgBox "「" & INPUT_SHEET & "」にデータ行がありません。", vbExclamation
        Exit Sub
    End If

    Dim src As Variant
    src = To2D(wsSrc.Range(wsSrc.Cells(HEADER_ROW, firstCol), wsSrc.Cells(lastRow, lastCol)).Value)
    ' src(1, *) が見出し、src(2〜, *) がデータ。列番号は 1 始まり（= シート列 - firstCol + 1）

    ' エラー値（#N/A など）のセルは空欄として扱う
    Dim rr As Long, cc As Long
    For rr = 1 To UBound(src, 1)
        For cc = 1 To UBound(src, 2)
            If IsError(src(rr, cc)) Then src(rr, cc) = Empty
        Next cc
    Next rr

    ' 列ごとの扱い: 0=初出優先(空欄なら補完) 1=数値合計 2=文字列結合 3=工種ごとの内訳
    Dim mode() As Long
    ReDim mode(1 To nCols)
    MarkColumns mode, src, SUM_COLS, 1
    MarkColumns mode, src, JOIN_COLS, 2
    MarkColumns mode, src, BREAKDOWN_COLS, 3

    ' 内訳列用：行ごと・列ごとに「工種(Chr2)セルの内容」を Chr(1) 区切りで溜めておく（セル内改行と衝突しないよう制御文字を使う）
    Dim items() As String
    ReDim items(1 To UBound(src, 1), 1 To nCols)

    ' key(正規化した業者名) → 出力配列上の行番号（Mac 版 Excel でも動くよう Collection を使用）
    Dim dict As Collection
    Set dict = New Collection

    Dim out() As Variant
    ReDim out(1 To UBound(src, 1), 1 To nCols + 1)
    Dim outCount As Long
    Dim companyOf() As String          ' 出力行ごとの業者キー（同じ業者の行を連続させるため）
    ReDim companyOf(1 To UBound(src, 1))
    Dim firstName As Collection        ' 業者キー → 最初に出てきた表記（行を分けても業者名の表記をそろえる）
    Set firstName = New Collection

    Dim r As Long, c As Long, key As String, idx As Long
    Dim company As String, work As String, nDataRows As Long

    For r = 2 To UBound(src, 1)
        company = Trim$(CStr(src(r, colCompany)))
        If Len(NormalizeKey(company)) > 0 Then
            nDataRows = nDataRows + 1
            work = Trim$(CStr(src(r, colWork)))
            key = NormalizeKey(company)
            For kc = 1 To nKeyCols
                key = key & Chr(1) & NormalizeKey(CStr(src(r, keyCols(kc))))
            Next kc

            If Not CollHas(dict, key) Then
                '---- 初出：その行を丸ごとコピー ----
                outCount = outCount + 1
                dict.Add outCount, key
                companyOf(outCount) = NormalizeKey(company)
                If CollHas(firstName, companyOf(outCount)) Then
                    company = firstName(companyOf(outCount))
                Else
                    firstName.Add company, companyOf(outCount)
                End If
                For c = 1 To nCols
                    out(outCount, c) = src(r, c)
                Next c
                out(outCount, colCompany) = company
                out(outCount, colWork) = work
                out(outCount, nCols + 1) = 1
                For c = 1 To nCols
                    If mode(c) = 3 Then items(outCount, c) = work & Chr(2) & CStr(src(r, c))
                Next c
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
                            Case 3   ' 工種ごとの内訳
                                items(idx, c) = items(idx, c) & Chr(1) & work & Chr(2) & CStr(src(r, c))
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
        MsgBox "「" & INPUT_SHEET & "」に業者名が入力された行がありません。", vbExclamation
        Exit Sub
    End If

    '---- 出力順：同じ業者の行が連続するように並べ替える（業者は初出順、同じ業者の中も初出順） ----
    Dim ord() As Long, placed() As Boolean, k As Long, m As Long
    ReDim ord(1 To outCount)
    ReDim placed(1 To outCount)
    k = 0
    For r = 1 To outCount
        If Not placed(r) Then
            For m = r To outCount
                If Not placed(m) Then
                    If companyOf(m) = companyOf(r) Then
                        k = k + 1
                        ord(k) = m
                        placed(m) = True
                    End If
                End If
            Next m
        End If
    Next r

    '---- 内訳列：2 件以上ある業者は「工種：金額」を 1 行ずつ並べ、最後に合計を付ける ----
    For c = 1 To nCols
        If mode(c) = 3 Then
            For r = 1 To outCount
                out(r, c) = BuildBreakdown(items(r, c), out(r, c))
            Next r
        End If
    Next c

    '---- 書き出し：「Sheet1」の見出し行より下を消して結果を書く（Sheet1 が無ければ依頼リストを複製して作る） ----
    ' 途中でエラーが起きても画面更新や警告表示の設定を元に戻す
    On Error GoTo Failed
    Application.ScreenUpdating = False

    Dim wsOut As Worksheet
    If SheetExists(wb, OUTPUT_SHEET) Then
        Set wsOut = wb.Worksheets(OUTPUT_SHEET)
    Else
        Dim nBefore As Long
        nBefore = wb.Worksheets.Count
        wsSrc.Copy After:=wsSrc
        Set wsOut = wb.Worksheets(wsSrc.Index + 1)
        If wb.Worksheets.Count = nBefore Or wsOut Is wsSrc Then
            Application.ScreenUpdating = True
            MsgBox "「" & OUTPUT_SHEET & "」シートの作成に失敗したため中止しました。", vbCritical
            Exit Sub
        End If
        wsOut.Name = OUTPUT_SHEET
    End If

    ' 出力シートの見出し行。見出しが無ければ依頼リストの見出しを同じ位置にコピーする
    Dim oFirst As Long, oLast As Long
    oFirst = FirstUsedCol(wsOut, HEADER_ROW)
    oLast = LastUsedCol(wsOut, HEADER_ROW)
    If oFirst = 0 Then
        wsSrc.Range(wsSrc.Cells(HEADER_ROW, firstCol), wsSrc.Cells(HEADER_ROW, lastCol)).Copy wsOut.Cells(HEADER_ROW, firstCol)
        oFirst = firstCol
        oLast = lastCol
    End If
    Dim ohdr As Variant
    ohdr = To2D(wsOut.Range(wsOut.Cells(HEADER_ROW, oFirst), wsOut.Cells(HEADER_ROW, oLast)).Value)

    ' 依頼リストの列 c → 出力シートの列番号（見出しの文字で対応づけ。0 = 出力シートに無い列）
    Dim colMap() As Long
    ReDim colMap(1 To nCols)
    Dim oc As Long
    For c = 1 To nCols
        If Len(HeaderText(hdr(1, c))) > 0 Then
            For oc = 1 To UBound(ohdr, 2)
                If HeaderText(ohdr(1, oc)) = HeaderText(hdr(1, c)) Then
                    colMap(c) = oFirst + oc - 1
                    Exit For
                End If
            Next oc
        End If
    Next c
    If colMap(colCompany) = 0 Or colMap(colWork) = 0 Then
        Application.ScreenUpdating = True
        MsgBox "「" & OUTPUT_SHEET & "」の " & HEADER_ROW & " 行目に「" & COMPANY_COL & "」「" & WORKTYPE_COL & "」の見出しが見つかりません。" & vbCrLf & _
               "「" & INPUT_SHEET & "」と同じ見出しを付けてください。", vbCritical
        Exit Sub
    End If

    ' 出力シートの既存データ行の範囲（書式のある範囲も含める）
    Dim oLastRow As Long, fmtLastRow As Long
    fmtLastRow = wsOut.UsedRange.Row + wsOut.UsedRange.Rows.Count - 1
    oLastRow = fmtLastRow
    If oLastRow < HEADER_ROW + outCount Then oLastRow = HEADER_ROW + outCount

    ' フィルタや行の非表示が残っていても結果がすべて見えるようにする
    On Error Resume Next
    If wsOut.FilterMode Then wsOut.ShowAllData
    On Error GoTo Failed
    wsOut.Range(wsOut.Rows(HEADER_ROW + 1), wsOut.Rows(oLastRow)).Hidden = False

    ' 前回の結果（データ行の値）だけ消す。書式は残す
    wsOut.Range(wsOut.Cells(HEADER_ROW + 1, oFirst), wsOut.Cells(oLastRow, oLast)).ClearContents

    ' 行数が増えて書式の無い行に書くときは、最初のデータ行の書式を広げる
    If fmtLastRow >= HEADER_ROW + 1 And HEADER_ROW + outCount > fmtLastRow Then
        wsOut.Range(wsOut.Cells(HEADER_ROW + 1, oFirst), wsOut.Cells(HEADER_ROW + 1, oLast)).Copy
        wsOut.Range(wsOut.Cells(fmtLastRow + 1, oFirst), wsOut.Cells(HEADER_ROW + outCount, oLast)).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If

    ' データ行の高さを標準的な高さ（出力シートのデータ行で最も多い高さ）に揃える
    Dim baseHeight As Double
    baseHeight = TypicalRowHeight(wsOut, HEADER_ROW + 1, oLastRow)
    If baseHeight > 0 Then wsOut.Range(wsOut.Rows(HEADER_ROW + 1), wsOut.Rows(oLastRow)).RowHeight = baseHeight

    ' 集約結果を列ごとに書き込む（出力シートに無い列は書かない）
    Dim colVals() As Variant
    ReDim colVals(1 To outCount, 1 To 1)
    For c = 1 To nCols
        If colMap(c) > 0 Then
            For r = 1 To outCount
                colVals(r, 1) = out(ord(r), c)
            Next r
            wsOut.Range(wsOut.Cells(HEADER_ROW + 1, colMap(c)), wsOut.Cells(HEADER_ROW + outCount, colMap(c))).Value = colVals
        End If
    Next c

    ' 「元の行数」列（任意）：見出しと各行の書式は隣の列をコピー
    If ADD_COUNT_COLUMN Then
        Dim cntCol As Long
        cntCol = oLast + 1
        wsOut.Range(wsOut.Cells(HEADER_ROW, oLast), wsOut.Cells(HEADER_ROW + outCount, oLast)).Copy
        wsOut.Cells(HEADER_ROW, cntCol).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        wsOut.Cells(HEADER_ROW, cntCol).Value = "元の行数"
        For r = 1 To outCount
            wsOut.Cells(HEADER_ROW + r, cntCol).Value = out(ord(r), nCols + 1)
        Next r
        wsOut.Columns(cntCol).ColumnWidth = 9
    End If

    ' 工種が長くなった行はセル内で折り返す（列幅は変えない）
    Dim needFit() As Boolean
    ReDim needFit(1 To outCount)
    Dim workColWidth As Double
    workColWidth = wsOut.Columns(colMap(colWork)).ColumnWidth
    If WRAP_WORKTYPE Then
        wsOut.Range(wsOut.Cells(HEADER_ROW + 1, colMap(colWork)), wsOut.Cells(HEADER_ROW + outCount, colMap(colWork))).WrapText = True
        For r = 1 To outCount
            If DisplayWidth(CStr(out(ord(r), colWork))) > workColWidth Then needFit(r) = True
        Next r
    End If

    ' 内訳列は折り返して、複数行になった行は高さを合わせる（列幅は変えない）
    For c = 1 To nCols
        If mode(c) = 3 And colMap(c) > 0 Then
            wsOut.Range(wsOut.Cells(HEADER_ROW + 1, colMap(c)), wsOut.Cells(HEADER_ROW + outCount, colMap(c))).WrapText = True
            For r = 1 To outCount
                If InStr(CStr(out(ord(r), c)), vbLf) > 0 Then needFit(r) = True
            Next r
        End If
    Next c
    For c = 1 To nCols
        If mode(c) = 2 And colMap(c) > 0 Then
            If wsOut.Cells(HEADER_ROW + 1, colMap(c)).WrapText Then
                For r = 1 To outCount
                    If InStr(CStr(out(ord(r), c)), JOIN_DELIMITER) > 0 Then needFit(r) = True
                Next r
            End If
        End If
    Next c
    For r = 1 To outCount
        If needFit(r) Then wsOut.Rows(HEADER_ROW + r).AutoFit
    Next r

    wsOut.Activate
    wsOut.Cells(HEADER_ROW, oFirst).Select
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "「" & OUTPUT_SHEET & "」に集約しました。" & vbCrLf & _
           "「" & INPUT_SHEET & "」 " & nDataRows & " 行 → " & outCount & IIf(nKeyCols > 0, " 行（業者と" & extraKeyCols & "の組み合わせ）", " 社"), vbInformation
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & _
           "（" & Err.Number & "：" & Err.Description & "）" & vbCrLf & vbCrLf & _
           "「" & OUTPUT_SHEET & "」シートが保護されていないか、結合セルが無いかを確認してください。", vbCritical
End Sub

' Range.Value が 1 セルのときは配列にならないので 1×1 の配列にそろえる
Private Function To2D(ByVal v As Variant) As Variant
    If IsArray(v) Then
        To2D = v
    Else
        Dim a(1 To 1, 1 To 1) As Variant
        a(1, 1) = v
        To2D = a
    End If
End Function

' 指定範囲の行で最も多く使われている行の高さ（データ行の「標準」の高さとみなす）
Private Function TypicalRowHeight(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long) As Double
    Dim heights As Collection, counts As Collection
    Set heights = New Collection: Set counts = New Collection
    Dim r As Long, h As Double, k As String, best As Double, bestN As Long, n As Long
    If r2 < r1 Then Exit Function
    If r2 - r1 > 2000 Then r2 = r1 + 2000         ' 行数が多いときは先頭 2000 行で判断
    For r = r1 To r2
        If ws.Rows(r).Hidden Then GoTo NextRow
        h = ws.Rows(r).RowHeight
        k = CStr(Round(h, 2))
        n = 0
        On Error Resume Next
        n = counts(k)
        On Error GoTo 0
        If n = 0 Then
            heights.Add h, k
            counts.Add 1, k
            n = 1
        Else
            counts.Remove k
            counts.Add n + 1, k
            n = n + 1
        End If
        If n > bestN Then bestN = n: best = h
NextRow:
    Next r
    TypicalRowHeight = best
End Function

' 文字列の表示幅のおおよそ（全角 2・半角 1 として数える。列幅との比較用）
Private Function DisplayWidth(ByVal s As String) As Double
    Dim i As Long, w As Double, code As Long
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If code < 0 Then code = code + 65536
        If code > 255 And Not (code >= &HFF61& And code <= &HFF9F&) Then
            w = w + 2          ' 全角（半角カナを除く）
        Else
            w = w + 1
        End If
    Next i
    DisplayWidth = w * 1.05
End Function

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

' 「工種(Chr2)セルの内容」(Chr1 区切り) から内訳テキストを作る（工種名は使わない）。
'   セルの内容に項目が複数書かれている場合（改行・「、」・「／」・「;」区切り）は項目ごとに分ける。
'   項目が全部で 1 つだけなら元の値をそのまま返す（数値・表示形式を維持）。
Private Function BuildBreakdown(ByVal packed As String, ByVal singleVal As Variant) As Variant
    Dim lines() As String, i As Long, j As Long, p As Long
    Dim raw As String, parts As Variant
    Dim label As String, amount As Double, ok As Boolean
    Dim total As Double, hasAny As Boolean, nItems As Long, nWorks As Long
    Dim result As String

    lines = Split(packed, Chr(1))
    nWorks = UBound(lines) + 1

    ' 項目の総数を数える（1 つだけならそのまま）
    For i = LBound(lines) To UBound(lines)
        p = InStr(lines(i), Chr(2))
        raw = Trim$(Mid$(lines(i), p + 1))
        If Len(raw) > 0 Then
            parts = SplitItems(raw)
            nItems = nItems + UBound(parts) + 1
        ElseIf BREAKDOWN_SHOW_EMPTY Then
            nItems = nItems + 1
        End If
    Next i
    If nItems <= 1 Then
        ' 金額が 1 件だけ → その値をそのまま返す（数値なら数値のまま。表示形式も元どおり）
        BuildBreakdown = singleVal
        For i = LBound(lines) To UBound(lines)
            p = InStr(lines(i), Chr(2))
            raw = Trim$(Mid$(lines(i), p + 1))
            If Len(raw) > 0 Then
                If IsNumeric(raw) Then BuildBreakdown = CDbl(raw) Else BuildBreakdown = raw
                Exit For
            End If
        Next i
        Exit Function
    End If

    ' 業者ごとの金額なので工種名は付けず、項目名（あれば）と金額だけを 1 行ずつ並べる
    For i = LBound(lines) To UBound(lines)
        p = InStr(lines(i), Chr(2))
        raw = Trim$(Mid$(lines(i), p + 1))

        If Len(raw) = 0 Then
            If BREAKDOWN_SHOW_EMPTY Then result = result & "未入力" & vbLf
        Else
            parts = SplitItems(raw)
            For j = LBound(parts) To UBound(parts)
                ParseAmount parts(j), label, amount, ok
                If ok Then
                    total = total + amount
                    hasAny = True
                    If Len(label) > 0 Then
                        result = result & label & "：" & Format(amount, BREAKDOWN_NUMFMT) & vbLf
                    Else
                        result = result & Format(amount, BREAKDOWN_NUMFMT) & vbLf
                    End If
                Else
                    result = result & Trim$(parts(j)) & vbLf      ' 金額として読めない文言はそのまま
                End If
            Next j
        End If
    Next i

    If Len(result) = 0 Then
        BuildBreakdown = Empty
    ElseIf hasAny Then
        BuildBreakdown = result & "合計：" & Format(total, BREAKDOWN_NUMFMT)
    Else
        BuildBreakdown = Left$(result, Len(result) - 1)
    End If
End Function

' セル内の複数項目を分割する。区切り：改行・「、」・「，」・「／」「/」・「;」「；」・
' 半角カンマ（ただし 1,250,000 のような桁区切りは分割しない）
Private Function SplitItems(ByVal raw As String) As Variant
    Dim buf As String, ch As String, i As Long, out As String
    buf = NarrowAscii(raw)                   ' 全角の数字・カンマ・記号を半角にそろえる（カナはそのまま）
    buf = Replace(buf, vbCrLf, vbLf)
    buf = Replace(buf, vbCr, vbLf)
    For i = 1 To Len(buf)
        ch = Mid$(buf, i, 1)
        Select Case ch
            Case vbLf, "、", ";"
                out = out & Chr(3)
            Case ","
                ' 1,250,000 のような桁区切りは分割しない
                If i > 1 And i + 3 <= Len(buf) Then
                    If IsDigitChar(Mid$(buf, i - 1, 1)) And IsDigitChar(Mid$(buf, i + 1, 1)) _
                       And IsDigitChar(Mid$(buf, i + 2, 1)) And IsDigitChar(Mid$(buf, i + 3, 1)) Then
                        out = out & ch
                    Else
                        out = out & Chr(3)
                    End If
                Else
                    out = out & Chr(3)
                End If
            Case "/"
                ' 10/25 のような日付の中のスラッシュは分割しない
                If i > 1 And i < Len(buf) Then
                    If IsDigitChar(Mid$(buf, i - 1, 1)) And IsDigitChar(Mid$(buf, i + 1, 1)) Then
                        out = out & ch
                    Else
                        out = out & Chr(3)
                    End If
                Else
                    out = out & Chr(3)
                End If
            Case Else
                out = out & ch
        End Select
    Next i
    ' 空の要素を除く
    Dim parts() As String, keep() As String, n As Long
    parts = Split(out, Chr(3))
    ReDim keep(0 To UBound(parts))
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            keep(n) = Trim$(parts(i))
            n = n + 1
        End If
    Next i
    If n = 0 Then
        ReDim keep(0 To 0)
        keep(0) = Trim$(raw)
    Else
        ReDim Preserve keep(0 To n - 1)
    End If
    SplitItems = keep
End Function

Private Function IsDigitChar(ByVal ch As String) As Boolean
    IsDigitChar = (ch >= "0" And ch <= "9")
End Function

' 全角の英数字・記号（！〜～）と全角スペース・全角円記号を半角にする（文字数は変わらない。カナは変換しない）
Private Function NarrowAscii(ByVal s As String) As String
    Dim i As Long, code As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        If code >= &HFF01& And code <= &HFF5E& Then
            out = out & ChrW(code - &HFF01& + &H21)
        ElseIf code = &H3000 Then
            out = out & " "
        ElseIf code = &HFFE5& Then
            out = out & "\"
        Else
            out = out & ch
        End If
    Next i
    NarrowAscii = out
End Function

' 「基礎工事 800,000」「￥450,000」「外構：450000円」などから項目名と金額を取り出す
Private Sub ParseAmount(ByVal item As String, ByRef label As String, ByRef amount As Double, ByRef ok As Boolean)
    Dim s As String, i As Long, ch As String, numEnd As Long, numStart As Long, numText As String
    s = NarrowAscii(Trim$(item))                ' 全角の数字・記号を半角に（カナはそのまま）
    s = Replace(s, "円", "")
    s = Replace(s, "￥", "")
    s = Replace(s, ChrW(&HA5), "")   ' 半角の円記号
    s = Replace(s, "\", "")
    s = Trim$(s)
    ok = False
    label = ""
    ' 末尾側にある数字列（カンマ・小数点込み）を探す
    numEnd = 0
    For i = Len(s) To 1 Step -1
        If IsDigitChar(Mid$(s, i, 1)) Then
            numEnd = i
            Exit For
        End If
    Next i
    If numEnd = 0 Then Exit Sub
    numStart = numEnd
    For i = numEnd - 1 To 1 Step -1
        ch = Mid$(s, i, 1)
        If IsDigitChar(ch) Or ch = "," Or ch = "." Then
            numStart = i
        Else
            Exit For
        End If
    Next i
    numText = Replace(Mid$(s, numStart, numEnd - numStart + 1), ",", "")
    If Not IsNumeric(numText) Then Exit Sub
    ' 「10/25」「2024-10-25」「3月」「50%」「2回目」のような数字は金額ではないので除外
    If numStart > 1 Then
        If Mid$(s, numStart - 1, 1) = "/" Then Exit Sub
        If Mid$(s, numStart - 1, 1) = "-" And numStart > 2 Then
            If IsDigitChar(Mid$(s, numStart - 2, 1)) Then Exit Sub
        End If
    End If
    If numEnd < Len(s) Then
        If InStr("/%月日年時回件名人台", Mid$(s, numEnd + 1, 1)) > 0 Then Exit Sub
    End If
    If numStart > 1 Then
        If Mid$(s, numStart - 1, 1) = "-" Then
            If numStart - 1 = 1 Then
                numStart = numStart - 1: numText = "-" & numText
            ElseIf InStr(" :=(", Mid$(s, numStart - 2, 1)) > 0 Then
                numStart = numStart - 1: numText = "-" & numText
            End If
        End If
    End If
    amount = CDbl(numText)
    ok = True
    label = Trim$(Left$(s, numStart - 1))
    ' 「：」「:」「=」などの末尾記号を落とす
    Do While Len(label) > 0 And InStr("：:=＝ 　-ー・", Right$(label, 1)) > 0
        label = Left$(label, Len(label) - 1)
    Loop
    ' 数字の後ろに残った文言（例「(税抜)」）はラベルに付ける
    If numEnd < Len(s) Then
        Dim tail As String
        tail = Trim$(Mid$(s, numEnd + 1))
        If Len(tail) > 0 Then label = Trim$(label & " " & tail)
    End If
End Sub

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
        s = NarrowAscii(s)               ' 全角英数記号を半角に
        On Error Resume Next
        s = StrConv(s, vbNarrow)         ' カナも半角にそろえる（対応していない環境では何もしない）
        On Error GoTo 0
        s = UCase$(s)                    ' 英字の大文字小文字を無視
    End If
    NormalizeKey = s
End Function

' 見出し文字列または列記号から、src 配列上の列番号を得る（見つからなければ 0）
Private Function ResolveColumn(ByVal ws As Worksheet, ByRef hdr As Variant, _
                               ByVal firstCol As Long, ByVal specList As String) As Long
    ' specList はカンマ区切りの候補。見出し文字列の一致 → 列記号の順に、先に見つかったものを返す
    Dim specs() As String, i As Long, c As Long, spec As String, sheetCol As Long
    specs = Split(Replace(specList, "，", ","), ",")
    For i = LBound(specs) To UBound(specs)
        spec = Trim$(specs(i))
        If Len(spec) > 0 Then
            For c = 1 To UBound(hdr, 2)
                If HeaderText(hdr(1, c)) = HeaderText(spec) Then
                    ResolveColumn = c
                    Exit Function
                End If
            Next c
        End If
    Next i
    For i = LBound(specs) To UBound(specs)
        spec = Trim$(specs(i))
        sheetCol = 0
        On Error Resume Next
        sheetCol = ws.Columns(spec).Column
        On Error GoTo 0
        If sheetCol >= firstCol And sheetCol - firstCol + 1 <= UBound(hdr, 2) And sheetCol > 0 Then
            ResolveColumn = sheetCol - firstCol + 1
            Exit Function
        End If
    Next i
End Function

' カンマ区切りの見出し名リストに該当する列へ mode を設定する
Private Sub MarkColumns(ByRef mode() As Long, ByRef src As Variant, ByVal names As String, ByVal value As Long)
    If Len(Trim$(names)) = 0 Then Exit Sub
    Dim parts() As String, i As Long, c As Long
    parts = Split(Replace(names, "，", ","), ",")
    For i = LBound(parts) To UBound(parts)
        For c = 1 To UBound(src, 2)
            If HeaderText(src(1, c)) = HeaderText(parts(i)) Then mode(c) = value
        Next c
    Next i
End Sub

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' Collection に指定キーがあるか
Private Function CollHas(ByVal c As Collection, ByVal k As String) As Boolean
    On Error Resume Next
    c.Item k
    CollHas = (Err.Number = 0)
    On Error GoTo 0
End Function

' 見出し比較用：エラー値は空、改行・空白（全角含む）を除く
Private Function HeaderText(ByVal v As Variant) As String
    Dim t As String
    If IsError(v) Then Exit Function
    t = CStr(v)
    t = Replace(t, vbLf, ""): t = Replace(t, vbCr, "")
    t = Replace(t, " ", ""): t = Replace(t, "　", ""): t = Replace(t, vbTab, "")
    HeaderText = t
End Function

Private Function FirstUsedCol(ByVal ws As Worksheet, ByVal headerRow As Long) As Long
    Dim c As Long
    For c = 1 To ws.Columns.Count
        If Len(HeaderText(ws.Cells(headerRow, c).Value)) > 0 Then
            FirstUsedCol = c
            Exit Function
        End If
        If c > 50 Then Exit For   ' 見出しが 50 列以上右にあることは想定しない
    Next c
End Function

Private Function LastUsedCol(ByVal ws As Worksheet, ByVal headerRow As Long) As Long
    LastUsedCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
End Function
