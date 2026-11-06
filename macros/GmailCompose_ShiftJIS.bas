Attribute VB_Name = "GmailCompose"
Option Explicit

'==============================================================================
' Gmail で作成画面を開く（Outlook を使わない方法）
'
'   シート「Sheet1」の 3 行目から、空行に当たるまで 1 行ずつ、ブラウザで Gmail の
'   メール作成画面（宛先・件名・本文が入った状態）を開きます。送信はしません。
'   作成画面を閉じると Gmail が自動的に下書きとして保存します。
'
'   件名・本文の取り方、差し込み（【会社名】【担当者名】【工種】、「御中」の扱い）は
'   「Outlook下書きメール作成」と同じです。シート「文面」の B1 が件名、B3 が本文です。
'
'   ※ ブラウザで Gmail にログインしている必要があります。
'   ※ 件数が多いとタブがたくさん開くので、実行前に件数を確認して進めるか尋ねます。
'==============================================================================

Public Sub Gmailで作成画面を開く()
    Const FIRST_ROW As Long = 3
    Const MAX_ROWS As Long = 100

    Dim wsData As Worksheet, wsText As Worksheet
    Dim subjectText As String, bodyTemplate As String, bodyText As String
    Dim rowNo As Long, lastRow As Long, n As Long, i As Long
    Dim companyName As String, personName As String, workType As String, toAddress As String
    Dim urls() As String

    On Error GoTo ErrorHandler

    Set wsData = ThisWorkbook.Worksheets("Sheet1")
    Set wsText = ThisWorkbook.Worksheets("文面")
    subjectText = CStr(wsText.Range("B1").Value)
    bodyTemplate = CStr(wsText.Range("B3").Value)

    If Len(Trim$(subjectText)) = 0 Then
        MsgBox "シート「文面」のセルB1に件名がありません。処理を中止します。", vbExclamation
        Exit Sub
    End If
    If Len(Trim$(bodyTemplate)) = 0 Then
        MsgBox "シート「文面」のセルB3に本文がありません。処理を中止します。", vbExclamation
        Exit Sub
    End If

    lastRow = FIRST_ROW + MAX_ROWS - 1
    ReDim urls(1 To MAX_ROWS)

    For rowNo = FIRST_ROW To lastRow
        If RowIsBlank(wsData, rowNo) Then Exit For

        companyName = Trim$(CStr(wsData.Cells(rowNo, "C").Value))
        personName = Trim$(CStr(wsData.Cells(rowNo, "D").Value))
        toAddress = NormalizeAddresses(CStr(wsData.Cells(rowNo, "E").Value))
        workType = Trim$(CStr(wsData.Cells(rowNo, "B").Value))

        If Len(companyName) = 0 Then
            MsgBox "シート「Sheet1」のC" & rowNo & "に会社名がありません。処理を中止します。", vbExclamation
            Exit Sub
        End If
        If Len(personName) = 0 Then
            MsgBox "シート「Sheet1」のD" & rowNo & "に担当者名がありません（会社宛てなら「御中」と入れてください）。処理を中止します。", vbExclamation
            Exit Sub
        End If
        If Len(toAddress) = 0 Then
            MsgBox "シート「Sheet1」のE" & rowNo & "にメールアドレスがありません。処理を中止します。", vbExclamation
            Exit Sub
        End If
        If Len(workType) = 0 Then
            MsgBox "シート「Sheet1」のB" & rowNo & "に工種がありません。処理を中止します。", vbExclamation
            Exit Sub
        End If

        bodyText = bodyTemplate
        bodyText = Replace(bodyText, "【会社名】", companyName)
        If personName = "御中" Then
            bodyText = Replace(bodyText, vbCrLf & "【担当者名】", "　【担当者名】")
            bodyText = Replace(bodyText, vbCr & "【担当者名】", "　【担当者名】")
            bodyText = Replace(bodyText, vbLf & "【担当者名】", "　【担当者名】")
        End If
        bodyText = Replace(bodyText, "【担当者名】", personName)
        bodyText = Replace(bodyText, "【工種】", workType)
        bodyText = Replace(bodyText, vbCrLf, vbLf)
        bodyText = Replace(bodyText, vbCr, vbLf)
        bodyText = Replace(bodyText, vbLf, vbCrLf)

        n = n + 1
        urls(n) = "https://mail.google.com/mail/?view=cm&fs=1" & _
                  "&to=" & UrlEncode(toAddress) & _
                  "&su=" & UrlEncode(subjectText) & _
                  "&body=" & UrlEncode(bodyText)
        If Len(urls(n)) > 7500 Then
            MsgBox "シート「Sheet1」の" & rowNo & "行目は本文が長すぎてブラウザで開けません（上限の目安は約 2,000 文字）。処理を中止します。", vbExclamation
            Exit Sub
        End If
    Next rowNo

    If n = 0 Then
        MsgBox "シート「Sheet1」の3行目以降にデータがありません。先に「業者を1行に集約」を押してください。", vbExclamation
        Exit Sub
    End If

    If MsgBox(n & " 件分の Gmail 作成画面をブラウザで開きます。" & vbCrLf & _
              "（ブラウザで Gmail にログインしておいてください。送信はしません）" & vbCrLf & vbCrLf & _
              "続けますか？", vbOKCancel + vbQuestion) <> vbOK Then Exit Sub

    For i = 1 To n
        ThisWorkbook.FollowHyperlink urls(i)
    Next i

    MsgBox n & " 件の作成画面を開きました。" & vbCrLf & _
           "各画面で内容を確認して送信するか、閉じれば下書きに残ります。", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Gmail 作成画面を開く途中でエラーが発生しました。" & vbCrLf & _
           Err.Number & ": " & Err.Description, vbCritical
End Sub

' URL 用に UTF-8 でパーセントエンコードする（英数字と - _ . ~ 以外をすべて変換）
'   Excel 2013 以降の ENCODEURL 関数を使い、使えない環境では自前で変換する
Private Function UrlEncode(ByVal s As String) As String
    On Error Resume Next
    UrlEncode = Application.WorksheetFunction.EncodeURL(s)
    If Err.Number = 0 And Len(UrlEncode) > 0 Then Exit Function
    Err.Clear
    On Error GoTo 0
    UrlEncode = UrlEncodeManual(s)
End Function

Private Function UrlEncodeManual(ByVal s As String) As String
    Dim i As Long, code As Long, ch As String, out As String
    Dim b1 As Long, b2 As Long, b3 As Long, b4 As Long, hi As Long, lo As Long
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        ' サロゲートペア（絵文字など）
        If code >= &HD800& And code <= &HDBFF& And i < Len(s) Then
            hi = code
            lo = AscW(Mid$(s, i + 1, 1))
            If lo < 0 Then lo = lo + 65536
            code = &H10000 + (hi - &HD800&) * &H400& + (lo - &HDC00&)
            i = i + 1
        End If
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) _
           Or code = 45 Or code = 95 Or code = 46 Or code = 126 Then
            out = out & ChrW(code)
        ElseIf code < &H80 Then
            out = out & "%" & Right$("0" & Hex$(code), 2)
        ElseIf code < &H800 Then
            b1 = &HC0 Or (code \ &H40): b2 = &H80 Or (code And &H3F)
            out = out & "%" & Hex$(b1) & "%" & Hex$(b2)
        ElseIf code < &H10000 Then
            b1 = &HE0 Or (code \ &H1000): b2 = &H80 Or ((code \ &H40) And &H3F): b3 = &H80 Or (code And &H3F)
            out = out & "%" & Hex$(b1) & "%" & Hex$(b2) & "%" & Hex$(b3)
        Else
            b1 = &HF0 Or (code \ &H40000): b2 = &H80 Or ((code \ &H1000) And &H3F)
            b3 = &H80 Or ((code \ &H40) And &H3F): b4 = &H80 Or (code And &H3F)
            out = out & "%" & Hex$(b1) & "%" & Hex$(b2) & "%" & Hex$(b3) & "%" & Hex$(b4)
        End If
        i = i + 1
    Loop
    UrlEncodeManual = out
End Function

Private Function RowIsBlank(ByVal ws As Worksheet, ByVal rowNo As Long) As Boolean
    RowIsBlank = _
        Len(Trim$(CStr(ws.Cells(rowNo, "B").Value))) = 0 And _
        Len(Trim$(CStr(ws.Cells(rowNo, "C").Value))) = 0 And _
        Len(Trim$(CStr(ws.Cells(rowNo, "D").Value))) = 0 And _
        Len(Trim$(CStr(ws.Cells(rowNo, "E").Value))) = 0
End Function

' セル内改行・カンマ・セミコロン区切りの宛先を、カンマ区切りにそろえる（Gmail 用）
Private Function NormalizeAddresses(ByVal rawText As String) As String
    Dim t As String, parts() As String, part As Variant, result As String
    t = Replace(rawText, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    t = Replace(t, ";", vbLf)
    t = Replace(t, "；", vbLf)
    t = Replace(t, "，", vbLf)
    t = Replace(t, ",", vbLf)
    parts = Split(t, vbLf)
    For Each part In parts
        If Len(Trim$(CStr(part))) > 0 Then
            If Len(result) > 0 Then result = result & ","
            result = result & Trim$(CStr(part))
        End If
    Next part
    NormalizeAddresses = result
End Function
