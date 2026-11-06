Attribute VB_Name = "MailDraftCreator"
Option Explicit

Public Sub Outlook下書きメール作成()
    Const FIRST_ROW As Long = 3
    Const MAX_ROWS As Long = 100
    Const OL_MAIL_ITEM As Long = 0

    Dim wsData As Worksheet
    Dim wsText As Worksheet
    Dim outlookApp As Object
    Dim mailItem As Object
    Dim subjectText As String
    Dim bodyTemplate As String
    Dim bodyText As String
    Dim rowNo As Long
    Dim lastRow As Long
    Dim createdCount As Long
    Dim companyName As String
    Dim personName As String
    Dim workType As String
    Dim toAddress As String

    On Error GoTo ErrorHandler

    ' メール作成に使用する入力シートと文面シートを取得します。
    Set wsData = ThisWorkbook.Worksheets("Sheet1")
    Set wsText = ThisWorkbook.Worksheets("文面")

    ' 全メール共通の件名と本文テンプレートをシート「文面」から取得します。
    subjectText = CStr(wsText.Range("B1").Value)
    bodyTemplate = CStr(wsText.Range("B3").Value)

    ' 共通項目が未設定の場合は、下書きを作らず処理を中止します。
    If Len(Trim$(subjectText)) = 0 Then
        MsgBox "シート「文面」のセルB1に件名がありません。処理を中止します。", vbExclamation
        Exit Sub
    End If

    If Len(Trim$(bodyTemplate)) = 0 Then
        MsgBox "シート「文面」のセルB3に本文がありません。処理を中止します。", vbExclamation
        Exit Sub
    End If

    ' 起動済みのOutlook(classic)を操作するためのアプリケーションオブジェクトを取得します。
    Set outlookApp = GetOutlookApplication()
    If outlookApp Is Nothing Then
        MsgBox "Outlook(classic)が起動していません。Outlook(classic)を起動してから、再度マクロを実行してください。", vbCritical
        Exit Sub
    End If

    lastRow = FIRST_ROW + MAX_ROWS - 1

    ' シート「Sheet1」の3行目から、空行に当たるまで1行ずつ下書きメールを作成します。
    For rowNo = FIRST_ROW To lastRow
        If IsInputRowBlank(wsData, rowNo) Then
            Exit For
        End If

        ' 1行分の送信先情報を取得します。
        companyName = Trim$(CStr(wsData.Cells(rowNo, "C").Value))
        personName = Trim$(CStr(wsData.Cells(rowNo, "D").Value))
        toAddress = NormalizeMailAddresses(CStr(wsData.Cells(rowNo, "E").Value))
        workType = Trim$(CStr(wsData.Cells(rowNo, "B").Value))

        ' 必須項目が欠けている行があれば、誤った下書きを作らず処理を中止します。
        If Len(companyName) = 0 Then
            MsgBox "シート「Sheet1」のC" & rowNo & "に会社名がありません。処理を中止します。", vbExclamation
            Exit Sub
        End If

        If Len(personName) = 0 Then
            MsgBox "シート「Sheet1」のD" & rowNo & "に担当者名がありません。処理を中止します。", vbExclamation
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

        ' 本文テンプレート内の差し込み文字を、対象行の内容で置き換えます。
        bodyText = bodyTemplate
        bodyText = Replace(bodyText, "【会社名】", companyName)

        ' 担当者名が「御中」の場合は、前の改行を全角スペースに変えて同じ行に続けます。
        If personName = "御中" Then
            bodyText = Replace(bodyText, vbCrLf & "【担当者名】", "　【担当者名】")
            bodyText = Replace(bodyText, vbCr & "【担当者名】", "　【担当者名】")
            bodyText = Replace(bodyText, vbLf & "【担当者名】", "　【担当者名】")
        End If

        bodyText = Replace(bodyText, "【担当者名】", personName)
        bodyText = Replace(bodyText, "【工種】", workType)

        ' Outlookにメールを1件作成し、送信せず下書きとして保存します。
        Set mailItem = outlookApp.CreateItem(OL_MAIL_ITEM)
        With mailItem
            .To = toAddress
            .Subject = subjectText
            .Body = bodyText
            .Save
        End With

        createdCount = createdCount + 1
    Next rowNo

    MsgBox createdCount & "件のメールをOutlookの下書きに保存しました。" & vbCrLf & _
           "メール送信は実行していません。", vbInformation
    Exit Sub

ErrorHandler:
    ' 想定外のエラーが出た場合も、どこで止まったかわかるように番号と内容を表示します。
    MsgBox "下書きメール作成中にエラーが発生しました。" & vbCrLf & _
           Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Function GetOutlookApplication() As Object
    ' 起動済みのOutlookだけを取得します。VBAからOutlookの自動起動は行いません。
    On Error Resume Next
    Set GetOutlookApplication = GetObject(, "Outlook.Application")
    On Error GoTo 0
End Function

Private Function IsInputRowBlank(ByVal ws As Worksheet, ByVal rowNo As Long) As Boolean
    ' 送信先情報の対象列がすべて空なら、入力行の終わりと判断します。
    IsInputRowBlank = _
        Len(Trim$(CStr(ws.Cells(rowNo, "A").Value))) = 0 And _
        Len(Trim$(CStr(ws.Cells(rowNo, "B").Value))) = 0 And _
        Len(Trim$(CStr(ws.Cells(rowNo, "C").Value))) = 0 And _
        Len(Trim$(CStr(ws.Cells(rowNo, "D").Value))) = 0 And _
        Len(Trim$(CStr(ws.Cells(rowNo, "E").Value))) = 0
End Function

Private Function NormalizeMailAddresses(ByVal rawText As String) As String
    ' セル内改行やカンマ区切りの宛先を、Outlook用のセミコロン区切りに整えます。
    Dim normalized As String
    Dim parts() As String
    Dim part As Variant
    Dim result As String

    normalized = Replace(rawText, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    normalized = Replace(normalized, ";", vbLf)
    normalized = Replace(normalized, ",", vbLf)

    parts = Split(normalized, vbLf)
    For Each part In parts
        If Len(Trim$(CStr(part))) > 0 Then
            If Len(result) > 0 Then
                result = result & "; "
            End If
            result = result & Trim$(CStr(part))
        End If
    Next part

    NormalizeMailAddresses = result
End Function
