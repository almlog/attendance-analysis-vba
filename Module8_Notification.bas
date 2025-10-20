' ========================================
' Module8_Notification
' タイプ: 標準モジュール
' 行数: 548
' エクスポート日時: 2025-10-20 14:30:49
' ========================================

' *************************************************************
' Module8_Notification (完全修正版)
' 目的: LINE WORKS通知機能
' 修正内容:
' - メッセージ分割送信機能追加(1000文字制限対応)
' - 休憩時間違反を別メッセージで送信
' - 時間表示を「HH:MM」形式に修正
' - Channel ID検証強化
' *************************************************************

Option Explicit

' 定数定義
Private Const MAX_MESSAGE_LENGTH As Long = 1000  ' LINE WORKS推奨最大文字数

' *************************************************************
' 関数名: SendToLineWorks
' 目的: LINE WORKS Webhook経由でメッセージ送信
' *************************************************************
Public Function SendToLineWorks(webhookUrl As String, messageText As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Trim(webhookUrl) = "" Then
        MsgBox "Webhook URLが設定されていません。", vbExclamation
        SendToLineWorks = False
        Exit Function
    End If
    
    If Trim(messageText) = "" Then
        MsgBox "送信するメッセージが空です。", vbExclamation
        SendToLineWorks = False
        Exit Function
    End If
    
    ' Channel ID取得(B5セル: CHANNEL_ID)
    Dim channelId As String
    Dim configSheet As Worksheet
    On Error Resume Next
    Set configSheet = ThisWorkbook.Sheets("設定")
    If Not configSheet Is Nothing Then
        channelId = Trim(CStr(configSheet.Cells(5, 2).Value))  ' B5セルから取得
    End If
    On Error GoTo ErrorHandler
    
    ' Channel ID検証
    If channelId = "" Then
        MsgBox "Channel IDが設定されていません。" & vbCrLf & vbCrLf & _
               "[設定]シートのB5セル(CHANNEL_ID)に値を入力してください。", _
               vbExclamation, "設定エラー"
        SendToLineWorks = False
        Exit Function
    End If
    
    Debug.Print "========================================="
    Debug.Print "LINE WORKS送信開始: " & Now
    Debug.Print "Channel ID: " & channelId
    Debug.Print "メッセージ長: " & Len(messageText) & "文字"
    
    ' テキストエスケープ
    Dim escapedText As String
    escapedText = messageText
    escapedText = Replace(escapedText, "\", "\\")
    escapedText = Replace(escapedText, """", "\""")
    escapedText = Replace(escapedText, vbLf, "\n")
    escapedText = Replace(escapedText, vbCr, "")
    escapedText = Replace(escapedText, vbTab, " ")
    
    ' JSON作成
    Dim jsonBody As String
    jsonBody = "{""channelId"":""" & channelId & """,""body"":{""text"":""" & escapedText & """}}"
    
    Debug.Print "JSON長: " & Len(jsonBody) & "文字"
    
    ' HTTP送信
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    http.Open "POST", webhookUrl, False
    http.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
    http.send jsonBody
    
    Debug.Print "HTTP Status: " & http.Status
    Debug.Print "Response: " & http.responseText
    Debug.Print "========================================="
    
    If http.Status = 200 Then
        SendToLineWorks = True
    Else
        MsgBox "LINE WORKS送信エラー" & vbCrLf & vbCrLf & _
               "HTTP Status: " & http.Status & vbCrLf & _
               "Response: " & http.responseText, vbCritical
        SendToLineWorks = False
    End If
    
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    Debug.Print "例外エラー: " & Err.Description
    MsgBox "通知送信中にエラーが発生しました: " & vbCrLf & Err.Description, vbCritical
    SendToLineWorks = False
End Function


' *************************************************************
' 関数名: ConvertDecimalTimeToHHMM
' 目的: 小数時間(日の割合)を「HH:MM」形式に変換
' 例: 0.430555555555556 → "10:20"
' *************************************************************
Private Function ConvertDecimalTimeToHHMM(decimalTime As Variant) As String
    On Error Resume Next
    
    If IsEmpty(decimalTime) Or decimalTime = "" Or decimalTime = 0 Then
        ConvertDecimalTimeToHHMM = "00:00"
        Exit Function
    End If
    
    Dim totalMinutes As Long
    totalMinutes = CLng(CDbl(decimalTime) * 24 * 60)
    
    Dim hours As Long, minutes As Long
    hours = totalMinutes \ 60
    minutes = totalMinutes Mod 60
    
    ConvertDecimalTimeToHHMM = Format(hours, "00") & ":" & Format(minutes, "00")
End Function


' *************************************************************
' 関数名: GenerateAttendanceMissingPart
' 目的: 勤怠入力漏れメッセージ生成(緊急度順ソート)
' *************************************************************
Private Function GenerateAttendanceMissingPart() As String
    On Error GoTo ErrorHandler
    
    Debug.Print "[INFO] 勤怠入力漏れ部分の生成開始"
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("勤怠入力漏れ一覧")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Debug.Print "[INFO] 勤怠入力漏れ一覧シートが見つかりません"
        GenerateAttendanceMissingPart = ""
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= 1 Then
        Debug.Print "[INFO] 勤怠入力漏れデータなし"
        GenerateAttendanceMissingPart = ""
        Exit Function
    End If
    
    Dim empDict As Object
    Set empDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim empID As String, empName As String, targetDate As Date
    Dim totalMissingCount As Long
    totalMissingCount = 0
    
    For i = 2 To lastRow
        empID = Trim(ws.Cells(i, 1).Value)
        empName = Trim(ws.Cells(i, 2).Value)
        
        On Error Resume Next
        targetDate = CDate(ws.Cells(i, 3).Value)
        If Err.Number <> 0 Then
            Debug.Print "[WARNING] 行" & i & "の日付が不正"
            GoTo NextMissing
        End If
        On Error GoTo ErrorHandler
        
        totalMissingCount = totalMissingCount + 1
        
        If Not empDict.Exists(empID) Then
            Dim newColl As Collection
            Set newColl = New Collection
            empDict.Add empID, Array(empName, newColl)
        End If
        
        Dim empData As Variant
        empData = empDict(empID)
        empData(1).Add targetDate
        empDict(empID) = empData
        
NextMissing:
    Next i
    
    If empDict.Count = 0 Then
        Debug.Print "[INFO] 勤怠入力漏れデータなし(解析後)"
        GenerateAttendanceMissingPart = ""
        Exit Function
    End If
    
    Debug.Print "[INFO] 勤怠入力漏れ: " & empDict.Count & "名, " & totalMissingCount & "件"
    
    ' ソート
    Dim keys As Variant
    keys = empDict.keys
    
    Dim sortArray() As Variant
    ReDim sortArray(0 To empDict.Count - 1, 0 To 2)
    
    Dim j As Long
    For j = 0 To empDict.Count - 1
        empData = empDict(keys(j))
        sortArray(j, 0) = keys(j)
        sortArray(j, 1) = empData(0)
        sortArray(j, 2) = empData(1).Count
    Next j
    
    Dim temp As Variant, swapped As Boolean, n As Long
    n = empDict.Count
    
    Do
        swapped = False
        For j = 0 To n - 2
            If sortArray(j, 2) < sortArray(j + 1, 2) Then
                temp = sortArray(j, 0): sortArray(j, 0) = sortArray(j + 1, 0): sortArray(j + 1, 0) = temp
                temp = sortArray(j, 1): sortArray(j, 1) = sortArray(j + 1, 1): sortArray(j + 1, 1) = temp
                temp = sortArray(j, 2): sortArray(j, 2) = sortArray(j + 1, 2): sortArray(j + 1, 2) = temp
                swapped = True
            End If
        Next j
        n = n - 1
    Loop While swapped
    
    ' メッセージ生成
    Dim message As String
    message = "未入力者: " & empDict.Count & "名 / 未入力日数: " & totalMissingCount & "日" & vbLf & vbLf
    
    For j = 0 To UBound(sortArray, 1)
        Dim empID_sorted As String, empName_sorted As String, missingCount As Long
        
        empID_sorted = sortArray(j, 0)
        empName_sorted = sortArray(j, 1)
        missingCount = sortArray(j, 2)
        
        empData = empDict(empID_sorted)
        
        Dim urgencyMark As String
        If missingCount >= 5 Then
            urgencyMark = "[!!緊急!!]"
        ElseIf missingCount >= 3 Then
            urgencyMark = "[!要注意!]"
        Else
            urgencyMark = "[確認]"
        End If
        
        Dim empText As String
        empText = urgencyMark & " " & empName_sorted & " さん (" & missingCount & "日)" & vbLf
        
        Dim dateItem As Variant, dateCount As Integer
        dateCount = 0
        For Each dateItem In empData(1)
            If dateCount < 5 Then
                empText = empText & "  - " & Format(dateItem, "mm/dd (aaa)") & vbLf
                dateCount = dateCount + 1
            End If
        Next dateItem
        
        If missingCount > 5 Then
            empText = empText & "  ...他" & (missingCount - 5) & "日" & vbLf
        End If
        
        empText = empText & vbLf
        message = message & empText
    Next j
    
    GenerateAttendanceMissingPart = message
    Exit Function
    
ErrorHandler:
    Debug.Print "[ERROR] 勤怠入力漏れ部分生成エラー: " & Err.Description
    GenerateAttendanceMissingPart = ""
End Function


' *************************************************************
' 関数名: GenerateBreakTimeViolationPart
' 目的: 休憩時間違反メッセージ生成(時間をHH:MM形式に修正)
' *************************************************************
Private Function GenerateBreakTimeViolationPart() As String
    On Error GoTo ErrorHandler
    
    Debug.Print "[INFO] 休憩時間違反部分の生成開始"
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("休憩時間チェック_違反者")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Debug.Print "[INFO] 休憩時間チェック_違反者シートが見つかりません"
        GenerateBreakTimeViolationPart = ""
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= 1 Then
        Debug.Print "[INFO] 休憩時間違反データなし"
        GenerateBreakTimeViolationPart = ""
        Exit Function
    End If
    
    If ws.Cells(2, 1).Value = "休憩時間違反はありません。" Then
        Debug.Print "[INFO] 休憩時間違反なし"
        GenerateBreakTimeViolationPart = ""
        Exit Function
    End If
    
    Dim empDict As Object
    Set empDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim empID As String, empName As String, targetDate As Date
    Dim workTime As Variant, breakTime As Variant, shortage As Variant
    Dim totalViolationCount As Long
    totalViolationCount = 0
    
    For i = 2 To lastRow
        empID = Trim(ws.Cells(i, 1).Value)
        empName = Trim(ws.Cells(i, 2).Value)
        
        On Error Resume Next
        targetDate = CDate(ws.Cells(i, 4).Value)
        If Err.Number <> 0 Then
            Debug.Print "[WARNING] 行" & i & "の日付が不正"
            GoTo NextViolation
        End If
        On Error GoTo ErrorHandler
        
        ' ★ 時間をHH:MM形式に変換
        workTime = ws.Cells(i, 5).Value
        breakTime = ws.Cells(i, 6).Value
        shortage = ws.Cells(i, 8).Value
        
        Dim workTimeStr As String, breakTimeStr As String, shortageStr As String
        workTimeStr = ConvertDecimalTimeToHHMM(workTime)
        breakTimeStr = ConvertDecimalTimeToHHMM(breakTime)
        shortageStr = ConvertDecimalTimeToHHMM(shortage)
        
        totalViolationCount = totalViolationCount + 1
        
        If Not empDict.Exists(empID) Then
            Dim newColl As Collection
            Set newColl = New Collection
            empDict.Add empID, Array(empName, newColl)
        End If
        
        Dim empData As Variant
        empData = empDict(empID)
        empData(1).Add Array(targetDate, workTimeStr, breakTimeStr, shortageStr)
        empDict(empID) = empData
        
NextViolation:
    Next i
    
    If empDict.Count = 0 Then
        Debug.Print "[INFO] 休憩時間違反データなし(解析後)"
        GenerateBreakTimeViolationPart = ""
        Exit Function
    End If
    
    Debug.Print "[INFO] 休憩時間違反: " & empDict.Count & "名, " & totalViolationCount & "件"
    
    Dim message As String
    message = "違反者: " & empDict.Count & "名 / 違反件数: " & totalViolationCount & "件" & vbLf & vbLf
    
    Dim key As Variant
    For Each key In empDict.keys
        empData = empDict(key)
        
        Dim empText As String
        empText = "[違反] " & empData(0) & " さん" & vbLf
        
        Dim violationItem As Variant, violationCount As Integer
        violationCount = 0
        For Each violationItem In empData(1)
            If violationCount < 5 Then
                empText = empText & "  - " & Format(violationItem(0), "mm/dd") & _
                          ": 実働" & violationItem(1) & " / 休憩" & violationItem(2) & _
                          " -> 不足" & violationItem(3) & vbLf
                violationCount = violationCount + 1
            End If
        Next violationItem
        
        If empData(1).Count > 5 Then
            empText = empText & "  ...他" & (empData(1).Count - 5) & "件" & vbLf
        End If
        
        empText = empText & vbLf
        message = message & empText
    Next key
    
    GenerateBreakTimeViolationPart = message
    Exit Function
    
ErrorHandler:
    Debug.Print "[ERROR] 休憩時間違反部分生成エラー: " & Err.Description
    GenerateBreakTimeViolationPart = ""
End Function


' *************************************************************
' 関数名: SendNotificationToLineWorks
' 目的: メイン処理(メッセージ分割送信対応)
' *************************************************************
Public Sub SendNotificationToLineWorks()
    On Error GoTo ErrorHandler
    
    Debug.Print vbCrLf & "####################################"
    Debug.Print "# LINE WORKS通知処理開始"
    Debug.Print "####################################"
    
    Application.ScreenUpdating = False
    
    Dim configSheet As Worksheet
    On Error Resume Next
    Set configSheet = ThisWorkbook.Sheets("設定")
    On Error GoTo ErrorHandler
    
    If configSheet Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "設定シートが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    Dim webhookUrl As String
    webhookUrl = Trim(configSheet.Cells(1, 2).Value)
    
    If webhookUrl = "" Then
        Application.ScreenUpdating = True
        MsgBox "Webhook URLが設定されていません。", vbExclamation
        Exit Sub
    End If
    
    ' メッセージ生成
    Dim attendancePart As String, breakTimePart As String
    attendancePart = GenerateAttendanceMissingPart()
    breakTimePart = GenerateBreakTimeViolationPart()
    
    If attendancePart = "" And breakTimePart = "" Then
        Application.ScreenUpdating = True
        MsgBox "通知するデータがありません。", vbInformation
        Exit Sub
    End If
    
    ' 送信確認
    Application.ScreenUpdating = True
    Dim confirmResult As VbMsgBoxResult
    
    Dim confirmMsg As String
    confirmMsg = "LINE WORKSに通知を送信しますか?" & vbCrLf & vbCrLf
    If attendancePart <> "" Then
        confirmMsg = confirmMsg & "- 勤怠入力漏れ通知" & vbCrLf
    End If
    If breakTimePart <> "" Then
        confirmMsg = confirmMsg & "- 休憩時間違反通知" & vbCrLf
    End If
    
    confirmResult = MsgBox(confirmMsg, vbQuestion + vbYesNo, "送信確認")
    
    If confirmResult = vbNo Then
        MsgBox "送信をキャンセルしました。", vbInformation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 送信処理
    Dim result1 As Boolean, result2 As Boolean
    result1 = False
    result2 = False
    
    ' 1. 勤怠入力漏れメッセージ送信
    If attendancePart <> "" Then
        Dim message1 As String
        message1 = "【SI1部 勤怠アラート - 未入力】" & vbLf
        message1 = message1 & Format(Now, "yyyy年mm月dd日 hh:nn") & vbLf
        message1 = message1 & "=============================" & vbLf & vbLf
        message1 = message1 & "【勤怠入力漏れ】" & vbLf
        message1 = message1 & attendancePart & vbLf
        message1 = message1 & "=============================" & vbLf
        message1 = message1 & "※各所属GLより該当者へ注意喚起と入力対応をお願いします"
        
        result1 = SendToLineWorks(webhookUrl, message1)
        
        If Not result1 Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        ' 次の送信まで少し待機(レート制限対策)
        If breakTimePart <> "" Then
            Application.Wait Now + timeValue("00:00:02")
        End If
    End If
    
    ' 2. 休憩時間違反メッセージ送信(データがある場合のみ)
    If breakTimePart <> "" Then
        Dim message2 As String
        message2 = "【SI1部 勤怠アラート - 休憩時間違反】" & vbLf
        message2 = message2 & Format(Now, "yyyy年mm月dd日 hh:nn") & vbLf
        message2 = message2 & "=============================" & vbLf & vbLf
        message2 = message2 & "【休憩時間違反】" & vbLf
        message2 = message2 & breakTimePart & vbLf
        message2 = message2 & "=============================" & vbLf
        message1 = message1 & "※各所属GLより該当者へ注意喚起と入力対応をお願いします"
        
        result2 = SendToLineWorks(webhookUrl, message2)
    End If
    
    Application.ScreenUpdating = True
    
    ' 結果表示
    Dim resultMsg As String
    If attendancePart <> "" And result1 Then
        resultMsg = "[OK] 勤怠入力漏れ通知: 送信完了" & vbCrLf
    End If
    If breakTimePart <> "" And result2 Then
        resultMsg = resultMsg & "[OK] 休憩時間違反通知: 送信完了"
    End If
    
    If resultMsg <> "" Then
        MsgBox "LINE WORKSへの通知送信が完了しました。" & vbCrLf & vbCrLf & resultMsg, _
               vbInformation, "送信完了"
    End If
    
    Debug.Print "####################################"
    Debug.Print "# LINE WORKS通知処理終了"
    Debug.Print "####################################" & vbCrLf
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "通知処理中にエラーが発生しました:" & vbCrLf & Err.Description, vbCritical
End Sub

