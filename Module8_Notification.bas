' ========================================
' Module8_Notification
' タイプ: 標準モジュール
' 行数: 623
' エクスポート日時: 2025-10-20 11:00:27
' ========================================

Option Explicit

' *************************************************************
' Module8_Notification
' 目的：LINE WORKS Webhook通知機能
' 作成日：2025-10-18
' 改版履歴：
' 2025/10/18 初版作成 - Webhook通知機能追加
' 2025/10/20 絵文字完全排除版 - ASCII文字のみ使用
' 2025/10/20 ADODB.Stream削除版 - シンプルな文字列送信
' *************************************************************

' *************************************************************
' 関数名: SendToLineWorks
' 目的: LINE WORKS WebhookにPOSTリクエストを送信
' 引数:
'   - webhookURL: Webhook URL
'   - messageText: 送信するメッセージ本文
' 戻り値:
'   - True: 送信成功(HTTP 200)
'   - False: 送信失敗
' *************************************************************
Public Function SendToLineWorks(webhookUrl As String, messageText As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 引数チェック
    If Trim(webhookUrl) = "" Then
        MsgBox "Webhook URLが設定されていません。" & vbCrLf & _
               "[設定]シートを確認してください。", vbExclamation, "設定エラー"
        SendToLineWorks = False
        Exit Function
    End If
    
    If Trim(messageText) = "" Then
        MsgBox "送信するメッセージが空です。", vbExclamation, "メッセージエラー"
        SendToLineWorks = False
        Exit Function
    End If
    
    ' Channel ID取得(GetChannelIDがModule1に存在しない場合のフォールバック)
    Dim channelID As String
    On Error Resume Next
    channelID = GetChannelID()
    If Err.Number <> 0 Then
        ' GetChannelID関数がない場合は設定シートから直接取得
        Err.Clear
        On Error GoTo ErrorHandler
        Dim configSheet As Worksheet
        Set configSheet = ThisWorkbook.Sheets("設定")
        channelID = Trim(configSheet.Cells(2, 2).Value)
    End If
    On Error GoTo ErrorHandler
    
    If channelID = "" Then
        MsgBox "Channel IDが設定されていません。" & vbCrLf & _
               "[設定]シートのB2セルにChannel IDを入力してください。", _
               vbExclamation, "設定エラー"
        SendToLineWorks = False
        Exit Function
    End If
    
    ' デバッグログ
    Debug.Print "========================================="
    Debug.Print "LINE WORKS送信開始: " & Now
    Debug.Print "URL: " & Left(webhookUrl, 50) & "..."
    Debug.Print "Channel ID: " & channelID
    Debug.Print "メッセージ長: " & Len(messageText) & "文字"
    
    ' テキストエスケープ処理
    Dim escapedText As String
    escapedText = messageText
    escapedText = Replace(escapedText, "\", "\\")      ' バックスラッシュ
    escapedText = Replace(escapedText, """", "\""")    ' ダブルクォート
    escapedText = Replace(escapedText, vbLf, "\n")     ' 改行(LF)
    escapedText = Replace(escapedText, vbCr, "")       ' 改行(CR)を削除
    escapedText = Replace(escapedText, vbTab, " ")     ' タブを空白に
    
    ' JSON構築
    Dim jsonBody As String
    jsonBody = "{""channelId"":""" & channelID & """,""body"":{""text"":""" & escapedText & """}}"
    
    Debug.Print "JSON長: " & Len(jsonBody) & "バイト"
    
    ' HTTP送信（文字列を直接送信）
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    http.Open "POST", webhookUrl, False
    http.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
    http.send jsonBody
    
    ' レスポンス確認
    Debug.Print "HTTP Status: " & http.Status
    Debug.Print "Response: " & http.responseText
    Debug.Print "========================================="
    
    If http.Status = 200 Then
        Debug.Print "送信成功"
        SendToLineWorks = True
    Else
        Debug.Print "送信失敗: HTTP " & http.Status
        MsgBox "LINE WORKS送信エラー" & vbCrLf & vbCrLf & _
               "HTTP Status: " & http.Status & vbCrLf & _
               "Response: " & http.responseText, _
               vbCritical, "送信エラー"
        SendToLineWorks = False
    End If
    
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    Debug.Print "例外エラー: " & Err.Description
    Debug.Print "エラー番号: " & Err.Number
    Debug.Print "========================================="
    MsgBox "通知送信中にエラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "エラー"
    SendToLineWorks = False
End Function


' *************************************************************
' 関数名: GenerateMessageFromSheet
' 目的: 統合メッセージ生成(勤怠入力漏れ + 休憩時間違反)
' 戻り値: フォーマット済みメッセージ(ASCII文字のみ)
' *************************************************************
Public Function GenerateMessageFromSheet() As String
    On Error GoTo ErrorHandler
    
    Debug.Print "========================================="
    Debug.Print "統合メッセージ生成開始: " & Now
    
    ' 各セクションのメッセージを生成
    Dim attendancePart As String
    Dim breakTimePart As String
    Dim finalMessage As String
    
    ' 1. 勤怠入力漏れ部分を生成
    attendancePart = GenerateAttendanceMissingPart()
    
    ' 2. 休憩時間違反部分を生成
    breakTimePart = GenerateBreakTimeViolationPart()
    
    ' 3. 両方とも空の場合
    If attendancePart = "" And breakTimePart = "" Then
        MsgBox "通知するデータがありません。" & vbCrLf & vbCrLf & _
               "・勤怠入力漏れ: なし" & vbCrLf & _
               "・休憩時間違反: なし", _
               vbInformation, "データなし"
        GenerateMessageFromSheet = ""
        Exit Function
    End If
    
    ' 4. ヘッダー作成
    finalMessage = "【SI1部 勤怠アラート】" & vbLf
    finalMessage = finalMessage & Format(Now, "yyyy年mm月dd日 hh:nn") & vbLf
    finalMessage = finalMessage & "=============================" & vbLf & vbLf
    
    ' 5. 勤怠入力漏れセクション追加
    If attendancePart <> "" Then
        finalMessage = finalMessage & "【勤怠入力漏れ】" & vbLf
        finalMessage = finalMessage & attendancePart & vbLf
    End If
    
    ' 6. 休憩時間違反セクション追加
    If breakTimePart <> "" Then
        If attendancePart <> "" Then
            finalMessage = finalMessage & vbLf ' セクション間の空行
        End If
        finalMessage = finalMessage & "【休憩時間違反】" & vbLf
        finalMessage = finalMessage & breakTimePart & vbLf
    End If
    
    ' 7. フッター追加
    finalMessage = finalMessage & vbLf
    finalMessage = finalMessage & "=============================" & vbLf
    finalMessage = finalMessage & "※各リーダーより該当者へ" & vbLf
    finalMessage = finalMessage & "  対応をお願いします" & vbLf
    finalMessage = finalMessage & "※申請決裁が未承認の場合も" & vbLf
    finalMessage = finalMessage & "  勤怠入力漏れと判定されます。" & vbLf
    finalMessage = finalMessage & "  承認漏れが無いかも確認してください。"
    
    Debug.Print "統合メッセージ生成完了"
    Debug.Print "メッセージ長: " & Len(finalMessage) & "文字"
    Debug.Print "========================================="
    
    GenerateMessageFromSheet = finalMessage
    Exit Function
    
ErrorHandler:
    Debug.Print "統合メッセージ生成エラー: " & Err.Description
    Debug.Print "========================================="
    MsgBox "メッセージ生成エラー: " & vbCrLf & vbCrLf & Err.Description, _
           vbCritical, "エラー"
    GenerateMessageFromSheet = ""
End Function


' *************************************************************
' 関数名: GenerateAttendanceMissingPart
' 目的: 勤怠入力漏れ部分のメッセージ生成
' 戻り値: フォーマット済みメッセージ(ASCII文字のみ)
' *************************************************************
Private Function GenerateAttendanceMissingPart() As String
    On Error GoTo ErrorHandler
    
    Debug.Print "[INFO] 勤怠入力漏れ部分の生成開始"
    
    ' 「勤怠入力漏れ一覧」シートを取得
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("勤怠入力漏れ一覧")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Debug.Print "[INFO] 勤怠入力漏れ一覧シートが見つかりません"
        GenerateAttendanceMissingPart = ""
        Exit Function
    End If
    
    ' データ行数確認(ヘッダー除く)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= 1 Then
        Debug.Print "[INFO] 勤怠入力漏れデータなし"
        GenerateAttendanceMissingPart = ""
        Exit Function
    End If
    
    ' Dictionary作成(社員ごとに未入力日をグループ化)
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
        
        ' 社員ごとに未入力日を集約
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
    
    ' メッセージ生成(緊急度判定付き)
    Dim message As String
    message = "未入力者: " & empDict.Count & "名 / 未入力日数: " & totalMissingCount & "日" & vbLf & vbLf
    
    Dim key As Variant
    For Each key In empDict.Keys
        empData = empDict(key)
        
        Dim missingCount As Integer
        missingCount = empData(1).Count
        
        ' 緊急度判定(絵文字なし)
        Dim urgencyMark As String
        If missingCount >= 5 Then
            urgencyMark = "[!!緊急!!]"
        ElseIf missingCount >= 3 Then
            urgencyMark = "[!要注意!]"
        Else
            urgencyMark = "[確認]"
        End If
        
        Dim empText As String
        empText = urgencyMark & " " & empData(0) & " さん (" & missingCount & "日)" & vbLf
        
        ' 日付リスト(最大5件まで)
        Dim dateItem As Variant
        Dim dateCount As Integer
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
    Next key
    
    GenerateAttendanceMissingPart = message
    Exit Function
    
ErrorHandler:
    Debug.Print "[ERROR] 勤怠入力漏れ部分生成エラー: " & Err.Description
    GenerateAttendanceMissingPart = ""
End Function


' *************************************************************
' 関数名: GenerateBreakTimeViolationPart
' 目的: 休憩時間違反部分のメッセージ生成
' 戻り値: フォーマット済みメッセージ(ASCII文字のみ)
' *************************************************************
Private Function GenerateBreakTimeViolationPart() As String
    On Error GoTo ErrorHandler
    
    Debug.Print "[INFO] 休憩時間違反部分の生成開始"
    
    ' 「休憩時間違反一覧」シートを取得
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("休憩時間違反一覧")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Debug.Print "[INFO] 休憩時間違反一覧シートが見つかりません"
        GenerateBreakTimeViolationPart = ""
        Exit Function
    End If
    
    ' データ行数確認(ヘッダー除く)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow <= 1 Then
        Debug.Print "[INFO] 休憩時間違反データなし"
        GenerateBreakTimeViolationPart = ""
        Exit Function
    End If
    
    ' Dictionary作成(社員ごとに違反日をグループ化)
    Dim empDict As Object
    Set empDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim empID As String, empName As String, targetDate As Date
    Dim workTime As String, breakTime As String, shortage As String
    Dim totalViolationCount As Long
    totalViolationCount = 0
    
    For i = 2 To lastRow
        empID = Trim(ws.Cells(i, 1).Value)
        empName = Trim(ws.Cells(i, 2).Value)
        
        On Error Resume Next
        targetDate = CDate(ws.Cells(i, 3).Value)
        If Err.Number <> 0 Then
            Debug.Print "[WARNING] 行" & i & "の日付が不正"
            GoTo NextViolation
        End If
        On Error GoTo ErrorHandler
        
        workTime = Trim(ws.Cells(i, 5).Value)   ' E列: 実働時間
        breakTime = Trim(ws.Cells(i, 6).Value)  ' F列: 休憩時間
        shortage = Trim(ws.Cells(i, 8).Value)   ' H列: 休憩不足時間
        
        totalViolationCount = totalViolationCount + 1
        
        ' 社員ごとに集約
        If Not empDict.Exists(empID) Then
            Dim newColl As Collection
            Set newColl = New Collection
            empDict.Add empID, Array(empName, newColl)
        End If
        
        Dim empData As Variant
        empData = empDict(empID)
        empData(1).Add Array(targetDate, workTime, breakTime, shortage)
        empDict(empID) = empData
        
NextViolation:
    Next i
    
    If empDict.Count = 0 Then
        Debug.Print "[INFO] 休憩時間違反データなし(解析後)"
        GenerateBreakTimeViolationPart = ""
        Exit Function
    End If
    
    Debug.Print "[INFO] 休憩時間違反: " & empDict.Count & "名, " & totalViolationCount & "件"
    
    ' メッセージ生成(絵文字なし)
    Dim message As String
    message = "違反者: " & empDict.Count & "名 / 違反件数: " & totalViolationCount & "件" & vbLf & vbLf
    
    Dim key As Variant
    For Each key In empDict.Keys
        empData = empDict(key)
        
        Dim empText As String
        empText = "[違反] " & empData(0) & " さん" & vbLf
        
        ' 違反詳細のリスト(最大5件まで)
        Dim violationItem As Variant
        Dim violationCount As Integer
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
' 目的: メイン処理 - メッセージ生成→送信
' 使用方法: ボタンから呼び出される公開関数
' *************************************************************
Public Sub SendNotificationToLineWorks()
    On Error GoTo ErrorHandler
    
    Debug.Print vbCrLf & "####################################"
    Debug.Print "# LINE WORKS通知処理開始"
    Debug.Print "####################################"
    
    Application.ScreenUpdating = False
    Application.StatusBar = "LINE WORKS通知を準備中..."
    
    ' 設定シートが存在するか確認
    Dim configSheet As Worksheet
    On Error Resume Next
    Set configSheet = ThisWorkbook.Sheets("設定")
    On Error GoTo ErrorHandler
    
    If configSheet Is Nothing Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "設定シートが見つかりません。" & vbCrLf & vbCrLf & _
               "初期セットアップを実行してください。", _
               vbExclamation, "設定エラー"
        Exit Sub
    End If
    
    ' 1. Webhook URL取得(GetWebhookURLがModule1に存在しない場合のフォールバック)
    Dim webhookUrl As String
    On Error Resume Next
    webhookUrl = GetWebhookURL()
    If Err.Number <> 0 Then
        ' GetWebhookURL関数がない場合は設定シートから直接取得
        Err.Clear
        On Error GoTo ErrorHandler
        webhookUrl = Trim(configSheet.Cells(1, 2).Value)
    End If
    On Error GoTo ErrorHandler
    
    If webhookUrl = "" Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "Webhook URLが設定されていません。" & vbCrLf & _
               "[設定]シートのB1セルにWebhook URLを入力してください。", _
               vbExclamation, "設定エラー"
        Exit Sub
    End If
    
    ' 2. メッセージ生成
    Application.StatusBar = "メッセージを生成中..."
    Dim message As String
    message = GenerateMessageFromSheet()
    
    If message = "" Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Exit Sub
    End If
    
    ' 3. 確認ダイアログ(プレビュー付き)
    Dim previewMsg As String
    previewMsg = Left(message, 300)
    If Len(message) > 300 Then
        previewMsg = previewMsg & vbLf & "..." & vbLf & "(以下省略)"
    End If
    
    Dim response As VbMsgBoxResult
    response = MsgBox("SI1部リーダーチャンネルに通知を送信しますか？" & vbCrLf & vbCrLf & _
                      "【プレビュー】" & vbCrLf & _
                      "=========================================" & vbCrLf & _
                      previewMsg & vbCrLf & _
                      "=========================================", _
                      vbQuestion + vbYesNo, "送信確認")
    
    If response <> vbYes Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "送信をキャンセルしました。", vbInformation, "キャンセル"
        Exit Sub
    End If
    
    ' 4. 送信
    Application.StatusBar = "LINE WORKSに送信中..."
    Dim result As Boolean
    result = SendToLineWorks(webhookUrl, message)
    
    ' 5. 完了メッセージ
    If result Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "LINE WORKSへの通知を送信しました。", vbInformation, "送信完了"
    Else
        Application.ScreenUpdating = True
        Application.StatusBar = False
        ' エラーは SendToLineWorks 内で表示済み
    End If
    
    Debug.Print "####################################"
    Debug.Print "# LINE WORKS通知処理終了"
    Debug.Print "####################################" & vbCrLf
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Debug.Print "[ERROR] メイン処理エラー: " & Err.Description
    MsgBox "通知処理でエラーが発生しました: " & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "エラー"
End Sub


' *************************************************************
' テスト用プロシージャ(絵文字なし)
' *************************************************************

Sub Test_SendToLineWorks_NoEmoji()
    Dim url As String
    On Error Resume Next
    url = GetWebhookURL()
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Dim configSheet As Worksheet
        Set configSheet = ThisWorkbook.Sheets("設定")
        url = Trim(configSheet.Cells(1, 2).Value)
    End If
    On Error GoTo 0
    
    If url = "" Then
        MsgBox "Webhook URLが設定されていません", vbExclamation
        Exit Sub
    End If
    
    Dim testMessage As String
    testMessage = "【テスト】絵文字なしメッセージ" & vbLf & _
                  "-----------------------------" & vbLf & _
                  "[!!緊急!!] 緊急レベル" & vbLf & _
                  "[!要注意!] 注意レベル" & vbLf & _
                  "[確認] 確認レベル" & vbLf & _
                  "[違反] 違反表示" & vbLf & _
                  "-----------------------------" & vbLf & _
                  "送信日時: " & Now
    
    Dim result As Boolean
    result = SendToLineWorks(url, testMessage)
    
    If result Then
        MsgBox "テスト送信成功！LINE WORKSで確認してください。", vbInformation
    End If
End Sub

Sub Test_GenerateMessage()
    Dim message As String
    message = GenerateMessageFromSheet()
    
    If message <> "" Then
        Debug.Print "========================================="
        Debug.Print "生成メッセージ:"
        Debug.Print message
        Debug.Print "========================================="
        MsgBox "メッセージ生成成功" & vbCrLf & _
               "イミディエイトウィンドウで確認してください", vbInformation
    End If
End Sub

