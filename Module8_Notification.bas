' ========================================
' Module8_Notification
' タイプ: 標準モジュール
' 行数: 551
' エクスポート日時: 2025-10-18 22:41:04
' ========================================

Option Explicit

' *************************************************************
' Module8_Notification
' 目的：LINE WORKS Webhook通知機能
' 作成日：2025-10-18
' 改版履歴：
' 2025/10/18 初版作成 - Webhook通知機能追加
' 2025/10/18 通知履歴機能を削除（LINE WORKS上で確認可能なため不要）
' *************************************************************

' *************************************************************
' 関数名: SendToLineWorks
' 目的: LINE WORKS WebhookにPOSTリクエストを送信
' 引数:
'   - webhookURL: Webhook URL
'   - messageText: 送信するメッセージ本文
' 戻り値:
'   - True: 送信成功（HTTP 200）
'   - False: 送信失敗
' 備考: LINE WORKS Webhook v2形式（channelId + body.text）
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
    
    ' Channel ID取得
    Dim channelID As String
    channelID = GetChannelID() ' Module1_Setup.bas の関数
    
    If channelID = "" Then
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
    escapedText = Replace(escapedText, vbLf, "\n")     ' 改行（LF）
    escapedText = Replace(escapedText, vbCr, "")       ' 改行（CR）を削除
    escapedText = Replace(escapedText, vbTab, " ")     ' タブを空白に
    
    ' JSON構築（LINE WORKS Webhook v2: channelId + body.text形式）
    Dim jsonBody As String
    jsonBody = "{""channelId"":""" & channelID & """,""body"":{""text"":""" & escapedText & """}}"
    
    Debug.Print "JSON: " & jsonBody
    Debug.Print "JSON長: " & Len(jsonBody) & "文字"
    
    ' HTTP通信オブジェクト作成
    Dim httpRequest As Object
    
    On Error Resume Next
    Set httpRequest = CreateObject("MSXML2.XMLHTTP.6.0")
    If httpRequest Is Nothing Then
        Set httpRequest = CreateObject("MSXML2.XMLHTTP.3.0")
    End If
    If httpRequest Is Nothing Then
        Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    End If
    If httpRequest Is Nothing Then
        Set httpRequest = CreateObject("Microsoft.XMLHTTP")
    End If
    On Error GoTo ErrorHandler
    
    If httpRequest Is Nothing Then
        MsgBox "HTTP通信オブジェクトを作成できませんでした。", vbCritical, "エラー"
        SendToLineWorks = False
        Exit Function
    End If
    
    Debug.Print "使用するHTTPオブジェクト: " & TypeName(httpRequest)
    
    ' HTTP POST送信
    httpRequest.Open "POST", webhookUrl, False
    httpRequest.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
    
    ' タイムアウト設定
    On Error Resume Next
    httpRequest.setTimeouts 30000, 30000, 30000, 30000
    On Error GoTo ErrorHandler
    
    httpRequest.send jsonBody
    
    ' レスポンス確認
    Debug.Print "HTTP Status: " & httpRequest.Status
    Debug.Print "レスポンス: " & httpRequest.responseText
    
    If httpRequest.Status = 200 Then
        Debug.Print "? 送信成功"
        Debug.Print "========================================="
        SendToLineWorks = True
    Else
        Debug.Print "? 送信失敗"
        Debug.Print "========================================="
        
        ' エラー詳細をユーザーに表示
        Dim errorMsg As String
        Select Case httpRequest.Status
            Case 400
                errorMsg = "パラメータエラー（HTTP 400）" & vbCrLf & _
                          "レスポンス: " & httpRequest.responseText & vbCrLf & vbCrLf & _
                          "Channel IDとWebhook URLを確認してください。"
            Case 401
                errorMsg = "認証エラー（HTTP 401）" & vbCrLf & _
                          "Webhook URLが無効です。再発行してください。"
            Case 404
                errorMsg = "URLが見つかりません（HTTP 404）" & vbCrLf & _
                          "Webhook URLまたはChannel IDを確認してください。"
            Case 429
                errorMsg = "レート制限超過（HTTP 429）" & vbCrLf & _
                          "5分待ってから再試行してください。"
            Case 500, 502, 503
                errorMsg = "サーバーエラー（HTTP " & httpRequest.Status & "）" & vbCrLf & _
                          "時間を置いて再試行してください。"
            Case Else
                errorMsg = "エラー（HTTP " & httpRequest.Status & "）" & vbCrLf & _
                          httpRequest.responseText
        End Select
        
        MsgBox "通知送信に失敗しました。" & vbCrLf & vbCrLf & errorMsg, _
               vbCritical, "送信エラー"
        SendToLineWorks = False
    End If
    
    Set httpRequest = Nothing
    Exit Function
    
ErrorHandler:
    Debug.Print "? 例外エラー: " & Err.Description
    Debug.Print "エラー番号: " & Err.Number
    Debug.Print "========================================="
    MsgBox "通知送信中にエラーが発生しました: " & vbCrLf & vbCrLf & _
           "エラー: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "エラー"
    SendToLineWorks = False
End Function


' *************************************************************
' 関数名: GenerateMessageFromSheet
' 目的: 「勤怠入力漏れ一覧」シートのデータを読み取ってメッセージ生成
' 戻り値: フォーマット済みメッセージ（文字列）
' 備考:
'   - 社員ごとにグループ化
'   - 緊急度別に分類（??5日以上、??3-4日、??1-2日）
' *************************************************************
Public Function GenerateMessageFromSheet() As String
    On Error GoTo ErrorHandler
    
    Debug.Print "========================================="
    Debug.Print "メッセージ生成開始: " & Now
    
    ' 「勤怠入力漏れ一覧」シートを取得
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("勤怠入力漏れ一覧")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "「勤怠入力漏れ一覧」シートが見つかりません。" & vbCrLf & _
               "先に勤怠チェックを実行してください。", _
               vbExclamation, "シートエラー"
        GenerateMessageFromSheet = ""
        Exit Function
    End If
    
    ' データ行数確認（ヘッダー除く）
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Debug.Print "データ行数: " & (lastRow - 1) & "行"
    
    If lastRow <= 1 Then
        MsgBox "未入力データがありません。", vbInformation, "データなし"
        GenerateMessageFromSheet = ""
        Exit Function
    End If
    
    ' データを読み取って社員ごとにグループ化
    Dim empDict As Object
    Set empDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim empID As String, empName As String, targetDate As Date, daysAgo As Long
    Dim comment As String
    Dim totalMissingCount As Long
    totalMissingCount = 0
    
    For i = 2 To lastRow ' 2行目からデータ開始（1行目はヘッダー）
        empID = Trim(ws.Cells(i, 1).Value)    ' A列: 社員番号
        empName = Trim(ws.Cells(i, 2).Value)  ' B列: 氏名
        
        ' C列が日付型かチェック
        On Error Resume Next
        targetDate = ws.Cells(i, 3).Value     ' C列: 日付
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo ErrorHandler
            GoTo NextRow ' 日付が不正な場合はスキップ
        End If
        On Error GoTo ErrorHandler
        
        comment = Trim(ws.Cells(i, 7).Value)  ' G列: コメント
        
        ' 未入力判定（コメントに「入力されていません」が含まれる）
        If InStr(comment, "入力されていません") > 0 Then
            daysAgo = DateDiff("d", targetDate, Date)
            totalMissingCount = totalMissingCount + 1
            
            ' 社員ごとに集約
            If Not empDict.Exists(empID) Then
                ' 構造: Array(氏名, Collection(日付配列), 最大日数)
                Dim newColl As Collection
                Set newColl = New Collection
                empDict.Add empID, Array(empName, newColl, 0)
            End If
            
            ' 未入力日を追加
            Dim empData As Variant
            empData = empDict(empID)
            empData(1).Add Array(targetDate, daysAgo)
            
            ' 最大日数を更新
            If daysAgo > empData(2) Then
                empData(2) = daysAgo
                empDict(empID) = empData
            End If
        End If
        
NextRow:
    Next i
    
    Debug.Print "抽出した社員数: " & empDict.Count & "名"
    Debug.Print "未入力件数: " & totalMissingCount & "件"
    
    ' データが0件の場合
    If empDict.Count = 0 Then
        MsgBox "未入力データがありません。", vbInformation, "データなし"
        GenerateMessageFromSheet = ""
        Exit Function
    End If
    
    ' メッセージ生成
    Dim message As String
    message = "【勤怠未入力アラート】" & Format(Date, "yyyy/mm/dd") & vbLf & vbLf
    message = message & "未入力者: " & empDict.Count & "名 / 未入力件数: " & totalMissingCount & "件" & vbLf & vbLf
    
    ' 緊急度別に分類
    Dim urgentList As String, warningList As String, normalList As String
    urgentList = ""
    warningList = ""
    normalList = ""
    
    Dim key As Variant
    For Each key In empDict.Keys
        empData = empDict(key)
        Dim maxDays As Long
        maxDays = empData(2)
        
        Dim empText As String
        empText = empData(0) & " さん" & vbLf
        
        ' 未入力日のリスト（最大5件まで表示）
        Dim dateItem As Variant
        Dim dateCount As Integer
        dateCount = 0
        For Each dateItem In empData(1)
            If dateCount < 5 Then
                empText = empText & "  ・" & Format(dateItem(0), "mm/dd") & _
                          "（" & dateItem(1) & "日前）" & vbLf
                dateCount = dateCount + 1
            End If
        Next dateItem
        
        ' 6件以上ある場合は省略表示
        If empData(1).Count > 5 Then
            empText = empText & "  ...他" & (empData(1).Count - 5) & "件" & vbLf
        End If
        
        empText = empText & vbLf
        
        ' 緊急度判定（絵文字なし）
        If maxDays >= 5 Then
            urgentList = urgentList & "[緊急] " & empText
        ElseIf maxDays >= 3 Then
            warningList = warningList & "[要注意] " & empText
        Else
            normalList = normalList & "[確認] " & empText
        End If
    Next key
    
    ' 緊急度順に追加
    If urgentList <> "" Then
        message = message & "■ 緊急対応（5日以上）" & vbLf & urgentList
    End If
    
    If warningList <> "" Then
        message = message & "■ 要注意（3-4日）" & vbLf & warningList
    End If
    
    If normalList <> "" Then
        message = message & "■ 確認（1-2日）" & vbLf & normalList
    End If
    
    message = message & "━━━━━━━━━━━━━━━" & vbLf
    message = message & "※各リーダーより該当者へ声掛けをお願いします" & vbLf
    message = message & "※申請決裁が未承認の場合も勤怠入力漏れと判定されます。" & vbLf
    message = message & "　承認漏れが無いかも確認してください。"
    
    Debug.Print "生成したメッセージ長: " & Len(message) & "文字"
    Debug.Print "========================================="
    
    GenerateMessageFromSheet = message
    Exit Function
    
ErrorHandler:
    Debug.Print "? メッセージ生成エラー: " & Err.Description
    Debug.Print "========================================="
    MsgBox "メッセージ生成エラー: " & vbCrLf & vbCrLf & Err.Description, _
           vbCritical, "エラー"
    GenerateMessageFromSheet = ""
End Function


' *************************************************************
' 関数名: SendNotificationToLineWorks
' 目的: メイン処理 - メッセージ生成→送信
' 使用方法: ボタンから呼び出される公開関数
' 備考: これが「?? LINE WORKS通知」ボタンに紐付けられる関数
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
               "初期セットアップを実行してください:" & vbCrLf & _
               "Module1_Setup の InitialSetup を実行", _
               vbExclamation, "設定エラー"
        Exit Sub
    End If
    
    ' 1. Webhook URL取得
    Dim webhookUrl As String
    webhookUrl = GetWebhookURL() ' Module1_Setup.bas の関数
    
    If webhookUrl = "" Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
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
    
    ' 3. 確認ダイアログ（プレビュー付き）
    Dim previewMsg As String
    previewMsg = Left(message, 300)
    If Len(message) > 300 Then
        previewMsg = previewMsg & vbLf & "..." & vbLf & "(以下省略)"
    End If
    
    Dim response As VbMsgBoxResult
    response = MsgBox("SI1部リーダーチャンネルに通知を送信しますか？" & vbCrLf & vbCrLf & _
                      "【プレビュー】" & vbCrLf & _
                      "━━━━━━━━━━━━━━━" & vbCrLf & _
                      previewMsg & vbCrLf & _
                      "━━━━━━━━━━━━━━━", _
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
    
    ' 5. 対象者数・未入力件数を集計（表示用）
    Dim targetCount As Long, totalCount As Long
    targetCount = 0
    totalCount = 0
    
    ' メッセージから抽出
    Dim lines As Variant
    lines = Split(message, vbLf)
    Dim line As Variant
    For Each line In lines
        If InStr(line, "未入力者:") > 0 Then
            ' 例: "未入力者: 3名 / 未入力件数: 5件"
            Dim parts As Variant
            parts = Split(line, " ")
            Dim j As Integer
            For j = 0 To UBound(parts)
                If InStr(parts(j), "名") > 0 Then
                    targetCount = Val(Replace(parts(j), "名", ""))
                ElseIf InStr(parts(j), "件") > 0 Then
                    totalCount = Val(Replace(parts(j), "件", ""))
                End If
            Next j
            Exit For
        End If
    Next line
    
    ' 6. 完了メッセージ
    If result Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "LINE WORKSへの通知を送信しました。" & vbCrLf & vbCrLf & _
               "対象者: " & targetCount & "名" & vbCrLf & _
               "未入力件数: " & totalCount & "件", _
               vbInformation, "送信完了"
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
    Debug.Print "? メイン処理エラー: " & Err.Description
    MsgBox "通知処理でエラーが発生しました: " & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "エラー"
End Sub


' *************************************************************
' テスト用プロシージャ
' *************************************************************

' テスト1: Webhook送信テスト
Sub Test_SendToLineWorks()
    Dim url As String
    url = GetWebhookURL()
    
    If url = "" Then Exit Sub
    
    Dim result As Boolean
    result = SendToLineWorks(url, "【テスト】システムテストメッセージ" & vbLf & "送信日時: " & Now)
    
    If result Then
        MsgBox "? Webhook送信テスト合格" & vbCrLf & _
               "LINE WORKSチャンネルでメッセージを確認してください。", _
               vbInformation, "テスト結果"
    End If
End Sub

' テスト2: メッセージ生成テスト
Sub Test_GenerateMessage()
    Dim message As String
    message = GenerateMessageFromSheet()
    
    If message <> "" Then
        MsgBox "? メッセージ生成テスト合格" & vbCrLf & vbCrLf & _
               "【生成されたメッセージ】" & vbCrLf & _
               "━━━━━━━━━━━━━━━" & vbCrLf & _
               Left(message, 500) & vbCrLf & _
               "━━━━━━━━━━━━━━━" & vbCrLf & vbCrLf & _
               "※イミディエイトウィンドウ（Ctrl+G）で全文を確認できます", _
               vbInformation, "テスト結果"
        Debug.Print vbCrLf & "【生成メッセージ全文】"
        Debug.Print message
    End If
End Sub

' テスト3: メイン処理を直接実行
Sub Test_SendNotificationToLineWorks()
    ' メイン処理を実行
    Call SendNotificationToLineWorks
End Sub

' テスト4: イミディエイトウィンドウにメッセージ全文を出力
Sub Test_ShowFullMessage()
    Dim message As String
    message = GenerateMessageFromSheet()
    
    If message <> "" Then
        Debug.Print "========================================="
        Debug.Print "生成されたメッセージ全文:"
        Debug.Print "========================================="
        Debug.Print message
        Debug.Print "========================================="
        Debug.Print "文字数: " & Len(message)
        Debug.Print "========================================="
        
        MsgBox "? メッセージ生成成功" & vbCrLf & vbCrLf & _
               "イミディエイトウィンドウ（Ctrl+G）で全文を確認できます。" & vbCrLf & vbCrLf & _
               "文字数: " & Len(message) & "文字", _
               vbInformation, "テスト結果"
    Else
        MsgBox "? メッセージ生成失敗", vbCritical, "テスト結果"
    End If
End Sub

