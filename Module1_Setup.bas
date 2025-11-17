Attribute VB_Name = "Module1_Setup"
' *************************************************************
' Module1_Setup
' 目的：設定管理と初期セットアップ
' 作成日：2025-10-18
' 改版履歴：
' 2025/10/18 初版作成 - LINE WORKS通知機能用の設定管理
' *************************************************************

Option Explicit

' *************************************************************
' 関数名: CreateConfigSheet
' 目的: 設定シートを作成してWebhook URLを格納
' 使用方法: VBエディタでF5実行、または他のプロシージャから呼び出し
' *************************************************************
Public Sub CreateConfigSheet()
    On Error GoTo ErrorHandler
    
    ' シート存在チェック
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("設定")
    On Error GoTo ErrorHandler
    
    If Not ws Is Nothing Then
        Debug.Print "設定シートは既に存在します"
        MsgBox "設定シートは既に存在します。", vbInformation, "確認"
        Exit Sub
    End If
    
    ' 画面更新停止
    Application.ScreenUpdating = False
    
    ' 新規シート作成
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.Name = "設定"
    
    ' 設定項目の書き込み
    With ws
        ' Webhook URL設定
        .Range("A1").Value = "WEBHOOK_URL"
        .Range("B1").Value = "" ' 後で手動入力
        .Range("A1").Font.Bold = True
        .Range("A1:A1").HorizontalAlignment = xlRight
        
        ' その他の設定項目
        .Range("A2").Value = "BOT_NAME"
        .Range("B2").Value = "勤怠リマインダー"
        .Range("A2").Font.Bold = True
        .Range("A2:A2").HorizontalAlignment = xlRight
        
        .Range("A3").Value = "VERSION"
        .Range("B3").Value = "2.1.0"
        .Range("A3").Font.Bold = True
        .Range("A3:A3").HorizontalAlignment = xlRight
        
        .Range("A4").Value = "LAST_UPDATE"
        .Range("B4").Value = Now
        .Range("A4").Font.Bold = True
        .Range("A4:A4").HorizontalAlignment = xlRight
        
        ' 列幅調整
        .Columns("A:B").AutoFit
        .Columns("B:B").ColumnWidth = 50 ' Webhook URL用に広めに
        
        ' 見出しに色を付ける
        .Range("A1:A4").Interior.Color = RGB(220, 230, 241)
        
        ' 非表示化
        .Visible = xlSheetVeryHidden
    End With
    
    ' 画面更新再開
    Application.ScreenUpdating = True
    
    Debug.Print "設定シートを作成しました: " & Now
    
    ' 完了メッセージ
    MsgBox "設定シートを作成しました。" & vbCrLf & vbCrLf & _
           "次のステップ:" & vbCrLf & _
           "1. この後、設定シートを再表示します" & vbCrLf & _
           "2. B1セルにWebhook URLを貼り付けてください" & vbCrLf & _
           "3. 保存して、再度非表示にします", _
           vbInformation, "設定シート作成完了"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "設定シートの作成に失敗しました: " & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "エラー"
End Sub


' *************************************************************
' 関数名: GetWebhookURL
' 目的: 設定シートからWebhook URLを取得
' 戻り値: Webhook URL（文字列）、未設定の場合は空文字
' 使用例:
'   Dim url As String
'   url = GetWebhookURL()
'   If url <> "" Then
'       ' URLが設定されている場合の処理
'   End If
' *************************************************************
Public Function GetWebhookURL() As String
    On Error GoTo ErrorHandler
    
    ' 設定シート取得
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("設定")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "設定シートが見つかりません。" & vbCrLf & _
               "CreateConfigSheetを実行してください。", _
               vbExclamation, "設定エラー"
        GetWebhookURL = ""
        Exit Function
    End If
    
    ' Webhook URL取得（B1セル）
    GetWebhookURL = Trim(ws.Range("B1").Value)
    
    ' 未設定チェック
    If GetWebhookURL = "" Then
        MsgBox "Webhook URLが設定されていません。" & vbCrLf & vbCrLf & _
               "設定方法:" & vbCrLf & _
               "1. VBAProject で [設定] シートを右クリック" & vbCrLf & _
               "2. [プロパティ] で Visible を [xlSheetVisible] に変更" & vbCrLf & _
               "3. Excelに戻り、[設定]シートのB1セルにWebhook URLを入力" & vbCrLf & _
               "4. 保存後、再度非表示に設定", _
               vbExclamation, "Webhook URL未設定"
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Webhook URLの取得に失敗しました: " & Err.Description, vbCritical, "エラー"
    GetWebhookURL = ""
End Function


' *************************************************************
' 関数名: CreateNotificationHistorySheet
' 目的: 通知履歴シートを作成
' 使用方法: VBエディタでF5実行、または他のプロシージャから呼び出し
' *************************************************************
Public Sub CreateNotificationHistorySheet()
    On Error GoTo ErrorHandler
    
    ' シート存在チェック
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("通知履歴")
    On Error GoTo ErrorHandler
    
    If Not ws Is Nothing Then
        Debug.Print "通知履歴シートは既に存在します"
        MsgBox "通知履歴シートは既に存在します。", vbInformation, "確認"
        Exit Sub
    End If
    
    ' 画面更新停止
    Application.ScreenUpdating = False
    
    ' 新規シート作成
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.Name = "通知履歴"
    
    ' ヘッダー設定
    With ws
        .Range("A1").Value = "送信日時"
        .Range("B1").Value = "対象者数"
        .Range("C1").Value = "未入力件数"
        .Range("D1").Value = "結果"
        .Range("E1").Value = "エラー詳細"
        
        ' 書式設定
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(217, 217, 217)
        .Range("A1:E1").HorizontalAlignment = xlCenter
        
        ' 列幅調整
        .Columns("A:A").ColumnWidth = 20 ' 送信日時
        .Columns("B:B").ColumnWidth = 12 ' 対象者数
        .Columns("C:C").ColumnWidth = 12 ' 未入力件数
        .Columns("D:D").ColumnWidth = 10 ' 結果
        .Columns("E:E").ColumnWidth = 40 ' エラー詳細
        
        ' 枠線
        .Range("A1:E1").Borders.LineStyle = xlContinuous
    End With
    
    ' 画面更新再開
    Application.ScreenUpdating = True
    
    Debug.Print "通知履歴シートを作成しました: " & Now
    
    MsgBox "通知履歴シートを作成しました。", vbInformation, "作成完了"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "通知履歴シートの作成に失敗しました: " & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "エラー"
End Sub


' *************************************************************
' 関数名: LogNotification
' 目的: 通知履歴シートに送信結果を記録
' 引数:
'   - targetCount: 対象者数
'   - totalCount: 未入力件数
'   - result: 結果（"成功" or "失敗"）
'   - errorDetail: エラー詳細（オプション）
' 使用例:
'   Call LogNotification(3, 5, "成功", "")
'   Call LogNotification(0, 0, "失敗", "HTTP 404 Not Found")
' *************************************************************
Public Sub LogNotification(targetCount As Integer, totalCount As Integer, _
                          result As String, Optional errorDetail As String = "")
    On Error GoTo ErrorHandler
    
    ' 通知履歴シート取得
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("通知履歴")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Debug.Print "通知履歴シートが見つかりません"
        Exit Sub
    End If
    
    ' 最終行取得（データを下に追加）
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    
    ' データ書き込み
    With ws
        .Cells(lastRow, 1).Value = Now ' 送信日時
        .Cells(lastRow, 2).Value = targetCount ' 対象者数
        .Cells(lastRow, 3).Value = totalCount ' 未入力件数
        .Cells(lastRow, 4).Value = result ' 結果
        .Cells(lastRow, 5).Value = errorDetail ' エラー詳細
        
        ' 結果に応じて色付け
        If result = "成功" Then
            .Cells(lastRow, 4).Interior.Color = RGB(198, 224, 180) ' 緑
        Else
            .Cells(lastRow, 4).Interior.Color = RGB(255, 199, 206) ' 赤
        End If
    End With
    
    Debug.Print "通知履歴を記録しました: " & Now & " | 結果: " & result
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "通知履歴の記録に失敗しました: " & Err.Description
End Sub


' *************************************************************
' 関数名: InitialSetup
' 目的: 初期セットアップをまとめて実行
' 使用方法: VBエディタでこの関数を開いてF5実行
' *************************************************************
Public Sub InitialSetup()
    On Error GoTo ErrorHandler
    
    MsgBox "LINE WORKS通知機能の初期セットアップを開始します。", vbInformation, "初期セットアップ"
    
    ' 設定シート作成
    Call CreateConfigSheet
    
    ' 通知履歴シート作成
    Call CreateNotificationHistorySheet
    
    MsgBox "初期セットアップが完了しました！" & vbCrLf & vbCrLf & _
           "次のステップ:" & vbCrLf & _
           "1. VBAエディタのプロジェクトエクスプローラーで" & vbCrLf & _
           "   [設定] シートを見つけてください" & vbCrLf & _
           "2. 右クリック → [プロパティ]" & vbCrLf & _
           "3. Visible を [0 - xlSheetVisible] に変更" & vbCrLf & _
           "4. Excelに戻り、B1セルにWebhook URLを入力" & vbCrLf & _
           "5. 保存後、Visibleを [2 - xlSheetVeryHidden] に戻す", _
           vbInformation, "セットアップ完了"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "初期セットアップでエラーが発生しました: " & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' *************************************************************
' テスト用プロシージャ
' *************************************************************

' テスト1: Webhook URL取得テスト
Sub Test_GetWebhookURL()
    Dim url As String
    url = GetWebhookURL()
    
    If url <> "" Then
        MsgBox "? テスト合格" & vbCrLf & vbCrLf & _
               "取得したURL: " & vbCrLf & url, _
               vbInformation, "Webhook URLテスト"
    Else
        MsgBox "? テスト失敗: URLが取得できませんでした", vbCritical, "テスト結果"
    End If
End Sub

' テスト2: 通知履歴記録テスト
Sub Test_LogNotification()
    ' テストデータを記録
    Call LogNotification(3, 5, "成功", "")
    Call LogNotification(0, 0, "失敗", "HTTP 404 Not Found")
    
    MsgBox "? テストデータを通知履歴シートに記録しました。" & vbCrLf & _
           "[通知履歴] シートを確認してください。", _
           vbInformation, "履歴記録テスト"
End Sub

' *************************************************************
' 関数名: GetChannelID
' 目的: 設定シートからChannel IDを取得
' 戻り値: Channel ID（文字列）、未設定の場合は空文字
' *************************************************************
Public Function GetChannelID() As String
    On Error GoTo ErrorHandler
    
    ' 設定シート取得
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("設定")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "設定シートが見つかりません。", vbExclamation, "設定エラー"
        GetChannelID = ""
        Exit Function
    End If
    
    ' Channel ID取得（B5セル）
    GetChannelID = Trim(ws.Range("B5").Value)
    
    ' 未設定チェック
    If GetChannelID = "" Then
        MsgBox "Channel IDが設定されていません。" & vbCrLf & _
               "[設定]シートのB5セルにChannel IDを入力してください。", _
               vbExclamation, "Channel ID未設定"
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Channel IDの取得に失敗しました: " & Err.Description, vbCritical, "エラー"
    GetChannelID = ""
End Function

