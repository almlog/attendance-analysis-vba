' ========================================
' Module2
' タイプ: 標準モジュール
' 行数: 136
' エクスポート日時: 2025-10-18 22:41:04
' ========================================

Option Explicit

' *************************************************************
' モジュール：休憩時間チェック（コア機能）
' 目的：実働時間に応じた休憩時間の取得確認および残業時間を計算する
' Copyright (c) 2025 SI1 shunpei.suzuki
' 作成日：2025年3月3日
'
' 改版履歴：
' 2025/03/03 初版作成_v1.5
' 2025/03/07 届出申請時の備考欄確認を追加①
' 2025/03/11 届出申請時の備考欄確認を追加②_v1.7
' 2025/03/21 除外社員機能・社員数カウント修正・パフォーマンス最適化_v2.0
' 2025/03/31 月末最終営業日の当日分チェック機能追加
' 2025/04/02 モジュール分割によるリファクタリング
' 2025/04/02 勤怠申請と入力内容の矛盾検知機能追加（午前有休・午後有休）
' 2025/10/18 LINE WORKS通知機能案内をメッセージに追加
' *************************************************************

' 定数定義
Const SHEET_NAME_MISSING_ENTRIES As String = "勤怠入力漏れ一覧"
Const COL_EMPLOYEE_ID As Integer = 1
Const COL_EMPLOYEE_NAME As Integer = 2
Const COL_DATE As Integer = 3
Const COL_DAY_TYPE As Integer = 4
Const COL_LEAVE_TYPE As Integer = 5
Const COL_MISSING_ENTRY_TYPE As Integer = 6
Const COL_COMMENT As Integer = 7
Const COL_ATTENDANCE_TIME As Integer = 8 ' 出勤時刻列を追加
Const COL_DEPARTURE_TIME As Integer = 9 ' 退勤時刻列を追加
Const COL_CONTRADICTION_TYPE As Integer = 10 ' 矛盾種別列を追加
Const DEBUG_MODE As Boolean = False ' デバッグモード設定 - 通常運用時はFalse

' グローバル変数
Public g_IncludeToday As Boolean

' 勤怠入力漏れチェックのメイン処理
Public Sub 勤怠入力漏れチェック()
    On Error GoTo ErrorHandler
    
    ' 当日分を含めるかどうかのオプション
    Dim includeToday As Boolean
    
    ' 月末の最後の5営業日以内かどうかを確認
    Dim isWithinLastFiveDays As Boolean
    isWithinLastFiveDays = IsWithinLastFiveBusinessDaysOfMonth(Date)
    
    ' ユーザーに確認（月末の最後の5営業日以内の場合はデフォルトでTrue）
    Dim promptMsg As String
    If isWithinLastFiveDays Then
        promptMsg = "本日は月末の最後の5営業日以内です。" & vbCrLf & _
                   "当日分の勤怠入力状態もチェックしますか？" & vbCrLf & _
                   "（「はい」：当日分を含む、「いいえ」：前日までのみ）"
        includeToday = (MsgBox(promptMsg, vbQuestion + vbYesNo + vbDefaultButton1, "勤怠入力漏れチェック") = vbYes)
    Else
        promptMsg = "当日分の勤怠入力状態もチェックしますか？" & vbCrLf & _
                   "（「はい」：当日分を含む、「いいえ」：前日までのみ）"
        includeToday = (MsgBox(promptMsg, vbQuestion + vbYesNo + vbDefaultButton2, "勤怠入力漏れチェック") = vbYes)
    End If
    
    ' グローバル変数に設定
    g_IncludeToday = includeToday
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual ' 計算モードを手動に設定
    
    ' 進捗状況表示
    Application.StatusBar = "勤怠入力漏れチェックを開始しています..."
    
    ' CSVデータシートを取得
    Dim wsCSVData As Worksheet
    Set wsCSVData = ThisWorkbook.Worksheets("CSVデータ")
    
    If wsCSVData Is Nothing Then
        MsgBox "CSVデータシートが見つかりません。先にCSVファイルを読み込んでください。", vbExclamation
        GoTo CleanExit
    End If
    
    ' 勤怠入力漏れチェックシートの準備
    Dim missingEntriesSheet As Worksheet
    Set missingEntriesSheet = PrepareOutputSheet()
    
    ' 勤怠入力漏れの検出と出力
    DetectMissingEntries wsCSVData, missingEntriesSheet
    
    ' 概要統計の計算と表示
    CalculateAndDisplaySummary missingEntriesSheet
    
    ' 成功メッセージ（LINE WORKS通知機能の案内を追加）
    MsgBox "勤怠入力漏れチェックが完了しました。" & vbCrLf & _
           "集計結果を表示します。" & vbCrLf & vbCrLf & _
           "【LINE WORKS通知について】" & vbCrLf & _
           "「勤怠入力漏れ一覧」シートのL列にある" & vbCrLf & _
           "「LINE WORKS通知」ボタンをクリックすると、" & vbCrLf & _
           "SI1部リーダーチャンネルに通知が送信されます。", vbInformation
    
CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic ' 計算モードを自動に戻す
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' すべてのチェックを一度に実行する関数
Public Sub すべてのチェック実行()
    ' 画面更新と計算を一時停止
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' 休憩時間チェック実行
    Call 休憩時間チェック
    
    ' 勤怠入力漏れチェック実行
    Call 勤怠入力漏れチェック
    
    ' 画面更新と計算を再開
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    ' 完了メッセージ（LINE WORKS通知機能の案内を追加）
    MsgBox "すべてのチェックが完了しました。" & vbCrLf & vbCrLf & _
           "【LINE WORKS通知について】" & vbCrLf & _
           "「勤怠入力漏れ一覧」シートのL列にある" & vbCrLf & _
           "「LINE WORKS通知」ボタンをクリックすると、" & vbCrLf & _
           "SI1部リーダーチャンネルに通知が送信されます。", vbInformation
End Sub

