Attribute VB_Name = "Module7"
Option Explicit

' *************************************************************
' モジュール：UI拡張機能
' 目的：ユーザーインターフェース関連の拡張機能を提供する
' Copyright (c) 2025 SI1 shunpei.suzuki
' 作成日：2025年4月2日
'
' 改版履歴：
' 2025/04/02 module2から分割作成
' *************************************************************

' 既存のメニューを拡張するためのメインメニュー処理の変更
Public Sub メイン処理拡張()
    ' 標準のCSVファイル読み込み処理を実行
    Call CSVファイル読み込み
    
    ' 追加で勤怠入力漏れチェックを実行するかの確認
    Dim response As Integer
    response = MsgBox("休憩時間チェックに加えて、勤怠入力漏れチェックも実行しますか？", _
                     vbQuestion + vbYesNo, "勤怠入力漏れチェック確認")
    
    If response = vbYes Then
        Call 勤怠入力漏れチェック
    End If
End Sub

' CSV読み込みシート作成時にボタンを追加する処理を拡張
Public Sub CSV読み込みシート作成拡張()
    ' 既存のCSV読み込みシート作成処理を実行
    Call CSV読み込みシート作成
    
    ' 勤怠入力漏れチェック用のボタンを追加
    Dim mainSheet As Worksheet
    On Error Resume Next
    Set mainSheet = Worksheets("CSV読み込みシート")
    If mainSheet Is Nothing Then
        MsgBox "CSV読み込みシートが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    ' 勤怠入力漏れチェックボタンを配置
    With mainSheet.Buttons.Add(528, 60, 150, 30)
        .OnAction = "勤怠入力漏れチェック"
        .Caption = "勤怠入力漏れチェック"
    End With
    
    ' 全チェック実行ボタンを配置
    With mainSheet.Buttons.Add(528, 100, 150, 30)
        .OnAction = "すべてのチェック実行"
        .Caption = "すべてのチェック実行"
    End With
End Sub




