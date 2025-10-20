' ========================================
' Module4
' タイプ: 標準モジュール
' 行数: 61
' エクスポート日時: 2025-10-20 09:55:14
' ========================================

Option Explicit

' *************************************************************
' モジュール：日付ユーティリティ関数
' 目的：日付関連の判定や計算を行う関数群
' Copyright (c) 2025 SI1 shunpei.suzuki
' 作成日：2025年4月2日
'
' 改版履歴：
' 2025/04/02 module2から分割作成
' *************************************************************

' 月末の最後の5営業日以内かどうかを判定する関数
Public Function IsWithinLastFiveBusinessDaysOfMonth(checkDate As Date) As Boolean
    Dim lastDayOfMonth As Date
    Dim businessDaysLeft As Integer
    Dim currentDate As Date
    Dim i As Integer
    
    ' 月の最終日を取得
    lastDayOfMonth = DateSerial(Year(checkDate), Month(checkDate) + 1, 0)
    
    ' 最終日から逆算して5営業日をカウント
    businessDaysLeft = 0
    currentDate = lastDayOfMonth
    
    ' 最大で月末から10日前まで遡って5営業日を探す
    For i = 0 To 10
        If currentDate < checkDate Then
            ' チェック日より前の日付になったら終了
            Exit For
        End If
        
        ' 土日と祝日をスキップ（祝日判定は簡易的）
        If Weekday(currentDate) <> vbSaturday And Weekday(currentDate) <> vbSunday And Not IsHoliday(currentDate) Then
            businessDaysLeft = businessDaysLeft + 1
            
            ' 5営業日以内ならTrue
            If businessDaysLeft <= 5 And currentDate = checkDate Then
                IsWithinLastFiveBusinessDaysOfMonth = True
                Exit Function
            End If
        End If
        
        ' 前日へ
        currentDate = currentDate - 1
    Next i
    
    ' 5営業日以内でなければFalse
    IsWithinLastFiveBusinessDaysOfMonth = False
End Function

' 祝日かどうかを判定する関数（実際の祝日判定はより複雑になるため、必要に応じて拡張）
Public Function IsHoliday(checkDate As Date) As Boolean
    ' ここに祝日判定のロジックを追加
    ' 例: 祝日マスタを参照するなど
    
    ' 簡易的な実装（実際の環境に合わせて修正が必要）
    IsHoliday = False
End Function
