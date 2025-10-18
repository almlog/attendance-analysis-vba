' ========================================
' Module1
' タイプ: 標準モジュール
' 行数: 2075
' エクスポート日時: 2025-10-18 23:37:16
' ========================================


Option Explicit
' *************************************************************
' モジュール：休憩時間チェック
' 目的：実働時間に応じた休憩時間の取得確認および残業時間を計算する
' Copyright (c) 2025 SI1 shunpei.suzuki
' 作成日：2025年3月3日
'
' 改版履歴：
' 2025/03/01 初版作成_v1.0
' 2025/03/03 勤怠入力漏れチェックと統合_v1.5
' 2025/03/07 届出申請時の備考欄確認を追加①
' 2025/03/11 届出申請時の備考欄確認を追加②_v1.7
' 2025/03/16 シート名を変更_v1.8
' 2025/03/21 除外社員機能・社員数カウント修正・パフォーマンス最適化_v2.0
' 2025/08/20 定時退社率計算機能を統合_v2.1
' *************************************************************
' グローバル変数
Public g_HeaderCheckError As Boolean
Sub 休憩時間チェック()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim 実働時間 As Double
    Dim 休憩時間 As Double
    Dim 必要休憩時間 As Double
    Dim 休憩不足 As Double
    Dim resultSheet As Worksheet
    Dim violationSheet As Worksheet
    Dim deliverySheet As Worksheet
    Dim overtimeSheet As Worksheet
    
    ' 除外社員番号を取得
    Dim excludeIDs As Variant
    excludeIDs = 除外社員番号取得()
    
    ' アクティブシートを使用
    Set ws = ActiveSheet
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' ヘッダー行を確認
    Dim 社員番号Col As Integer
    Dim 氏名Col As Integer
    Dim 部門Col As Integer
    Dim 休憩時間Col As Integer
    Dim 実働時間Col As Integer
    Dim 届出Col As Integer
    Dim 状況区分Col As Integer
    Dim 法定外休出Col As Integer  ' 追加
    
    ' 各列のインデックスを特定
    Dim 備考Col As Integer
    備考Col = 0
    
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Select Case ws.Cells(1, i).Value
            Case "社員番号"
                社員番号Col = i
            ' 他の列（略）
            Case "備考"
                備考Col = i
            Case "氏名"
                氏名Col = i
            Case "部門"
                部門Col = i
            Case "休憩時間"
                休憩時間Col = i
            Case "実働時間"
                実働時間Col = i
            Case "届出内容"
                届出Col = i
            Case "状況区分"
                状況区分Col = i
            Case "法定外休出"
                法定外休出Col = i  ' 追加
        End Select
    Next i
    
    ' 備考列がヘッダーで見つからない場合、デフォルトでBH列（通常は60列目）を使用
    If 備考Col = 0 Then
        備考Col = 60  ' BH列
    End If
    ' 必要な列が存在するか確認
    If 社員番号Col = 0 Or 氏名Col = 0 Or 部門Col = 0 Or 休憩時間Col = 0 Or 実働時間Col = 0 Then
        MsgBox "必要な列（社員番号、氏名、部門、休憩時間、実働時間）が見つかりませんでした。", vbExclamation
        Exit Sub
    End If
    
    ' 結果用の全体分析シートを作成
    On Error Resume Next
    Set resultSheet = Worksheets("時間チェック_全体")
    If resultSheet Is Nothing Then
        Set resultSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        resultSheet.Name = "時間チェック_全体"
    Else
        resultSheet.Cells.Clear
    End If
    
    ' 違反者用のシートを作成
    Set violationSheet = Worksheets("休憩時間チェック_違反者")
    If violationSheet Is Nothing Then
        Set violationSheet = Worksheets.Add(After:=resultSheet)
        violationSheet.Name = "休憩時間チェック_違反者"
    Else
        violationSheet.Cells.Clear
    End If
    
    ' 届出一覧用のシートを作成 - 名前変更
    Set deliverySheet = Worksheets("休憩時間、備考一覧")
    If deliverySheet Is Nothing Then
        Set deliverySheet = Worksheets.Add(After:=violationSheet)
        deliverySheet.Name = "休憩時間、備考一覧"
    Else
        deliverySheet.Cells.Clear
    End If
    ' 残業一覧用のシートを作成
    Set overtimeSheet = Worksheets("残業一覧")
    If overtimeSheet Is Nothing Then
        Set overtimeSheet = Worksheets.Add(After:=deliverySheet)
        overtimeSheet.Name = "残業一覧"
    Else
        overtimeSheet.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 結果シートのヘッダーを設定
    With resultSheet
        .Cells(1, 1).Value = "社員番号"
        .Cells(1, 2).Value = "氏名"
        .Cells(1, 3).Value = "部門"
        .Cells(1, 4).Value = "日付"
        .Cells(1, 5).Value = "実働時間"
        .Cells(1, 6).Value = "休憩時間"
        .Cells(1, 7).Value = "必要休憩時間"
        .Cells(1, 8).Value = "休憩不足時間"
        .Cells(1, 9).Value = "残業時間"
        .Cells(1, 10).Value = "ステータス"
        
        ' ヘッダー行の書式設定
        .Range("A1:J1").Font.Bold = True
        .Range("A1:J1").Interior.Color = RGB(200, 200, 200)
    End With
    
    ' 違反者シートのヘッダーを設定
    With violationSheet
        .Cells(1, 1).Value = "社員番号"
        .Cells(1, 2).Value = "氏名"
        .Cells(1, 3).Value = "部門"
        .Cells(1, 4).Value = "日付"
        .Cells(1, 5).Value = "実働時間"
        .Cells(1, 6).Value = "休憩時間"
        .Cells(1, 7).Value = "必要休憩時間"
        .Cells(1, 8).Value = "休憩不足時間"
        .Cells(1, 9).Value = "残業時間"
        .Cells(1, 10).Value = "ステータス"
        .Cells(1, 11).Value = "備考"    ' 備考欄を追加
        
        ' ヘッダー行の書式設定
        .Range("A1:K1").Font.Bold = True
        .Range("A1:K1").Interior.Color = RGB(200, 200, 200)
    End With
    
    ' 届出一覧シートのヘッダーを設定
    With deliverySheet
        .Cells(1, 1).Value = "社員番号"
        .Cells(1, 2).Value = "氏名"
        .Cells(1, 3).Value = "部門"
        .Cells(1, 4).Value = "日付"
        .Cells(1, 5).Value = "届出"
        .Cells(1, 6).Value = "状況区分"
        .Cells(1, 7).Value = "実働時間"
        .Cells(1, 8).Value = "休憩時間"
        .Cells(1, 9).Value = "必要休憩時間"
        .Cells(1, 10).Value = "休憩不足時間"
        .Cells(1, 11).Value = "残業時間"
        .Cells(1, 12).Value = "ステータス"
        .Cells(1, 13).Value = "備考"  ' 備考欄を追加
        ' ヘッダー行の書式設定
        .Range("A1:M1").Font.Bold = True
        .Range("A1:M1").Interior.Color = RGB(200, 200, 200)
    End With
    
    ' 残業一覧シートのヘッダーを設定
    With overtimeSheet
        .Cells(1, 1).Value = "社員番号"
        .Cells(1, 2).Value = "氏名"
        .Cells(1, 3).Value = "部門"
        .Cells(1, 4).Value = "年月"
        .Cells(1, 5).Value = "総残業時間"
        .Cells(1, 6).Value = "ステータス"
        
        ' ヘッダー行の書式設定
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").Interior.Color = RGB(200, 200, 200)
    End With
    
    ' 行ごとに処理
    Dim resultRow As Long
    Dim violationRow As Long
    Dim deliveryRow As Long
    resultRow = 2
    violationRow = 2
    deliveryRow = 2
    
    ' 概要情報用の変数
    Dim totalCount As Long
    Dim violationCount As Long
    Dim deliveryCount As Long
    Dim processedCount As Long ' 実働時間が0以外のレコード数
    Dim overtimeCount As Long ' 残業時間が発生したレコード数
    Dim holidayWorkCount As Long ' 休日出勤の件数
    totalCount = 0
    violationCount = 0
    deliveryCount = 0
    processedCount = 0
    overtimeCount = 0
    holidayWorkCount = 0
    
    ' 残業集計用の配列
    Dim 社員番号Array() As String
    Dim 氏名Array() As String
    Dim 部門Array() As String
    Dim 年月Array() As String
    Dim 残業時間Array() As Double
    Dim 集計数 As Long
    集計数 = 0
    ReDim 社員番号Array(1000) ' 十分な大きさで初期化
    ReDim 氏名Array(1000)
    ReDim 部門Array(1000)
    ReDim 年月Array(1000)
    ReDim 残業時間Array(1000)
    
    ' 日付列の特定
    Dim 日付Col As Integer
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If ws.Cells(1, i).Value = "日付" Then
            日付Col = i
            Exit For
        End If
    Next i
    
    If 日付Col = 0 Then 日付Col = 5  ' デフォルトで5列目を日付と仮定
    ' カレンダー列をチェックして法定外休日の書式設定を行う
    Dim カレンダーCol As Integer
    Dim 曜日Col As Integer
    ' カレンダー列と曜日列を特定
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Select Case ws.Cells(1, i).Value
            Case "カレンダー"
                カレンダーCol = i
            Case "曜日"
                曜日Col = i
        End Select
    Next i
    ' カレンダー列が見つかった場合、法定外休日の書式設定
    If カレンダーCol > 0 Then
        For i = 2 To lastRow
            Dim カレンダー値 As String
            カレンダー値 = Trim(ws.Cells(i, カレンダーCol).Value)
            
            ' 法定外休日（「法定外」または「休」という文字を含む）をチェック
            If InStr(1, カレンダー値, "法定外", vbTextCompare) > 0 Or _
               InStr(1, カレンダー値, "休", vbTextCompare) > 0 Then
                
                ' 日付列と曜日列を赤文字に
                If 日付Col > 0 Then
                    ws.Cells(i, 日付Col).Font.Color = RGB(255, 0, 0)
                End If
                
                If 曜日Col > 0 Then
                    ws.Cells(i, 曜日Col).Font.Color = RGB(255, 0, 0)
                End If
                
                ' 行全体を薄いグレーに
                ws.Rows(i).Interior.Color = RGB(240, 240, 240)
            End If
        Next i
    End If
    For i = 2 To lastRow ' ヘッダー行をスキップ
        ' 社員番号を取得
        Dim currentEmployeeID As String
        currentEmployeeID = Trim(CStr(ws.Cells(i, 社員番号Col).Value))
        
        ' 除外社員番号のチェック
        Dim isExcluded As Boolean
        isExcluded = False
        
        For j = LBound(excludeIDs) To UBound(excludeIDs)
            If excludeIDs(j) <> "" And currentEmployeeID = excludeIDs(j) Then
                isExcluded = True
                Exit For
            End If
        Next j
        
        ' 除外社員の場合はスキップ
        If isExcluded Then
            GoTo NextIteration
        End If
        ' 実働時間と休憩時間を取得
        実働時間 = ConvertTimeToMinutes(ws.Cells(i, 実働時間Col).Value)
        休憩時間 = ConvertTimeToMinutes(ws.Cells(i, 休憩時間Col).Value)
        
        ' 総レコード数をカウント
        totalCount = totalCount + 1
        
        ' 休日出勤の判定
        Dim 休日出勤フラグ As Boolean
        Dim 法定外休出時間 As Double
        
        休日出勤フラグ = False
        法定外休出時間 = 0
        
        ' 法定外休出時間の取得
        If 法定外休出Col > 0 Then
            法定外休出時間 = ConvertTimeToMinutes(ws.Cells(i, 法定外休出Col).Value)
            If 法定外休出時間 > 0 Then
                休日出勤フラグ = True
            End If
        End If
        
        ' 届出内容を確認
        Dim 届出内容 As String
        Dim 状況区分 As String
        Dim 届出追加フラグ As Boolean
        
        届出追加フラグ = False
        届出内容 = ""
        状況区分 = ""
        
        If 届出Col > 0 Then
            届出内容 = Trim(ws.Cells(i, 届出Col).Value)
            If 届出内容 <> "" Then
                届出追加フラグ = True
                
                ' 届出内容に「休日出勤」「休出」などが含まれている場合も休日出勤とみなす
                If InStr(1, 届出内容, "休日出勤", vbTextCompare) > 0 Or _
                   InStr(1, 届出内容, "休出", vbTextCompare) > 0 Then
                    休日出勤フラグ = True
                End If
            End If
        End If
        
        If 状況区分Col > 0 Then
            状況区分 = Trim(ws.Cells(i, 状況区分Col).Value)
            If 状況区分 <> "" Then
                届出追加フラグ = True
            End If
        End If
        
        ' 休日出勤をカウント
        If 休日出勤フラグ Then
            holidayWorkCount = holidayWorkCount + 1
        End If
        
        ' 残業時間の計算
        Dim 残業時間 As Double
        If 休日出勤フラグ Then
            ' 休日出勤の場合は実働時間すべてを残業時間とする
            残業時間 = 実働時間
        ElseIf 実働時間 > 480 Then
            ' 通常勤務で8時間超過分
            残業時間 = 実働時間 - 480
        Else
            残業時間 = 0
        End If
        
        ' 残業時間がある場合は残業カウントを増やす
        If 残業時間 > 0 Then
            overtimeCount = overtimeCount + 1
        End If
        
        ' 残業集計用のデータを準備
        Dim 年月 As String
        Dim 日付Value As String
        
        ' 日付の取得と年月形式への変換
        日付Value = ws.Cells(i, 日付Col).Value
        
        ' 日付が正しい形式であることを確認
        If IsDate(日付Value) Then
            年月 = Format(CDate(日付Value), "yyyy/mm")
        Else
            ' 日付形式でない場合は "不明" とする
            年月 = "不明"
        End If
        
        ' 残業時間を社員ごと月ごとに集計
        Dim found As Boolean
        found = False
        
        ' 既存のエントリがあるか検索
        For j = 0 To 集計数 - 1
            If 社員番号Array(j) = ws.Cells(i, 社員番号Col).Value And _
               年月Array(j) = 年月 Then
                ' 既存エントリに加算
                残業時間Array(j) = 残業時間Array(j) + 残業時間
                found = True
                Exit For
            End If
        Next j
        
        ' 新規エントリの追加
        If Not found And 残業時間 > 0 Then
            社員番号Array(集計数) = ws.Cells(i, 社員番号Col).Value
            氏名Array(集計数) = ws.Cells(i, 氏名Col).Value
            部門Array(集計数) = ws.Cells(i, 部門Col).Value
            年月Array(集計数) = 年月
            残業時間Array(集計数) = 残業時間
            集計数 = 集計数 + 1
            
            ' 配列サイズの確認と拡張
            If 集計数 >= UBound(社員番号Array) Then
                ReDim Preserve 社員番号Array(UBound(社員番号Array) + 1000)
                ReDim Preserve 氏名Array(UBound(氏名Array) + 1000)
                ReDim Preserve 部門Array(UBound(部門Array) + 1000)
                ReDim Preserve 年月Array(UBound(年月Array) + 1000)
                ReDim Preserve 残業時間Array(UBound(残業時間Array) + 1000)
            End If
        End If
        
        ' 届出一覧に追加
        If 届出追加フラグ Then
            deliveryCount = deliveryCount + 1
            
            ' 届出一覧シートに追加
            With deliverySheet
                ' 社員番号を文字列として保持するために書式設定
                .Cells(deliveryRow, 1).NumberFormat = "@"
                .Cells(deliveryRow, 1).Value = ws.Cells(i, 社員番号Col).Value
                .Cells(deliveryRow, 2).Value = ws.Cells(i, 氏名Col).Value
                .Cells(deliveryRow, 3).Value = ws.Cells(i, 部門Col).Value
                .Cells(deliveryRow, 4).Value = ws.Cells(i, 日付Col).Value
                .Cells(deliveryRow, 5).Value = 届出内容
                .Cells(deliveryRow, 6).Value = 状況区分
                .Cells(deliveryRow, 7).Value = MinutesToTime(実働時間)
                .Cells(deliveryRow, 8).Value = MinutesToTime(休憩時間)
                
                ' 必要な休憩時間を計算
                Dim 届出必要休憩時間 As Double
                届出必要休憩時間 = 必要休憩時間計算(実働時間)
                .Cells(deliveryRow, 9).Value = MinutesToTime(届出必要休憩時間)
                
                ' 休憩不足時間を計算
                Dim 届出休憩不足 As Double
                届出休憩不足 = IIf(届出必要休憩時間 > 休憩時間, 届出必要休憩時間 - 休憩時間, 0)
                .Cells(deliveryRow, 10).Value = MinutesToTime(届出休憩不足)
                
                ' 残業時間を追加
                .Cells(deliveryRow, 11).Value = MinutesToTime(残業時間)
                
                ' ステータスを設定
                If 届出休憩不足 > 0 Then
                    .Cells(deliveryRow, 12).Value = "違反"
                    ' 違反行を赤色でハイライト
                    .Range(.Cells(deliveryRow, 1), .Cells(deliveryRow, 13)).Interior.Color = RGB(255, 200, 200)
                Else
                    .Cells(deliveryRow, 12).Value = "適正"
                End If
                
                ' 備考欄を追加
                Dim 備考 As String
                備考 = Trim(ws.Cells(i, 備考Col).Value)
                .Cells(deliveryRow, 13).Value = 備考
                
        
                ' 備考欄のチェックとハイライト
                If 届出内容 <> "有休" Then ' 有休以外の場合のみチェック
                    If 備考 = "" Then
                        .Cells(deliveryRow, 13).Interior.Color = RGB(255, 0, 0) ' 赤色でハイライト
                    Else
                        .Cells(deliveryRow, 13).Interior.Color = xlNone ' 色をクリア
                    End If
                Else
                   .Cells(deliveryRow, 13).Interior.Color = xlNone ' 有休の場合は色をクリア
                End If
        
            End With
        
            deliveryRow = deliveryRow + 1
        End If
        ' 実働時間が0の場合はスキップ（休日として処理）
        If 実働時間 <= 0 Then
            GoTo NextIteration
        End If
        
        ' 実働時間が0より大きい場合のみカウント
        processedCount = processedCount + 1
        
        ' 必要な休憩時間を計算
        必要休憩時間 = 必要休憩時間計算(実働時間)
        
        ' 休憩不足時間を計算
        休憩不足 = IIf(必要休憩時間 > 休憩時間, 必要休憩時間 - 休憩時間, 0)
        
        ' 全体分析シートに追加
        With resultSheet
            ' 社員番号を文字列として保持するために書式設定
            .Cells(resultRow, 1).NumberFormat = "@"
            .Cells(resultRow, 1).Value = ws.Cells(i, 社員番号Col).Value
            .Cells(resultRow, 2).Value = ws.Cells(i, 氏名Col).Value
            .Cells(resultRow, 3).Value = ws.Cells(i, 部門Col).Value
            .Cells(resultRow, 4).Value = ws.Cells(i, 日付Col).Value
            .Cells(resultRow, 5).Value = MinutesToTime(実働時間)
            .Cells(resultRow, 6).Value = MinutesToTime(休憩時間)
            .Cells(resultRow, 7).Value = MinutesToTime(必要休憩時間)
            .Cells(resultRow, 8).Value = MinutesToTime(休憩不足)
            .Cells(resultRow, 9).Value = MinutesToTime(残業時間)
            
            If 休憩不足 > 0 Then
                .Cells(resultRow, 10).Value = "違反"
                ' 違反行を赤色でハイライト
                .Range(.Cells(resultRow, 1), .Cells(resultRow, 10)).Interior.Color = RGB(255, 200, 200)
                violationCount = violationCount + 1
                
                ' 違反者シートにも追加
                With violationSheet
                    ' 社員番号を文字列として保持するために書式設定
                    .Cells(violationRow, 1).NumberFormat = "@"
                    .Cells(violationRow, 1).Value = ws.Cells(i, 社員番号Col).Value
                    .Cells(violationRow, 2).Value = ws.Cells(i, 氏名Col).Value
                    .Cells(violationRow, 3).Value = ws.Cells(i, 部門Col).Value
                    .Cells(violationRow, 4).Value = ws.Cells(i, 日付Col).Value
                    .Cells(violationRow, 5).Value = MinutesToTime(実働時間)
                    .Cells(violationRow, 6).Value = MinutesToTime(休憩時間)
                    .Cells(violationRow, 7).Value = MinutesToTime(必要休憩時間)
                    .Cells(violationRow, 8).Value = MinutesToTime(休憩不足)
                    .Cells(violationRow, 9).Value = MinutesToTime(残業時間)
                    .Cells(violationRow, 10).Value = "違反"
                    ' 備考欄の値を追加
                    .Cells(violationRow, 11).Value = ws.Cells(i, 備考Col).Value
                    ' 行全体を赤色でハイライト
                    .Range(.Cells(violationRow, 1), .Cells(violationRow, 11)).Interior.Color = RGB(255, 200, 200)
                End With
                violationRow = violationRow + 1
            Else
                .Cells(resultRow, 10).Value = "適正"
            End If
        End With
        
        resultRow = resultRow + 1
NextIteration:
    Next i
    
    ' 残業一覧シートにデータを出力
    Dim overtimeRow As Long
    overtimeRow = 2
    
    ' 集計した残業時間データを残業一覧シートに書き込む
    For j = 0 To 集計数 - 1
        ' 社員情報を取得
        Dim 社員番号 As String
        Dim 氏名 As String
        Dim 部門 As String
        Dim 集計年月 As String
        
        社員番号 = 社員番号Array(j)
        氏名 = 氏名Array(j)
        部門 = 部門Array(j)
        集計年月 = 年月Array(j)
        
        ' 総残業時間を取得
        Dim 総残業時間 As Double
        総残業時間 = 残業時間Array(j)
        
        ' ステータスを設定
        Dim 残業ステータス As String
        
        If 総残業時間 >= 70 * 60 Then ' 70時間以上
            残業ステータス = "親会社報告"
        ElseIf 総残業時間 >= 60 * 60 Then ' 60時間以上
            残業ステータス = "残業抑止要請"
        ElseIf 総残業時間 >= 45 * 60 Then ' 45時間以上
            残業ステータス = "年6回まで"
        Else
            残業ステータス = "適正"
        End If
        
        ' 残業一覧シートに追加
        With overtimeSheet
            ' 社員番号を文字列として保持
            .Cells(overtimeRow, 1).NumberFormat = "@"
            .Cells(overtimeRow, 1).Value = 社員番号
            .Cells(overtimeRow, 2).Value = 氏名
            .Cells(overtimeRow, 3).Value = 部門
            .Cells(overtimeRow, 4).Value = 集計年月
            .Cells(overtimeRow, 5).Value = MinutesToTime(総残業時間)
            .Cells(overtimeRow, 6).Value = 残業ステータス
            
            ' 45時間以上は赤背景、黒太字
            If 総残業時間 >= 45 * 60 Then
                .Range(.Cells(overtimeRow, 1), .Cells(overtimeRow, 6)).Interior.Color = RGB(255, 200, 200)
                .Range(.Cells(overtimeRow, 6), .Cells(overtimeRow, 6)).Font.Bold = True
                .Range(.Cells(overtimeRow, 6), .Cells(overtimeRow, 6)).Font.Color = RGB(0, 0, 0)
            End If
        End With
        
        overtimeRow = overtimeRow + 1
    Next j
    
    ' 概要情報の作成
    Dim summarySheet As Worksheet
    On Error Resume Next
    Set summarySheet = Worksheets("勤怠情報分析結果")
    If summarySheet Is Nothing Then
        Set summarySheet = Worksheets.Add(After:=overtimeSheet)
        summarySheet.Name = "勤怠情報分析結果"
    Else
        summarySheet.Cells.Clear
    End If
    On Error GoTo 0
    With summarySheet
        .Cells(1, 1).Value = "項目"
        .Cells(1, 2).Value = "数値"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 2).Font.Bold = True
        
        .Cells(2, 1).Value = "総レコード数"
        .Cells(2, 2).Value = totalCount
        
        .Cells(3, 1).Value = "処理対象レコード数（実働時間>0）"
        .Cells(3, 2).Value = processedCount
        
        .Cells(4, 1).Value = "適正レコード数"
        .Cells(4, 2).Value = processedCount - violationCount
        
        .Cells(5, 1).Value = "違反レコード数"
        .Cells(5, 2).Value = violationCount
        
        .Cells(6, 1).Value = "違反率"
        If processedCount > 0 Then
            .Cells(6, 2).Value = Format(violationCount / processedCount, "0.0%")
        Else
            .Cells(6, 2).Value = "0%"
        End If
        
        .Cells(7, 1).Value = "届出レコード数"
        .Cells(7, 2).Value = deliveryCount
        
        .Cells(8, 1).Value = "残業発生レコード数"
        .Cells(8, 2).Value = overtimeCount
        
        .Cells(9, 1).Value = "残業発生率"
        If processedCount > 0 Then
            .Cells(9, 2).Value = Format(overtimeCount / processedCount, "0.0%")
        Else
            .Cells(9, 2).Value = "0%"
        End If
        
        ' 休日出勤情報を追加
        .Cells(10, 1).Value = "休日出勤レコード数"
        .Cells(10, 2).Value = holidayWorkCount
        
        .Cells(11, 1).Value = "休日出勤率"
        If processedCount > 0 Then
            .Cells(11, 2).Value = Format(holidayWorkCount / processedCount, "0.0%")
        Else
            .Cells(11, 2).Value = "0%"
        End If
        
        ' 書式設定
        .Range("A1:B1").Interior.Color = RGB(200, 200, 200)
        .Columns("B:D").AutoFit ' B列からD列はAutoFit
        .Columns("A").ColumnWidth = 32 ' A列のみ幅を32に設定
    End With
    
    ' 社員番号列を文字列形式に設定（全てのシート）
    resultSheet.Columns("A").NumberFormat = "@"
    violationSheet.Columns("A").NumberFormat = "@"
    deliverySheet.Columns("A").NumberFormat = "@"
    overtimeSheet.Columns("A").NumberFormat = "@"
    
    ' 結果シートの時間関連列の書式を設定
    FormatTimeColumns resultSheet, 5, 6, 7, 8, 9
    FormatTimeColumns violationSheet, 5, 6, 7, 8, 9
    FormatTimeColumns deliverySheet, 7, 8, 9, 10, 11
    FormatTimeColumns overtimeSheet, 5
    
    ' 結果の列幅を自動調整
    resultSheet.Columns("B:J").AutoFit
    violationSheet.Columns("B:K").AutoFit
    deliverySheet.Columns("B:M").AutoFit
    overtimeSheet.Columns("B:F").AutoFit
    
    ' 違反者がいない場合のメッセージ
    If violationRow = 2 Then
        violationSheet.Cells(2, 1).Value = "休憩時間違反はありません。"
        violationSheet.Range("A2:J2").Merge
        violationSheet.Range("A2:J2").HorizontalAlignment = xlCenter
    End If
    
    ' 届出がない場合のメッセージ
    If deliveryRow = 2 Then
        deliverySheet.Cells(2, 1).Value = "届出の記録はありません。"
        deliverySheet.Range("A2:L2").Merge
        deliverySheet.Range("A2:L2").HorizontalAlignment = xlCenter
    End If
    
    ' 残業がない場合のメッセージ
    If overtimeRow = 2 Then
        overtimeSheet.Cells(2, 1).Value = "残業時間の記録はありません。"
        overtimeSheet.Range("A2:F2").Merge
        overtimeSheet.Range("A2:F2").HorizontalAlignment = xlCenter
    End If
    
    ' 概要シートをアクティブにする
    summarySheet.Activate
    
    ' 残業一覧シートから部門別残業時間を集計
    Call 部門別残業集計
        
        MsgBox "休憩時間・残業時間チェックが完了しました。" & vbCrLf & _
           "総レコード数: " & totalCount & vbCrLf & _
           "処理対象レコード数: " & processedCount & vbCrLf & _
           "違反レコード数: " & violationCount & vbCrLf & _
           "違反率: " & IIf(processedCount > 0, Format(violationCount / processedCount, "0.0%"), "0%") & vbCrLf & _
           "届出レコード数: " & deliveryCount & vbCrLf & _
           "残業発生レコード数: " & overtimeCount & vbCrLf & _
           "休日出勤レコード数: " & holidayWorkCount & vbCrLf & vbCrLf & _
           "勤怠入力漏れチェックを行い、" & vbCrLf & _
           "特別休暇申請が出ている場合はそのレコードを表示します。", _
           vbInformation, "休憩時間・残業時間・届出チェック結果"
       
End Sub
' 時間文字列（HH:MM）を分に変換する関数
Function ConvertTimeToMinutes(timeStr As Variant) As Double
    If IsEmpty(timeStr) Or timeStr = "" Then
        ConvertTimeToMinutes = 0
        Exit Function
    End If
    
    If IsNumeric(timeStr) Then
        ' すでに時間値として格納されている場合（Excelの時間は日の割合で格納）
        ConvertTimeToMinutes = timeStr * 24 * 60
        Exit Function
    End If
    
    Dim timeParts As Variant
    Dim hours As Double, minutes As Double
    
    ' HH:MM形式を想定
    timeParts = Split(CStr(timeStr), ":")
    
    If UBound(timeParts) >= 1 Then
        If IsNumeric(timeParts(0)) And IsNumeric(timeParts(1)) Then
            hours = CDbl(timeParts(0))
            minutes = CDbl(timeParts(1))
            ConvertTimeToMinutes = hours * 60 + minutes
        Else
            ConvertTimeToMinutes = 0
        End If
    Else
        ConvertTimeToMinutes = 0
    End If
End Function
' 分を時間文字列（HH:MM）に変換する関数
Function MinutesToTime(minutes As Double) As String
    Dim hours As Integer
    Dim mins As Integer
    
    hours = Int(minutes / 60)
    mins = minutes Mod 60
    
    ' 必ず2桁表示になるようフォーマット
    MinutesToTime = Format(hours, "00") & ":" & Format(mins, "00")
End Function
' 実働時間に基づいて必要な休憩時間（分）を計算する関数
Function 必要休憩時間計算(実働時間分 As Double) As Double
    If 実働時間分 < 360 Then
        ' 6時間未満
        必要休憩時間計算 = 0
    ElseIf 実働時間分 >= 360 And 実働時間分 < 480 Then
        ' 6時間以上8時間未満
        必要休憩時間計算 = 45
    Else
        ' 8時間以上
        必要休憩時間計算 = 60
    End If
End Function
' CSVファイルを読み込む関数
Sub CSVファイル読み込み()
    Dim filePath As Variant
    Dim ws As Worksheet
    Dim existingData As Boolean
    
    ' グローバル変数の初期化
    g_HeaderCheckError = False
    
    ' CSVデータシートが既に存在し、データがあるか確認
    On Error Resume Next
    Set ws = Worksheets("CSVデータ")
    If Not ws Is Nothing Then
        If ws.Cells(1, 1).Value <> "" Then
            existingData = True
        End If
    End If
    On Error GoTo 0
    
    ' 既存データがある場合、サイド分析の確認
    If existingData Then
        Dim response As Integer
        response = MsgBox("既にCSVデータが存在します。このデータを使用して分析を実行しますか？" & vbCrLf & _
                          "「はい」：現在のデータで分析を実行" & vbCrLf & _
                          "「いいえ」：新しいCSVファイルを読み込む", _
                          vbQuestion + vbYesNo, "データ分析確認")
        
        If response = vbYes Then
            ' 現在のデータで分析を実行
            Call 統合分析実行
            Exit Sub
        End If
        
        ' 「いいえ」を選択した場合は、新しいCSVファイルを読み込む処理を続行
    End If
    
    ' ファイル選択ダイアログを表示
    filePath = Application.GetOpenFilename("CSVファイル (*.csv),*.csv", , "CSVファイルを選択してください")
    
    If filePath = False Then
        MsgBox "ファイルが選択されませんでした。", vbExclamation
        Exit Sub
    End If
    
    ' 新しいシートを作成
    On Error Resume Next
    Application.DisplayAlerts = False
    
    ' 既存のシートを削除
    Dim sheetNames As Variant
    sheetNames = Array("CSVデータ", "時間チェック_全体", "休憩時間チェック_違反者", "勤怠情報分析結果", "休憩時間、備考一覧", "残業一覧", "勤怠入力漏れ一覧", "申請詳細分析一覧")
    Dim i As Integer
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = Worksheets(sheetNames(i))
        If Not ws Is Nothing Then
            ws.Delete
        End If
        On Error GoTo 0
    Next i
    
    Application.DisplayAlerts = True
    
    ' CSVデータ用シートを作成
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "CSVデータ"
    On Error GoTo 0
    
    ' CSVファイルを開く
    With ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A1"))
        .Name = "CSVインポート"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 932 ' 日本語Shift-JIS
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        
        ' 社員番号列を文字列として扱うための設定
        Dim fieldTypes() As Integer
        ReDim fieldTypes(1 To 100) ' 最大100列を想定
        For i = 1 To 100
            fieldTypes(i) = 1 ' デフォルトはGeneralとして扱う
        Next i
        
        ' 1列目（通常は社員番号）を文字列として扱う
        fieldTypes(1) = 2 ' 2はテキスト形式
        
        .TextFileColumnDataTypes = fieldTypes
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    ' 社員番号列を文字列形式に設定（先頭の0が削除されるのを防ぐ）
    Dim 社員番号Col As Integer
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If ws.Cells(1, i).Value = "社員番号" Then
            社員番号Col = i
            Exit For
        End If
    Next i
    If 社員番号Col > 0 Then
        ws.Columns(社員番号Col).NumberFormat = "@"
        ' 既存データを文字列として再設定
        For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            Dim 社員番号 As String
            社員番号 = Trim(CStr(ws.Cells(i, 社員番号Col).Value))
            ' 先頭ゼロが失われている場合、復元を試みる
            If Len(社員番号) < 7 And IsNumeric(社員番号) Then
                社員番号 = Right("0000000" & 社員番号, 7)
            End If
            ws.Cells(i, 社員番号Col).Value = 社員番号
        Next i
    End If
    
    ' 必要なヘッダーが存在するかチェック（休憩時間チェック用）
    Dim 休憩時間Col As Integer, 実働時間Col As Integer, 届出Col As Integer
    休憩時間Col = 0
    実働時間Col = 0
    届出Col = 0
    
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Select Case ws.Cells(1, i).Value
            Case "休憩時間"
                休憩時間Col = i
            Case "実働時間"
                実働時間Col = i
            Case "届出内容"
                届出Col = i
        End Select
    Next i
    
    ' ヘッダーチェック - 休憩時間チェック用のヘッダーがない場合はフラグを設定
    If 休憩時間Col = 0 Or 実働時間Col = 0 Then
        g_HeaderCheckError = True
        MsgBox "休憩時間チェックに必要なヘッダー（休憩時間、実働時間）が見つかりませんでした。" & vbCrLf & _
               "休憩時間・勤怠入力漏れチェックはスキップして申請分析のみ実行可能です。", vbExclamation
    End If
    
    ' 分析実行の確認
    Dim analysisResponse As Integer
    If g_HeaderCheckError Then
        ' ヘッダーエラーがある場合は、申請分析のみ実行するか確認
        Dim applicationAnalysisResponse As Integer
        applicationAnalysisResponse = MsgBox("申請分析を行いますか？" & vbCrLf & _
                                           "申請分析は申請決裁画面から全申請を対象として保存したcsvファイルを選択してください。", _
                                           vbQuestion + vbYesNo, "申請分析確認")
        
        If applicationAnalysisResponse = vbYes Then
            ' 申請分析を実行
            On Error Resume Next
            Call 申請詳細分析実行
            If Err.Number <> 0 Then
                MsgBox "申請分析の実行中にエラーが発生しました: " & Err.Description, vbExclamation
                Err.Clear
            End If
            On Error GoTo 0
        Else
            MsgBox "CSVファイルの読み込みが完了しました。必要に応じて各分析ボタンを押してください。", vbInformation
        End If
    Else
        ' 通常の分析実行確認
        analysisResponse = MsgBox("CSVファイルの読み込みが完了しました。分析を実行しますか？", _
                                vbQuestion + vbYesNo, "分析実行確認")
                                
        If analysisResponse = vbYes Then
            Call 統合分析実行
        Else
            MsgBox "CSVファイルの読み込みが完了しました。必要に応じて各分析ボタンを押してください。", vbInformation
        End If
    End If
End Sub
' すべての分析を統合して実行する関数
Public Sub 統合分析実行()
    On Error Resume Next
    
    ' CSVデータシートの存在確認
    Dim dataSheet As Worksheet
    Set dataSheet = ThisWorkbook.Worksheets("CSVデータ")
    
    If dataSheet Is Nothing Then
        MsgBox "CSVデータシートが見つかりません。先にCSVファイルを読み込んでください。", vbExclamation
        Exit Sub
    End If
    
    ' データの存在確認（最低限の検証）
    If dataSheet.Cells(1, 1).Value = "" Then
        MsgBox "CSVデータが空です。先にCSVファイルを読み込んでください。", vbExclamation
        Exit Sub
    End If
    
    ' ヘッダーエラーがある場合は休憩時間・勤怠入力漏れチェックをスキップ
    If g_HeaderCheckError Then
        ' 申請分析のみ実行するか確認
        Dim applicationAnalysisResponse As Integer
        applicationAnalysisResponse = MsgBox("申請分析を行いますか？" & vbCrLf & _
                                           "申請分析は申請決裁画面から全申請を対象として保存したcsvファイルを選択してください。", _
                                           vbQuestion + vbYesNo, "申請分析確認")
        
        If applicationAnalysisResponse = vbYes Then
            ' 申請分析を実行
            On Error Resume Next
            Call 申請詳細分析実行
            If Err.Number <> 0 Then
                MsgBox "申請分析の実行中にエラーが発生しました: " & Err.Description, vbExclamation
                Err.Clear
            End If
            On Error GoTo 0
        End If
        Exit Sub
    End If
    
    ' 進行状況表示
    Application.StatusBar = "分析を実行しています..."
    
    ' 重要な修正: アクティブシートを確実にCSVデータシートに設定
    dataSheet.Activate
    
    ' エラー発生フラグと実行成功フラグ
    Dim hasError As Boolean
    Dim breakTimeSuccess As Boolean
    Dim attendanceSuccess As Boolean
    
    hasError = False
    breakTimeSuccess = False
    attendanceSuccess = False
    
    ' 休憩時間チェック実行
    On Error Resume Next
    Err.Clear ' エラー状態をクリア
    Call 休憩時間チェック
    If Err.Number <> 0 Then
        ' エラーの詳細を記録（デバッグ用）
        Debug.Print "休憩時間チェックでエラー発生: " & Err.Number & " - " & Err.Description
        ' エラーがあっても続行する（サイレントに失敗）
        Err.Clear
    Else
        breakTimeSuccess = True
    End If
    
    ' 勤怠入力漏れチェック実行
    On Error Resume Next
    Err.Clear ' エラー状態をクリア
    Call 勤怠入力漏れチェック
    If Err.Number <> 0 Then
        ' エラーの詳細を記録（デバッグ用）
        Debug.Print "勤怠入力漏れチェックでエラー発生: " & Err.Number & " - " & Err.Description
        ' エラーがあっても続行する（サイレントに失敗）
        Err.Clear
    Else
        attendanceSuccess = True
    End If
    On Error GoTo 0
    
    Application.StatusBar = False
    
    ' いずれかのチェックが成功していれば概要シートを表示
    If breakTimeSuccess Then
        ' 勤怠情報分析結果シートをアクティブにする
        Dim summarySheet As Worksheet
        On Error Resume Next
        Set summarySheet = ThisWorkbook.Worksheets("勤怠情報分析結果")
        If Not summarySheet Is Nothing Then
            summarySheet.Activate
        End If
        On Error GoTo 0
    ElseIf attendanceSuccess Then
        ' 勤怠入力漏れ一覧シートをアクティブにする
        Dim missingEntriesSheet As Worksheet
        On Error Resume Next
        Set missingEntriesSheet = ThisWorkbook.Worksheets("勤怠入力漏れ一覧")
        If Not missingEntriesSheet Is Nothing Then
            missingEntriesSheet.Activate
        End If
        On Error GoTo 0
    End If
    ' 分析結果サマリーを表示（引数を渡す）
    Call DisplayAnalysisSummary(breakTimeSuccess, attendanceSuccess)
End Sub
' 分析結果のサマリーを表示する関数
Private Sub DisplayAnalysisSummary(ByVal breakTimeSuccess As Boolean, ByVal attendanceSuccess As Boolean)
    ' 各シートから情報を収集
    Dim breakViolationCount As Long
    Dim attendanceViolationCount As Long
    Dim overtimeCount As Long
    Dim sheetsCreated As String
    Dim activatedSheet As String
    
    sheetsCreated = ""
    activatedSheet = ""  ' 初期化
    
    ' 休憩時間違反数
    On Error Resume Next
    Dim violationSheet As Worksheet
    Set violationSheet = ThisWorkbook.Worksheets("休憩時間チェック_違反者")
    If Not violationSheet Is Nothing Then
        ' 2行目以降にデータがあるかチェック
        If Not IsEmpty(violationSheet.Cells(2, 1).Value) And violationSheet.Cells(2, 1).Value <> "休憩時間違反はありません。" Then
            ' 最終行を取得してカウント
            breakViolationCount = violationSheet.Cells(violationSheet.Rows.Count, "A").End(xlUp).Row - 1
        End If
    End If
    On Error GoTo 0
    
    ' 勤怠入力漏れ数
    On Error Resume Next
    Dim missingEntriesSheet As Worksheet
    Set missingEntriesSheet = ThisWorkbook.Worksheets("勤怠入力漏れ一覧")
    If Not missingEntriesSheet Is Nothing Then
        ' 2行目以降にデータがあるかチェック
        If Not IsEmpty(missingEntriesSheet.Cells(2, 1).Value) And missingEntriesSheet.Cells(2, 1).Value <> "勤怠入力漏れは検出されませんでした。" Then
            ' 最終行を取得してカウント
            attendanceViolationCount = missingEntriesSheet.Cells(missingEntriesSheet.Rows.Count, "A").End(xlUp).Row - 1
        End If
    End If
    On Error GoTo 0
    
    ' 残業発生数
    On Error Resume Next
    Dim overtimeSheet As Worksheet
    Set overtimeSheet = ThisWorkbook.Worksheets("残業一覧")
    If Not overtimeSheet Is Nothing Then
        ' 2行目以降にデータがあるかチェック
        If Not IsEmpty(overtimeSheet.Cells(2, 1).Value) And overtimeSheet.Cells(2, 1).Value <> "残業時間の記録はありません。" Then
            ' 最終行を取得してカウント
            overtimeCount = overtimeSheet.Cells(overtimeSheet.Rows.Count, "A").End(xlUp).Row - 1
        End If
    End If
    On Error GoTo 0
    
    If breakTimeSuccess Then
        sheetsCreated = sheetsCreated & "・休憩時間チェック" & vbCrLf
        activatedSheet = "「勤怠情報分析結果」シート"
    End If
    If attendanceSuccess Then
        sheetsCreated = sheetsCreated & "・勤怠入力漏れチェック" & vbCrLf
        If breakTimeSuccess = False Then
            activatedSheet = "「勤怠入力漏れ一覧」シート"
        End If
    End If
    
    ' サマリーメッセージの作成
    Dim message As String
    
    If breakTimeSuccess Or attendanceSuccess Then
        message = "分析が完了しました。" & vbCrLf & vbCrLf
        message = message & "【実行した分析】" & vbCrLf & sheetsCreated & vbCrLf
        message = message & "【分析結果サマリー】" & vbCrLf
        
        If breakTimeSuccess Then
            message = message & "・休憩時間違反: " & breakViolationCount & "件" & vbCrLf
            message = message & "・残業発生: " & overtimeCount & "件" & vbCrLf
        End If
        
        If attendanceSuccess Then
            message = message & "・勤怠入力漏れ: " & attendanceViolationCount & "件" & vbCrLf
        End If
        
        message = message & vbCrLf & "詳細は各シートをご確認ください。" & vbCrLf
        If activatedSheet <> "" Then
            message = message & "現在、" & activatedSheet & "を表示しています。"
        End If
    Else
        message = "必要なデータ列が不足しているため、分析を実行できませんでした。" & vbCrLf & vbCrLf
        message = message & "CSVデータが正しくフォーマットされていることを確認してください。" & vbCrLf
        message = message & "必要な列：社員番号、氏名、日付、（勤怠内容に応じて）休憩時間、実働時間など"
    End If
    
    ' メッセージを表示
    MsgBox message, vbInformation, "分析完了"
    
    ' 申請分析の実行確認
    Dim applicationAnalysisResponse As Integer
    applicationAnalysisResponse = MsgBox("申請分析を行いますか？" & vbCrLf & _
                                       "申請分析は申請決裁画面から全申請を対象として保存したcsvファイルを選択してください。", _
                                       vbQuestion + vbYesNo, "申請分析確認")
    
    If applicationAnalysisResponse = vbYes Then
        ' 申請分析を実行
        On Error Resume Next
        Call 申請詳細分析実行
        If Err.Number <> 0 Then
            MsgBox "申請分析の実行中にエラーが発生しました: " & Err.Description, vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
    Else
        ' 何もせず終了
        MsgBox "すべての分析が完了しました。", vbInformation, "分析終了"
        ' 分析が完了したので勤怠情報分析結果シートをアクティブにする
        On Error Resume Next
        Dim finalSummarySheet As Worksheet
        Set finalSummarySheet = ThisWorkbook.Worksheets("勤怠情報分析結果")
        If Not finalSummarySheet Is Nothing Then
            finalSummarySheet.Activate
        End If
        On Error GoTo 0
    End If
End Sub
' 特別休暇リストを表示する関数
Private Sub AddSpecialLeaveList(summarySheet As Worksheet, NextRow As Long)
    ' CSVデータシートを取得
    Dim wsCSVData As Worksheet
    On Error Resume Next
    Set wsCSVData = ThisWorkbook.Worksheets("CSVデータ")
    If wsCSVData Is Nothing Then Exit Sub
    
    ' 最終行を取得
    Dim lastRow As Long
    lastRow = wsCSVData.Cells(wsCSVData.Rows.Count, "A").End(xlUp).Row
    
    ' 列インデックスの特定
    Dim 社員番号Col As Integer, 氏名Col As Integer, 部門Col As Integer
    Dim 役職Col As Integer, 日付Col As Integer, 曜日Col As Integer
    Dim カレンダーCol As Integer, 届出Col As Integer, 備考Col As Integer
    
    ' 各列のインデックスを特定
    Dim i As Long
    For i = 1 To wsCSVData.Cells(1, wsCSVData.Columns.Count).End(xlToLeft).Column
        Select Case wsCSVData.Cells(1, i).Value
            Case "社員番号": 社員番号Col = i
            Case "氏名": 氏名Col = i
            Case "部門": 部門Col = i
            Case "役職": 役職Col = i
            Case "日付": 日付Col = i
            Case "曜日": 曜日Col = i
            Case "カレンダー": カレンダーCol = i
            Case "届出内容": 届出Col = i
            Case "備考": 備考Col = i
        End Select
    Next i
    
    ' 必要な列が見つからない場合はデフォルト値を設定
    If 社員番号Col = 0 Then 社員番号Col = 1
    If 氏名Col = 0 Then 氏名Col = 2
    If 部門Col = 0 Then 部門Col = 3
    If 役職Col = 0 Then 役職Col = 4
    If 日付Col = 0 Then 日付Col = 5
    If 曜日Col = 0 Then 曜日Col = 6
    If カレンダーCol = 0 Then カレンダーCol = 7
    If 届出Col = 0 Then 届出Col = 8
    If 備考Col = 0 Then 備考Col = 60 ' デフォルトでBH列
    
    ' 特別休暇レコードを収集
    Dim specialLeaves As New Collection
    Dim leaveRecord As Object
    
    ' CSV各行をチェック
    For i = 2 To lastRow
        ' 届出内容が「特別休暇」のレコードを抽出
        If Trim(wsCSVData.Cells(i, 届出Col).Value) = "特別休暇" Then
            Set leaveRecord = CreateObject("Scripting.Dictionary")
            leaveRecord.Add "社員番号", wsCSVData.Cells(i, 社員番号Col).Value
            leaveRecord.Add "氏名", wsCSVData.Cells(i, 氏名Col).Value
            leaveRecord.Add "部門", wsCSVData.Cells(i, 部門Col).Value
            leaveRecord.Add "役職", wsCSVData.Cells(i, 役職Col).Value
            leaveRecord.Add "日付", wsCSVData.Cells(i, 日付Col).Value
            leaveRecord.Add "曜日", wsCSVData.Cells(i, 曜日Col).Value
            leaveRecord.Add "カレンダー", wsCSVData.Cells(i, カレンダーCol).Value
            leaveRecord.Add "届出内容", wsCSVData.Cells(i, 届出Col).Value
            leaveRecord.Add "備考", wsCSVData.Cells(i, 備考Col).Value
            leaveRecord.Add "備考空欄", (Trim(wsCSVData.Cells(i, 備考Col).Value) = "")
            
            ' コレクションに追加
            specialLeaves.Add leaveRecord
        End If
    Next i
    
    ' 特別休暇がなければ終了
    If specialLeaves.Count = 0 Then Exit Sub
    
    ' 特別休暇リストの表示位置（勤怠入力漏れ概要の2行下）
    Dim listRow As Long
    listRow = NextRow + 8
    
    ' ヘッダー行を設定
    With summarySheet
        .Cells(listRow, 1).Value = "特別休暇リスト"
        .Cells(listRow, 1).Font.Bold = True
        .Cells(listRow, 1).Interior.Color = RGB(200, 200, 200)
        .Range(.Cells(listRow, 1), .Cells(listRow, 9)).Merge
        
        listRow = listRow + 1
        
        ' カラムヘッダー
        .Cells(listRow, 1).Value = "社員番号"
        .Cells(listRow, 2).Value = "氏名"
        .Cells(listRow, 3).Value = "部門"
        .Cells(listRow, 4).Value = "役職"
        .Cells(listRow, 5).Value = "日付"
        .Cells(listRow, 6).Value = "曜日"
        .Cells(listRow, 7).Value = "カレンダー"
        .Cells(listRow, 8).Value = "届出内容"
        .Cells(listRow, 9).Value = "備考"
        
        ' ヘッダー行の書式設定
        .Range(.Cells(listRow, 1), .Cells(listRow, 9)).Font.Bold = True
        .Range(.Cells(listRow, 1), .Cells(listRow, 9)).Interior.Color = RGB(200, 200, 200)
        
        listRow = listRow + 1
        
        ' 特別休暇レコードを表示
        Dim hasEmptyRemarks As Boolean
        hasEmptyRemarks = False
        
        Dim leaveItem As Object
        For Each leaveItem In specialLeaves
            .Cells(listRow, 1).NumberFormat = "@"
            .Cells(listRow, 1).Value = leaveItem("社員番号")
            .Cells(listRow, 2).Value = leaveItem("氏名")
            .Cells(listRow, 3).Value = leaveItem("部門")
            .Cells(listRow, 4).Value = leaveItem("役職")
            .Cells(listRow, 5).Value = leaveItem("日付")
            .Cells(listRow, 6).Value = leaveItem("曜日")
            .Cells(listRow, 7).Value = leaveItem("カレンダー")
            .Cells(listRow, 8).Value = leaveItem("届出内容")
            .Cells(listRow, 9).Value = leaveItem("備考")
            
            ' 備考欄が空欄の場合は優しい黄色でハイライト
            If leaveItem("備考空欄") Then
                .Cells(listRow, 9).Interior.Color = RGB(255, 255, 200)  ' より優しい黄色
                hasEmptyRemarks = True
            End If
            
            listRow = listRow + 1
        Next leaveItem
        
        ' コメントを追加
        .Cells(listRow + 1, 1).Value = "届出内容が明確、かつ確実に備考欄で説明がなされていること。"
        .Cells(listRow + 2, 1).Value = "備考欄の記載不備は修正が必要です。"
        
        If hasEmptyRemarks Then
            .Range(.Cells(listRow + 1, 1), .Cells(listRow + 2, 9)).Font.Color = RGB(255, 0, 0)
            .Range(.Cells(listRow + 1, 1), .Cells(listRow + 2, 9)).Font.Bold = True
        End If
        
        ' 表のボーダーを設定
        Dim tableRange As Range
        Set tableRange = .Range(.Cells(listRow - specialLeaves.Count, 1), .Cells(listRow - 1, 9))
        tableRange.Borders.LineStyle = xlContinuous
        tableRange.Borders.Weight = xlThin
        
        ' 列幅の自動調整
        .Columns("A:I").AutoFit
    End With
End Sub
' メイン処理（CSVファイル読み込みボタンから呼び出される関数）を統合
Sub メイン処理()
    Call CSVファイル読み込み
End Sub
' シートクリア処理（拡張版）
Sub シートクリア()
    Dim sheetNames As Variant
    Dim i As Integer
    Dim ws As Worksheet
    
    ' 削除対象のシート名を配列で定義（新しいシート名も含む）
    sheetNames = Array("CSVデータ", "時間チェック_全体", "休憩時間チェック_違反者", "勤怠情報分析結果", "休憩時間、備考一覧", "残業一覧", "勤怠入力漏れ一覧", "申請詳細分析一覧")
    Application.DisplayAlerts = False
    
    ' 各シートを確認し、存在すれば削除
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(sheetNames(i))
        If Not ws Is Nothing Then
            ws.Delete
        End If
        On Error GoTo 0
    Next i
    
    Application.DisplayAlerts = True
    
    MsgBox "すべての分析シートがクリアされました。", vbInformation
End Sub
' CSV読み込みシート作成時にボタンを追加する処理（統合版）
Public Sub CSV読み込みシート作成統合()
    ' CSV読み込みシート作成
    Call CSV読み込みシート作成
    
    ' メインシートを取得
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
    
    ' 申請分析ボタンを配置
    With mainSheet.Buttons.Add(528, 100, 150, 30)
        .OnAction = "申請詳細分析実行"
        .Caption = "申請分析"
    End With
    
    MsgBox "ボタンが追加されました。", vbInformation
End Sub
' 時間列の書式を設定するヘルパー関数
Sub FormatTimeColumns(ws As Worksheet, ParamArray columnIndexes() As Variant)
    Dim lastRow As Long
    Dim i As Long, j As Long
    
    ' シートの最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub ' データがない場合は終了
    
    ' 指定された列の書式を設定
    For i = 0 To UBound(columnIndexes)
        For j = 2 To lastRow ' ヘッダー行をスキップ
            ' 時間形式を HH:MM に設定
            ws.Cells(j, columnIndexes(i)).NumberFormat = "[hh]:mm"
        Next j
    Next i
End Sub
' 別名で保存する関数
Sub 別名保存()
    Dim filePath As Variant
    
    ' 保存ダイアログを表示（マクロ有効ブック形式）
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="勤之助明細チェック_" & Format(Date, "yyyymmdd") & ".xlsm", _
        FileFilter:="マクロ有効ブック (*.xlsm),*.xlsm", _
        Title:="別名で保存")
    If filePath = False Then
        MsgBox "保存がキャンセルされました。", vbExclamation
        Exit Sub
    End If
    
    ' 現在のブックを保存
    On Error Resume Next
    ThisWorkbook.SaveAs filePath, xlOpenXMLWorkbookMacroEnabled
    
    If Err.Number <> 0 Then
        MsgBox "保存中にエラーが発生しました: " & Err.Description, vbCritical
    Else
        MsgBox "ファイルが正常に保存されました: " & filePath, vbInformation
    End If
    On Error GoTo 0
End Sub

' *************************************************************
' CSV読み込みシートの説明文を更新（完全版）
' 目的: SI1部専用ツールの情報追加、LINEWORKS通知機能説明追加
' 作成日: 2025-10-18
' 更新内容:
' - SI1部専用ツールであることを明記
' - LINEWORKS通知機能の説明追加
' - 連絡先情報の追加
' - 除外社員番号の入力欄をA51に移動（元のA46から変更）
' *************************************************************

Sub CSV読み込みシート作成()
    Dim ws As Worksheet
    Dim mainSheet As Worksheet
    Dim initSheet As Worksheet
    Dim btn As Button
    Dim sh As Worksheet
    
    Application.DisplayAlerts = False
    
    ' 初期化シートとCSV読み込みシート以外のシートを削除
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name <> "初期化シート" And sh.Name <> "CSV読み込みシート" Then
            sh.Delete
        End If
    Next sh
    
    Application.DisplayAlerts = True
    
    ' 初期化シートが存在するか確認
    On Error Resume Next
    Set initSheet = Worksheets("初期化シート")
    If initSheet Is Nothing Then
        ' シートが存在しない場合は作成
        Set initSheet = Worksheets.Add(Before:=Worksheets(1))
        initSheet.Name = "初期化シート"
    End If
    On Error GoTo 0
    
    ' CSV読み込みシートが存在するか確認
    On Error Resume Next
    Set mainSheet = Worksheets("CSV読み込みシート")
    If mainSheet Is Nothing Then
        ' シートが存在しない場合は作成
        Set mainSheet = Worksheets.Add(After:=initSheet)
        mainSheet.Name = "CSV読み込みシート"
    End If
    On Error GoTo 0
    
    ' メインシートをクリア
    mainSheet.Cells.Clear
    
    ' タイトルを設定
    With mainSheet
        .Range("A1").Value = "休憩時間チェックツール（SI1部専用）"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Color = RGB(0, 102, 204) ' 青色
    End With
    
    ' 説明文を8行目以降に配置
    With mainSheet
        .Range("A8").Value = "説明："
        .Range("A8").Font.Bold = True
        
        .Range("A9").Value = "実働時間に対して適切な休憩時間がとられているかをチェックします。"
        .Range("A10").Value = "残業時間について正確に算出します。"
        .Range("A11").Value = "勤怠の入力漏れについてチェックします。"
        .Range("A12").Value = "申請の分析を行い、社員ごとの有休日数などを集計します。"
        
        ' 区切り線
        .Range("A13").Value = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        .Range("A13").Font.Color = RGB(128, 128, 128)
        
        ' ★★★ SI1部専用ツール情報（追加） ★★★
        .Range("A14").Value = "【SI1部専用ツール】"
        .Range("A14").Font.Bold = True
        .Range("A14").Font.Size = 12
        .Range("A14").Font.Color = RGB(255, 0, 0) ' 赤色で強調
        
        .Range("A15").Value = "このツールはSI1部専用にカスタマイズされています。"
        .Range("A15").Font.Color = RGB(255, 0, 0)
        
        ' 区切り線
        .Range("A16").Value = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        .Range("A16").Font.Color = RGB(128, 128, 128)
        
        ' ★★★ LINEWORKS通知機能（追加） ★★★
        .Range("A17").Value = "【LINEWORKS通知機能】"
        .Range("A17").Font.Bold = True
        .Range("A17").Font.Size = 11
        .Range("A17").Font.Color = RGB(0, 153, 0) ' 緑色
        
        .Range("A18").Value = "・勤怠未入力者の情報をLINE WORKS「SI1部リーダーチャンネル」に自動通知"
        .Range("A19").Value = "・[勤怠入力漏れチェック]ボタンで未入力者リストを取得"
        .Range("A20").Value = "・未入力者リストが表示された後、[管理者通知]ボタンで通知を送信"
        .Range("A21").Value = "・緊急度別に自動分類されます："
        
        ' 緊急度の視覚的説明（色付きセル）
        .Range("B22").Value = "【緊急】"
        .Range("B22").Interior.Color = RGB(255, 200, 200) ' 薄い赤
        .Range("B22").Font.Color = RGB(192, 0, 0) ' 濃い赤
        .Range("B22").Font.Bold = True
        .Range("C22").Value = "5日以上未入力"
        
        .Range("B23").Value = "【要注意】"
        .Range("B23").Interior.Color = RGB(255, 235, 156) ' 薄い黄色
        .Range("B23").Font.Color = RGB(204, 102, 0) ' オレンジ
        .Range("B23").Font.Bold = True
        .Range("C23").Value = "3-4日未入力"
        
        .Range("B24").Value = "【確認】"
        .Range("B24").Interior.Color = RGB(198, 239, 206) ' 薄い緑
        .Range("B24").Font.Color = RGB(0, 128, 0) ' 濃い緑
        .Range("B24").Font.Bold = True
        .Range("C24").Value = "1-2日未入力"
        
        ' 枠線追加
        .Range("B22:C24").Borders.LineStyle = xlContinuous
        .Range("B22:C24").Borders.Weight = xlThin
        
        ' 区切り線
        .Range("A25").Value = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        .Range("A25").Font.Color = RGB(128, 128, 128)
        
        ' ★★★ 機能追加・修正の連絡先（追加） ★★★
        .Range("A26").Value = "【機能追加・修正のご依頼】"
        .Range("A26").Font.Bold = True
        .Range("A26").Font.Size = 11
        .Range("A26").Font.Color = RGB(0, 102, 204)
        
        .Range("A27").Value = "機能追加や不具合修正のご要望は下記までご連絡ください："
        .Range("A28").Value = "  連絡先: suzuki.shunpei@altx.co.jp"
        .Range("A28").Font.Bold = True
        .Range("A28").Font.Size = 11
        .Range("A28").Font.Color = RGB(0, 102, 204)
        
        .Range("A29").Value = "※ご連絡の際は、具体的な内容や動作環境をお知らせください"
        .Range("A29").Font.Size = 9
        .Range("A29").Font.Color = RGB(128, 128, 128)
        
        ' 区切り線
        .Range("A30").Value = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        .Range("A30").Font.Color = RGB(128, 128, 128)
        
        ' 既存の説明（元のA14～A22の内容をA31以降に移動）
        .Range("A31").Value = "勤之助の月別出勤簿ページから集計したい社員のCSV明細ファイルを出力してください。"
        .Range("A31").Font.Bold = True
        .Range("A32").Value = "個人でも複数人でも対応します。部署が分かれていても問題ありません。"
        
        .Range("A33").Value = "CSVファイルを読み込んでCSV明細ファイルを選択してください。"
        
        .Range("A35").Value = "休憩時間の基準："
        .Range("A35").Font.Bold = True
        .Range("A36").Value = "・実働6時間未満: 休憩なしでも可"
        .Range("A37").Value = "・実働6～8時間: 45分以上の休憩が必要"
        .Range("A38").Value = "・実働8時間以上: 1時間以上の休憩が必要"
        
        ' 定時退社の基準（元のA23～A26をA40～A43に移動）
        .Range("A40").Value = "定時退社の基準："
        .Range("A40").Font.Bold = True
        .Range("A41").Value = "・定時退社：退社時刻が17:45より前、または有休等の休暇取得"
        .Range("A42").Value = "・定時退社率 = (定時退社日数 ÷ 総勤務日数) × 100"
        .Range("A43").Value = "・総勤務日：実働1時間以上または有休等の休暇取得日（振替休暇除く）"
        
        ' 届出,備考欄の基準（元のA28～A30をA45～A47に移動）
        .Range("A45").Value = "届出,備考欄の基準"
        .Range("A45").Font.Bold = True
        .Range("A46").Value = "・年休以外の届けについて備考欄チェックを行なう"
        .Range("A47").Value = "・特別休暇がある場合は概要シートに一覧を表示させ備考欄チェックを行う"
        
        ' 勤怠入力漏れの基準（元のA32～A35をA49～A52に移動）
        .Range("A49").Value = "勤怠入力漏れの基準："
        .Range("A49").Font.Bold = True
        .Range("A50").Value = "・前日以前の平日の勤怠入力が行なわれているかチェック"
        .Range("A51").Value = "・休暇届けが出ている場合は入力チェックを行わない"
        .Range("A52").Value = "・月末までの5営業日は当日入力の確認を行うかのポップアップを表示"
        
        ' 申請分析の基準（元のA37～A40をA54～A57に移動）
        .Range("A54").Value = "申請分析の基準："
        .Range("A54").Font.Bold = True
        .Range("A55").Value = "・申請決裁画面から全申請を対象としてCSVエクスポートしたファイルを解析"
        .Range("A56").Value = "・社員ごとの有休日数、時間有休時間、午前/午後有休回数を集計"
        .Range("A57").Value = "・5日以上の休暇取得状況を確認"
        
        ' 除外社員設定項目（元のA42～A46をA59～A63に移動）★★★重要：A51に変更★★★
        .Range("A59").Value = "除外社員の設定："
        .Range("A59").Font.Bold = True
        .Range("A60").Value = "・退職済み、移動済みなど分析から除外したい社員の社員番号を以下のグレーのセルに入力してください。"
        .Range("A61").Value = "・複数の社員を除外する場合は社員番号をカンマ区切りで入力してください。例: 1234567,2345678"
        
        .Range("A62").Value = "除外社員番号："
        .Range("A63").Value = "" ' ★★★ 入力欄をA51に変更 ★★★
        .Range("A63").BorderAround ColorIndex:=1
        .Range("A63").Interior.Color = RGB(217, 217, 217)
        .Range("A63").NumberFormat = "@" ' 文字列形式
        
        ' ロジック説明（元のA48～A51をA65～A68に移動）
        .Range("A65").Value = "ロジック："
        .Range("A65").Font.Bold = True
        .Range("A66").Value = "・休憩時間：実働時間に応じて必要な休憩時間を計算し、取得している休憩時間と比較"
        .Range("A67").Value = "・残業時間：平日の場合は8時間（480分）を超えた時間、休日出勤は全時間を残業としてカウント"
        .Range("A68").Value = "・部門別集計：部門ごとの残業時間、平均残業時間、休日出勤数を集計"
        .Range("A69").Value = "・遅刻・早退：正確にカウント"
        .Range("A70").Value = "・有休申請：半休・全休を正確に判定"
        .Range("A71").Value = "・定時退社：遅刻・早退・欠勤・休日出勤を除き、17:45前退社または休暇取得"
        .Range("A72").Value = "・違反検出：必要な休憩時間を取得していない場合は「違反」として表示"
        .Range("A73").Value = "・違反検出：必要な勤怠入力をしていない場合は「出退勤時刻なし」として表示"
        
        ' Copyright（元のA56をA75に移動）
        .Range("A75").Value = "Copyright (c) 2025 SI1 shunpei.suzuki"
        .Range("A75").Font.Italic = True
        .Range("A75").Font.Size = 8
        .Range("A75").Font.Color = RGB(128, 128, 128)
        
    End With
    
    ' 列幅を調整
    mainSheet.Columns("A").ColumnWidth = 100
    
    ' ボタンを配置（既存のボタンは削除）
    On Error Resume Next
    For Each btn In mainSheet.Buttons
        btn.Delete
    Next btn
    On Error GoTo 0
    
    ' CSVファイル読み込みボタンを配置
    With mainSheet.Buttons.Add(96, 60, 120, 30)
        .OnAction = "メイン処理"
        .Caption = "CSVファイル読み込み"
    End With
    
    ' 別名保存ボタンを配置
    With mainSheet.Buttons.Add(240, 60, 120, 30)
        .OnAction = "別名保存"
        .Caption = "別名で保存"
    End With
    
    ' シートクリアボタンを配置
    With mainSheet.Buttons.Add(384, 60, 120, 30)
        .OnAction = "シートクリア"
        .Caption = "シートをクリア"
    End With
    
    ' 初期化シートを非表示にする
    initSheet.Visible = xlSheetVeryHidden
    
    ' このシートをアクティブにする
    mainSheet.Activate
    
    MsgBox "CSV読み込みシートの初期化が完了しました。", vbInformation
End Sub

' ★★★ 除外社員番号取得関数を修正（A46 → A63に変更）★★★
Function 除外社員番号取得() As Variant
    Dim mainSheet As Worksheet
    Dim excludeNumbersStr As String
    Dim excludeNumbers As Variant
    Dim i As Long
    
    On Error Resume Next
    Set mainSheet = ThisWorkbook.Worksheets("CSV読み込みシート")
    If mainSheet Is Nothing Then
        ' シートが見つからない場合は空の配列を返す
        ReDim excludeNumbers(0)
        excludeNumbers(0) = ""
        除外社員番号取得 = excludeNumbers
        Exit Function
    End If
    
    ' ★★★ 除外社員番号欄の値を取得（A46 → A63に変更）★★★
    excludeNumbersStr = Trim(mainSheet.Range("A63").Value)
    
    ' デバッグ出力
    Debug.Print "除外社員番号文字列: [" & excludeNumbersStr & "]"
    
    If excludeNumbersStr = "" Then
        ' 入力がない場合は空の配列を返す
        ReDim excludeNumbers(0)
        excludeNumbers(0) = ""
    Else
        ' カンマ区切りで分割
        excludeNumbers = Split(excludeNumbersStr, ",")
        
        ' 各番号からスペースを削除し、文字列として整形
        For i = LBound(excludeNumbers) To UBound(excludeNumbers)
            excludeNumbers(i) = Trim(CStr(excludeNumbers(i)))
            ' デバッグ出力
            Debug.Print "除外社員番号[" & i & "]: [" & excludeNumbers(i) & "]"
        Next i
    End If
    
    除外社員番号取得 = excludeNumbers
End Function

' 初期化シートを表示するプロシージャ（開発者モード用）
Sub 初期化シート表示()
    Dim initSheet As Worksheet
    
    On Error Resume Next
    Set initSheet = Worksheets("初期化シート")
    
    If Not initSheet Is Nothing Then
        initSheet.Visible = xlSheetVisible
        initSheet.Activate
    Else
        MsgBox "初期化シートが見つかりません。", vbExclamation
    End If
    On Error GoTo 0
End Sub
' 初期化シートを非表示にするプロシージャ
Sub 初期化シート非表示()
    Dim initSheet As Worksheet
    
    On Error Resume Next
    Set initSheet = Worksheets("初期化シート")
    
    If Not initSheet Is Nothing Then
        initSheet.Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0
End Sub
' 部門別残業時間を集計する改良版関数
Sub 部門別残業集計()
    Dim wsCSVData As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim summaryRow As Long
    Dim dict As Object
    Dim dept As Variant
    Dim deptCode As String
    Dim deptName As String
    Dim employeeDict As Object
    Dim empID As String
    Dim empName As String
    Dim emp As Variant ' For Each ループのための変数追加
    
    ' 必要なシートの取得
    On Error Resume Next
    Set wsCSVData = ThisWorkbook.Worksheets("CSVデータ")
    Set wsSummary = ThisWorkbook.Worksheets("勤怠情報分析結果")
    
    If wsCSVData Is Nothing Then
        MsgBox "CSVデータシートが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    If wsSummary Is Nothing Then
        Set wsSummary = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsSummary.Name = "勤怠情報分析結果"
    End If
    On Error GoTo 0
    
    ' CSVデータの最終行を取得
    lastRow = wsCSVData.Cells(wsCSVData.Rows.Count, "A").End(xlUp).Row
    
    If lastRow <= 1 Then
        MsgBox "CSVデータが存在しません。", vbExclamation
        Exit Sub
    End If
    
    ' 列インデックスの特定
    Dim 社員番号Col As Integer
    Dim 氏名Col As Integer
    Dim 部門Col As Integer
    Dim 届出内容Col As Integer
    Dim 実働時間Col As Integer
    
    ' 各列のインデックスを特定（ヘッダー行から）
    For i = 1 To wsCSVData.Cells(1, wsCSVData.Columns.Count).End(xlToLeft).Column
        Select Case wsCSVData.Cells(1, i).Value
            Case "社員番号"
                社員番号Col = i
            Case "氏名"
                氏名Col = i
            Case "部門"
                部門Col = i
            Case "届出内容"
                届出内容Col = i
            Case "実働時間"
                実働時間Col = i
        End Select
    Next i
    
    ' 必要な列が見つからない場合はデフォルト値を設定
    If 社員番号Col = 0 Then 社員番号Col = 1
    If 氏名Col = 0 Then 氏名Col = 2
    If 部門Col = 0 Then 部門Col = 3
    If 届出内容Col = 0 Then 届出内容Col = 9
    If 実働時間Col = 0 Then 実働時間Col = 41
    
    ' デバッグ用：列情報を表示
    Debug.Print "社員番号Col: " & 社員番号Col & ", 氏名Col: " & 氏名Col & ", 部門Col: " & 部門Col & ", 届出内容Col: " & 届出内容Col & ", 実働時間Col: " & 実働時間Col
    
    ' 部門データ集計用の辞書を作成
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' CSVデータの各行を処理
    For i = 3 To lastRow ' ヘッダー行をスキップ
        deptCode = Trim(wsCSVData.Cells(i, 部門Col).Value)
        empID = Trim(wsCSVData.Cells(i, 社員番号Col).Value)
        empName = Trim(wsCSVData.Cells(i, 氏名Col).Value)
        
        ' 空の部門コードをスキップ
        If deptCode <> "" And empID <> "" Then  ' ← empIDも空でないことを確認
            ' この部門が辞書にない場合は追加
            If Not dict.Exists(deptCode) Then
                Set dict(deptCode) = CreateObject("Scripting.Dictionary")
                dict(deptCode)("TotalOvertime") = 0        ' 合計残業時間（分）
                dict(deptCode)("OccurrenceCount") = 0      ' 残業発生件数
                dict(deptCode)("HolidayWorkCount") = 0     ' 休日出勤件数
                Set dict(deptCode)("Employees") = CreateObject("Scripting.Dictionary") ' 社員リスト
                dict(deptCode)("DepartmentName") = deptCode ' 部門名
            End If
            
            ' ★修正箇所: CSVに存在するすべての社員を追加（重複チェックあり）
            ' 社員番号が存在し、まだ追加されていない場合のみ追加
            If Not dict(deptCode)("Employees").Exists(empID) Then
                dict(deptCode)("Employees").Add empID, empName
            End If
            
            ' 以下、残業時間の計算処理は従来通り
            ' 休日出勤かどうかを判定
            Dim isHolidayWork As Boolean
            Dim deliveryContent As String
            
            deliveryContent = Trim(wsCSVData.Cells(i, 届出内容Col).Value)
            
            ' 届出内容で休日出勤判定
            isHolidayWork = (InStr(1, deliveryContent, "休日出勤", vbTextCompare) > 0) Or _
                           (InStr(1, deliveryContent, "休出", vbTextCompare) > 0)
            
            ' 実働時間を取得して分に変換
            Dim workingMinutes As Double
            Dim overtimeMinutes As Double
            Dim rawWorkingTime As Variant
            
            rawWorkingTime = wsCSVData.Cells(i, 実働時間Col).Value
            
            ' 実働時間を適切に変換
            If IsNumeric(rawWorkingTime) Then
                If rawWorkingTime < 1 Then
                    workingMinutes = rawWorkingTime * 24 * 60
                Else
                    workingMinutes = 0
                End If
            Else
                workingMinutes = ConvertTimeToMinutes(rawWorkingTime)
            End If
            
            ' 残業時間を計算
            If isHolidayWork Then
                overtimeMinutes = workingMinutes
            ElseIf workingMinutes > 480 Then
                overtimeMinutes = workingMinutes - 480
            Else
                overtimeMinutes = 0
            End If
            
            ' 残業時間がある場合のみ残業関連の集計
            If overtimeMinutes > 0 Then
                dict(deptCode)("TotalOvertime") = dict(deptCode)("TotalOvertime") + overtimeMinutes
                dict(deptCode)("OccurrenceCount") = dict(deptCode)("OccurrenceCount") + 1
                
                If isHolidayWork Then
                    dict(deptCode)("HolidayWorkCount") = dict(deptCode)("HolidayWorkCount") + 1
                End If
            End If
        End If
    Next i

    ' 概要シートの残業集計部分をクリア
    wsSummary.Range("A14:F100").ClearContents
    ' 最後の行番号を取得（最低でも13行目から始める）
    Dim headerRow As Long
    headerRow = 13
    For i = 1 To 30
        If IsEmpty(wsSummary.Cells(i, 1).Value) Then
            headerRow = i + 2  ' 空の行を見つけたら2行空けて配置
            Exit For
        End If
    Next i
    ' ヘッダーを設定
    wsSummary.Cells(headerRow, 1).Value = "部署"
    wsSummary.Cells(headerRow, 2).Value = "合計残業時間"
    wsSummary.Cells(headerRow, 3).Value = "平均残業時間/回"
    wsSummary.Cells(headerRow, 4).Value = "平均月残業時間/人"
    wsSummary.Cells(headerRow, 5).Value = "休日出勤回数"
    wsSummary.Cells(headerRow, 6).Value = "人数"
    ' ヘッダー行の書式設定
    With wsSummary.Range("A" & headerRow & ":F" & headerRow)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(200, 200, 200)
    End With
    ' 部門ごとのデータを出力
    summaryRow = headerRow + 1  ' ヘッダーの次の行から
    ' 集計用の変数
    Dim totalOvertimeAll As Double
    Dim totalOccurrencesAll As Long
    Dim totalPersonsAll As Long
    Dim totalHolidayWorkAll As Long
    Dim allEmployees As Object
    
    totalOvertimeAll = 0
    totalOccurrencesAll = 0
    totalHolidayWorkAll = 0
    Set allEmployees = CreateObject("Scripting.Dictionary")
    
    ' 各部門の集計を出力
    For Each dept In dict.Keys
        Dim totalOvertime As Double
        Dim occurrenceCount As Long
        Dim holidayWorkCount As Long
        Dim personCount As Long
        
        totalOvertime = dict(dept)("TotalOvertime")
        occurrenceCount = dict(dept)("OccurrenceCount")
        holidayWorkCount = dict(dept)("HolidayWorkCount")
        personCount = dict(dept)("Employees").Count
        
        ' 合計値に加算
        totalOvertimeAll = totalOvertimeAll + totalOvertime
        totalOccurrencesAll = totalOccurrencesAll + occurrenceCount
        totalHolidayWorkAll = totalHolidayWorkAll + holidayWorkCount
        
        ' 全社員リストに追加
        For Each emp In dict(dept)("Employees").Keys
            If Not allEmployees.Exists(emp) Then
                allEmployees.Add emp, dict(dept)("Employees")(emp)
            End If
        Next emp
        
        ' 部門名
        wsSummary.Cells(summaryRow, 1).Value = dict(dept)("DepartmentName")
        
        ' 合計残業時間
        wsSummary.Cells(summaryRow, 2).Value = MinutesToTime(totalOvertime)
        
        ' 平均残業時間/回
        If occurrenceCount > 0 Then
            wsSummary.Cells(summaryRow, 3).Value = MinutesToTime(totalOvertime / occurrenceCount)
        Else
            wsSummary.Cells(summaryRow, 3).Value = "0:00"
        End If
        
        ' 平均残業時間/人
        If personCount > 0 Then
            wsSummary.Cells(summaryRow, 4).Value = MinutesToTime(totalOvertime / personCount)
        Else
            wsSummary.Cells(summaryRow, 4).Value = "0:00"
        End If
        
        ' 休日出勤回数
        wsSummary.Cells(summaryRow, 5).Value = holidayWorkCount
        
        ' 人数
        wsSummary.Cells(summaryRow, 6).Value = personCount
        
        summaryRow = summaryRow + 1
    Next dept
    
    ' 全社の合計行
    totalPersonsAll = allEmployees.Count
    
    wsSummary.Cells(summaryRow, 1).Value = "合計"
    wsSummary.Cells(summaryRow, 1).Font.Bold = True
    
    ' 合計残業時間
    wsSummary.Cells(summaryRow, 2).Value = MinutesToTime(totalOvertimeAll)
    
    ' 全体平均残業時間/回
    If totalOccurrencesAll > 0 Then
        wsSummary.Cells(summaryRow, 3).Value = MinutesToTime(totalOvertimeAll / totalOccurrencesAll)
    Else
        wsSummary.Cells(summaryRow, 3).Value = "0:00"
    End If
    
    ' 全体平均残業時間/人
    If totalPersonsAll > 0 Then
        wsSummary.Cells(summaryRow, 4).Value = MinutesToTime(totalOvertimeAll / totalPersonsAll)
    Else
        wsSummary.Cells(summaryRow, 4).Value = "0:00"
    End If
    
    ' 全体休日出勤回数
    wsSummary.Cells(summaryRow, 5).Value = totalHolidayWorkAll
    
    ' 全体人数
    wsSummary.Cells(summaryRow, 6).Value = totalPersonsAll
    
    ' 合計行の書式設定
    With wsSummary.Range(wsSummary.Cells(summaryRow, 1), wsSummary.Cells(summaryRow, 6))
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Interior.Color = RGB(240, 240, 240)
    End With
    
    ' 結果の書式設定
    With wsSummary.Range("A" & headerRow + 1 & ":F" & summaryRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
    End With
    
    ' 列幅の自動調整
    wsSummary.Columns("A:F").AutoFit
' =====================================================
    ' 定時退社率計算機能を呼び出し（2025/08/20追加）
    ' =====================================================
    On Error Resume Next  ' エラーが発生しても処理を継続
    
    Debug.Print "定時退社率計算を開始します..."
    
    ' 直接呼び出し（モジュール名なしで）
    Application.Run "CalculateAndOutputRate"
    
    ' エラーが発生した場合の処理
    If Err.Number <> 0 Then
        Debug.Print "定時退社率計算でエラー: " & Err.Description
        Err.Clear
    Else
        Debug.Print "定時退社率計算が正常に完了しました"
    End If
    On Error GoTo 0
    
    MsgBox "部門別残業時間の集計が完了しました。" & vbCrLf & _
    "休憩時間・残業時間チェック結果を表示します。", vbInformation, "集計完了"
End Sub



