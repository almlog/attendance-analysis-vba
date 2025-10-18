' ========================================
' Module8
' タイプ: 標準モジュール
' 行数: 479
' エクスポート日時: 2025-10-18 23:37:17
' ========================================

Option Explicit

' *************************************************************
' モジュール：定時退社率分析
' 目的：部門ごとの定時退社率を自動計算し、分析結果シートに出力する
' Copyright (c) 2025 SI1 shunpei.suzuki
' 作成日：2025年8月20日
'
' 改版履歴：
' 2025/08/20 初版作成_v1.0
' *************************************************************

' =====================================================
' 定数定義
' =====================================================
' シート名
Private Const SHEET_NAME_ANALYSIS_RESULT As String = "勤怠情報分析結果"

' 列名
Private Const COL_NAME_DEPT As String = "部門"
Private Const COL_NAME_WORK_TIME As String = "実働時間"
Private Const COL_NAME_LEAVE_TIME As String = "退社"
Private Const COL_NAME_DELIVERY As String = "届出内容"

' 判定基準（分単位）
Private Const MIN_WORK_MINUTES As Double = 60          ' 総勤務日の最小実働時間（1時間）
Private Const ON_TIME_LEAVE_MINUTES As Double = 1065   ' 定時退社の退社時刻基準（17:45）

' =====================================================
' メインプロシージャ：定時退社率の計算と出力
' =====================================================
Public Sub CalculateAndOutputRate()
    On Error GoTo ErrorHandler
    
    ' デバッグ用メッセージ追加
    Debug.Print "CalculateAndOutputRate: 開始"
    
    ' 画面更新を停止して処理を高速化
    Application.ScreenUpdating = False
    Application.StatusBar = "定時退社率を計算しています..."
    
    ' シートオブジェクトの取得
    Dim wsCSVData As Worksheet  ' ← CSVデータシートを使用
    Dim wsAnalysisResult As Worksheet
    
    ' シートの存在確認と設定
    On Error Resume Next
    Set wsCSVData = ThisWorkbook.Worksheets("CSVデータ")  ' ← 変更
    Set wsAnalysisResult = ThisWorkbook.Worksheets(SHEET_NAME_ANALYSIS_RESULT)
    On Error GoTo ErrorHandler
    
    ' シートの存在チェック
    If wsCSVData Is Nothing Then
        MsgBox "「CSVデータ」シートが見つかりません。" & vbCrLf & _
               "CSVファイルを先に読み込んでください。", vbExclamation, "エラー"
        GoTo CleanExit
    End If
    
    If wsAnalysisResult Is Nothing Then
        MsgBox "「" & SHEET_NAME_ANALYSIS_RESULT & "」シートが見つかりません。", vbExclamation, "エラー"
        GoTo CleanExit
    End If
    
    ' 部門ごとの集計データを取得（CSVデータシートから）
    Dim deptData As Object
    Set deptData = GetOnTimeDepartureData(wsCSVData)  ' ← CSVデータシートを渡す
    
    ' データが空の場合のチェック
    If deptData Is Nothing Or deptData.Count = 0 Then
        MsgBox "集計可能なデータが見つかりませんでした。", vbInformation, "情報"
        GoTo CleanExit
    End If
    
    ' 結果をシートに書き出し
    Call WriteResultsToSheet(deptData, wsAnalysisResult)
    
    ' 処理完了メッセージ（サイレント実行のためコメントアウト）
    ' MsgBox "定時退社率の計算が完了しました。", vbInformation, "完了"
    Debug.Print "CalculateAndOutputRate: 完了"
    
CleanExit:
    ' 画面更新を再開
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Debug.Print "CalculateAndOutputRate エラー: " & Err.Description
    ' MsgBox "エラーが発生しました。" & vbCrLf & _
    '        "エラー内容: " & Err.Description & vbCrLf & _
    '        "エラー番号: " & Err.Number, vbCritical, "エラー"
    Resume CleanExit
End Sub

' =====================================================
' 計算用関数：部門ごとの定時退社データを取得
' =====================================================
Private Function GetOnTimeDepartureData(ByVal sourceSheet As Worksheet) As Object
    On Error GoTo ErrorHandler
    
    ' Dictionaryオブジェクトの作成
    Dim deptDict As Object
    Set deptDict = CreateObject("Scripting.Dictionary")
    
    ' ヘッダー行から必要な列インデックスを取得
    Dim colDept As Long, colWorkTime As Long, colLeaveTime As Long
    Dim colDelivery As Long
    Dim lastCol As Long
    Dim i As Long
    
    lastCol = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column
    
    ' 列名から列番号を特定
    For i = 1 To lastCol
        Select Case sourceSheet.Cells(1, i).Value
            Case COL_NAME_DEPT
                colDept = i
            Case COL_NAME_WORK_TIME
                colWorkTime = i
            Case COL_NAME_LEAVE_TIME
                colLeaveTime = i
            Case COL_NAME_DELIVERY
                colDelivery = i
        End Select
    Next i
    
    ' 必要な列が見つからない場合のエラー処理
    If colDept = 0 Or colWorkTime = 0 Or colLeaveTime = 0 Then
        MsgBox "必要な列（部門、実働時間、退社）が見つかりません。", _
               vbExclamation, "エラー"
        Set GetOnTimeDepartureData = Nothing
        Exit Function
    End If
    
    ' データ範囲を取得
    Dim lastRow As Long
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, colDept).End(xlUp).Row
    
    If lastRow <= 1 Then
        Set GetOnTimeDepartureData = Nothing
        Exit Function
    End If
    
    ' データを配列に一括読み込み
    Dim dataArray As Variant
    dataArray = sourceSheet.Range(sourceSheet.Cells(2, 1), _
                                  sourceSheet.Cells(lastRow, lastCol)).Value
    
    ' 配列をループして集計
    Dim rowIndex As Long
    Dim deptName As String
    Dim workTimeStr As String
    Dim leaveTimeStr As String
    Dim deliveryContent As String
    Dim workMinutes As Double
    Dim leaveMinutes As Double
    Dim deptStats As Variant
    Dim isOnTimeDeparture As Boolean
    Dim isWorkDay As Boolean
    
    ' デバッグ用カウンター
    Dim debugOnTimeCount As Long
    Dim debugWorkDayCount As Long
    debugOnTimeCount = 0
    debugWorkDayCount = 0
    
    For rowIndex = 1 To UBound(dataArray, 1)
        ' 部門名の取得
        deptName = Trim(CStr(dataArray(rowIndex, colDept)))
        
        ' 空の部門名はスキップ
        If deptName <> "" Then
            ' 実働時間の取得と変換
            workTimeStr = CStr(dataArray(rowIndex, colWorkTime))
            workMinutes = Round(ConvertTimeToMinutesLocal(workTimeStr), 0)
            
            ' 退社時刻の取得
            leaveTimeStr = CStr(dataArray(rowIndex, colLeaveTime))
            leaveMinutes = ConvertTimeToMinutesLocal(leaveTimeStr)
            
            ' 届出内容の取得（届出列がない場合も考慮）
            If colDelivery > 0 Then
                deliveryContent = Trim(CStr(dataArray(rowIndex, colDelivery)))
            Else
                deliveryContent = ""
            End If
            
            ' 部門が初めて出現した場合は初期化
            If Not deptDict.Exists(deptName) Then
                deptDict.Add deptName, Array(0, 0)
            End If
            
            ' 現在の統計値を取得
            deptStats = deptDict(deptName)
            
            ' ===== 総勤務日の判定 =====
            isWorkDay = False
            
            ' 振替休暇は勤務日としてカウントしない
            If deliveryContent = "振替休暇" Then
                isWorkDay = False
                
            ' 休暇系の届出（実働0でも勤務日とする）
            ElseIf deliveryContent = "有休" Or _
                   deliveryContent = "午前有休" Or _
                   deliveryContent = "午後有休" Or _
                   deliveryContent = "時間有休" Or _
                   deliveryContent = "子の看護休暇" Or _
                   deliveryContent = "生理休暇" Or _
                   deliveryContent = "特別休暇" Then
                isWorkDay = True
                
            ' その他の届出（欠勤、遅刻、早退、休日出勤、振替出勤、電車遅延、休憩修正）
            ElseIf deliveryContent = "欠勤" Or _
                   deliveryContent = "遅刻" Or _
                   deliveryContent = "早退" Or _
                   deliveryContent = "休日出勤" Or _
                   deliveryContent = "振替出勤" Or _
                   deliveryContent = "電車遅延" Or _
                   deliveryContent = "休憩修正" Then
                ' 実働がある場合のみ勤務日
                If workMinutes >= 60 Then
                    isWorkDay = True
                End If
                
            ' 届出なしで実働がある場合
            ElseIf deliveryContent = "" And workMinutes >= 60 Then
                isWorkDay = True
            End If
            
            If isWorkDay Then
                deptStats(0) = deptStats(0) + 1  ' 総勤務日数を増加
                debugWorkDayCount = debugWorkDayCount + 1
                
                ' ===== 定時退社の判定 =====
                isOnTimeDeparture = False
                
                ' 定時退社から除外する届出
                If deliveryContent = "遅刻" Or _
                   deliveryContent = "早退" Or _
                   deliveryContent = "欠勤" Or _
                   deliveryContent = "休日出勤" Then
                    isOnTimeDeparture = False
                    
                ' 休暇系の届出は全て定時退社
                ElseIf deliveryContent = "有休" Or _
                       deliveryContent = "午前有休" Or _
                       deliveryContent = "午後有休" Or _
                       deliveryContent = "時間有休" Or _
                       deliveryContent = "子の看護休暇" Or _
                       deliveryContent = "生理休暇" Or _
                       deliveryContent = "特別休暇" Then
                    isOnTimeDeparture = True
                    Debug.Print "休暇で定時退社: " & deptName & " - " & deliveryContent
                    
                ' その他の届出または届出なしの場合、退社時刻で判定
                Else
                    ' 退社時刻が17:45より前（17:45 = 1065分）
                    If leaveMinutes > 0 And leaveMinutes < 1065 Then
                        isOnTimeDeparture = True
                        Debug.Print "17:45前退社で定時退社: " & deptName & " - 届出:" & deliveryContent & " 退社:" & leaveTimeStr
                    End If
                End If
                
                If isOnTimeDeparture Then
                    deptStats(1) = deptStats(1) + 1  ' 定時退社日数を増加
                    debugOnTimeCount = debugOnTimeCount + 1
                End If
            End If
            
            ' 更新した値を辞書に戻す
            deptDict(deptName) = deptStats
        End If
    Next rowIndex
    
    Debug.Print "総勤務日数: " & debugWorkDayCount
    Debug.Print "定時退社総数: " & debugOnTimeCount
    
    ' 結果を返す
    Set GetOnTimeDepartureData = deptDict
    Exit Function
    
ErrorHandler:
    MsgBox "データ集計中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
    Set GetOnTimeDepartureData = Nothing
End Function
' =====================================================
' 書き出し用プロシージャ：計算結果を既存の部門別残業集計表に追加
' =====================================================
Private Sub WriteResultsToSheet(ByVal results As Object, ByVal targetSheet As Worksheet)
    On Error GoTo ErrorHandler
    
    Debug.Print "WriteResultsToSheet: 開始"
    Debug.Print "結果件数: " & results.Count
    
    ' 部門別残業集計表を探す
    Dim headerRow As Long
    Dim i As Long, j As Long
    headerRow = 0
    Debug.Print "ヘッダー行: " & headerRow
    ' 「部署」というヘッダーを探す（最大30行まで検索）
    For i = 1 To 30
        If targetSheet.Cells(i, 1).Value = "部署" Then
            headerRow = i
            Exit For
        End If
    Next i
    
    ' ヘッダーが見つからない場合はエラー
    If headerRow = 0 Then
        MsgBox "部門別残業集計表が見つかりません。", vbExclamation, "エラー"
        Exit Sub
    End If
    
    ' 既存のヘッダーの最後に新しい列を追加（7列目から）
    targetSheet.Cells(headerRow, 7).Value = "総勤務日数"
    targetSheet.Cells(headerRow, 8).Value = "定時退社日数"
    targetSheet.Cells(headerRow, 9).Value = "定時退社率"
    
    ' 新しいヘッダーの書式設定（罫線なし、背景色と太字のみ）
    With targetSheet.Range(targetSheet.Cells(headerRow, 7), targetSheet.Cells(headerRow, 9))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(200, 200, 200)
        ' 罫線は設定しない
    End With
    
    ' 各部門のデータを追加
    Dim currentRow As Long
    Dim deptName As String
    Dim deptStats As Variant
    Dim totalWorkDays As Long
    Dim onTimeDays As Long
    Dim onTimeRate As Double
    
    ' 全社合計用の変数
    Dim totalWorkDaysAll As Long
    Dim totalOnTimeDaysAll As Long
    totalWorkDaysAll = 0
    totalOnTimeDaysAll = 0
    
    currentRow = headerRow + 1
    
    ' 部署列の値を読み取って、対応するデータを追加
    Do While targetSheet.Cells(currentRow, 1).Value <> ""
        deptName = Trim(targetSheet.Cells(currentRow, 1).Value)
        
    ' 「合計」行の場合
    If deptName = "合計" Then
        ' 全社合計を計算して出力
        If totalWorkDaysAll > 0 Then
            onTimeRate = (totalOnTimeDaysAll / totalWorkDaysAll) * 100
        Else
            onTimeRate = 0
        End If
        
        targetSheet.Cells(currentRow, 7).Value = totalWorkDaysAll
        targetSheet.Cells(currentRow, 8).Value = totalOnTimeDaysAll
        targetSheet.Cells(currentRow, 9).Value = Format(onTimeRate, "0.0") & "%"
        
        ' 合計行の書式設定（中央揃えを追加）
        With targetSheet.Range(targetSheet.Cells(currentRow, 7), targetSheet.Cells(currentRow, 9))
            .Font.Bold = True
            .Interior.Color = RGB(240, 240, 240)
            .HorizontalAlignment = xlCenter  ' 中央揃え
        End With
        Exit Do
    End If
        
        ' 通常の部門データ
        If results.Exists(deptName) Then
            deptStats = results(deptName)
            totalWorkDays = deptStats(0)
            onTimeDays = deptStats(1)
            
            ' 全社合計に加算
            totalWorkDaysAll = totalWorkDaysAll + totalWorkDays
            totalOnTimeDaysAll = totalOnTimeDaysAll + onTimeDays
            
            ' 定時退社率の計算
            If totalWorkDays > 0 Then
                onTimeRate = (onTimeDays / totalWorkDays) * 100
            Else
                onTimeRate = 0
            End If
            
            ' データの書き込み
            targetSheet.Cells(currentRow, 7).Value = totalWorkDays
            targetSheet.Cells(currentRow, 8).Value = onTimeDays
            targetSheet.Cells(currentRow, 9).Value = Format(onTimeRate, "0.0") & "%"
        Else
            ' データがない場合は0を表示
            targetSheet.Cells(currentRow, 7).Value = 0
            targetSheet.Cells(currentRow, 8).Value = 0
            targetSheet.Cells(currentRow, 9).Value = "0.0%"
        End If
        
        currentRow = currentRow + 1
    Loop
    
    ' データ範囲全体の罫線設定
    ' With targetSheet.Range(targetSheet.Cells(headerRow + 1, 7), _
    '                       targetSheet.Cells(currentRow - 1, 9))
    '     .Borders.LineStyle = xlContinuous
    '     .Borders.Weight = xlThin
    '     .HorizontalAlignment = xlCenter
    ' End With
    
    ' データ範囲のみに罫線設定（ヘッダーは含まない）
    With targetSheet.Range(targetSheet.Cells(headerRow + 1, 7), _
                          targetSheet.Cells(currentRow, 9))  ' currentRowまで（合計行含む）
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
    End With
    
    ' 列幅の自動調整
    targetSheet.Columns("G:I").AutoFit
    
    Exit Sub
    
ErrorHandler:
    MsgBox "結果の出力中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub
' =====================================================
' ローカル時間変換関数（Module1の関数が使えない場合の代替）
' =====================================================
Private Function ConvertTimeToMinutesLocal(timeStr As Variant) As Double
    On Error GoTo ErrorHandler
    
    ' デバッグ用
    Dim originalValue As String
    originalValue = CStr(timeStr)
    
    ' 空文字またはNullの場合
    If IsEmpty(timeStr) Or timeStr = "" Then
        ConvertTimeToMinutesLocal = 0
        Exit Function
    End If
    
    ' 数値形式（Excelの時間値）の場合
    If IsNumeric(timeStr) Then
        ' Excelの時間は日の割合で格納されている
        ConvertTimeToMinutesLocal = CDbl(timeStr) * 24 * 60
        ' デバッグ出力
        If ConvertTimeToMinutesLocal >= 450 And ConvertTimeToMinutesLocal <= 470 Then
            Debug.Print "時間変換(数値): " & originalValue & " → " & ConvertTimeToMinutesLocal & "分"
        End If
        Exit Function
    End If
    
    ' HH:MM形式の文字列の場合
    Dim timeParts As Variant
    timeParts = Split(CStr(timeStr), ":")
    
    If UBound(timeParts) >= 1 Then
        If IsNumeric(timeParts(0)) And IsNumeric(timeParts(1)) Then
            Dim hours As Double, minutes As Double
            hours = CDbl(timeParts(0))
            minutes = CDbl(timeParts(1))
            ConvertTimeToMinutesLocal = hours * 60 + minutes
            ' デバッグ出力
            If ConvertTimeToMinutesLocal >= 450 And ConvertTimeToMinutesLocal <= 470 Then
                Debug.Print "時間変換(文字列): " & originalValue & " → " & ConvertTimeToMinutesLocal & "分"
            End If
        Else
            ConvertTimeToMinutesLocal = 0
        End If
    Else
        ConvertTimeToMinutesLocal = 0
    End If
    
    Exit Function
    
ErrorHandler:
    ConvertTimeToMinutesLocal = 0
End Function

