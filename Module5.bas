' ========================================
' Module5
' タイプ: 標準モジュール
' 行数: 671
' エクスポート日時: 2025-10-20 14:30:49
' ========================================

Option Explicit

' *************************************************************
' モジュール：勤怠入力漏れ検出
' 目的：勤怠入力漏れを検出する関数群
' Copyright (c) 2025 SI1 shunpei.suzuki
' 作成日：2025年4月2日
'
' 改版履歴：
' 2025/04/02 module2から分割作成
' 2025/04/04 お昼休憩時間の矛盾チェック機能追加
' 2025/04/05 お昼休憩時間（退勤）の矛盾チェックロジック修正 - 12:00は許容、12:01～12:59のみ矛盾として検出
' *************************************************************

' 定数定義（module2_coreと同じ定数を定義）
Private Const COL_EMPLOYEE_ID As Integer = 1
Private Const COL_EMPLOYEE_NAME As Integer = 2
Private Const COL_DATE As Integer = 3
Private Const COL_DAY_TYPE As Integer = 4
Private Const COL_LEAVE_TYPE As Integer = 5
Private Const COL_MISSING_ENTRY_TYPE As Integer = 6
Private Const COL_COMMENT As Integer = 7
Private Const COL_ATTENDANCE_TIME As Integer = 8 ' 出勤時刻列を追加
Private Const COL_DEPARTURE_TIME As Integer = 9 ' 退勤時刻列を追加
Private Const COL_CONTRADICTION_TYPE As Integer = 10 ' 矛盾種別列を追加
Private Const DEBUG_MODE As Boolean = False ' デバッグモード設定 - 通常運用時はFalse

' グローバル変数の参照（module2_coreで定義されているものを参照）
' Public g_IncludeToday As Boolean

' 勤怠入力漏れを検出して出力する - 最適化版
Public Sub DetectMissingEntries(wsCSVData As Worksheet, outputSheet As Worksheet)
    ' 当日分を含めるかどうかのオプションを取得
    Dim includeToday As Boolean
    includeToday = g_IncludeToday ' グローバル変数から取得
    
    ' デバッグ情報
    If DEBUG_MODE Then
        Debug.Print "当日分を含める設定: " & includeToday
    End If
    
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "勤怠入力漏れを検出しています..."
    
    ' 最終行を取得
    Dim lastRow As Long
    Dim i As Long, j As Long
    lastRow = wsCSVData.Cells(wsCSVData.Rows.Count, "A").End(xlUp).Row
    
    If lastRow <= 1 Then
        MsgBox "CSVデータが存在しません。", vbExclamation
        Exit Sub
    End If
    
    ' 各列のインデックスを特定
    Dim 社員番号Col As Integer, 氏名Col As Integer, 部門Col As Integer
    Dim 日付Col As Integer, カレンダーCol As Integer, 曜日Col As Integer
    Dim 届出Col As Integer, 状況区分Col As Integer
    Dim 出勤時刻Col As Integer, 退勤時刻Col As Integer, 備考Col As Integer
    
    社員番号Col = 0: 氏名Col = 0: 部門Col = 0: 日付Col = 0
    カレンダーCol = 0: 曜日Col = 0: 届出Col = 0: 状況区分Col = 0
    出勤時刻Col = 0: 退勤時刻Col = 0: 備考Col = 0
    
    ' 列インデックスの特定 - 一度に取得して高速化
    Dim headerRow As Range
    Set headerRow = wsCSVData.Range(wsCSVData.Cells(1, 1), wsCSVData.Cells(1, wsCSVData.Cells(1, wsCSVData.Columns.Count).End(xlToLeft).Column))
    
    For i = 1 To headerRow.Columns.Count
        Select Case headerRow.Cells(1, i).Value
            Case "社員番号": 社員番号Col = i
            Case "氏名": 氏名Col = i
            Case "部門": 部門Col = i
            Case "日付": 日付Col = i
            Case "カレンダー": カレンダーCol = i
            Case "曜日": 曜日Col = i
            Case "届出内容": 届出Col = i
            Case "状況区分": 状況区分Col = i
            Case "出社": 出勤時刻Col = i
            Case "退社": 退勤時刻Col = i
            Case "備考": 備考Col = i
        End Select
    Next i
    
    ' 必要な列が存在するか確認
    If 社員番号Col = 0 Or 氏名Col = 0 Or 日付Col = 0 Then
        MsgBox "必要な列（社員番号、氏名、日付）が見つかりませんでした。", vbExclamation
        Exit Sub
    End If
    
    ' 出勤・退勤時刻列が見つからない場合のデフォルト値
    If 出勤時刻Col = 0 Then 出勤時刻Col = 10
    If 退勤時刻Col = 0 Then 退勤時刻Col = 11
    
    ' 出力行カウンター
    Dim outputRow As Long
    outputRow = 2 ' ヘッダー行の次から始める
    
    ' 入力漏れタイプのカウント
    Dim missingAttendanceCount As Long, missingDepartureCount As Long
    Dim missingBothCount As Long, totalMissingCount As Long
    
    missingAttendanceCount = 0: missingDepartureCount = 0
    missingBothCount = 0: totalMissingCount = 0
    
    ' 対象従業員辞書 - 社員番号をキーに
    Dim employeeDict As Object
    Set employeeDict = CreateObject("Scripting.Dictionary")
    employeeDict.CompareMode = vbTextCompare ' 大文字小文字を区別しない
    
    ' 除外社員番号を取得
    Dim excludeIDs As Variant
    excludeIDs = 除外社員番号取得()
    
    ' 高速化のため除外IDを辞書に変換
    Dim excludeDict As Object
    Set excludeDict = CreateObject("Scripting.Dictionary")
    excludeDict.CompareMode = vbTextCompare ' 大文字小文字を区別しない

    For j = LBound(excludeIDs) To UBound(excludeIDs)
        If excludeIDs(j) <> "" Then
            ' 文字列型に明示的に変換
            Dim excludeKey As String
            excludeKey = Trim(CStr(excludeIDs(j)))
            ' 重複チェック
            If Not excludeDict.Exists(excludeKey) Then
                excludeDict.Add excludeKey, True
                Debug.Print "除外辞書に追加: [" & excludeKey & "]"
            End If
        End If
    Next j
    
    ' データをバッファに取得して高速化
    Dim dataRange As Range
    Set dataRange = wsCSVData.Range(wsCSVData.Cells(2, 1), wsCSVData.Cells(lastRow, wsCSVData.Cells(1, wsCSVData.Columns.Count).End(xlToLeft).Column))
    Dim dataArray As Variant
    dataArray = dataRange.Value
    
    ' 今日の日付
    Dim todayDate As Date
    todayDate = Date
    
    ' 各行を処理
    Dim isExcluded As Boolean ' ← ここで isExcluded 変数を宣言

    For i = 1 To UBound(dataArray, 1)
        ' CSVデータから必要な情報を取得
        Dim employeeID As String, employeeName As String
        Dim entryDate As Date, dayType As String
        Dim hasAttendanceTime As Boolean, hasDepartureTime As Boolean
        Dim deliveryContent As String ' 届出内容
        
        ' 値の取得 - 配列から直接取得して高速化
        employeeID = Trim(CStr(dataArray(i, 社員番号Col)))

        ' 除外社員番号のチェック - 厳密な文字列比較
        isExcluded = (employeeID <> "" And excludeDict.Exists(employeeID))

        ' デバッグ情報
        If DEBUG_MODE Then
            Debug.Print "社員番号チェック: [" & employeeID & "] 除外判定: " & isExcluded
        End If

        ' 除外社員の場合はスキップ
        If isExcluded Then
            If DEBUG_MODE Then Debug.Print "==> 除外社員のためスキップします: " & employeeID
            GoTo NextRow
        End If
        
        employeeName = CStr(dataArray(i, 氏名Col))
        
        ' 日付の変換確認
        If IsDate(dataArray(i, 日付Col)) Then
            entryDate = CDate(dataArray(i, 日付Col))
        Else
            ' 日付が不正な場合はスキップ
            GoTo NextRow
        End If
        
        ' 曜日区分の取得
        If 曜日Col > 0 Then
            dayType = CStr(dataArray(i, 曜日Col))
        Else
            dayType = "不明"
        End If
        
        ' カレンダー種別の取得（主な判定基準）
        Dim calendarType As String
        calendarType = ""
        If カレンダーCol > 0 Then
            calendarType = CStr(dataArray(i, カレンダーCol))
        End If
        
        ' 届出内容の取得
        deliveryContent = ""
        If 届出Col > 0 Then
            If Not IsEmpty(dataArray(i, 届出Col)) Then
                deliveryContent = Trim(CStr(dataArray(i, 届出Col)))
            End If
        End If
        
        ' 出勤・退勤時刻の有無を確認
        hasAttendanceTime = Not IsEmpty(dataArray(i, 出勤時刻Col)) And _
                            Trim(CStr(dataArray(i, 出勤時刻Col))) <> ""
        
        hasDepartureTime = Not IsEmpty(dataArray(i, 退勤時刻Col)) And _
                            Trim(CStr(dataArray(i, 退勤時刻Col))) <> ""
        
        ' 入力が必要かどうかを判断
        ' 当日分を含めるかどうかの設定に基づいて条件を変更
        If (DateDiff("d", entryDate, todayDate) > 0 Or (includeToday And DateDiff("d", entryDate, todayDate) = 0)) Then
            ' 入力漏れの種類を判断
            Dim missingEntryType As String
            Dim comment As String
            Dim contradictionType As String ' 矛盾種別
            Dim attendanceTime As String ' 出勤時刻
            Dim departureTime As String ' 退勤時刻
            
            ' 出勤・退勤時刻を取得
            attendanceTime = ""
            departureTime = ""
            If hasAttendanceTime Then
                attendanceTime = Trim(CStr(dataArray(i, 出勤時刻Col)))
            End If
            If hasDepartureTime Then
                departureTime = Trim(CStr(dataArray(i, 退勤時刻Col)))
            End If
            
            ' 矛盾チェック
            contradictionType = ""
            
            ' デバッグ情報
            If DEBUG_MODE Then
                Debug.Print "レコード: " & entryDate & ", 届出: " & deliveryContent & ", 出勤時刻: " & attendanceTime & ", 退勤時刻: " & departureTime
            End If
            
            ' 午前有休の場合、出勤時刻が13時より前であれば矛盾
            If Trim(deliveryContent) = "午前有休" And hasAttendanceTime Then
                Dim attendanceHour As Integer
                attendanceHour = GetHourFromTimeString(attendanceTime)
                
                ' デバッグ情報
                If DEBUG_MODE Then
                    Debug.Print "  午前有休チェック - 出勤時刻: " & attendanceTime & ", 解析結果: " & attendanceHour & "時"
                End If
                
                ' 出勤時刻が13時より前の場合のみ矛盾として検出
                If attendanceHour < 13 Then
                    ' 数値形式の場合は表示用に変換
                    Dim displayTime As String
                    If IsNumeric(attendanceTime) Then
                        displayTime = Format(CDbl(attendanceTime), "h:mm")
                    Else
                        displayTime = attendanceTime
                    End If
                    
                    contradictionType = "1"  ' 午前有休矛盾
                    comment = "午前有休なのに出勤時刻が13時より前（" & displayTime & "）になっています"
                End If
            End If
            
            ' 午後有休の場合、退勤時刻が12時より後であれば矛盾
            If Trim(deliveryContent) = "午後有休" And hasDepartureTime Then
                Dim departureHour As Integer
                Dim departureMinute As Integer
                departureHour = GetHourFromTimeString(departureTime)
                departureMinute = GetMinuteFromTimeString(departureTime)
                
                ' デバッグ情報
                If DEBUG_MODE Then
                    Debug.Print "  午後有休チェック - 退勤時刻: " & departureTime & ", 解析結果: " & departureHour & "時" & departureMinute & "分"
                End If
                
                ' 退勤時刻が12時より後の場合のみ矛盾として検出（12:00は許容）
                If departureHour > 12 Or (departureHour = 12 And departureMinute > 0) Then
                    ' 数値形式の場合は表示用に変換
                    Dim displayDepartureTime As String
                    If IsNumeric(departureTime) Then
                        displayDepartureTime = Format(CDbl(departureTime), "h:mm")
                    Else
                        displayDepartureTime = departureTime
                    End If
                    
                    contradictionType = "2"  ' 午後有休矛盾
                    comment = "午後有休なのに退勤時刻が12時より後（" & displayDepartureTime & "）になっています"
                End If
            End If

            ' お昼休憩時間（12:00～12:59）に業務開始/終了している場合の矛盾チェック
            If contradictionType = "" And (hasAttendanceTime Or hasDepartureTime) Then
                ' 出勤時刻が12時台の場合
                If hasAttendanceTime Then
                    Dim attendanceHourLunch As Integer
                    attendanceHourLunch = GetHourFromTimeString(attendanceTime)

                    ' デバッグ情報
                    If DEBUG_MODE Then
                        Debug.Print "  お昼休憩チェック(出勤) - 出勤時刻: " & attendanceTime & ", 解析結果: " & attendanceHourLunch & "時"
                    End If

                    ' 出勤時刻が12時台の場合
                    If attendanceHourLunch = 12 Then
                        ' 数値形式の場合は表示用に変換
                        Dim displayLunchAttendTime As String
                        If IsNumeric(attendanceTime) Then
                            displayLunchAttendTime = Format(CDbl(attendanceTime), "h:mm")
                        Else
                            displayLunchAttendTime = attendanceTime
                        End If

                        contradictionType = "お昼休憩矛盾"
                        comment = "お昼休憩時間(12:00～12:59)に出勤（" & displayLunchAttendTime & "）しています"
                    End If
                End If

                ' 退勤時刻が12時台の場合（既に矛盾がない場合のみチェック）
                If contradictionType = "" And hasDepartureTime Then
                    Dim departureHourLunch As Integer
                    Dim departureMinuteLunch As Integer
                    departureHourLunch = GetHourFromTimeString(departureTime)
                    departureMinuteLunch = GetMinuteFromTimeString(departureTime)

                    ' デバッグ情報
                    If DEBUG_MODE Then
                        Debug.Print "  お昼休憩チェック(退勤) - 退勤時刻: " & departureTime & ", 解析結果: " & departureHourLunch & "時" & departureMinuteLunch & "分"
                    End If

                    ' 退勤時刻が12時台かつ00分でない場合（12:01～12:59）のみ矛盾として検出
                    If departureHourLunch = 12 And departureMinuteLunch > 0 Then
                        ' 数値形式の場合は表示用に変換
                        Dim displayLunchDepartTime As String
                        If IsNumeric(departureTime) Then
                            displayLunchDepartTime = Format(CDbl(departureTime), "h:mm")
                        Else
                            displayLunchDepartTime = departureTime
                        End If

                        contradictionType = "お昼休憩矛盾"
                        comment = "お昼休憩時間(12:01～12:59)に退勤（" & displayLunchDepartTime & "）しています"
                    End If
                End If
            End If

            ' 矛盾がある場合は出力
            If contradictionType <> "" Then
                ' 社員情報を辞書に追加（重複なし）
                If Not employeeDict.Exists(employeeID) Then
                    employeeDict.Add employeeID, employeeName
                End If
                
                ' 結果をシートに出力
                With outputSheet
                    .Cells(outputRow, COL_EMPLOYEE_ID).Value = employeeID
                    .Cells(outputRow, COL_EMPLOYEE_NAME).Value = employeeName
                    .Cells(outputRow, COL_DATE).Value = entryDate
                    .Cells(outputRow, COL_DAY_TYPE).Value = dayType
                    .Cells(outputRow, COL_LEAVE_TYPE).Value = deliveryContent
                    .Cells(outputRow, COL_MISSING_ENTRY_TYPE).Value = ""
                    .Cells(outputRow, COL_COMMENT).Value = comment
                    
                    ' 時間形式を設定
                    .Cells(outputRow, COL_ATTENDANCE_TIME).Value = attendanceTime
                    .Cells(outputRow, COL_ATTENDANCE_TIME).NumberFormat = "h:mm"
                    .Cells(outputRow, COL_DEPARTURE_TIME).Value = departureTime
                    .Cells(outputRow, COL_DEPARTURE_TIME).NumberFormat = "h:mm"
                    
                    ' 矛盾の行を強調表示（矛盾種別列には値を設定しない）
                    .Range(.Cells(outputRow, 1), .Cells(outputRow, COL_DEPARTURE_TIME)).Interior.Color = RGB(255, 200, 200)
                End With
                
                outputRow = outputRow + 1
                totalMissingCount = totalMissingCount + 1
            End If
            
            ' 入力漏れチェック（矛盾がない場合も含めて）
            If IsEntryRequired(calendarType, deliveryContent) Then
                If Not hasAttendanceTime And Not hasDepartureTime Then
                    missingEntryType = "3"  ' 出退勤時刻なし
                    comment = "出勤時刻と退勤時刻の両方が入力されていません"
                    missingBothCount = missingBothCount + 1
                ElseIf Not hasAttendanceTime Then
                    missingEntryType = "1"  ' 出勤時刻なし
                    comment = "出勤時刻が入力されていません"
                    missingAttendanceCount = missingAttendanceCount + 1
                ElseIf Not hasDepartureTime Then
                    missingEntryType = "2"  ' 退勤時刻なし
                    comment = "退勤時刻が入力されていません"
                    missingDepartureCount = missingDepartureCount + 1
                Else
                    ' 入力漏れなし
                    missingEntryType = ""
                    comment = ""
                End If
                
                ' 入力漏れがある場合のみ出力
                If missingEntryType <> "" Then
                    ' 社員情報を辞書に追加（重複なし）
                    If Not employeeDict.Exists(employeeID) Then
                        employeeDict.Add employeeID, employeeName
                    End If
                    
                    ' 結果をシートに出力
                    With outputSheet
                        .Cells(outputRow, COL_EMPLOYEE_ID).Value = employeeID
                        .Cells(outputRow, COL_EMPLOYEE_NAME).Value = employeeName
                        .Cells(outputRow, COL_DATE).Value = entryDate
                        .Cells(outputRow, COL_DAY_TYPE).Value = dayType
                        .Cells(outputRow, COL_LEAVE_TYPE).Value = deliveryContent
                    .Cells(outputRow, COL_MISSING_ENTRY_TYPE).Value = ""
                        .Cells(outputRow, COL_COMMENT).Value = comment
                        
                        ' 時間形式を設定
                        .Cells(outputRow, COL_ATTENDANCE_TIME).Value = attendanceTime
                        .Cells(outputRow, COL_ATTENDANCE_TIME).NumberFormat = "h:mm"
                        .Cells(outputRow, COL_DEPARTURE_TIME).Value = departureTime
                        .Cells(outputRow, COL_DEPARTURE_TIME).NumberFormat = "h:mm"
                        
                        ' 入力漏れの行を強調表示（矛盾種別列には値を設定しない）
                        .Range(.Cells(outputRow, 1), .Cells(outputRow, COL_DEPARTURE_TIME)).Interior.Color = RGB(255, 200, 200)
                    End With
                    
                    outputRow = outputRow + 1
                    totalMissingCount = totalMissingCount + 1
                End If
            End If
        End If
        
NextRow:
    Next i
    
    ' 結果が空の場合のメッセージ
    If outputRow = 2 Then
        With outputSheet
            .Cells(2, 1).Value = "勤怠入力漏れは検出されませんでした。"
            .Range(.Cells(2, 1), .Cells(2, COL_COMMENT)).Merge
            .Range(.Cells(2, 1), .Cells(2, COL_COMMENT)).HorizontalAlignment = xlCenter
        End With
    End If
    
    ' 統計情報の保存（後で見えなくする部分）
    outputSheet.Range("J2").Value = totalMissingCount
    outputSheet.Range("J3").Value = missingAttendanceCount
    outputSheet.Range("J4").Value = missingDepartureCount
    outputSheet.Range("J5").Value = missingBothCount
    outputSheet.Range("J6").Value = employeeDict.Count  ' ここで一意の社員数を保存
    
    ' 不要な計算データを非表示にする（白色の文字にする）
    outputSheet.Range("J2").Font.Color = RGB(255, 255, 255)
    outputSheet.Range("J3").Font.Color = RGB(200, 200, 200)

    ' 列幅の自動調整（J列は幅0に設定）
    outputSheet.Columns("A:I").AutoFit
    outputSheet.Columns("J").ColumnWidth = 0
    
    Exit Sub
    
ErrorHandler:
    MsgBox "勤怠入力漏れの検出中にエラーが発生しました: " & Err.Description, vbCritical
    Resume Next
End Sub

' 勤怠入力が必要かどうかを判断する関数 - 最適化版
Public Function IsEntryRequired(calendarType As String, deliveryContent As String) As Boolean
    ' デフォルトでは入力が必要
    IsEntryRequired = True
    
    ' カレンダー種別に基づく判断
    If InStr(1, calendarType, "法定外", vbTextCompare) > 0 Then
        ' 法定外は休日とみなして入力不要
        IsEntryRequired = False
        
        ' ただし、休日出勤がある場合は入力が必要
        If InStr(1, deliveryContent, "休日出勤", vbTextCompare) > 0 Or _
           InStr(1, deliveryContent, "休出", vbTextCompare) > 0 Then
            IsEntryRequired = True
        End If
    ElseIf InStr(1, calendarType, "平日", vbTextCompare) > 0 Then
        ' 平日の場合、届出内容に基づいて判断
        
        ' 届出内容が空の場合は通常勤務と判断（入力必要）
        If Trim(deliveryContent) = "" Then
            IsEntryRequired = True
            
        ' 届出内容に基づく判断
        Else
            Select Case Trim(deliveryContent)
                ' 入力不要な届出
                Case "有休", "欠勤", "振替休暇", "特別休暇"
                    IsEntryRequired = False
                    
                ' 入力必要な届出
                Case "時間有休", "午前有休", "午後有休"
                    IsEntryRequired = True
                    
                ' その他の届出は基本的に入力必要
                Case Else
                    ' 特定のキーワードを含む場合の判断
                    If InStr(1, deliveryContent, "休日出勤", vbTextCompare) > 0 Or _
                       InStr(1, deliveryContent, "振替出勤", vbTextCompare) > 0 Then
                        IsEntryRequired = True
                    End If
            End Select
        End If
    End If
    
    ' デバッグ情報（デバッグモードがONの場合のみ出力）
    If DEBUG_MODE Then
        Debug.Print "カレンダー: " & calendarType & ", 届出: " & deliveryContent & ", 入力必要: " & IsEntryRequired
    End If
End Function

' 時間文字列から時間部分を取得する関数
Private Function GetHourFromTimeString(timeStr As String) As Integer
    If timeStr = "" Then
        GetHourFromTimeString = 0
        Exit Function
    End If
    
    ' 変数宣言を関数の先頭にまとめる
    Dim numericTime As Double
    Dim hour As Integer
    Dim timeParts As Variant
    
    ' デバッグ情報
    If DEBUG_MODE Then
        Debug.Print "GetHourFromTimeString - 入力: " & timeStr
    End If
    
    ' 数値形式の場合（例：0.54166...）
    If IsNumeric(timeStr) Then
        numericTime = CDbl(timeStr)
        
        ' 特定の値の処理
        ' 0.541666... は 13:00 を表す
        If Abs(numericTime - 0.541666) < 0.01 Then
            ' デバッグ情報
            If DEBUG_MODE Then
                Debug.Print "  特殊ケース: 0.541666... → 13時"
            End If
            GetHourFromTimeString = 13
            Exit Function
        End If
        
        ' 0.375 は 9:00 を表す
        If Abs(numericTime - 0.375) < 0.01 Then
            ' デバッグ情報
            If DEBUG_MODE Then
                Debug.Print "  特殊ケース: 0.375 → 9時"
            End If
            GetHourFromTimeString = 9
            Exit Function
        End If
        
        ' その他の数値は24時間形式の小数として解釈
        ' 例：0.5 = 12時間 = 12:00
        hour = Int(numericTime * 24)
        
        ' デバッグ情報
        If DEBUG_MODE Then
            Debug.Print "  数値変換: " & numericTime & " → " & hour & "時"
        End If
        
        GetHourFromTimeString = hour
        Exit Function
    End If
    
    ' HH:MM形式の場合
    timeParts = Split(timeStr, ":")
    
    If UBound(timeParts) >= 0 Then
        If IsNumeric(timeParts(0)) Then
            hour = CInt(timeParts(0))
            
            ' デバッグ情報
            If DEBUG_MODE Then
                Debug.Print "  HH:MM形式: " & timeStr & " → " & hour & "時"
            End If
            
            GetHourFromTimeString = hour
        Else
            GetHourFromTimeString = 0
        End If
    Else
        GetHourFromTimeString = 0
    End If
End Function

' 時間文字列から分部分を取得する関数
Private Function GetMinuteFromTimeString(timeStr As String) As Integer
    If timeStr = "" Then
        GetMinuteFromTimeString = 0
        Exit Function
    End If
    
    ' 変数宣言を関数の先頭にまとめる
    Dim numericTime As Double
    Dim minute As Integer
    Dim timeParts As Variant
    
    ' デバッグ情報
    If DEBUG_MODE Then
        Debug.Print "GetMinuteFromTimeString - 入力: " & timeStr
    End If
    
    ' 数値形式の場合（例：0.54166...）
    If IsNumeric(timeStr) Then
        numericTime = CDbl(timeStr)
        
        ' 特定の値の処理
        ' 0.541666... は 13:00 を表す
        If Abs(numericTime - 0.541666) < 0.01 Then
            ' デバッグ情報
            If DEBUG_MODE Then
                Debug.Print "  特殊ケース: 0.541666... → 0分"
            End If
            GetMinuteFromTimeString = 0
            Exit Function
        End If
        
        ' 0.375 は 9:00 を表す
        If Abs(numericTime - 0.375) < 0.01 Then
            ' デバッグ情報
            If DEBUG_MODE Then
                Debug.Print "  特殊ケース: 0.375 → 0分"
            End If
            GetMinuteFromTimeString = 0
            Exit Function
        End If
        
        ' その他の数値は24時間形式の小数として解釈
        ' 例：0.5 = 12時間 = 12:00
        minute = Int((numericTime * 24 - Int(numericTime * 24)) * 60 + 0.5)
        
        ' デバッグ情報
        If DEBUG_MODE Then
            Debug.Print "  数値変換: " & numericTime & " → " & minute & "分"
        End If
        
        GetMinuteFromTimeString = minute
        Exit Function
    End If
    
    ' HH:MM形式の場合
    timeParts = Split(timeStr, ":")
    
    If UBound(timeParts) >= 1 Then
        If IsNumeric(timeParts(1)) Then
            minute = CInt(timeParts(1))
            
            ' デバッグ情報
            If DEBUG_MODE Then
                Debug.Print "  HH:MM形式: " & timeStr & " → " & minute & "分"
            End If
            
            GetMinuteFromTimeString = minute
        Else
            GetMinuteFromTimeString = 0
        End If
    Else
        GetMinuteFromTimeString = 0
    End If
End Function

' 除外社員番号を取得する関数
' この関数は他のモジュールで定義されていると仮定
' Public Function 除外社員番号取得() As Variant
'     ' 実装は別モジュールにあると仮定
' End Function

