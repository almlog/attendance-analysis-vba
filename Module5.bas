Attribute VB_Name = "Module5"
' ========================================
' Module5 (機能統合・完全修正版)
' タイプ: 標準モジュール
' 修正内容: 旧ロジック復元 + 通知連携用修正 + データ浄化
' ========================================

Option Explicit

' *************************************************************
' モジュール：勤怠入力漏れ検出
' 目的：勤怠入力漏れを検出する関数群
' *************************************************************

' 定数定義
Private Const COL_EMPLOYEE_ID As Integer = 1
Private Const COL_EMPLOYEE_NAME As Integer = 2
Private Const COL_DATE As Integer = 3
Private Const COL_DAY_TYPE As Integer = 4
Private Const COL_LEAVE_TYPE As Integer = 5
Private Const COL_MISSING_ENTRY_TYPE As Integer = 6
Private Const COL_COMMENT As Integer = 7
Private Const COL_ATTENDANCE_TIME As Integer = 8
Private Const COL_DEPARTURE_TIME As Integer = 9
Private Const COL_CONTRADICTION_TYPE As Integer = 10 ' J列（通知用コード）
Private Const DEBUG_MODE As Boolean = False

' グローバル変数の参照（module2_coreで定義）
' Public g_IncludeToday As Boolean

' *************************************************************
' 関数名: DetectMissingEntries
' 目的: 勤怠入力漏れ検出（旧ロジックベース + データ浄化 + 通知連携）
' *************************************************************
Public Sub DetectMissingEntries(wsCSVData As Worksheet, outputSheet As Worksheet)
    Dim includeToday As Boolean
'    includeToday = g_IncludeToday
    
    On Error GoTo ErrorHandler
    Application.StatusBar = "勤怠入力漏れを検出しています..."
    
    ' 最終行を取得
    Dim lastRow As Long
    lastRow = wsCSVData.Cells(wsCSVData.Rows.count, "A").End(xlUp).Row
    
    If lastRow <= 1 Then
        MsgBox "CSVデータが存在しません。", vbExclamation
        Exit Sub
    End If
    
    ' 列インデックスの特定
    Dim 社員番号Col As Integer, 氏名Col As Integer, 日付Col As Integer
    Dim カレンダーCol As Integer, 曜日Col As Integer, 届出Col As Integer
    Dim 出勤時刻Col As Integer, 退勤時刻Col As Integer
    
    社員番号Col = 0: 氏名Col = 0: 日付Col = 0
    カレンダーCol = 0: 曜日Col = 0: 届出Col = 0
    出勤時刻Col = 0: 退勤時刻Col = 0
    
    Dim headerRow As Range
    Set headerRow = wsCSVData.Range(wsCSVData.Cells(1, 1), wsCSVData.Cells(1, wsCSVData.Cells(1, wsCSVData.Columns.count).End(xlToLeft).Column))
    
    Dim i As Long, j As Long
    For i = 1 To headerRow.Columns.count
        ' ヘッダーもCleanStringを通すことで、微妙なスペース違いを吸収
        Select Case CleanString(headerRow.Cells(1, i).Value)
            Case "社員番号": 社員番号Col = i
            Case "氏名": 氏名Col = i
            Case "日付": 日付Col = i
            Case "カレンダー": カレンダーCol = i
            Case "曜日": 曜日Col = i
            Case "届出内容": 届出Col = i
            Case "出社": 出勤時刻Col = i
            Case "退社": 退勤時刻Col = i
        End Select
    Next i
    
    If 社員番号Col = 0 Or 氏名Col = 0 Or 日付Col = 0 Then
        MsgBox "必要な列（社員番号、氏名、日付）が見つかりませんでした。", vbExclamation
        Exit Sub
    End If
    
    ' デフォルト値
    If 出勤時刻Col = 0 Then 出勤時刻Col = 10
    If 退勤時刻Col = 0 Then 退勤時刻Col = 11
    
    Dim outputRow As Long
    outputRow = 2
    
    Dim missingAttendanceCount As Long, missingDepartureCount As Long
    Dim missingBothCount As Long, totalMissingCount As Long, contradictionCount As Long
    missingAttendanceCount = 0: missingDepartureCount = 0
    missingBothCount = 0: totalMissingCount = 0: contradictionCount = 0
    
    Dim employeeDict As Object
    Set employeeDict = CreateObject("Scripting.Dictionary")
    employeeDict.CompareMode = vbTextCompare
    
    ' 除外社員番号の取得（エラーハンドリング追加）
    Dim excludeDict As Object
    Set excludeDict = CreateObject("Scripting.Dictionary")
    excludeDict.CompareMode = vbTextCompare
    
    On Error Resume Next
    Dim excludeIDs As Variant
    excludeIDs = 除外社員番号取得() ' ※この関数が存在しない場合は空として扱う
    If Err.Number = 0 And IsArray(excludeIDs) Then
        For j = LBound(excludeIDs) To UBound(excludeIDs)
            If excludeIDs(j) <> "" Then
                excludeDict(Trim(CStr(excludeIDs(j)))) = True
            End If
        Next j
    End If
    On Error GoTo ErrorHandler
    
    ' データ読み込み
    Dim dataRange As Range
    Set dataRange = wsCSVData.Range(wsCSVData.Cells(2, 1), wsCSVData.Cells(lastRow, wsCSVData.Cells(1, wsCSVData.Columns.count).End(xlToLeft).Column))
    Dim dataArray As Variant
    dataArray = dataRange.Value
    
    Dim todayDate As Date
    todayDate = Date
    
    ' メインループ
    For i = 1 To UBound(dataArray, 1)
        ' ★★★ 修正点: CleanStringを使ってゴミを除去 ★★★
        Dim employeeID As String
        employeeID = CleanString(dataArray(i, 社員番号Col))
        
        ' 除外チェック
        If employeeID <> "" And excludeDict.Exists(employeeID) Then GoTo NextRow
        
        Dim entryDate As Date
        If IsDate(dataArray(i, 日付Col)) Then
            entryDate = CDate(dataArray(i, 日付Col))
        Else
            GoTo NextRow
        End If
        
        Dim employeeName As String: employeeName = CleanString(dataArray(i, 氏名Col))
        Dim dayType As String: If 曜日Col > 0 Then dayType = CleanString(dataArray(i, 曜日Col))
        Dim calendarType As String: If カレンダーCol > 0 Then calendarType = CleanString(dataArray(i, カレンダーCol))
        Dim deliveryContent As String: If 届出Col > 0 Then deliveryContent = CleanString(dataArray(i, 届出Col))
        
        Dim attendanceTime As String, departureTime As String
        attendanceTime = CleanString(dataArray(i, 出勤時刻Col))
        departureTime = CleanString(dataArray(i, 退勤時刻Col))
        
        Dim hasAttendanceTime As Boolean: hasAttendanceTime = (attendanceTime <> "")
        Dim hasDepartureTime As Boolean: hasDepartureTime = (departureTime <> "")
        
        ' 日付チェック
        If (DateDiff("d", entryDate, todayDate) > 0 Or (includeToday And DateDiff("d", entryDate, todayDate) = 0)) Then
            Dim contradictionType As String
            Dim comment As String
            contradictionType = ""
            comment = ""
            
            ' --- 矛盾チェック (元のロジックを維持) ---
            
            ' 1. 午前有休
            If deliveryContent = "午前有休" And hasAttendanceTime Then
                If GetHourFromTimeString(attendanceTime) < 13 Then
                    contradictionType = "1" ' Module8用コード
                    comment = "午前有休なのに出勤時刻が13時より前（" & FormatTimeDisplay(attendanceTime) & "）になっています"
                End If
            End If
            
            ' 2. 午後有休
            If deliveryContent = "午後有休" And hasDepartureTime Then
                Dim depH As Integer, depM As Integer
                depH = GetHourFromTimeString(departureTime)
                depM = GetMinuteFromTimeString(departureTime)
                If depH > 12 Or (depH = 12 And depM > 0) Then
                    contradictionType = "2" ' Module8用コード
                    comment = "午後有休なのに退勤時刻が12時より後（" & FormatTimeDisplay(departureTime) & "）になっています"
                End If
            End If
            
            ' 3. お昼休憩 (矛盾なしの場合のみ)
            If contradictionType = "" Then
                ' 出勤
                If hasAttendanceTime Then
                    If GetHourFromTimeString(attendanceTime) = 12 Then
                        contradictionType = "3" ' Module8用コード (旧: "お昼休憩矛盾")
                        comment = "お昼休憩時間(12:00～12:59)に出勤（" & FormatTimeDisplay(attendanceTime) & "）しています"
                    End If
                End If
                ' 退勤
                If contradictionType = "" And hasDepartureTime Then
                    Dim lDepH As Integer, lDepM As Integer
                    lDepH = GetHourFromTimeString(departureTime)
                    lDepM = GetMinuteFromTimeString(departureTime)
                    If lDepH = 12 And lDepM > 0 Then
                        contradictionType = "3" ' Module8用コード (旧: "お昼休憩矛盾")
                        comment = "お昼休憩時間(12:01～12:59)に退勤（" & FormatTimeDisplay(departureTime) & "）しています"
                    End If
                End If
            End If
            
            ' 矛盾出力
            If contradictionType <> "" Then
                If Not employeeDict.Exists(employeeID) Then employeeDict.Add employeeID, employeeName
                
                With outputSheet
                    .Cells(outputRow, COL_EMPLOYEE_ID).Value = employeeID
                    .Cells(outputRow, COL_EMPLOYEE_NAME).Value = employeeName
                    .Cells(outputRow, COL_DATE).Value = entryDate
                    .Cells(outputRow, COL_DAY_TYPE).Value = dayType
                    .Cells(outputRow, COL_LEAVE_TYPE).Value = deliveryContent
                    .Cells(outputRow, COL_MISSING_ENTRY_TYPE).Value = ""
                    .Cells(outputRow, COL_COMMENT).Value = comment
                    .Cells(outputRow, COL_ATTENDANCE_TIME).Value = FormatTimeDisplay(attendanceTime)
                    .Cells(outputRow, COL_DEPARTURE_TIME).Value = FormatTimeDisplay(departureTime)
                    
                    ' ★★★ 修正点: Module8のために矛盾コードを出力する ★★★
                    .Cells(outputRow, COL_CONTRADICTION_TYPE).Value = contradictionType
                    
                    .Range(.Cells(outputRow, 1), .Cells(outputRow, COL_DEPARTURE_TIME)).Interior.Color = COLOR_CONTRADICTION
                End With
                outputRow = outputRow + 1
                totalMissingCount = totalMissingCount + 1
                contradictionCount = contradictionCount + 1
            End If
            
            ' --- 入力漏れチェック ---
            If IsEntryRequired(calendarType, deliveryContent) Then
                Dim missingType As String: missingType = ""
                
                If Not hasAttendanceTime And Not hasDepartureTime Then
                    missingType = "3" ' 両方なし
                    comment = "出勤時刻と退勤時刻の両方が入力されていません"
                    missingBothCount = missingBothCount + 1
                ElseIf Not hasAttendanceTime Then
                    missingType = "1" ' 出勤なし
                    comment = "出勤時刻が入力されていません"
                    missingAttendanceCount = missingAttendanceCount + 1
                ElseIf Not hasDepartureTime Then
                    missingType = "2" ' 退勤なし
                    comment = "退勤時刻が入力されていません"
                    missingDepartureCount = missingDepartureCount + 1
                End If
                
                ' 漏れ出力 (矛盾と重複して出力しない制御が必要ならここで行うが、元のロジックに従い両方チェック)
                ' ※ただし同じ行に上書きするとおかしくなるため、行を分けるか、矛盾がない場合のみ出力する
                ' ここでは「矛盾がない場合」に出力するロジックとします
                If missingType <> "" And contradictionType = "" Then
                    If Not employeeDict.Exists(employeeID) Then employeeDict.Add employeeID, employeeName
                    
                    With outputSheet
                        .Cells(outputRow, COL_EMPLOYEE_ID).Value = employeeID
                        .Cells(outputRow, COL_EMPLOYEE_NAME).Value = employeeName
                        .Cells(outputRow, COL_DATE).Value = entryDate
                        .Cells(outputRow, COL_DAY_TYPE).Value = dayType
                        .Cells(outputRow, COL_LEAVE_TYPE).Value = deliveryContent
                        .Cells(outputRow, COL_MISSING_ENTRY_TYPE).Value = missingType
                        .Cells(outputRow, COL_COMMENT).Value = comment
                        .Cells(outputRow, COL_ATTENDANCE_TIME).Value = FormatTimeDisplay(attendanceTime)
                        .Cells(outputRow, COL_DEPARTURE_TIME).Value = FormatTimeDisplay(departureTime)
                        
                        ' ★★★ 修正点: 漏れの場合は0を出力 ★★★
                        .Cells(outputRow, COL_CONTRADICTION_TYPE).Value = "0"
                        
                        .Range(.Cells(outputRow, 1), .Cells(outputRow, COL_DEPARTURE_TIME)).Interior.Color = COLOR_MISSING
                    End With
                    outputRow = outputRow + 1
                    totalMissingCount = totalMissingCount + 1
                End If
            End If
        End If
NextRow:
    Next i
    
    ' 結果表示
    If outputRow = 2 Then
        outputSheet.Cells(2, 1).Value = "勤怠入力漏れ・矛盾は検出されませんでした。"
    End If
    
    ' 統計
    outputSheet.Range("J2").Value = totalMissingCount
    outputSheet.Range("J3").Value = missingAttendanceCount
    outputSheet.Range("J4").Value = missingDepartureCount
    outputSheet.Range("J5").Value = missingBothCount
    outputSheet.Range("J6").Value = employeeDict.count
    outputSheet.Range("J7").Value = contradictionCount
    
    ' 見た目の調整
    outputSheet.Range("J2:J7").Font.Color = RGB(255, 255, 255)
    outputSheet.Columns("A:I").AutoFit
    outputSheet.Columns("J").ColumnWidth = 0
    
    Exit Sub
ErrorHandler:
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub

' *************************************************************
' 関数名: 勤怠入力漏れチェック (自動シート特定ラッパー)
' 目的: シート名違いのエラーを回避する
' *************************************************************
Public Sub 勤怠入力漏れチェック()
    Dim wsData As Worksheet
    Dim wsOut As Worksheet
    
    On Error Resume Next
    ' データシート特定 (全データ > 勤怠データ > Sheet1 > アクティブシート)
    Set wsData = ThisWorkbook.Sheets("全データ")
    If wsData Is Nothing Then Set wsData = ThisWorkbook.Sheets("勤怠データ")
    If wsData Is Nothing Then Set wsData = ThisWorkbook.Sheets("Sheet1")
    
    If wsData Is Nothing Then
        If ActiveSheet.Name <> "設定" And ActiveSheet.Name <> "勤怠入力漏れ一覧" Then
            Set wsData = ActiveSheet
        End If
    End If
    On Error GoTo 0
    
    If wsData Is Nothing Then
        MsgBox "【エラー】勤怠データシートが見つかりません。", vbCritical
        Exit Sub
    End If
    
    ' 出力シート特定
    On Error Resume Next
    Set wsOut = ThisWorkbook.Sheets("勤怠入力漏れ一覧")
    On Error GoTo 0
    
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsOut.Name = "勤怠入力漏れ一覧"
        wsOut.Range("A1:J1").Value = Array("社員番号", "氏名", "日付", "カレンダー", "届出内容", "入力漏れ種別", "コメント", "出社", "退社", "矛盾コード")
    Else
        If wsOut.Cells(wsOut.Rows.count, 1).End(xlUp).Row >= 2 Then
            wsOut.Rows("2:" & wsOut.Rows.count).ClearContents
            wsOut.Rows("2:" & wsOut.Rows.count).Interior.ColorIndex = xlNone
        End If
    End If
    
    Call DetectMissingEntries(wsData, wsOut)
End Sub

' *************************************************************
' 関数名: CleanString (新規追加: ゴミ文字除去)
' *************************************************************
Private Function CleanString(val As Variant) As String
    If IsError(val) Or IsNull(val) Or IsEmpty(val) Then
        CleanString = ""
        Exit Function
    End If
    Dim s As String: s = CStr(val)
    s = Replace(s, ChrW(160), " ") ' NBSP除去
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    ' 制御文字除去
    Dim i As Long, res As String
    For i = 1 To Len(s)
        If AscW(Mid(s, i, 1)) >= 32 Then res = res & Mid(s, i, 1)
    Next i
    CleanString = Trim(res)
End Function

' *************************************************************
' 以下、既存のヘルパー関数群 (IsEntryRequired, GetHour..., etc)
' *************************************************************
Public Function IsEntryRequired(calendarType As String, deliveryContent As String) As Boolean
    IsEntryRequired = True
    If InStr(1, calendarType, "法定外", vbTextCompare) > 0 Then
        IsEntryRequired = False
        If InStr(1, deliveryContent, "休出") > 0 Or InStr(1, deliveryContent, "休日出勤") > 0 Then IsEntryRequired = True
    ElseIf InStr(1, calendarType, "平日", vbTextCompare) > 0 Then
        If Trim(deliveryContent) = "" Then
            IsEntryRequired = True
        Else
            Select Case Trim(deliveryContent)
                Case "有休", "欠勤", "振替休暇", "特別休暇": IsEntryRequired = False
                Case "時間有休", "午前有休", "午後有休": IsEntryRequired = True
                Case Else
                    If InStr(1, deliveryContent, "休日出勤") > 0 Or InStr(1, deliveryContent, "振替出勤") > 0 Then IsEntryRequired = True
            End Select
        End If
    End If
End Function

Private Function GetHourFromTimeString(timeStr As String) As Integer
    If timeStr = "" Then GetHourFromTimeString = 0: Exit Function
    If IsNumeric(timeStr) Then
        Dim d As Double: d = CDbl(timeStr)
        If Abs(d - 0.541666) < 0.01 Then GetHourFromTimeString = 13: Exit Function
        If Abs(d - 0.375) < 0.01 Then GetHourFromTimeString = 9: Exit Function
        GetHourFromTimeString = Int(d * 24)
    Else
        Dim p: p = Split(timeStr, ":")
        If UBound(p) >= 0 And IsNumeric(p(0)) Then GetHourFromTimeString = CInt(p(0)) Else GetHourFromTimeString = 0
    End If
End Function

Private Function GetMinuteFromTimeString(timeStr As String) As Integer
    If timeStr = "" Then GetMinuteFromTimeString = 0: Exit Function
    If IsNumeric(timeStr) Then
        Dim d As Double: d = CDbl(timeStr)
        GetMinuteFromTimeString = Int((d * 24 - Int(d * 24)) * 60 + 0.5)
    Else
        Dim p: p = Split(timeStr, ":")
        If UBound(p) >= 1 And IsNumeric(p(1)) Then GetMinuteFromTimeString = CInt(p(1)) Else GetMinuteFromTimeString = 0
    End If
End Function

Private Function FormatTimeDisplay(timeStr As String) As String
    On Error Resume Next
    If IsNumeric(timeStr) Then FormatTimeDisplay = Format(CDbl(timeStr), "h:mm") Else FormatTimeDisplay = timeStr
End Function

' 除外社員番号取得がない場合のダミー（既存モジュールにある場合は削除可）
#If False Then
Private Function 除外社員番号取得() As Variant
    除外社員番号取得 = Array()
End Function
#End If
