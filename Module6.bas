' ========================================
' Module6
' タイプ: 標準モジュール
' 行数: 395
' エクスポート日時: 2025-10-17 14:37:26
' ========================================

Option Explicit

' *************************************************************
' モジュール：勤怠入力漏れレポート生成
' 目的：勤怠入力漏れのレポートを生成する関数群
' Copyright (c) 2025 SI1 shunpei.suzuki
' 作成日：2025年4月2日
'
' 改版履歴：
' 2025/04/02 module2から分割作成
' *************************************************************

' 定数定義（module2_coreと同じ定数を定義）
Private Const SHEET_NAME_MISSING_ENTRIES As String = "勤怠入力漏れ一覧"
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

' 出力シートを準備する
Public Function PrepareOutputSheet() As Worksheet
    On Error Resume Next
    
    Application.StatusBar = "出力シートを準備しています..."
    
    ' 既存シートがあれば削除
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME_MISSING_ENTRIES)
    If Not ws Is Nothing Then
        ws.Delete
    End If
    
    ' 残業一覧シートを取得
    Dim overtimeSheet As Worksheet
    On Error Resume Next
    Set overtimeSheet = ThisWorkbook.Worksheets("残業一覧")
    On Error GoTo 0
    
    ' 新しいシートを作成（残業一覧シートの右隣に）
    If Not overtimeSheet Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=overtimeSheet)
    Else
        Set ws = ThisWorkbook.Sheets.Add
    End If
    ws.Name = SHEET_NAME_MISSING_ENTRIES
    
    ' ヘッダー行の設定
    ws.Cells(1, COL_EMPLOYEE_ID).Value = "社員番号"
    ws.Cells(1, COL_EMPLOYEE_NAME).Value = "氏名"
    ws.Cells(1, COL_DATE).Value = "日付"
    ws.Cells(1, COL_DAY_TYPE).Value = "曜日区分"
    ws.Cells(1, COL_LEAVE_TYPE).Value = "届出内容"
    ws.Cells(1, COL_COMMENT).Value = "コメント"
    ws.Cells(1, COL_ATTENDANCE_TIME).Value = "出勤時刻"
    ws.Cells(1, COL_DEPARTURE_TIME).Value = "退勤時刻"
    
    ' 入力漏れ種別と矛盾種別の列は非表示にする
    ws.Columns(COL_MISSING_ENTRY_TYPE).Hidden = True
    ws.Columns(COL_CONTRADICTION_TYPE).Hidden = True
    
    ' 矛盾種別の説明は概要統計の下に配置するため、ここでは追加しない
    
    ' ヘッダー行の書式設定
    ws.Range(ws.Cells(1, 1), ws.Cells(1, COL_CONTRADICTION_TYPE)).Interior.Color = RGB(200, 200, 200)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, COL_CONTRADICTION_TYPE)).Font.Bold = True
    
    ' 列幅の自動調整
    ws.Columns("B:M").AutoFit
    
    ' 社員番号列を文字列形式に設定
    ws.Columns("A").NumberFormat = "@"
    
    Set PrepareOutputSheet = ws
End Function

' 概要統計を計算して表示する
Public Sub CalculateAndDisplaySummary(missingEntriesSheet As Worksheet)
    On Error GoTo ErrorHandler
    
    Application.StatusBar = "概要統計を計算しています..."
    
    ' 保存された統計情報の取得
    Dim totalMissing As Long
    Dim missingAttendance As Long
    Dim missingDeparture As Long
    Dim missingBoth As Long
    Dim employeeCount As Long
    Dim i As Long
    Dim lastRow As Long
    Dim nextRow As Long
    
    totalMissing = missingEntriesSheet.Range("J2").Value
    missingAttendance = missingEntriesSheet.Range("J3").Value
    missingDeparture = missingEntriesSheet.Range("J4").Value
    missingBoth = missingEntriesSheet.Range("J5").Value
    employeeCount = missingEntriesSheet.Range("J6").Value
    
    ' 入力漏れ一覧のシートに概要統計エリアを作成
    With missingEntriesSheet
        ' 不要な計算データセルは白色文字にしておく（J2-J6）
        .Range("J2:J3").Font.Color = RGB(255, 255, 255)
        
        .Cells(3, 12).Value = "概要統計"
        .Cells(3, 12).Font.Bold = True
        
        .Cells(4, 12).Value = "検出された入力漏れ"
        .Cells(4, 13).Value = totalMissing & "件"
        
        .Cells(5, 12).Value = "出勤時刻なし"
        .Cells(5, 13).Value = missingAttendance & "件"
        
        .Cells(6, 12).Value = "退勤時刻なし"
        .Cells(6, 13).Value = missingDeparture & "件"
        
        .Cells(7, 12).Value = "出退勤時刻なし"
        .Cells(7, 13).Value = missingBoth & "件"
        
        .Cells(8, 12).Value = "対象従業員数"
        .Cells(8, 13).Value = employeeCount & "名"
        
        ' 書式設定
        .Range(.Cells(3, 12), .Cells(8, 13)).Borders.LineStyle = xlNone
        .Range(.Cells(3, 12), .Cells(3, 13)).Interior.Color = RGB(200, 200, 200)
        
        ' 説明は不要
        
        ' 列幅の自動調整
        .Columns("L:M").AutoFit
    End With
    
    ' 勤怠情報分析結果シートにも情報を追加（シート名を修正）
    Dim summarySheet As Worksheet
    On Error Resume Next
    Set summarySheet = ThisWorkbook.Worksheets("勤怠情報分析結果")
    
    If Not summarySheet Is Nothing Then
        ' 既存の最終行を見つける（部門別残業集計の最終行以降）
        lastRow = 0
        
        ' 部署列（A列）を下方向にスキャンして最後の非空セルを見つける
        For i = 1 To 100
            If Not IsEmpty(summarySheet.Cells(i, 1).Value) Then
                lastRow = i
            End If
        Next i
        
        ' 最終行から3行空けて開始
        nextRow = lastRow + 3
        
        ' 勤怠入力漏れ情報のヘッダー
        summarySheet.Cells(nextRow, 1).Value = "勤怠入力漏れ概要"
        summarySheet.Cells(nextRow, 1).Font.Bold = True
        summarySheet.Cells(nextRow, 1).Interior.Color = RGB(200, 200, 200)
        summarySheet.Range(summarySheet.Cells(nextRow, 1), summarySheet.Cells(nextRow, 2)).Merge
        
        ' 詳細情報
        summarySheet.Cells(nextRow + 1, 1).Value = "検出された入力漏れ"
        summarySheet.Cells(nextRow + 1, 2).Value = totalMissing & "件"
        
        summarySheet.Cells(nextRow + 2, 1).Value = "出勤時刻なし"
        summarySheet.Cells(nextRow + 2, 2).Value = missingAttendance & "件"
        
        summarySheet.Cells(nextRow + 3, 1).Value = "退勤時刻なし"
        summarySheet.Cells(nextRow + 3, 2).Value = missingDeparture & "件"
        
        summarySheet.Cells(nextRow + 4, 1).Value = "出退勤時刻なし"
        summarySheet.Cells(nextRow + 4, 2).Value = missingBoth & "件"
        
        summarySheet.Cells(nextRow + 5, 1).Value = "対象従業員数"
        summarySheet.Cells(nextRow + 5, 2).Value = employeeCount & "名"
        
        ' 書式設定
        summarySheet.Range(summarySheet.Cells(nextRow + 1, 1), summarySheet.Cells(nextRow + 5, 2)).Borders.LineStyle = xlContinuous
    End If
    
    ' 特別休暇リストを表示（勤怠入力漏れ概要の下）
    Call AddSpecialLeaveList(summarySheet, nextRow)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "概要統計の計算中にエラーが発生しました: " & Err.Description, vbCritical
End Sub

' 特別休暇リストを表示する - 最適化版
Public Sub AddSpecialLeaveList(summarySheet As Worksheet, nextRow As Long)
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
    
    ' 高速化のためにヘッダー行をバッファに取得
    Dim headerRange As Range
    Set headerRange = wsCSVData.Range(wsCSVData.Cells(1, 1), wsCSVData.Cells(1, wsCSVData.Cells(1, wsCSVData.Columns.Count).End(xlToLeft).Column))
    Dim headerArray As Variant
    headerArray = headerRange.Value
    
    ' 各列のインデックスを特定
    Dim i As Long, j As Long
    For i = 1 To UBound(headerArray, 2)
        Select Case headerArray(1, i)
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
    
    ' 除外社員番号を取得
    Dim excludeIDs As Variant
    excludeIDs = 除外社員番号取得()
    
    ' 高速化のため除外IDを辞書に変換
    Dim excludeDict As Object
    Set excludeDict = CreateObject("Scripting.Dictionary")
    excludeDict.CompareMode = vbTextCompare
    
    For j = LBound(excludeIDs) To UBound(excludeIDs)
        If excludeIDs(j) <> "" Then
            excludeDict.Add excludeIDs(j), True
        End If
    Next j
    
    ' データをバッファに取得して高速化
    Dim dataRange As Range
    Set dataRange = wsCSVData.Range(wsCSVData.Cells(2, 1), wsCSVData.Cells(lastRow, wsCSVData.Cells(1, wsCSVData.Columns.Count).End(xlToLeft).Column))
    Dim dataArray As Variant
    dataArray = dataRange.Value
    
    ' 特別休暇レコードを収集
    Dim specialLeaves As New Collection
    Dim leaveRecord As Object
    
    ' CSV各行をチェック - 高速化
    For i = 1 To UBound(dataArray, 1)
        ' 社員番号を取得
        Dim employeeID As String
        employeeID = Trim(CStr(dataArray(i, 社員番号Col)))
        
        ' 除外社員の場合はスキップ - 厳密な文字列比較
        If excludeDict.Exists(employeeID) Then
            Debug.Print "==> 特別休暇リストから除外: " & employeeID
            GoTo NextSpecialLeave
        End If
        
        ' 届出内容が「特別休暇」のレコードを抽出
        If 届出Col > 0 And Trim(CStr(dataArray(i, 届出Col))) = "特別休暇" Then
            Set leaveRecord = CreateObject("Scripting.Dictionary")
            leaveRecord.Add "部門", dataArray(i, 部門Col) ' 部門を最初に
            leaveRecord.Add "社員番号", employeeID
            leaveRecord.Add "氏名", dataArray(i, 氏名Col)
            leaveRecord.Add "役職", dataArray(i, 役職Col)
            leaveRecord.Add "日付", dataArray(i, 日付Col)
            leaveRecord.Add "曜日", dataArray(i, 曜日Col)
            leaveRecord.Add "カレンダー", dataArray(i, カレンダーCol)
            leaveRecord.Add "届出内容", dataArray(i, 届出Col)
            leaveRecord.Add "備考", dataArray(i, 備考Col)
            leaveRecord.Add "備考空欄", (Trim(CStr(dataArray(i, 備考Col))) = "")
            
            ' コレクションに追加
            specialLeaves.Add leaveRecord
        End If
NextSpecialLeave:
    Next i
    
    ' 特別休暇がなければ終了
    If specialLeaves.Count = 0 Then Exit Sub
    
    ' 特別休暇リストの表示位置（勤怠入力漏れ概要の2行下）
    Dim listRow As Long
    listRow = nextRow + 8
    
    ' ヘッダー行を設定
    With summarySheet
        .Cells(listRow, 1).Value = "特別休暇リスト"
        .Cells(listRow, 1).Font.Bold = True
        .Cells(listRow, 1).Interior.Color = RGB(200, 200, 200)
        .Range(.Cells(listRow, 1), .Cells(listRow, 9)).Merge
        
        listRow = listRow + 1
        
        ' カラムヘッダー (並び順を変更)
        .Cells(listRow, 1).Value = "部署" ' 部門を部署に変更
        .Cells(listRow, 2).Value = "社員番号"
        .Cells(listRow, 3).Value = "氏名"
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
        
        ' 特別休暇レコードを表示 - バッファに一度にデータを準備して高速化
        Dim outputData() As Variant
        ReDim outputData(1 To specialLeaves.Count, 1 To 9)
        Dim hasEmptyRemarks As Boolean
        hasEmptyRemarks = False
        
        Dim idx As Long
        idx = 1
        
        For Each leaveRecord In specialLeaves
            outputData(idx, 1) = leaveRecord("部門")
            outputData(idx, 2) = leaveRecord("社員番号")
            outputData(idx, 3) = leaveRecord("氏名")
            outputData(idx, 4) = leaveRecord("役職")
            outputData(idx, 5) = leaveRecord("日付")
            outputData(idx, 6) = leaveRecord("曜日")
            outputData(idx, 7) = leaveRecord("カレンダー")
            outputData(idx, 8) = leaveRecord("届出内容")
            outputData(idx, 9) = leaveRecord("備考")
            
            If leaveRecord("備考空欄") Then
                hasEmptyRemarks = True
            End If
            
            idx = idx + 1
        Next leaveRecord
        
        ' データを一括でシートに書き込み
        .Range(.Cells(listRow, 1), .Cells(listRow + specialLeaves.Count - 1, 9)).Value = outputData
        
        ' 社員番号列を文字列形式に設定
        .Range(.Cells(listRow, 2), .Cells(listRow + specialLeaves.Count - 1, 2)).NumberFormat = "@"
        
        ' 備考欄が空欄の行をハイライト
        For i = 1 To specialLeaves.Count
            If outputData(i, 9) = "" Then
                .Cells(listRow + i - 1, 9).Interior.Color = RGB(255, 255, 200)
            End If
        Next i
        
        ' コメントを追加
        .Cells(listRow + specialLeaves.Count + 1, 1).Value = "届出内容に対して備考欄の記載が明確、かつ確実に説明がなされていることを確認すること。"
        .Cells(listRow + specialLeaves.Count + 2, 1).Value = "備考欄の記載不備は修正が必要です。"
        .Cells(listRow + specialLeaves.Count + 3, 1).Value = "入力、報告不備が原因で指摘を受けた場合は報告書対応となります。"
        .Cells(listRow + specialLeaves.Count + 5, 1).Value = "【指摘あり実績】"
        .Cells(listRow + specialLeaves.Count + 6, 1).Value = "　2025年3月 慶弔休暇申請について、「慶弔休暇」という備考欄の記載は認められない。"
        .Cells(listRow + specialLeaves.Count + 7, 1).Value = "　2025年3月 慶弔休暇申請について、「慶事」なのか「弔事」なのか明確に記載ががあることを確認すること。"
        
        If hasEmptyRemarks Then
            .Range(.Cells(listRow + specialLeaves.Count + 1, 1), .Cells(listRow + specialLeaves.Count + 2, 9)).Font.Color = RGB(255, 0, 0)
            .Range(.Cells(listRow + specialLeaves.Count + 1, 1), .Cells(listRow + specialLeaves.Count + 2, 9)).Font.Bold = True
        End If
        
        ' 表のボーダーを設定
        Dim tableRange As Range
        Set tableRange = .Range(.Cells(listRow, 1), .Cells(listRow + specialLeaves.Count - 1, 9))
        tableRange.Borders.LineStyle = xlContinuous
        tableRange.Borders.Weight = xlThin
        
        ' 列幅の自動調整
        .Columns("B:I").AutoFit
    End With
End Sub
