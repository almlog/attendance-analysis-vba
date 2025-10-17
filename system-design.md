# システム設計書

## ドキュメント情報

| 項目 | 内容 |
|------|------|
| プロジェクト名 | 勤怠未入力通知システム |
| バージョン | 1.0 |
| 作成日 | 2025-10-17 |
| 最終更新日 | 2025-10-17 |
| ステータス | レビュー中 |

---

## 1. システムアーキテクチャ

### 1.1 全体構成

```
┌─────────────────────────────────────────────────────────┐
│                    ユーザー層                            │
│  ┌──────────────┐  ┌──────────────┐  ┌─────────────┐  │
│  │ 人事担当者    │  │ 部門管理者    │  │ 一般社員     │  │
│  │ (Excel操作)  │  │ (LINE WORKS) │  │(LINE WORKS) │  │
│  └──────┬───────┘  └──────┬───────┘  └──────┬──────┘  │
└─────────┼──────────────────┼──────────────────┼──────────┘
          │                  │                  │
          │                  │                  │
┌─────────┼──────────────────┼──────────────────┼──────────┐
│         │     アプリケーション層                │          │
│         ↓                  ↓                  ↓          │
│  ┌────────────────────────────────────────────────────┐ │
│  │         Excel VBA アプリケーション                  │ │
│  ├────────────────────────────────────────────────────┤ │
│  │  ┌──────────┐  ┌──────────┐  ┌────────────────┐  │ │
│  │  │ UI層     │  │ ビジネス  │  │ データアクセス │  │ │
│  │  │          │  │ ロジック層│  │ 層              │  │ │
│  │  │ ・ボタン │  │          │  │                │  │ │
│  │  │ ・シート │  │ ・CSV処理│  │ ・シート操作   │  │ │
│  │  │ ・ダイアログ│ │ ・抽出   │  │ ・マスタ管理   │  │ │
│  │  │          │  │ ・通知生成│  │ ・履歴記録     │  │ │
│  │  └─────┬────┘  └─────┬────┘  └───────┬────────┘  │ │
│  └────────┼──────────────┼─────────────────┼───────────┘ │
└───────────┼──────────────┼─────────────────┼─────────────┘
            │              │                 │
            │              │                 │
┌───────────┼──────────────┼─────────────────┼─────────────┐
│           │      外部サービス層            │              │
│           ↓              ↓                 ↓              │
│  ┌─────────────┐  ┌──────────────┐  ┌──────────────┐  │
│  │ 勤怠システム │  │ LINE WORKS   │  │ GitHub       │  │
│  │             │  │              │  │              │  │
│  │ ・CSV出力   │  │ ・Webhook API│  │ ・VBAコード  │  │
│  │             │  │ ・Bot API    │  │  管理        │  │
│  └─────────────┘  └──────────────┘  └──────────────┘  │
└─────────────────────────────────────────────────────────┘
```

### 1.2 レイヤー構成

#### UI層（プレゼンテーション層）
- **責務**: ユーザーとの対話、入力受付、結果表示
- **コンポーネント**:
  - ボタンコントロール（CSV読込、通知送信、履歴表示）
  - シートUI（勤怠未入力、社員マスタ、通知履歴）
  - ダイアログ（ファイル選択、確認、エラー表示）

#### ビジネスロジック層
- **責務**: 業務ロジックの実装、データ変換、判定処理
- **コンポーネント**:
  - CSV解析エンジン
  - 未入力者抽出エンジン
  - 通知メッセージ生成エンジン
  - 段階的エスカレーションロジック

#### データアクセス層
- **責務**: データの永続化、シート操作の抽象化
- **コンポーネント**:
  - シートアクセサ（CRUD操作）
  - 社員マスタリポジトリ
  - 通知履歴リポジトリ

#### 外部サービス統合層
- **責務**: 外部APIとの通信、エラーハンドリング
- **コンポーネント**:
  - Webhook クライアント
  - Bot API クライアント（将来実装）

---

## 2. モジュール設計

### 2.1 モジュール構成図

```
VBAProject (勤怠管理.xlsm)
│
├─ Module1_Setup
│   └─ 初期セットアップ、設定管理
│
├─ Module2_Import
│   └─ CSV読込、データ検証
│
├─ Module3_Notification
│   └─ 通知送信、履歴記録
│
├─ Module4_Master (将来実装)
│   └─ 社員マスタ管理、検索
│
├─ Module5_Utils (将来実装)
│   └─ 共通関数、ヘルパー
│
└─ ThisWorkbook
    └─ イベントハンドラ
```

### 2.2 Module1_Setup

**責務**: システムの初期設定とシート構造の生成

**主要関数**:

```vba
Public Sub InitialSetup()
  ' 初回セットアップ処理のエントリーポイント
  ' - シート作成
  ' - 設定初期化
  ' - ボタン配置
End Sub

Private Sub CreateSheets()
  ' 必要なシートを作成
  ' - 勤怠未入力
  ' - 社員マスタ
  ' - 通知履歴
  ' - 設定（非表示）
End Sub

Private Sub AddButtons()
  ' 操作ボタンを配置
  ' - CSV読込
  ' - 管理者通知
  ' - 個別通知
  ' - 履歴表示
End Sub

Private Function GetConfig(key As String) As String
  ' 設定値取得
  ' 引数: key - 設定キー（例: "WEBHOOK_URL_MANAGER"）
  ' 戻り値: 設定値
End Function
```

**データフロー**:
```
InitialSetup()
  ├─ CreateSheets()
  │   ├─ CreateSheet("勤怠未入力")
  │   ├─ CreateSheet("社員マスタ")
  │   ├─ CreateSheet("通知履歴")
  │   └─ CreateSheet("設定") → 非表示化
  ├─ SetupConfig()
  │   └─ 設定シートに初期値書き込み
  └─ AddButtons()
      └─ 各シートにボタン配置
```

---

### 2.3 Module2_Import

**責務**: CSV読込とデータ検証

**主要関数**:

```vba
Public Sub ImportCSV()
  ' CSV読込のエントリーポイント
  ' 1. ファイル選択
  ' 2. CSV解析
  ' 3. データ検証
  ' 4. シート書き込み
  ' 5. 社員マスタ更新
End Sub

Private Function ParseCSVLine(line As String) As Variant
  ' CSV行の解析
  ' 引数: line - CSV1行分の文字列
  ' 戻り値: 配列（列ごとの値）
End Function

Private Function ValidateData(data As Variant) As Boolean
  ' データ妥当性検証
  ' - 必須列の存在確認
  ' - 日付形式チェック
  ' - 重複チェック
End Function

Private Sub UpdateMasterFromImport(dataWs As Worksheet, masterWs As Worksheet)
  ' 社員マスタの自動更新
  ' - 未登録社員の抽出
  ' - マスタへの追加
End Sub
```

**処理フロー**:
```
ImportCSV()
  ├─ ファイル選択ダイアログ表示
  ├─ ファイルオープン
  ├─ 各行を処理
  │   ├─ ParseCSVLine() → 配列に変換
  │   ├─ ValidateData() → 妥当性チェック
  │   └─ シートに書き込み
  ├─ UpdateMasterFromImport()
  │   └─ 新規社員を社員マスタに追加
  └─ 完了メッセージ表示
```

**エラーハンドリング**:
- ファイルが選択されない → 処理中断、メッセージなし
- CSV形式エラー → 行番号と内容を表示、処理継続
- 必須列欠損 → エラーダイアログ、処理中断
- 日付形式エラー → 警告表示、該当行スキップ

---

### 2.4 Module3_Notification

**責務**: 通知の生成・送信と履歴記録

**主要関数**:

```vba
Public Sub SendManagerNotification()
  ' 管理者チャンネル通知のエントリーポイント
  ' 1. データ収集
  ' 2. 緊急度判定
  ' 3. メッセージ生成
  ' 4. 送信
  ' 5. 履歴記録
End Sub

Private Function GenerateManagerMessage(empData As Object) As String
  ' 管理者向けメッセージ生成
  ' 引数: empData - 社員別の未入力データ（Dictionary）
  ' 戻り値: フォーマット済みメッセージ
End Function

Public Function SendToLineWorks(webhookURL As String, messageText As String) As Boolean
  ' LINE WORKS Webhook送信
  ' 引数: 
  '   - webhookURL: Webhook URL
  '   - messageText: 送信メッセージ
  ' 戻り値: 成功=True、失敗=False
End Function

Private Sub LogNotification(notifyMethod As String, targetCount As Integer, result As String)
  ' 通知履歴の記録
  ' 引数:
  '   - notifyMethod: 通知方法（"管理者Ch" / "個別トーク"）
  '   - targetCount: 対象者数
  '   - result: 結果（"成功" / "失敗"）
End Sub

Public Sub SendIndividualNotifications()
  ' 個別トーク通知（フェーズ3）
  ' - Bot API認証
  ' - ユーザーID取得
  ' - 個別送信
End Sub
```

**通知送信フロー**:
```
SendManagerNotification()
  ├─ [勤怠未入力]シートからデータ取得
  ├─ 社員ごとにグルーピング
  │   └─ Dictionary<社員番号, {氏名, 未入力日リスト, 最大日数}>
  ├─ 緊急度判定
  │   ├─ 5日以上 → 🔴 緊急
  │   ├─ 3-4日 → 🟡 要注意
  │   └─ 1-2日 → 🟢 確認
  ├─ GenerateManagerMessage()
  │   └─ メッセージフォーマット適用
  ├─ 送信確認ダイアログ
  ├─ SendToLineWorks()
  │   ├─ JSON作成
  │   ├─ HTTP POST
  │   └─ レスポンス確認
  └─ LogNotification()
      └─ 通知履歴シートに記録
```

**メッセージフォーマット**:
```
【勤怠未入力アラート】{日付}

未入力者: {対象者数}名 / 未入力件数: {件数}件

■ 緊急対応（5日以上）
🔴 {氏名} さん
  ・{日付}（{経過日数}日前）
  ...

■ 要注意（3-4日）
🟡 {氏名} さん
  ...

■ 確認（1-2日）
🟢 {氏名} さん
  ...

━━━━━━━━━━━━━━━
※各管理者より該当者へ声掛けをお願いします
```

---

### 2.5 Module4_Master（将来実装）

**責務**: 社員マスタの高度な管理機能

**主要関数**:

```vba
Public Function SearchEmployee(empNo As String) As Variant
  ' 社員検索
  ' 引数: empNo - 社員番号
  ' 戻り値: 社員情報（配列）
End Function

Public Sub UpdateEmployee(empNo As String, email As String)
  ' 社員情報更新
End Sub

Public Function GetEmailByEmpNo(empNo As String) As String
  ' LINEWORKSメール取得
End Function

Public Sub ValidateEmailFormat(email As String) As Boolean
  ' メールアドレス形式検証
End Sub
```

---

## 3. データベース設計

### 3.1 論理データモデル（ER図）

```
┌─────────────────┐
│   社員マスタ     │
├─────────────────┤
│ 社員番号 (PK)   │───┐
│ 氏名            │   │
│ LINEWORKSメール │   │
│ 更新日          │   │
└─────────────────┘   │
                      │ 1
                      │
                      │
                      │ N
┌─────────────────┐   │
│  勤怠未入力     │───┘
├─────────────────┤
│ 社員番号 (FK)   │
│ 氏名            │
│ 日付            │
│ 曜日            │
│ コメント        │
│ 未入力日数      │
│ メール有無      │
└─────────────────┘
        │
        │ 1
        │
        │ N
        │
┌─────────────────┐
│   通知履歴      │
├─────────────────┤
│ 送信日時 (PK)   │
│ 通知方法        │
│ 対象者数        │
│ 対象日付        │
│ 結果            │
└─────────────────┘
```

### 3.2 物理データモデル（シート設計）

#### シート1: 勤怠未入力

| 列 | 列名 | データ型 | 制約 | 説明 | 計算式 |
|----|------|---------|------|------|--------|
| A | 社員番号 | Text | NOT NULL | 7桁の数字 | - |
| B | 氏名 | Text | NOT NULL | 最大50文字 | - |
| C | 日付 | Date | NOT NULL | yyyy/mm/dd | - |
| D | 曜日 | Text | - | 月/火/水/木/金/土/日 | - |
| E | コメント | Text | - | 未入力理由 | - |
| F | 未入力日数 | Number | - | 経過日数 | `=TODAY()-C2` |
| G | メール有無 | Text | - | ✓ or - | `=IFERROR(IF(VLOOKUP(A2,社員マスタ!A:C,3,FALSE)<>"","✓","-"),"-")` |

**インデックス**: A列（社員番号）でソート推奨
**データ保持**: 当日分のみ（次回CSV読込時に上書き）

---

#### シート2: 社員マスタ

| 列 | 列名 | データ型 | 制約 | 説明 | 計算式 |
|----|------|---------|------|------|--------|
| A | 社員番号 | Text | PRIMARY KEY | 7桁、一意 | - |
| B | 氏名 | Text | NOT NULL | 最大50文字 | - |
| C | LINEWORKSメール | Text | UNIQUE | メール形式 | - |
| D | 更新日 | Date | - | 最終更新日 | 手動/自動 |

**追加フィールド（参考）**:
- F1: `入力済み人数：`
- G1: `=COUNTA(C2:C1000)`

**インデックス**: A列（社員番号）主キー
**データ保持**: 永続

---

#### シート3: 通知履歴

| 列 | 列名 | データ型 | 制約 | 説明 |
|----|------|---------|------|------|
| A | 送信日時 | DateTime | NOT NULL | yyyy/mm/dd hh:mm:ss |
| B | 通知方法 | Text | NOT NULL | "管理者Ch" / "個別トーク" |
| C | 対象者数 | Number | NOT NULL | 正の整数 |
| D | 対象日付 | Date | NOT NULL | 未入力対象日 |
| E | 結果 | Text | NOT NULL | "成功" / "失敗" |

**インデックス**: A列（送信日時）降順
**データ保持**: 6ヶ月（手動削除）

---

#### シート4: 設定（非表示）

| 列 | 設定キー | 設定値 | 説明 |
|----|---------|--------|------|
| A1 | WEBHOOK_URL_MANAGER | (URL) | 管理者チャンネルWebhook URL |
| A2 | BOT_CLIENT_ID | (ID) | Bot APIクライアントID（将来実装） |
| A3 | BOT_CLIENT_SECRET | (Secret) | Bot APIシークレット（将来実装） |
| A4 | THRESHOLD_MENTION | 3 | メンション開始日数 |
| A5 | THRESHOLD_URGENT | 5 | 緊急判定日数 |

**セキュリティ**: シートを `xlSheetVeryHidden` で非表示化

---

## 4. API設計

### 4.1 LINE WORKS Webhook API

#### エンドポイント
```
POST https://talk.worksmobile.com/...（ユーザー固有）
```

#### リクエスト仕様

**Headers**:
```http
Content-Type: application/json; charset=UTF-8
```

**Body**:
```json
{
  "content": {
    "type": "text",
    "text": "メッセージ本文（最大10,000文字）"
  }
}
```

**VBA実装例**:
```vba
Function SendToLineWorks(webhookURL As String, messageText As String) As Boolean
    Dim httpRequest As Object
    Dim jsonBody As String
    Dim escapedText As String
    
    ' テキストエスケープ処理
    escapedText = Replace(messageText, "\", "\\")
    escapedText = Replace(escapedText, """", "\""")
    escapedText = Replace(escapedText, vbLf, "\n")
    escapedText = Replace(escapedText, vbCr, "")
    
    ' JSON構築
    jsonBody = "{""content"":{""type"":""text"",""text"":""" & escapedText & """}}"
    
    ' HTTP POST送信
    Set httpRequest = CreateObject("MSXML2.XMLHTTP.6.0")
    httpRequest.Open "POST", webhookURL, False
    httpRequest.setRequestHeader "Content-Type", "application/json; charset=UTF-8"
    httpRequest.send jsonBody
    
    ' レスポンス確認
    If httpRequest.Status = 200 Then
        SendToLineWorks = True
    Else
        Debug.Print "Error: " & httpRequest.Status & " - " & httpRequest.responseText
        SendToLineWorks = False
    End If
    
    Set httpRequest = Nothing
End Function
```

#### レスポンス仕様

**成功時**:
```http
HTTP/1.1 200 OK
Content-Type: application/json

{}
```

**エラー時**:
```http
HTTP/1.1 400 Bad Request
Content-Type: application/json

{
  "code": "INVALID_PARAMETER",
  "message": "Invalid content format"
}
```

**エラーコード一覧**:
| コード | HTTP Status | 説明 | 対処方法 |
|--------|------------|------|----------|
| INVALID_PARAMETER | 400 | パラメータ不正 | JSON形式を確認 |
| UNAUTHORIZED | 401 | 認証エラー | Webhook URL再発行 |
| NOT_FOUND | 404 | URLが無効 | Webhook URL確認 |
| RATE_LIMIT_EXCEEDED | 429 | レート制限超過 | 送信間隔を調整 |
| INTERNAL_SERVER_ERROR | 500 | サーバーエラー | 時間を置いて再試行 |

---

### 4.2 LINE WORKS Bot API（将来実装）

#### 認証フロー

```
┌─────────┐                              ┌──────────────┐
│ VBA App │                              │ LINE WORKS   │
└────┬────┘                              └──────┬───────┘
     │                                          │
     │ 1. Client Credentials 送信              │
     │─────────────────────────────────────────>│
     │                                          │
     │ 2. Access Token 返却                    │
     │<─────────────────────────────────────────│
     │                                          │
     │ 3. API リクエスト（Token付き）           │
     │─────────────────────────────────────────>│
     │                                          │
     │ 4. レスポンス                            │
     │<─────────────────────────────────────────│
```

#### エンドポイント

**トークン取得**:
```
POST https://auth.worksmobile.com/oauth2/v2.0/token
```

**メッセージ送信**:
```
POST https://www.worksapis.com/v1.0/bots/{botId}/users/{userId}/messages
```

#### リクエスト例（トークン取得）

```http
POST /oauth2/v2.0/token HTTP/1.1
Host: auth.worksmobile.com
Content-Type: application/x-www-form-urlencoded

grant_type=client_credentials
&client_id={CLIENT_ID}
&client_secret={CLIENT_SECRET}
&scope=bot
```

#### リクエスト例（メッセージ送信）

```http
POST /v1.0/bots/{botId}/users/{userId}/messages HTTP/1.1
Host: www.worksapis.com
Authorization: Bearer {ACCESS_TOKEN}
Content-Type: application/json

{
  "content": {
    "type": "text",
    "text": "個別メッセージ本文"
  }
}
```

---

## 5. セキュリティ設計

### 5.1 脅威モデル

#### 脅威分析（STRIDE）

| 脅威 | カテゴリ | リスク | 対策 |
|------|---------|--------|------|
| Webhook URL漏洩 | 情報漏洩 | 高 | 非表示シート、定期再発行 |
| VBAコード改ざん | 改ざん | 中 | VBAプロジェクトのパスワード保護 |
| 社員情報の不正アクセス | 情報漏洩 | 高 | ファイルアクセス権限設定 |
| 通信傍受 | 情報漏洩 | 低 | HTTPS通信（標準） |
| なりすまし送信 | なりすまし | 中 | Webhook URLの厳重管理 |
| サービス妨害 | DoS | 低 | レート制限遵守 |

### 5.2 セキュリティ対策

#### データ保護

```vba
' 設定シートの非表示化
Sub HideConfigSheet()
    ThisWorkbook.Sheets("設定").Visible = xlSheetVeryHidden
End Sub

' VBAプロジェクトの保護
' ツール > VBAProjectのプロパティ > 保護タブ
' ☑ プロジェクトを表示用にロックする
' パスワード設定
```

#### 通信セキュリティ

- **HTTPS必須**: Webhook URLは必ずhttpsスキーム
- **証明書検証**: MSXML2.XMLHTTP.6.0が自動検証
- **タイムアウト設定**: 30秒でタイムアウト

```vba
' タイムアウト設定例
httpRequest.setTimeouts 30000, 30000, 30000, 30000
' 引数: 解決, 接続, 送信, 受信（ミリ秒）
```

#### アクセス制御

| リソース | 読取 | 書込 | 実行 |
|---------|------|------|------|
| Excel ファイル | 全社員 | 人事担当のみ | 人事担当のみ |
| 設定シート | システムのみ | 管理者のみ | - |
| 社員マスタ | 人事担当 | 人事担当 | - |
| VBAコード | 開発者のみ | 開発者のみ | 全社員 |

#### 監査ログ

```vba
' デバッグログ出力
Sub LogActivity(action As String, details As String)
    Debug.Print Now & " - " & action & ": " & details
End Sub

' 使用例
Call LogActivity("CSV読込", "30件のデータを読み込み")
Call LogActivity("通知送信", "管理者チャンネルに送信成功")
Call LogActivity("エラー", "Webhook送信失敗: HTTP 404")
```

---

## 6. エラーハンドリング設計

### 6.1 エラー分類

| カテゴリ | 重要度 | 対応方針 |
|---------|--------|----------|
| 入力エラー | 低 | ユーザーに修正を促す |
| データエラー | 中 | 該当データをスキップ |
| 通信エラー | 高 | リトライ可能性を判定 |
| システムエラー | 最高 | 処理中断、管理者通知 |

### 6.2 エラーハンドリングパターン

#### パターン1: 基本エラーハンドリング

```vba
Sub SampleFunction()
    On Error GoTo ErrorHandler
    
    ' 処理本体
    ' ...
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Debug.Print "Error " & Err.Number & ": " & Err.Description
End Sub
```

#### パターン2: リトライ付きエラーハンドリング

```vba
Function SendWithRetry(url As String, message As String, maxRetries As Integer) As Boolean
    Dim attempt As Integer
    Dim success As Boolean
    
    For attempt = 1 To maxRetries
        success = SendToLineWorks(url, message)
        
        If success Then
            SendWithRetry = True
            Exit Function
        End If
        
        ' 指数バックオフ
        Application.Wait Now + TimeValue("00:00:" & (2 ^ attempt))
    Next attempt
    
    SendWithRetry = False
End Function
```

#### パターン3: カスタムエラーメッセージ

```vba
Function ShowFriendlyError(errNumber As Long, errDescription As String) As String
    Dim message As String
    
    Select Case errNumber
        Case -2147352567 ' VBAアクセスエラー
            message = "VBAプロジェクトへのアクセスが拒否されました。" & vbLf & vbLf & _
                     "【対処方法】" & vbLf & _
                     "1. Excelオプション > トラストセンター > マクロの設定" & vbLf & _
                     "2. 'VBAプロジェクトオブジェクトモデルへのアクセスを信頼する' にチェック"
        
        Case 53 ' ファイルが見つからない
            message = "指定されたファイルが見つかりません。" & vbLf & _
                     "ファイルパスを確認してください。"
        
        Case Else
            message = "エラーが発生しました: " & errDescription
    End Select
    
    ShowFriendlyError = message
End Function
```

### 6.3 エラーコード定義

| エラーコード | 説明 | 対処方法 |
|------------|------|----------|
| ERR_CSV_001 | CSV形式エラー | CSV形式を確認 |
| ERR_CSV_002 | 必須列欠損 | CSVテンプレート使用 |
| ERR_WH_001 | Webhook送信失敗 | URL確認、再送信 |
| ERR_WH_002 | レート制限超過 | 送信間隔を調整 |
| ERR_MASTER_001 | 社員マスタ不整合 | マスタ再構築 |
| ERR_SYS_001 | システムエラー | 管理者に連絡 |

---

## 7. パフォーマンス設計

### 7.1 パフォーマンス目標

| 処理 | 目標時間 | 測定条件 |
|------|---------|----------|
| CSV読込 | 5秒以内 | 1,000件 |
| 通知送信 | 3秒以内 | 1回 |
| 社員マスタ検索 | 0.1秒以内 | 200件中1件 |
| 画面更新 | 1秒以内 | ボタンクリック後 |

### 7.2 最適化手法

#### 画面更新の停止

```vba
Sub OptimizedProcess()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 処理本体
    ' ...
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
```

#### 配列処理の活用

```vba
' ❌ 遅い方法（セル単位）
For i = 2 To lastRow
    ws.Cells(i, 1).Value = data(i - 2, 0)
Next i

' ✅ 速い方法（配列一括）
ws.Range("A2:A" & lastRow).Value = data
```

#### Dictionary による高速検索

```vba
' 社員マスタをDictionaryにキャッシュ
Function LoadMasterToDictionary() As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets("社員マスタ")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        dict.Add ws.Cells(i, 1).Value, ws.Cells(i, 3).Value ' 社員番号 → メール
    Next i
    
    Set LoadMasterToDictionary = dict
End Function
```

---

## 8. テスト設計

### 8.1 テストレベル

```
┌────────────────────────────────┐
│     受入テスト                  │
│  ・エンドツーエンド             │
│  ・ユーザーシナリオ             │
└───────────┬────────────────────┘
            │
┌───────────┴────────────────────┐
│   統合テスト                    │
│  ・モジュール間連携             │
│  ・外部API統合                 │
└───────────┬────────────────────┘
            │
┌───────────┴────────────────────┐
│   単体テスト                    │
│  ・関数レベル                   │
│  ・ロジック検証                 │
└────────────────────────────────┘
```

### 8.2 テストケース設計

#### CSV読込機能のテストケース

| ID | テスト項目 | 入力 | 期待結果 | 優先度 |
|----|-----------|------|----------|--------|
| TC-CSV-001 | 正常なCSV読込 | 標準形式CSV | 全件読込成功 | 高 |
| TC-CSV-002 | 空ファイル | 0件CSV | エラーメッセージ | 中 |
| TC-CSV-003 | 大量データ | 1,000件CSV | 5秒以内に完了 | 高 |
| TC-CSV-004 | 文字コード | UTF-8 CSV | 文字化けなし | 中 |
| TC-CSV-005 | 日付形式エラー | 不正日付含む | 該当行スキップ | 高 |
| TC-CSV-006 | 重複データ | 同一行が2回 | 警告表示 | 低 |

#### 通知送信機能のテストケース

| ID | テスト項目 | 入力 | 期待結果 | 優先度 |
|----|-----------|------|----------|--------|
| TC-NTF-001 | 正常送信 | 3名の未入力者 | 送信成功 | 高 |
| TC-NTF-002 | Webhook URLエラー | 無効なURL | エラーメッセージ | 高 |
| TC-NTF-003 | 緊急度分類 | 1日/3日/5日前 | 正しく分類 | 高 |
| TC-NTF-004 | メッセージ形式 | - | フォーマット正常 | 中 |
| TC-NTF-005 | 履歴記録 | 送信成功後 | 履歴追加 | 中 |
| TC-NTF-006 | ネットワークエラー | 通信断 | リトライ可能表示 | 低 |

### 8.3 テストデータ

#### 標準テストデータ（test_data.csv）

```csv
社員番号,氏名,日付,曜日区分,届出内容,コメント,出勤時刻,退勤時刻
1904010,テスト１,2025/10/15,水,,出勤時刻と退勤時刻の両方が入力されていません,,
1910003,テスト２,2025/10/15,水,,出勤時刻と退勤時刻の両方が入力されていません,,
1710002,テスト３,2025/10/14,火,,出勤時刻と退勤時刻の両方が入力されていません,,
1710002,テスト３,2025/10/15,水,,出勤時刻と退勤時刻の両方が入力されていません,,
```

#### 異常系テストデータ

```csv
社員番号,氏名,日付,曜日区分,届出内容,コメント,出勤時刻,退勤時刻
1904010,テスト１,2025/13/40,水,,不正な日付形式,,
,テスト２,2025/10/15,水,,社員番号欠損,,
1710002,テスト３,invalid_date,火,,不正な日付,,
```

---

## 9. デプロイ設計

### 9.1 デプロイメントアーキテクチャ

```
┌──────────────────────────────────────────┐
│         開発環境                          │
│  ┌────────────────────────────────────┐  │
│  │  勤怠管理_dev.xlsm                  │  │
│  │  ・テストデータ                     │  │
│  │  ・デバッグ機能有効                 │  │
│  └────────────────────────────────────┘  │
└────────────┬─────────────────────────────┘
             │ VBAエクスポート
             ↓
┌──────────────────────────────────────────┐
│         GitHub                            │
│  ┌────────────────────────────────────┐  │
│  │  VBAコードリポジトリ                 │  │
│  │  ・Module1_Setup.bas                │  │
│  │  ・Module2_Import.bas               │  │
│  │  ・Module3_Notification.bas         │  │
│  │  ・ドキュメント                      │  │
│  └────────────────────────────────────┘  │
└────────────┬─────────────────────────────┘
             │ VBAインポート
             ↓
┌──────────────────────────────────────────┐
│         本番環境                          │
│  ┌────────────────────────────────────┐  │
│  │  勤怠管理.xlsm                       │  │
│  │  ・本番データ                        │  │
│  │  ・デバッグ機能無効                  │  │
│  └────────────────────────────────────┘  │
└──────────────────────────────────────────┘
```

### 9.2 デプロイ手順

#### Step 1: VBAコードのエクスポート（開発環境）

```bash
# Pythonツールを使用
python extract_vba_code.py

# 出力先確認
# 勤怠管理_VBA_YYYYMMDD_HHMMSS/
#   ├─ Module1_Setup.bas
#   ├─ Module2_Import.bas
#   ├─ Module3_Notification.bas
#   └─ _README.txt
```

#### Step 2: Gitへのコミット

```bash
cd vba_src
git add .
git commit -m "feat: 管理者通知機能実装"
git push origin main
```

#### Step 3: 本番環境への適用

```bash
# 1. Gitから最新版取得
git pull origin main

# 2. 本番xlsmファイルのバックアップ
copy 勤怠管理.xlsm 勤怠管理_backup_YYYYMMDD.xlsm

# 3. VBAコードのインポート
python import_vba_code.py 勤怠管理.xlsm ./vba_src
```

#### Step 4: 動作確認

- [ ] CSV読込テスト
- [ ] テスト通知送信
- [ ] エラーハンドリング確認
- [ ] パフォーマンス測定

### 9.3 ロールバック手順

```bash
# バックアップファイルから復元
copy 勤怠管理_backup_YYYYMMDD.xlsm 勤怠管理.xlsm
```

---

## 10. 運用設計

### 10.1 日常運用フロー

#### 毎日（平日）9:00

```
1. 勤怠システムにログイン
2. [レポート] > [勤怠未入力レポート] > [CSV出力]
3. 勤怠管理.xlsm を開く
4. [📥 CSV読込] ボタンをクリック
5. ダウンロードしたCSVファイルを選択
6. 未入力者リストを確認
7. [📤 管理者通知] ボタンをクリック
8. 確認ダイアログで [はい]
9. LINE WORKSで送信確認
```

**所要時間**: 約3分

#### 週1回（金曜日）

```
1. 社員マスタシートを開く
2. 新入社員のLINEWORKSメール入力
3. 入力状況確認（G1セル）
4. 通知履歴シートで送信状況確認
```

**所要時間**: 約5分

### 10.2 監視項目

| 項目 | 確認方法 | 閾値 | 対応 |
|------|---------|------|------|
| 通知送信成功率 | 通知履歴シート | 95%以上 | 失敗時は原因調査 |
| CSV読込時間 | 体感 | 10秒以内 | 超過時はデータ量確認 |
| 未入力者数 | 日次推移 | 前週比+20%以内 | 超過時は要因分析 |
| エラー発生件数 | デバッグログ | 週3件以内 | 超過時は改善検討 |

### 10.3 保守タスク

#### 月次

- [ ] Webhook URLの有効性確認
- [ ] 通知履歴の整理（6ヶ月以上経過分削除）
- [ ] VBAコードのバックアップ
- [ ] パフォーマンス測定

#### 四半期

- [ ] 社員マスタの整合性チェック
- [ ] セキュリティパッチ適用（Excel）
- [ ] 運用マニュアルの更新

#### 年次

- [ ] システム全体レビュー
- [ ] ユーザー満足度調査
- [ ] 改善計画策定

---

## 11. 移行設計（将来のフェーズ）

### 11.1 Phase 2 → Phase 3 移行（個別通知機能追加）

#### 前提条件チェックリスト

- [ ] Bot API設定完了
- [ ] Developer Console アクセス権取得
- [ ] 社員マスタの80%以上でメール入力完了
- [ ] テスト環境での動作確認完了

#### 移行手順

```
1. Bot作成（LINE WORKS管理画面）
2. Client ID / Secret 取得
3. 設定シートに認証情報登録
4. Module3_Notification に個別送信機能追加
5. 段階的ロールアウト
   - Week 1: テスト部署（5名）
   - Week 2: 本部（30名）
   - Week 3: 全社展開（200名）
```

### 11.2 データ移行計画

**現状**: Excelシート
**将来**: データベース化（検討中）

```
┌─────────────┐      移行       ┌─────────────┐
│  Excel      │  ──────────>    │  SQLite     │
│  ・シート管理│                 │  ・テーブル管理│
│  ・手動管理 │                 │  ・SQL操作   │
└─────────────┘                 └─────────────┘
```

**移行シナリオ**（Phase 4以降）:
1. SQLiteデータベース作成
2. 社員マスタ移行
3. 通知履歴移行
4. Excel → DB接続モジュール実装
5. 並行稼働期間（1ヶ月）
6. 完全移行

---

## 12. ドキュメント管理

### 12.1 ドキュメント一覧

| ドキュメント | ファイル名 | 更新頻度 | 管理者 |
|------------|-----------|---------|--------|
| README | README.md | リリース毎 | 開発者 |
| 要件定義書 | requirements.md | Phase毎 | PM |
| ユーザーストーリー | user-stories.md | Sprint毎 | PO |
| システム設計書 | system-design.md | 機能追加時 | アーキテクト |
| API仕様書 | api-specification.md | API変更時 | 開発者 |
| 運用マニュアル | operation-manual.md | 四半期 | 運用担当 |
| トラブルシューティング | troubleshooting.md | 随時 | サポート |

### 12.2 変更管理プロセス

```
┌──────────────┐
│ 変更要求受付 │
└──────┬───────┘
       │
       ↓
┌──────────────┐
│ 影響範囲分析 │
└──────┬───────┘
       │
       ↓
┌──────────────┐
│ 設計書更新   │
└──────┬───────┘
       │
       ↓
┌──────────────┐
│ レビュー     │
└──────┬───────┘
       │
       ↓
┌──────────────┐
│ 承認・実装   │
└──────────────┘
```

---

## 改訂履歴

| バージョン | 日付 | 変更内容 | 作成者 |
|-----------|------|----------|--------|
| 1.0 | 2025/10/17 | 初版作成 | [担当者名] |

---

## 承認

| 役割 | 氏名 | 承認日 | 署名 |
|------|------|--------|------|
| システムアーキテクト | [氏名] | 2025/10/17 | |
| 開発リーダー | [氏名] | 2025/10/17 | |
| 品質保証責任者 | [氏名] | 2025/10/17 | |

