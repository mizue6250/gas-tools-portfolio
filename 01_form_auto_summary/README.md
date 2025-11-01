# 🧾 フォーム回答 自動要約ツール（Form Auto Summary）

Google フォームやスプレッドシートに蓄積された回答を自動で収集し、  
ChatGPT API（gpt-4o-mini）を使って Markdown 形式で要約、  
Google ドキュメント＋PDF化してメール送信するツールです。  

社内アンケート、問い合わせログ、ヒアリング記録などを  
**自動で整理・共有するレポート生成ワークフロー**として活用できます。

---

## 🚀 主な機能

| 機能 | 説明 |
|------|------|
| ✅ 回答の自動収集 | 指定シート（`FormResponses`）から最新日付分の回答を抽出 |
| ✅ ChatGPTによる要約 | gpt-4o-mini API を呼び出して要約文を生成 |
| ✅ Markdown → Google Docs 変換 | 自動生成された要約をドキュメント化 |
| ✅ PDFレポート添付メール送信 | Gmail 経由で PDF を自動送信 |
| ✅ 自動トリガー設定 | 毎朝9時に自動実行（`setupTrigger_Summarizer0900()`） |
| ✅ デモデータ生成 | `seedFormResponsesDemo()` で即動作テスト可能 |

---

## 🧰 使用技術

- **Google Apps Script (GAS)**
  - SpreadsheetApp / GmailApp / DocumentApp / DriveApp / UrlFetchApp  
- **OpenAI API**
  - Chat Completions エンドポイント（gpt-4o-mini）  
- **出力形式**
  - Markdown → Google Docs → PDF  

---

## ⚙️ セットアップ手順

1. **スプレッドシートを準備**
   - シート名：`FormResponses`
   - 列構成：
     | A | B | C |
     |---|---|---|
     | Timestamp | 名前 | 本文 |

2. **Apps Script エディタを開く**
   - スプレッドシート上で「拡張機能 → Apps Script」を開き、
     `src/Code.gs` の内容を貼り付けて保存。

3. **APIキー設定**
   - メニュー：`ファイル → プロジェクトのプロパティ → スクリプトのプロパティ`
   - `OPENAI_API_KEY` を追加し、OpenAIのAPIキーを入力。

4. **メール送信先設定**
   ```js
   mail: {
     to: 'yourname@example.com', // 送信先メール
     subjectPrefix: '【要約レポート】'
   }

---

5. **動作確認**

seedFormResponsesDemo() を実行し、ダミーデータを生成。
summarizeFormResponsesAndSend() を実行し、PDF付きメールを確認。

---

6. **自動化設定**

初回のみ setupTrigger_Summarizer0900() を実行。
→ 以後、毎朝9時に自動レポート送信。

---

## 🧪 デモ結果サンプル

| 出力物 | 内容 |
|--------|------|
| 📄 Googleドキュメント | Markdown要約を自動整形した日次レポート |
| 📎 PDF添付メール | 同レポートをPDF化し、自動送信 |

---

## 🧱 設定オプション（`SUMCFG`）

| キー | 説明 | 例 |
|------|------|------|
| `sheetName` | 回答シート名 | `"FormResponses"` |
| `bodyColIndex` | 本文列の番号 | `3` |
| `fromDays` | 何日前を要約対象にするか | `0`（当日） |
| `tz` | タイムゾーン | `"Asia/Tokyo"` |
| `model` | 使用モデル | `"gpt-4o-mini"` |
| `reportTitle` | ドキュメントのタイトル | `"フォーム回答 自動要約レポート"` |

---