# 📧 日報自動送信ツール（Daily Report Mailer）

Google スプレッドシート上の「DailyLog」シートを読み取り、  
チームメンバーの日報を HTML テーブル形式で自動集計し、Gmail 経由で送信または下書き作成するツールです。  
デモデータ生成・毎朝9時の自動送信トリガー設定も含まれています。

---

## 🚀 主な機能

| 機能 | 説明 |
|------|------|
| ✅ ダミーデータ生成 | 30日分×6名のサンプルデータを自動投入（`seedDailyLogDemoData()`） |
| ✅ 日報HTML生成 | 当日のログのみ抽出してHTMLテーブルを構築（`buildDailyReportHtml_()`） |
| ✅ 自動メール送信 | Gmailで日報を自動送信または下書き作成（`sendDailyReport()`） |
| ✅ トリガー設定 | 毎朝9時に自動送信（`setupTriggerEveryMorning0900()`） |
| ✅ テストモード | `sendAsDraft: true` に設定すると、まず下書きのみ作成 |

---

## 🧰 使用技術

- Google Apps Script (GAS)
  - SpreadsheetApp / GmailApp / ScriptApp / Utilities
- HTML + CSS によるメール整形
- ローカルタイムゾーン: `Asia/Tokyo`

---

## ⚙️ セットアップ手順

1. **スプレッドシートを用意**
   - シート名：`DailyLog`
   - 列構成：  
     | A | B | C | D |
     |---|---|---|---|
     | 日付 | 担当 | タスク | 進捗/メモ |

2. **Apps Scriptエディタを開く**  
   - スプレッドシート上で「拡張機能 → Apps Script」を開く  
   - `src/Code.gs` の内容を貼り付けて保存

3. **メール設定を変更**
   ```js
   const CFG = {
     mail: {
       to: 'メールアドレス',
       cc: '',
       subjectPrefix: '【日報自動送信】',
       sendAsDraft: true // ←最初はtrue（下書き確認モード）
     },
     businessName: '日報デモ'
   };
