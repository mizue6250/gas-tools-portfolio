# gas-tools-portfolio
Portfolio of Google Apps Script tools for business automation / 日報・在庫・フォーム要約の自動化ツール集

### by Mizue Oishi

Google Apps Script（GAS）とChatGPT APIを活用した業務自動化ツールのデモ集です。  
実際の業務でよくある「日報・フォーム集計・在庫管理」などの手間を、  
GAS × AI でどこまで効率化できるかをテーマに制作しています。

---

## 💼 プロジェクト一覧

| No | ツール名 | 概要 | 主な技術 |
|----|-----------|-------|-----------|
| ① | [AI_Report_AutoSummary](./01_AI_Report_AutoSummary) | Googleフォーム回答をChatGPTで要約し、PDFレポートとしてGmail送信 | GAS / ChatGPT API / DocsApp / GmailApp |
| ② | [DailyReport_AutoMailer](./02_DailyReport_AutoMailer) | スプレッドシートの日報を自動集計し、HTMLメールで送信 | GAS / GmailApp / Trigger |
| ③ | [Inventory_Sales_AutoDashboard](./03_Inventory_Sales_AutoDashboard) | 在庫・受注データを自動集計してダッシュボード化 | GAS / SpreadsheetApp / Custom Functions |

---

## ⚙️ 使用技術・API
- **Google Apps Script（JavaScriptベース）**
- **ChatGPT API（OpenAI GPT-4）**
- **Google Workspace 標準API**
  - GmailApp / SpreadsheetApp / DocsApp / DriveApp

---

## 🧩 ツールの特徴
- **ノーコード運用可**：GASトリガーで自動処理  
- **軽量設計**：中小規模チーム・在宅業務でも動作可能  
- **AI活用例**：ChatGPT APIを業務文脈で活用（要約・レポート・自然言語出力）

---

## 🖼️ デモプレビュー
現在、スクリーンショットを準備中です。  
近日中に以下のデモ例を追加予定です 👇
- `AI_Report_AutoSummary/screenshots/mail_output.png`
- `DailyReport_AutoMailer/screenshots/daily_mail_sample.png`
- `Inventory_Sales_AutoDashboard/screenshots/dashboard_view.png`

---

## 🧠 詳細解説（Notionで見る）
より詳しい解説・背景・設計意図はこちら：
👉 [Notion ポートフォリオページ](https://www.notion.so/AI-Mizue-Oishi-291e5d8709fb809fb941f9eac5a2f5f8)

---

## 📬 Contact
**Mizue Oishi**  
📧 osoranokanatahe@gmail.com  
🌐 [GitHub Profile](https://github.com/mizue6250)

---

## 📝 License
このリポジトリのソースコードは学習・ポートフォリオ目的で公開しています。  
商用利用・再配布を行う場合は、事前にご相談ください。
