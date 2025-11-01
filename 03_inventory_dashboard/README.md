# 📦 在庫・売上集計ダッシュボード（Inventory Pipeline）

Amazon / BASE の受注を正規化して1枚の `Sales` シートに集計し、  
**KPI（昨日売上・昨日利益・今月売上・在庫アラート数）** を自動算出。  
在庫が閾値を下回った SKU を **メール（任意で Slack）** に通知します。

---

## 🚀 主な機能

| 機能 | 説明 |
|------|------|
| ✅ ダミーデータ投入 | `seedInventoryDemoData()` で各シートのサンプルを自動作成 |
| ✅ 受注の正規化 | `normalizeOrdersToSales()` で Amazon/BASE → `Sales` に統合 |
| ✅ KPI算出 | `buildInventoryDashboardAndAlert()` が昨日/今月の売上・利益を集計 |
| ✅ 在庫アラート | 閾値以下の SKU を抽出し、メール＆Slack通知（任意） |
| ✅ 一括実行 | `runAll_InventoryPipeline()` で「正規化→集計→通知」を一発実行 |
| ✅ 自動化 | `setupTrigger_Inventory0900()` で毎朝9時に自動実行 |

---

## 🧰 使用技術

- **Google Apps Script (GAS)**  
  SpreadsheetApp / GmailApp / UrlFetchApp / Utilities / ScriptApp
- **通知**  
  Gmail（必須）／ Slack Incoming Webhook（任意）

---

## ⚙️ セットアップ手順

1. **スプレッドシートを用意**（空でOK）
2. **Apps Script を開く**  
   スプレッドシート →「拡張機能 → Apps Script」→ `src/Code.gs` を貼り付けて保存
3. **設定（公開用）を見直す**  
   `INV` オブジェクト内を必要に応じて変更
   ```js
   const INV = {
     tz: 'Asia/Tokyo',
     mail: { to: 'yourname@example.com', subjectPrefix: '【在庫アラート】', sendAsDraft: true },
     slackWebhookUrl: '', // 使う場合のみURLを入れる
     sheet: {
       products: 'Products', ordersAmazon: 'Orders_Amazon',
       ordersBASE: 'Orders_BASE', sales: 'Sales',
       stockMov: 'StockMovements', dashboard: 'Dashboard'
     }
   };
