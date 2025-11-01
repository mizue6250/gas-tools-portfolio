const INV = {
  tz: 'Asia/Tokyo',
  mail: {
    to: 'yourname@example.com', // デモ用宛先（公開時に個人アドレスは入れない）
    subjectPrefix: '【在庫アラート】',
    sendAsDraft: true           // true=下書きのみ、false=本送信
  },
  slackWebhookUrl: '',          // 任意：Slack Incoming Webhook（空なら送信しない）
  sheet: {
    products: 'Products',
    ordersAmazon: 'Orders_Amazon',
    ordersBASE: 'Orders_BASE',
    sales: 'Sales',
    stockMov: 'StockMovements',
    dashboard: 'Dashboard'
  }
};

/**
 * =============================
 * A. ダミーデータ投入（最初に1回）
 * =============================
 */
function seedInventoryDemoData() {
  const ss = SpreadsheetApp.getActive();
  const shP = getOrCreateSheet_(ss, INV.sheet.products);
  const shA = getOrCreateSheet_(ss, INV.sheet.ordersAmazon);
  const shB = getOrCreateSheet_(ss, INV.sheet.ordersBASE);
  const shM = getOrCreateSheet_(ss, INV.sheet.stockMov);
  const shS = getOrCreateSheet_(ss, INV.sheet.sales);
  const shD = getOrCreateSheet_(ss, INV.sheet.dashboard);

  // Products
  shP.clear();
  shP.getRange(1, 1, 1, 6).setValues([['SKU', '商品名', '仕入単価', '販売単価', '閾値', '初期在庫']]);
  const prows = [
    ['A-001', 'コーヒー豆 200g', 600, 1200, 10, 100],
    ['A-002', 'カフェラテ粉 250g', 700, 1300, 8, 60],
    ['B-001', 'マグカップ', 400, 900, 5, 40]
  ];
  shP.getRange(2, 1, prows.length, 6).setValues(prows);

  // Orders_Amazon
  shA.clear();
  shA.getRange(1, 1, 1, 6).setValues([['注文日', '注文ID', 'SKU', '数量', '販売価格', '手数料(%)']]);
  const tz = INV.tz;
  const today = new Date();
  const recentA = [];
  for (let i = 0; i < 10; i++) {
    const d = Utilities.formatDate(new Date(today.getTime() - i * 86400000), tz, 'yyyy-MM-dd');
    recentA.push([d, `AMZ-${1000 + i}`, 'A-001', 1 + Math.floor(Math.random() * 2), 1200, 10]);
  }
  shA.getRange(2, 1, recentA.length, 6).setValues(recentA);

  // Orders_BASE
  shB.clear();
  shB.getRange(1, 1, 1, 7).setValues([['注文日', '受注番号', 'SKU', '数量', '単価', '送料', '割引']]);
  const recentB = [];
  for (let i = 0; i < 10; i++) {
    const d = Utilities.formatDate(new Date(today.getTime() - i * 86400000), tz, 'yyyy-MM-dd');
    recentB.push([d, `BASE-${2000 + i}`, 'B-001', 1, 900, 0, 0]);
  }
  shB.getRange(2, 1, recentB.length, 7).setValues(recentB);

  // StockMovements
  shM.clear();
  shM.getRange(1, 1, 1, 5).setValues([['日付', 'SKU', '区分', '数量', '備考']]);
  shM.getRange(2, 1, 3, 5).setValues([
    [Utilities.formatDate(today, tz, 'yyyy-MM-dd'), 'A-001', '入庫', 20, '補充'],
    [Utilities.formatDate(today, tz, 'yyyy-MM-dd'), 'A-002', '調整', -2, '破損'],
    [Utilities.formatDate(today, tz, 'yyyy-MM-dd'), 'B-001', '入庫', 10, '補充']
  ]);

  // Sales/Dashboard ヘッダー
  shS.clear();
  shS.getRange(1, 1, 1, 9).setValues([['注文日', 'チャネル', '注文ID', 'SKU', '数量', '売上金額', '手数料', '送料', '割引']]);
  shD.clear();
  shD.getRange(1, 1, 1, 2).setValues([['KPI', '値']]);

  Logger.log('Seed completed.');
}

/**
 * =============================
 * B. 正規化（Orders_* → Sales）
 * =============================
 */
function normalizeOrdersToSales() {
  const ss = SpreadsheetApp.getActive();
  const shA = ss.getSheetByName(INV.sheet.ordersAmazon);
  const shB = ss.getSheetByName(INV.sheet.ordersBASE);
  const shS = ss.getSheetByName(INV.sheet.sales);
  if (!shA || !shB || !shS) throw new Error('Order or Sales sheet missing');

  const aVals = readData_(shA);
  const bVals = readData_(shB);

  // Amazon → Sales
  const aRows = aVals.map(r => {
    const [date, id, sku, qty, price, feePct] = r;
    const q = toNum_(qty), p = toNum_(price), fPct = toNum_(feePct);
    const amount = p * q;
    const fee = Math.round(amount * (fPct / 100));
    return [date, 'Amazon', id, sku, q, amount, fee, 0, 0];
  });

  // BASE → Sales
  const bRows = bVals.map(r => {
    const [date, id, sku, qty, price, shipping, discount] = r;
    const q = toNum_(qty), p = toNum_(price);
    return [date, 'BASE', id, sku, q, p * q, 0, toNum_(shipping), toNum_(discount)];
  });

  const rows = [...aRows, ...bRows].sort((x, y) => String(x[0]).localeCompare(String(y[0])));

  // 既存データ消去 → 書き込み
  const last = shS.getLastRow();
  if (last > 1) shS.getRange(2, 1, last - 1, 9).clearContent();
  if (rows.length) shS.getRange(2, 1, rows.length, 9).setValues(rows);

  Logger.log(`Normalized rows: ${rows.length}`);
}

/**
 * =============================
 * C. 集計 & KPI & 在庫アラート
 * =============================
 */
function buildInventoryDashboardAndAlert() {
  const ss = SpreadsheetApp.getActive();
  const tz = INV.tz;
  const shP = ss.getSheetByName(INV.sheet.products);
  const shS = ss.getSheetByName(INV.sheet.sales);
  const shM = ss.getSheetByName(INV.sheet.stockMov);
  const shD = ss.getSheetByName(INV.sheet.dashboard);

  const products = objectBySku_(readData_(shP, true)); // {SKU:{...}}
  const sales = readData_(shS);
  const moves = readData_(shM);

  // 売上・利益集計（昨日/今月）
  const today = new Date();
  const todayStr = fmtDate_(today, tz);
  const yesterdayStr = fmtDate_(new Date(today.getTime() - 86400000), tz);
  const monthPrefix = Utilities.formatDate(today, tz, 'yyyy-MM');

  let salesYesterday = 0, profitYesterday = 0, salesMonth = 0;

  sales.forEach(r => {
    const [date, , , sku, qty, amount, fee, ship, disc] = r;
    const p = products[sku] || {};
    const q = toNum_(qty), a = toNum_(amount), f = toNum_(fee), s = toNum_(ship), d = toNum_(disc);
    const cost = toNum_(p['仕入単価']) * q;
    const gross = a - f - s - d;
    const profit = gross - cost;

    if (String(date) === yesterdayStr) {
      salesYesterday += a;
      profitYesterday += profit;
    }
    if (String(date).startsWith(monthPrefix)) salesMonth += a;
  });

  // 在庫計算
  const soldBySku = {};
  sales.forEach(r => {
    const sku = r[3]; const q = toNum_(r[4]);
    soldBySku[sku] = (soldBySku[sku] || 0) + q;
  });

  const moveBySku = {};
  moves.forEach(r => {
    const sku = r[1]; const q = toNum_(r[3]);
    moveBySku[sku] = (moveBySku[sku] || 0) + q;
  });

  // 現在庫＆アラート
  const alerts = [];
  Object.keys(products).forEach(sku => {
    const p = products[sku];
    const init = toNum_(p['初期在庫']);
    const moved = toNum_(moveBySku[sku]);
    const sold = toNum_(soldBySku[sku]);
    const stock = init + moved - sold;
    const threshold = toNum_(p['閾値']);
    if (stock <= threshold) alerts.push({ sku, name: p['商品名'], stock, threshold });
  });

  // KPI書き込み（上から4行）
  shD.getRange(1, 1, 4, 2).setValues([
    ['昨日売上', salesYesterday],
    ['昨日利益(概算)', profitYesterday],
    ['今月売上', salesMonth],
    ['在庫アラートSKU数', alerts.length]
  ]);

  // 在庫アラート通知（メール／Slack）
  if (alerts.length) {
    const lines = alerts.map(a => `・${a.sku} ${a.name} 残数:${a.stock}（閾値:${a.threshold}）`).join('\n');
    const subject = `${INV.mail.subjectPrefix} 閾値割れ ${todayStr}`;
    const body = `在庫アラート一覧\n\n${lines}\n\n（このメールはGASで自動生成されています）`;

    if (INV.mail.sendAsDraft) GmailApp.createDraft(INV.mail.to, subject, body);
    else GmailApp.sendEmail(INV.mail.to, subject, body);

    if (INV.slackWebhookUrl) {
      UrlFetchApp.fetch(INV.slackWebhookUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ text: `:warning: ${subject}\n${lines}` })
      });
    }
  }

  Logger.log(`Alerts: ${alerts.length}`);
}

/**
 * =============================
 * D. パイプライン（正規化→集計→通知）
 * =============================
 */
function runAll_InventoryPipeline() {
  normalizeOrdersToSales();
  buildInventoryDashboardAndAlert();
}

/**
 * =============================
 * E. トリガー設定（毎朝9:00）
 * =============================
 */
function setupTrigger_Inventory0900() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runAll_InventoryPipeline')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('runAll_InventoryPipeline')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  Logger.log('Inventory trigger set at 09:00 JST');
}

/**
 * =============================
 * 内部ユーティリティ
 * =============================
 */
function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function readData_(sh, includeHeaderMap = false) {
  if (!sh) return includeHeaderMap ? [] : [];
  const last = sh.getLastRow();
  if (last < 2) return includeHeaderMap ? [] : [];
  const vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  if (!includeHeaderMap) return vals;

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return vals.map(row => {
    const obj = {};
    header.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function objectBySku_(rows) {
  const obj = {};
  rows.forEach(r => { const sku = r['SKU']; if (sku) obj[sku] = r; });
  return obj;
}

function fmtDate_(d, tz) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
function toNum_(v) { const n = Number(v); return isFinite(n) ? n : 0; }
