/**
 * 📊 ダミーデータ生成
 * DailyLog(A:D) に 30日分 × 6名 × 1〜3件/日 のデータを自動投入
 * A: 日付(yyyy-mm-dd), B: 担当, C: タスク, D: 進捗/メモ
 */
function seedDailyLogDemoData() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('DailyLog') || ss.insertSheet('DailyLog');

  // ヘッダー作成 & 既存データ削除
  sh.clear();
  sh.getRange(1, 1, 1, 4).setValues([['日付', '担当', 'タスク', '進捗/メモ']]);

  const tz = 'Asia/Tokyo';
  const today = new Date();
  const days = 30; // 生成する日数（14〜30で見栄え良好）

  const members = ['小林', '斉藤', '田中', '鈴木', '高橋', '佐藤'];
  const tasks = [
    '広告レポート集計', '在庫表クリーニング', '顧客対応ログ更新', '請求書チェック',
    '出品データ整形', '商品画像差し替え', 'QA回答まとめ', 'キャンペーン反映',
    '配送遅延アラート確認', '売上ダッシュボード更新', '返品処理', 'フォーム不備修正'
  ];
  const notes = ['完了', '80%', '50%', '要確認', '明日対応', '保留（依頼待ち）', '追加データ待ち'];

  const rows = [];
  for (let d = 0; d < days; d++) {
    const date = new Date(today.getTime() - d * 24 * 3600 * 1000);
    const dateStr = Utilities.formatDate(date, tz, 'yyyy-MM-dd');

    members.forEach(m => {
      const itemCount = 1 + Math.floor(Math.random() * 3);
      for (let i = 0; i < itemCount; i++) {
        const task = tasks[Math.floor(Math.random() * tasks.length)];
        const note = notes[Math.floor(Math.random() * notes.length)];
        rows.push([dateStr, m, task, note]);
      }
    });
  }

  rows.sort((a, b) => a[0].localeCompare(b[0]));
  if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
  Logger.log(`Inserted demo rows: ${rows.length}`);
}

/**
 * 基本設定（CFGオブジェクトを変更するだけで使い回し可能）
 */
const CFG = {
  sheetName: 'DailyLog',
  headerRow: 1,
  dateColIndex: 1,
  mail: {
    to: 'yourname@example.com', // デモ用宛先
    cc: '',
    subjectPrefix: '【日報自動送信】',
    sendAsDraft: true // true = 下書きモード, false = 本番送信
  },
  businessName: '日報デモ',
  tz: 'Asia/Tokyo'
};

/**
 * 📩 日報HTML生成
 * 当日のログのみ抽出し、HTMLメール形式に整形
 */
function buildDailyReportHtml_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.sheetName);
  if (!sh) throw new Error(`Sheet not found: ${CFG.sheetName}`);

  const lastRow = sh.getLastRow();
  const values = lastRow > CFG.headerRow
    ? sh.getRange(CFG.headerRow + 1, 1, lastRow - CFG.headerRow, 4).getValues()
    : [];

  const today = new Date();
  const tz = CFG.tz;
  const start = new Date(today); start.setHours(0, 0, 0, 0);
  const end = new Date(start); end.setDate(end.getDate() + 1);

  const rows = values.filter(r => {
    const d = r[0] instanceof Date ? r[0] : new Date(String(r[0]));
    return !isNaN(d) && d >= start && d < end;
  });

  const todayStr = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
  const esc = s => String(s ?? '').replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));

  const tableRows = rows.length
    ? rows.sort((a, b) => String(a[1]).localeCompare(String(b[1])))
      .map(r => `<tr><td>${esc(r[0])}</td><td>${esc(r[1])}</td><td>${esc(r[2])}</td><td>${esc(r[3])}</td></tr>`)
      .join('')
    : `<tr><td colspan="4" style="text-align:center;color:#888;padding:14px">本日の入力はありません</td></tr>`;

  const styles = `
    <style>
      body{font-family:system-ui, -apple-system,"Segoe UI",Roboto,sans-serif;color:#202124;}
      h1{margin:0 0 6px;font-size:18px}
      table{border-collapse:collapse;font-size:13px}
      th,td{border:1px solid #e0e3e7;padding:8px 10px;vertical-align:top}
      thead th{background:#f2f5f9;text-align:center}
      .note{color:#666;font-size:12px;margin-top:10px}
    </style>`;

  const header = `<h1>${esc(CFG.businessName)} ${todayStr}</h1>
<p>お疲れさまです。本日の進捗を自動集計しました。</p>`;

  const table = `<table>
  <thead><tr><th>日付</th><th>担当</th><th>タスク</th><th>進捗/メモ</th></tr></thead>
  <tbody>${tableRows}</tbody>
</table>`;

  const footer = `<p class="note">※このメールはGASで自動生成されています</p>`;

  return styles + header + table + footer;
}

/**
 * HTML → テキスト変換（プレーンテキストメール用）
 */
function stripHtml_(html) {
  return html
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

/**
 * 📬 テスト送信（下書きモード強制）
 */
function previewDailyReport() {
  CFG.mail.sendAsDraft = true;
  sendDailyReport();
}

/**
 * ✉️ 日報メール送信
 */
function sendDailyReport() {
  const html = buildDailyReportHtml_();
  const text = stripHtml_(html);
  const todayFmt = Utilities.formatDate(new Date(), CFG.tz, 'yyyy-MM-dd (E)');
  const subject = `${CFG.mail.subjectPrefix}${CFG.businessName} ${todayFmt} チーム進捗レポート`;

  if (CFG.mail.sendAsDraft) {
    GmailApp.createDraft(CFG.mail.to, subject, text, { htmlBody: html, cc: CFG.mail.cc });
    Logger.log('Draft created.');
  } else {
    GmailApp.sendEmail(CFG.mail.to, subject, text, { htmlBody: html, cc: CFG.mail.cc });
    Logger.log('Mail sent.');
  }
}

/**
 * ⏰ トリガー設定：毎朝9時に送信
 */
function setupTriggerEveryMorning0900() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailyReport')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('sendDailyReport')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  Logger.log('Trigger set: 09:00 JST');
}

/**
 * 🧪 デモ用関数：今日分のダミーデータを3〜6件追加
 */
function seedDailyLog_forToday() {
  const tz = CFG.tz || 'Asia/Tokyo';
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.sheetName) || ss.insertSheet(CFG.sheetName);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 4).setValues([['日付', '担当', 'タスク', '進捗/メモ']]);
  }

  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const members = ['小林', '斉藤', '田中', '鈴木', '高橋', '佐藤'];
  const tasks = ['広告レポート集計', '在庫表クリーニング', '顧客対応ログ更新', '請求書チェック', '出品データ整形', '商品画像差し替え'];
  const notes = ['完了', '80%', '50%', '要確認', '明日対応', '保留（依頼待ち）', '追加データ待ち'];

  // 今日分の既存データを削除
  const last = sh.getLastRow();
  if (last > 1) {
    const range = sh.getRange(2, 1, last - 1, 4).getValues();
    const remain = range.filter(r => String(r[0]) !== todayStr);
    if (remain.length !== range.length) {
      sh.getRange(2, 1, last - 1, 4).clearContent();
      if (remain.length) sh.getRange(2, 1, remain.length, 4).setValues(remain);
    }
  }

  const rows = [];
  const n = 3 + Math.floor(Math.random() * 4);
  for (let i = 0; i < n; i++) {
    rows.push([
      todayStr,
      members[Math.floor(Math.random() * members.length)],
      tasks[Math.floor(Math.random() * tasks.length)],
      notes[Math.floor(Math.random() * notes.length)],
    ]);
  }

  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  Logger.log(`Seeded today's rows: ${rows.length}`);
}
