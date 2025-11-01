/**
 * 📋 設定
 * 使用するスプレッドシート、メール、要約モデルなどを定義
 */
const SUMCFG = {
  sheetName: 'FormResponses',       // 回答が溜まるシート名
  bodyColIndex: 3,                  // 「本文」列の番号（A=1, B=2, C=3）
  fromDays: 0,                      // 要約対象日数（昨日=1）
  tz: 'Asia/Tokyo',
  mail: {
    to: 'yourname@example.com',     // デモ用宛先（公開時は個人アドレスを避ける）
    subjectPrefix: '【要約レポート】'
  },
  model: 'gpt-4o-mini',             // 高速・低コストな要約モデル
  reportTitle: 'フォーム回答 自動要約レポート'
};

/**
 * 🧠 メイン処理：フォーム回答を要約 → PDF化 → Gmail送信
 */
function summarizeFormResponsesAndSend() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SUMCFG.sheetName);
  if (!sh) throw new Error('Sheet not found: ' + SUMCFG.sheetName);

  const last = sh.getLastRow();
  if (last < 2) {
    Logger.log('No responses.');
    return;
  }

  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const today = new Date();
  const tz = SUMCFG.tz;
  const targetDate = new Date(today.getTime() - SUMCFG.fromDays * 86400000);
  const targetStr = Utilities.formatDate(targetDate, tz, 'yyyy-MM-dd');

  // 指定日の回答行を抽出（Timestamp列はA列想定）
  const yRows = values.filter(r => {
    const d = new Date(r[0]);
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd') === targetStr;
  });

  if (!yRows.length) {
    Logger.log(`No rows for ${targetStr}`);
    return;
  }

  // 本文のみ抽出・結合
  const bodies = yRows.map(r => String(r[SUMCFG.bodyColIndex - 1] || '').trim()).filter(Boolean);
  const joined = bodies.join('\n\n---\n\n');

  // ChatGPTに渡すプロンプトを組み立て
  const prompt = [
    `以下は${targetStr}に集まったフォーム回答の本文です。管理者が全体を把握できるよう、日本語で要約してください。`,
    `出力形式はMarkdownで、以下の3セクションを含めてください：`,
    `1) 概要（3〜5行）`,
    `2) 主要トピック（箇条書き）`,
    `3) アクションアイテム（担当や期日があれば抽出）`,
    `---`,
    joined
  ].join('\n');

  const summaryMd = callOpenAI_(prompt);
  const docUrl = createDocFromMarkdown_(targetStr, summaryMd);
  const doc = DocumentApp.openByUrl(docUrl);
  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');

  const subject = `${SUMCFG.mail.subjectPrefix}${SUMCFG.reportTitle} ${targetStr}`;
  const bodyText = `自動生成された要約レポートです。\n${docUrl}\n\n（このメールはGASで自動送信されています）`;

  GmailApp.sendEmail(SUMCFG.mail.to, subject, bodyText, {
    attachments: [pdfBlob]
  });

  Logger.log(`Summary sent. Doc: ${docUrl}`);
}

/**
 * 🤖 OpenAI API呼び出し（Chat Completions）
 */
function callOpenAI_(userPrompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY not set in Script Properties.');

  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    model: SUMCFG.model,
    messages: [
      { role: 'system', content: 'あなたは簡潔で的確な要約者です。重要点と行動を整理します。' },
      { role: 'user', content: userPrompt }
    ],
    temperature: 0.2
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: { Authorization: `Bearer ${apiKey}` },
    payload: JSON.stringify(payload)
  });

  if (res.getResponseCode() !== 200) {
    throw new Error(`OpenAI API error: ${res.getResponseCode()} ${res.getContentText()}`);
  }

  const data = JSON.parse(res.getContentText());
  const text = data.choices?.[0]?.message?.content?.trim();
  if (!text) throw new Error('No content from OpenAI.');
  return text;
}

/**
 * 📝 Markdown → Googleドキュメント変換（簡易パーサー）
 */
function createDocFromMarkdown_(dateStr, md) {
  const doc = DocumentApp.create(`${SUMCFG.reportTitle} ${dateStr}`);
  const body = doc.getBody();
  const lines = md.split(/\r?\n/);

  lines.forEach(line => {
    if (/^#\s+/.test(line)) {
      body.appendParagraph(line.replace(/^#\s+/, '')).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    } else if (/^##\s+/.test(line)) {
      body.appendParagraph(line.replace(/^##\s+/, '')).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    } else if (/^[-*]\s+/.test(line)) {
      body.appendListItem(line.replace(/^[-*]\s+/, '')).setGlyphType(DocumentApp.GlyphType.BULLET);
    } else if (line.trim() === '---') {
      body.appendHorizontalRule();
    } else {
      body.appendParagraph(line);
    }
  });

  doc.saveAndClose();
  return doc.getUrl();
}

/**
 * ⏰ トリガー設定（毎朝9時に自動実行）
 */
function setupTrigger_Summarizer0900() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'summarizeFormResponsesAndSend')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('summarizeFormResponsesAndSend')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  Logger.log('Trigger set: summarizeFormResponsesAndSend at 09:00 JST');
}

/**
 * 🧪 デモデータ生成
 * FormResponses シートに「昨日・今日」のサンプル回答を投入
 */
function seedFormResponsesDemo() {
  const tz = SUMCFG.tz || 'Asia/Tokyo';
  const ss = SpreadsheetApp.getActive();
  const name = SUMCFG.sheetName || 'FormResponses';
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);

  // ヘッダー作成（既存は残す）
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 3).setValues([['Timestamp', '名前', '本文']]);
  }

  // 昨日・今日のダミー回答
  const now = new Date();
  const today0 = new Date(now); today0.setHours(10, 0, 0, 0);
  const today1 = new Date(now); today1.setHours(15, 30, 0, 0);
  const yest0 = new Date(now); yest0.setDate(yest0.getDate() - 1); yest0.setHours(11, 10, 0, 0);
  const yest1 = new Date(now); yest1.setDate(yest1.getDate() - 1); yest1.setHours(16, 45, 0, 0);

  const rows = [
    [yest0, '佐藤', 'サイト導線の改善要望。FAQ追記で問い合わせ削減の見込み。'],
    [yest1, '田中', '在庫連携の不具合。SKU A-001が二重計上。影響調査が必要。'],
    [today0, '鈴木', '広告費の入札単価を10%調整。CVRは横ばい。週次で再確認。'],
    [today1, '高橋', '顧客A社の要望ヒアリング完了。次回、要件定義に進めたい。']
  ];

  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).setWrap(true);
  Logger.log('Demo responses seeded.');
}
