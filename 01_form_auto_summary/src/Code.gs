/***** 設定 *****/
const SUMCFG = {
  sheetName: 'FormResponses',       // 回答が溜まるシート名
  bodyColIndex: 3,                  // 「本文」列の番号（A=1, B=2, C=3）
  fromDays: 0,                      // 直近何日分を要約するか（昨日=1）
  tz: 'Asia/Tokyo',
  mail: {
    to: 'osoranokanatahe@gmail.com',   // 送信先
    subjectPrefix: '【要約レポート】'
  },
  model: 'gpt-4o-mini',             // 速い＆安価な要約向けモデル
  // ↓ レポートの見出し
  reportTitle: 'フォーム回答 自動要約レポート'
};

/***** メイン：昨日分を要約→Doc作成→メール *****/
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
  const y = new Date(today.getTime() - SUMCFG.fromDays * 86400000);
  const yStr = Utilities.formatDate(y, tz, 'yyyy-MM-dd');

  // 「昨日」の行だけ抽出（Timestamp列はA想定）
  const yRows = values.filter(r => {
    const d = new Date(r[0]);
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd') === yStr;
  });

  if (!yRows.length) {
    Logger.log('No rows for ' + yStr);
    return;
  }

  // 本文だけを結合
  const bodies = yRows.map(r => String(r[SUMCFG.bodyColIndex - 1] || '').trim()).filter(Boolean);
  const joined = bodies.join('\n\n---\n\n');

  // 要約プロンプト
  const prompt = [
    `以下は${yStr}に集まった回答の本文です。管理者が素早く把握できるように日本語で要約してください。`,
    `出力フォーマットはMarkdownで、次の3セクションを必ず含めてください:`,
    `1) 概要（3～5行）`,
    `2) 主要トピック（箇条書き）`,
    `3) アクションアイテム（担当/期日があれば抽出）`,
    `---`,
    joined
  ].join('\n');

  const summaryMd = callOpenAI_(prompt);      // ← ChatGPTに要約してもらう
const docUrl = createDocFromMarkdown_(yStr, summaryMd);
const doc    = DocumentApp.openByUrl(docUrl);           // ← 公式APIで開く
const pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');  // ← 正式にID取得


  const subject = `${SUMCFG.mail.subjectPrefix}${SUMCFG.reportTitle} ${yStr}`;
  const bodyText = `自動生成された要約レポートです。\n${docUrl}\n\n（本メールはGASで自動送信）`;

  GmailApp.sendEmail(SUMCFG.mail.to, subject, bodyText, {
    attachments: [pdfBlob]
  });

  Logger.log('Summary sent. Doc: ' + docUrl);
}

/***** OpenAI呼び出し（Chat Completions） *****/
function callOpenAI_(userPrompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('OPENAI_API_KEY not set in Script Properties.');
  const url = 'https://api.openai.com/v1/chat/completions'; // 公式のチャット補完API

  const payload = {
    model: SUMCFG.model, // gpt-4o-mini（要約向け・安価）公式モデルページ参照
    messages: [
      { role: 'system', content: 'あなたは簡潔で的確な要約者です。重要点と行動を抜け漏れなく整理します。' },
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
    throw new Error('OpenAI API error: ' + res.getResponseCode() + ' ' + res.getContentText());
  }
  const data = JSON.parse(res.getContentText());
  const text = data.choices?.[0]?.message?.content?.trim();
  if (!text) throw new Error('No content from OpenAI.');
  return text;
}

/***** Markdown → Googleドキュメント（超簡易） *****/
function createDocFromMarkdown_(dateStr, md) {
  // Markdownを簡易パース（# 見出し / 箇条書きのみ）
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

/***** トリガー（毎朝9時に実行） *****/
function setupTrigger_Summarizer0900() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'summarizeFormResponsesAndSend')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('summarizeFormResponsesAndSend')
    .timeBased().atHour(9).everyDays(1).create();
}

/**
 * デモ用：FormResponses シートに「昨日と今日」のダミー回答を投入
 * ヘッダー：A:Timestamp, B:名前, C:本文
 */
function seedFormResponsesDemo() {
  const tz = SUMCFG.tz || 'Asia/Tokyo';
  const ss = SpreadsheetApp.getActive();
  const name = SUMCFG.sheetName || 'FormResponses';
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);

  // ヘッダー作成（既存は残す）
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,3).setValues([['Timestamp','名前','本文']]);
  }

  // 昨日・今日の日時を数件ずつ
  const now = new Date();
  const today0 = new Date(now);  today0.setHours(10, 0, 0, 0);
  const today1 = new Date(now);  today1.setHours(15, 30, 0, 0);
  const yest0  = new Date(now);  yest0.setDate(yest0.getDate()-1); yest0.setHours(11, 10, 0, 0);
  const yest1  = new Date(now);  yest1.setDate(yest1.getDate()-1); yest1.setHours(16, 45, 0, 0);

  const rows = [
    [yest0, '佐藤', 'サイト導線の改善要望。FAQ追記で問い合わせ削減の見込み。'],
    [yest1, '田中', '在庫連携の不具合。SKU A-001が二重計上。影響調査が必要。'],
    [today0,'鈴木', '広告費の入札単価を10%調整。CVRは横ばい。週次で再確認。'],
    [today1,'高橋', '顧客A社の要望ヒアリング完了。次回、要件定義に進めたい。'],
  ];

  // 末尾に追記（TimestampはDate型でOK）
  sh.getRange(sh.getLastRow()+1, 1, rows.length, 3).setValues(rows);
  // 見やすさ調整
  sh.getRange(1,1,1,3).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2,1,sh.getLastRow()-1,3).setWrap(true);
  Logger.log('Seeded demo responses.');
}

