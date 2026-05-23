// ================================================================
// mono.create — LPお問い合わせバックエンド (Google Apps Script)
// ================================================================
// 【デプロイ手順】
//   1. https://script.google.com で新規プロジェクトを作成
//   2. このファイルの全内容をコード.gsに貼り付け
//   3. ADMIN_KEY を任意の文字列に変更
//      → admin.html の ADMIN_PASSWORD と同じ値に揃える
//   4. CHATWORK_TOKEN / CHATWORK_ROOM_ID を設定
//   5. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」
//      実行ユーザー: 自分  /  アクセス: 全員（匿名含む）
//   6. デプロイURLを editor.html・admin.html の GAS_URL に貼る
// ================================================================

// ─── 設定 ───────────────────────────────────────────────────────
var ADMIN_KEY         = 'monocreate2025';
var SPREADSHEET_ID    = '13RESWCy5tuOqVzzG5aoeIFtyDk--OrLnpaPDep5yjj0';
var CHATWORK_TOKEN    = 'f79405b3d71215d721e6a9d3f86f55a6';
var CHATWORK_ROOM_ID  = '437407663';  // HPお問い合わせ
var PAYMENT_ROOM_ID   = '437439208';  // 振込確認依頼
var CHATWORK_MENTION  = 9377370;
var SHEET_NAME        = 'inquiries';
// ────────────────────────────────────────────────────────────────

// ── POST: フォーム受信 / ポートフォリオ追加・更新 ───────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // ポートフォリオ追加
    if (data.action === 'portfolio_add') {
      if (data.key !== ADMIN_KEY) return jsonResponse({ error: 'unauthorized' });
      return addPortfolio(data);
    }

    // ポートフォリオ更新
    if (data.action === 'portfolio_update') {
      if (data.key !== ADMIN_KEY) return jsonResponse({ error: 'unauthorized' });
      return updatePortfolio(data);
    }

    // 振込完了通知
    if (data.action === 'payment_notify') {
      return notifyPayment(data);
    }

    // 契約同意記録
    if (data.type === 'contract') {
      return saveContract(data);
    }

    // ヒアリングシート回答
    if (data.type === 'hearing') {
      return saveHearing(data);
    }

    // 売上記録（青色申告対応）
    if (data.type === 'sales') {
      return saveSales(data);
    }

    // お問い合わせフォーム送信
    var sheet = getOrCreateSheet();
    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    sheet.appendRow([
      now,
      data.name        || '',
      data.company     || '',
      data.email       || '',
      data.chatwork_id || '',
      data.plan        || '',
      data.message     || '',
      '未対応'
    ]);

    if (CHATWORK_ROOM_ID) {
      notifyChatwork(data, now);
    }

    return jsonResponse({ success: true });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ── GET: 管理画面向けAPI ──────────────────────────────────────
function doGet(e) {
  var action = (e.parameter.action || '').toLowerCase();
  var key    = e.parameter.key || '';

  if (key !== ADMIN_KEY) {
    return jsonResponse({ error: 'unauthorized' });
  }

  if (action === 'list') {
    return listInquiries();
  }

  if (action === 'update') {
    var row    = parseInt(e.parameter.row, 10);
    var status = e.parameter.status || '';
    return updateStatus(row, status);
  }

  if (action === 'portfolio') {
    return listPortfolio();
  }

  if (action === 'portfolio_delete') {
    var row = parseInt(e.parameter.row, 10);
    return deletePortfolio(row);
  }

  if (action === 'contracts') {
    return listContracts();
  }

  if (action === 'hearings') {
    return listHearings();
  }

  if (action === 'sales') {
    return listSales();
  }

  return jsonResponse({ error: 'unknown action' });
}

// ── 一覧取得 ─────────────────────────────────────────────────
function listInquiries() {
  var sheet = getOrCreateSheet();
  var values = sheet.getDataRange().getValues();

  // 1行目がヘッダーの場合はスキップ（初回appendRowより前に手動でヘッダーを入れた場合）
  var startRow = 2; // Spreadsheetの実際の行番号（1始まり、1行目=ヘッダー）

  var data = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    data.push({
      row:         i + 1,          // Spreadsheet行番号（ステータス更新時に使用）
      date:        row[0] || '',
      name:        row[1] || '',
      company:     row[2] || '',
      email:       row[3] || '',
      chatwork_id: row[4] || '',
      plan:        row[5] || '',
      message:     row[6] || '',
      status:      row[7] || '未対応'
    });
  }

  // 新しい順に並び替え
  data.sort(function(a, b) { return b.date > a.date ? 1 : -1; });

  return jsonResponse({ data: data });
}

// ── ステータス更新 ────────────────────────────────────────────
function updateStatus(row, status) {
  var allowed = ['未対応', '対応中', '完了'];
  if (!row || allowed.indexOf(status) === -1) {
    return jsonResponse({ error: 'invalid params' });
  }
  var sheet = getOrCreateSheet();
  sheet.getRange(row, 8).setValue(status); // 8列目=ステータス
  return jsonResponse({ success: true });
}

// ── 振込完了通知 ──────────────────────────────────────────────
function notifyPayment(data) {
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  // スプレッドシートに記録
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('payments');
  if (!sheet) {
    sheet = ss.insertSheet('payments');
    sheet.appendRow(['報告日時', '名前/振込名義', 'CW/メール', 'プラン', '金額', '振込日', '備考', 'ステータス']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  sheet.appendRow([
    now,
    data.name    || '',
    data.contact || '',
    data.plan    || '',
    data.amount  || '',
    data.date    || '',
    data.note    || '',
    '入金確認待ち'
  ]);

  // Chatwork通知
  if (PAYMENT_ROOM_ID) {
    var msg = [
      '[To:' + CHATWORK_MENTION + '] 中村航汰',
      '',
      '━━━━━━━━━━━━━━━━━━━━',
      '💳 振込完了のご報告 — mono.create',
      '━━━━━━━━━━━━━━━━━━━━',
      '報告日時：' + now,
      '振込名義：' + (data.name    || '未入力'),
      '連絡先  ：' + (data.contact || '未入力'),
      'プラン  ：' + (data.plan    || '未入力'),
      '振込金額：' + (data.amount  || '未入力'),
      '振込日  ：' + (data.date    || '未入力'),
      '備考    ：' + (data.note    || 'なし'),
      '━━━━━━━━━━━━━━━━━━━━',
      '▶ 振込確認後、Chatworkにて編集開始のご連絡をお願いします。',
    ].join('\n');

    var url = 'https://api.chatwork.com/v2/rooms/' + PAYMENT_ROOM_ID + '/messages';
    UrlFetchApp.fetch(url, {
      method:  'post',
      headers: { 'X-ChatWorkToken': CHATWORK_TOKEN },
      payload: { body: msg }
    });
  }

  return jsonResponse({ success: true });
}

// ── Chatwork通知 ──────────────────────────────────────────────
function notifyChatwork(data, now) {
  var msg = [
    '[To:' + CHATWORK_MENTION + '] 中村航汰',
    '',
    '━━━━━━━━━━━━━━━━━━━━',
    '📩 新着お問い合わせ — mono.create LP',
    '━━━━━━━━━━━━━━━━━━━━',
    '受信日時：' + now,
    'お名前  ：' + (data.name || '未入力'),
    '会社名  ：' + (data.company || '未入力'),
    'メール  ：' + (data.email || '未入力'),
    'CW ID   ：' + (data.chatwork_id || '未入力'),
    'プラン  ：' + (data.plan || '未入力'),
    '',
    '【相談内容】',
    data.message || '',
    '━━━━━━━━━━━━━━━━━━━━',
    '管理画面: YOUR_ADMIN_URL',
  ].join('\n');

  var url = 'https://api.chatwork.com/v2/rooms/' + CHATWORK_ROOM_ID + '/messages';
  UrlFetchApp.fetch(url, {
    method:  'post',
    headers: { 'X-ChatWorkToken': CHATWORK_TOKEN },
    payload: { body: msg }
  });
}

// ── ポートフォリオ操作 ────────────────────────────────────────
function listPortfolio() {
  var sheet = getOrCreatePortfolioSheet();
  var values = sheet.getDataRange().getValues();
  var data = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[1]) continue; // URL空はスキップ
    data.push({
      row:   i + 1,
      date:  row[0] || '',
      url:   row[1] || '',
      title: row[2] || '',
      genre: row[3] || '',
      type:  row[4] || 'ショート',
      order: row[5] || 99
    });
  }
  data.sort(function(a, b) { return (a.order - b.order) || (a.date < b.date ? 1 : -1); });
  return jsonResponse({ data: data });
}

function addPortfolio(data) {
  var sheet = getOrCreatePortfolioSheet();
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  var count = sheet.getLastRow() - 1;
  sheet.appendRow([now, data.url || '', data.title || '', data.genre || '', data.type || 'ショート', count + 1]);
  return jsonResponse({ success: true });
}

function updatePortfolio(data) {
  var row = parseInt(data.row, 10);
  if (!row) return jsonResponse({ error: 'invalid row' });
  var sheet = getOrCreatePortfolioSheet();
  sheet.getRange(row, 2).setValue(data.url   || '');
  sheet.getRange(row, 3).setValue(data.title || '');
  sheet.getRange(row, 4).setValue(data.genre || '');
  sheet.getRange(row, 5).setValue(data.type  || 'ショート');
  sheet.getRange(row, 6).setValue(data.order || 99);
  return jsonResponse({ success: true });
}

function deletePortfolio(row) {
  if (!row) return jsonResponse({ error: 'invalid row' });
  var sheet = getOrCreatePortfolioSheet();
  sheet.deleteRow(row);
  return jsonResponse({ success: true });
}

// ── ヘルパー ─────────────────────────────────────────────────
function getOrCreateSheet() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['受信日時', 'お名前', '会社名', 'メール', 'CW ID', 'プラン', '相談内容', 'ステータス']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  }
  return sheet;
}

function getOrCreatePortfolioSheet() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('portfolio');
  if (!sheet) {
    sheet = ss.insertSheet('portfolio');
    sheet.appendRow(['追加日時', 'URL', 'タイトル', 'ジャンル', 'タイプ', '表示順']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sheet;
}

// ================================================================
// 契約同意記録
// ================================================================
function saveContract(data) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('contracts');
  if (!sheet) {
    sheet = ss.insertSheet('contracts');
    sheet.appendRow(['同意日時', 'お名前', 'メールアドレス', 'プラン', '契約バージョン', 'IPメモ']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  sheet.appendRow([
    now,
    data.name        || '',
    data.email       || '',
    data.plan        || '',
    data.contractVer || '',
    ''
  ]);

  // Chatwork通知
  if (CHATWORK_ROOM_ID) {
    var msg = [
      '[To:' + CHATWORK_MENTION + '] 中村航汰',
      '',
      '━━━━━━━━━━━━━━━━━━━━',
      '📝 契約同意完了 — mono.create LP',
      '━━━━━━━━━━━━━━━━━━━━',
      '同意日時：' + now,
      'お名前  ：' + (data.name  || ''),
      'メール  ：' + (data.email || ''),
      'プラン  ：' + (data.plan  || ''),
      '━━━━━━━━━━━━━━━━━━━━',
      '▶ 次のヒアリングシート回答をお待ちください。',
    ].join('\n');
    var url = 'https://api.chatwork.com/v2/rooms/' + CHATWORK_ROOM_ID + '/messages';
    UrlFetchApp.fetch(url, { method:'post', headers:{'X-ChatWorkToken':CHATWORK_TOKEN}, payload:{body:msg} });
  }

  return jsonResponse({ success: true });
}

function listContracts() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('contracts');
  if (!sheet) return jsonResponse({ data: [] });
  var values = sheet.getDataRange().getValues();
  var data = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    data.push({ row: i+1, date: row[0]||'', name: row[1]||'', email: row[2]||'', plan: row[3]||'', ver: row[4]||'' });
  }
  data.sort(function(a,b){ return b.date > a.date ? 1 : -1; });
  return jsonResponse({ data: data });
}

// ================================================================
// ヒアリングシート
// ================================================================
function saveHearing(data) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('hearings');
  if (!sheet) {
    sheet = ss.insertSheet('hearings');
    sheet.appendRow(['受信日時','お名前','メール','プラン','回答JSON','ステータス']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  sheet.appendRow([
    now,
    data.name  || '',
    data.email || '',
    data.plan  || '',
    JSON.stringify(data.answers || {}),
    '未対応'
  ]);

  // Chatwork通知
  if (CHATWORK_ROOM_ID) {
    var lines = ['[To:' + CHATWORK_MENTION + '] 中村航汰', '',
      '━━━━━━━━━━━━━━━━━━━━',
      '📋 ヒアリングシート回答 — mono.create LP',
      '━━━━━━━━━━━━━━━━━━━━',
      '受信日時：' + now,
      'お名前  ：' + (data.name  || ''),
      'メール  ：' + (data.email || ''),
      'プラン  ：' + (data.plan  || ''),
      '━━━━━━━━━━━━━━━━━━━━'
    ];
    var ans = data.answers || {};
    Object.keys(ans).forEach(function(k){ lines.push(k + '：' + ans[k]); });
    lines.push('━━━━━━━━━━━━━━━━━━━━');
    var url = 'https://api.chatwork.com/v2/rooms/' + CHATWORK_ROOM_ID + '/messages';
    UrlFetchApp.fetch(url, { method:'post', headers:{'X-ChatWorkToken':CHATWORK_TOKEN}, payload:{body:lines.join('\n')} });
  }

  return jsonResponse({ success: true });
}

function listHearings() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('hearings');
  if (!sheet) return jsonResponse({ data: [] });
  var values = sheet.getDataRange().getValues();
  var data = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var answers = {};
    try { answers = JSON.parse(row[4]); } catch(e) {}
    data.push({ row:i+1, date:row[0]||'', name:row[1]||'', email:row[2]||'', plan:row[3]||'', answers:answers, status:row[5]||'未対応' });
  }
  data.sort(function(a,b){ return b.date > a.date ? 1 : -1; });
  return jsonResponse({ data: data });
}

// ================================================================
// 売上帳（青色申告対応）
// 帳簿種別: 売上帳・経費帳
// ================================================================
function saveSales(data) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('uriage');
  if (!sheet) {
    sheet = ss.insertSheet('uriage');
    // 青色申告・売上帳フォーマット
    sheet.appendRow(['取引日','取引先名','取引先住所','摘要（プラン）','売上金額（税抜）','消費税額','売上金額（税込）','入金日','入金確認','備考']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    sheet.setColumnWidth(1, 100);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 120);
    sheet.setColumnWidth(8, 100);
    sheet.setColumnWidth(9, 80);
    sheet.setColumnWidth(10, 200);
  }

  var taxInc   = parseFloat(data.amount) || 0;
  var taxRate  = 0.10;
  var taxExc   = Math.round(taxInc / (1 + taxRate));
  var taxAmt   = taxInc - taxExc;

  sheet.appendRow([
    data.date        || '',   // 取引日
    data.client      || '',   // 取引先名
    data.address     || '',   // 取引先住所
    data.plan        || '',   // 摘要
    taxExc,                   // 売上金額（税抜）
    taxAmt,                   // 消費税額
    taxInc,                   // 売上金額（税込）
    data.payDate     || '',   // 入金日
    data.confirmed   || '未確認', // 入金確認
    data.note        || ''    // 備考
  ]);
  return jsonResponse({ success: true });
}

function listSales() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('uriage');
  if (!sheet) return jsonResponse({ data: [] });
  var values = sheet.getDataRange().getValues();
  var data = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    data.push({
      row:       i+1,
      date:      row[0]||'',
      client:    row[1]||'',
      plan:      row[3]||'',
      taxExc:    row[4]||0,
      taxAmt:    row[5]||0,
      taxInc:    row[6]||0,
      payDate:   row[7]||'',
      confirmed: row[8]||'未確認',
      note:      row[9]||''
    });
  }
  data.sort(function(a,b){ return b.date > a.date ? 1 : -1; });
  return jsonResponse({ data: data });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
