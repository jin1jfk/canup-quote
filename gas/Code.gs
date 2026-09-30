// 株式会社Canup 見積書アプリ
// スプレッドシート「株式会社Canup_物件管理リスト」に組み込むコンテナバインドスクリプト。
// シート: 設定 / 見積書テンプレ / 見積台帳 / 選択肢(F列=見積品目) を使う。

const SHEET_SETTINGS = '設定';
const SHEET_TEMPLATE = '見積書テンプレ';
const SHEET_LEDGER = '見積台帳';
const SHEET_CHOICES = '選択肢';
const SHEET_ANKEN = '案件管理';
const ITEM_FIRST_ROW = 15;
const ITEM_ROWS = 15;
const TAX_RATE = 0.1;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('見積書')
    .addItem('見積書を作成', 'openQuoteDialog')
    .addSeparator()
    .addItem('初期設定(最初に1回)', 'setup')
    .addToUi();
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Canup 見積書')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- Web版(GitHub Pages)から呼ぶAPI ----------
// POST本文(JSON文字列): { key, action: 'init' | 'issue' | 'attachPdf', ... }

function doPost(e) {
  let out;
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const st = readSettings_(ss);
    const key = String(st['アプリの合言葉'] || '');
    if (!key || body.key !== key) throw new Error('合言葉が違います。スプレッドシートの「設定」シートにある「アプリの合言葉」を入れ直してください。');
    if (body.action === 'init') out = apiInit_(ss, st);
    else if (body.action === 'issue') out = apiIssue_(ss, st, body.data || {});
    else if (body.action === 'attachPdf') out = apiAttachPdf_(ss, st, body);
    else throw new Error('不明な操作です: ' + body.action);
    out.ok = true;
  } catch (err) {
    out = { ok: false, error: err.message || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function apiInit_(ss, st) {
  const init = getInitData();
  return {
    issuer: issuerInfo_(st),
    validDays: init.validDays,
    recipients: init.recipients,
    items: init.items,
    ready: init.ready && !!ss.getSheetByName(SHEET_LEDGER),
  };
}

function apiIssue_(ss, st, data) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ledger = ss.getSheetByName(SHEET_LEDGER);
    if (!ledger) throw new Error('見積台帳シートがありません。スプレッドシートの「見積書」メニューから「初期設定」を実行してください。');
    const q = calcQuote_(data, st);
    const no = nextQuoteNo_(ledger, st['見積番号の頭'] || 'Q', q.issue);
    ledger.appendRow([no, q.issue, data.to, data.subject || '', q.taxMode, q.subtotal, q.tax, q.total, q.valid, '(PDF保存待ち)']);
    return {
      no: no,
      issueDate: Utilities.formatDate(q.issue, 'Asia/Tokyo', 'yyyy-MM-dd'),
      validDate: Utilities.formatDate(q.valid, 'Asia/Tokyo', 'yyyy-MM-dd'),
      subtotal: q.subtotal, tax: q.tax, total: q.total,
      issuer: issuerInfo_(st),
    };
  } finally {
    lock.releaseLock();
  }
}

function apiAttachPdf_(ss, st, body) {
  if (!body.no || !body.pdf) throw new Error('見積番号かPDFがありません。');
  const ledger = ss.getSheetByName(SHEET_LEDGER);
  const nos = ledger.getRange(2, 1, Math.max(ledger.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]));
  const idx = nos.indexOf(String(body.no));
  if (idx < 0) throw new Error('見積台帳に ' + body.no + ' が見つかりません。');
  const to = ledger.getRange(idx + 2, 3).getValue();
  const blob = Utilities.newBlob(Utilities.base64Decode(body.pdf), 'application/pdf',
    body.no + '_' + safeName_(to) + '_見積書.pdf');
  const file = getFolder_(ss, st).createFile(blob);
  ledger.getRange(idx + 2, 10).setValue(file.getUrl());
  return { url: file.getUrl() };
}

function issuerInfo_(st) {
  return {
    company: st['社名'] || '', zip: st['郵便番号'] || '', address: st['住所'] || '',
    tel: st['電話'] || '', mail: st['メール'] || '', regNo: st['登録番号'] || '', bank: st['振込先'] || '',
  };
}

function calcQuote_(data, st) {
  const items = (data.items || []).filter(it => it.name && Number(it.qty) && it.price !== '' && it.price !== null);
  if (!data.to) throw new Error('宛名が空です。');
  if (!items.length) throw new Error('品目・数量・単価がそろった行が1つもありません。');
  if (items.length > ITEM_ROWS) throw new Error('品目は' + ITEM_ROWS + '行までです。');
  const issue = data.date ? new Date(data.date + 'T00:00:00+09:00') : new Date();
  const validDays = Number(st['有効期限(日)']) || 30;
  const valid = new Date(issue.getTime() + validDays * 86400000);
  const sum = items.reduce((a, it) => a + Math.round(Number(it.qty) * Number(it.price)), 0);
  const taxMode = data.taxMode === '内税' ? '内税' : '外税';
  let subtotal, tax, total;
  if (taxMode === '内税') { total = sum; tax = Math.floor(sum * TAX_RATE / (1 + TAX_RATE)); subtotal = total - tax; }
  else { subtotal = sum; tax = Math.floor(sum * TAX_RATE); total = subtotal + tax; }
  return { items, issue, valid, subtotal, tax, total, taxMode };
}

function openQuoteDialog() {
  const html = HtmlService.createTemplateFromFile('Index').evaluate()
    .setWidth(900).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, '見積書を作成');
}

// ---------- 初期設定 ----------

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSettings_(ss);
  setupLedger_(ss);
  setupTemplate_(ss);
  setupItemChoices_(ss);
  SpreadsheetApp.getUi().alert('初期設定が終わりました。「設定」シートの空欄(電話・メール・登録番号・振込先)を埋めてください。');
}

function setupSettings_(ss) {
  if (ss.getSheetByName(SHEET_SETTINGS)) return;
  const sh = ss.insertSheet(SHEET_SETTINGS);
  const rows = [
    ['項目', '内容', '説明'],
    ['社名', '株式会社Canup', ''],
    ['代表者', '山下真我', '見積書には載せない(控え)'],
    ['郵便番号', '〒566-0062', ''],
    ['住所', '大阪府摂津市鳥飼上1丁目16-5', ''],
    ['電話', '', '空欄なら見積書に出ない'],
    ['メール', '', '空欄なら見積書に出ない'],
    ['登録番号', '', 'インボイス登録番号(T+13桁)。空欄なら出ない'],
    ['振込先', '', '備考欄の末尾に載る。空欄なら出ない'],
    ['有効期限(日)', 30, '発行日からの日数'],
    ['見積番号の頭', 'Q', '見積番号 = 頭-YYMMDD-連番'],
    ['保存フォルダID', '', '空欄なら初回作成時にマイドライブに「Canup見積書」フォルダを作って自動で入る'],
    ['アプリの合言葉', Utilities.getUuid().replace(/-/g, '').slice(0, 12), 'Web版アプリに1回だけ入力する。他人に教えない。変えると全端末で入れ直し'],
  ];
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange('A1:C1').setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
  sh.getRange('A2:A' + rows.length).setFontWeight('bold');
  sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 360); sh.setColumnWidth(3, 420);
  sh.getRange('B2:B' + rows.length).setBackground('#FFF9E6');
  sh.setFrozenRows(1);
}

function setupLedger_(ss) {
  if (ss.getSheetByName(SHEET_LEDGER)) return;
  const sh = ss.insertSheet(SHEET_LEDGER);
  const head = ['見積番号', '発行日', '宛名', '件名', '税区分', '小計', '消費税', '合計', '有効期限', 'PDF'];
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  [130, 100, 220, 260, 70, 100, 90, 110, 100, 320].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('B:B').setNumberFormat('yyyy/mm/dd');
  sh.getRange('I:I').setNumberFormat('yyyy/mm/dd');
  sh.getRange('F:H').setNumberFormat('#,##0');
}

function setupItemChoices_(ss) {
  const sh = ss.getSheetByName(SHEET_CHOICES);
  if (!sh || sh.getRange('F1').getValue() !== '') return;
  const items = ['見積品目', '残地物回収', 'エアコンクリーニング', 'エアコン(天カセ)', '防犯カメラ',
    '定期清掃', 'ハウスクリーニング', '内装工事', 'ステージング', '出張費'];
  sh.getRange(1, 6, items.length, 1).setValues(items.map(v => [v]));
  sh.getRange('F1').setFontWeight('bold');
  sh.setColumnWidth(6, 180);
}

function setupTemplate_(ss) {
  if (ss.getSheetByName(SHEET_TEMPLATE)) return;
  const sh = ss.insertSheet(SHEET_TEMPLATE);
  const lastItemRow = ITEM_FIRST_ROW + ITEM_ROWS - 1; // 29
  sh.getRange('A1:F60').setFontSize(10).setVerticalAlignment('middle');
  [40, 250, 60, 50, 90, 110].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setHiddenGridlines(true);

  sh.getRange('A1:F1').merge().setValue('御 見 積 書')
    .setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(1, 44);

  sh.getRange('A3:C3').merge().setFontSize(14).setFontWeight('bold')
    .setBorder(null, null, true, null, null, null);
  sh.getRange('E3').setValue('見積番号'); sh.getRange('E4').setValue('発行日'); sh.getRange('E5').setValue('有効期限');
  sh.getRange('F4:F5').setNumberFormat('yyyy年m月d日');
  sh.getRange('E3:E5').setFontColor('#555555');
  sh.getRange('F3:F5').setHorizontalAlignment('right');

  sh.getRange('A6:C6').merge().setValue('下記のとおりお見積り申し上げます。');
  sh.getRange('A7').setValue('件名').setFontWeight('bold');
  sh.getRange('B7:C7').merge().setBorder(null, null, true, null, null, null);
  sh.getRange('A9').setValue('御見積金額').setFontWeight('bold');
  sh.getRange('B9:C9').merge().setFontSize(16).setFontWeight('bold').setNumberFormat('"¥"#,##0')
    .setHorizontalAlignment('left').setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange('A10:C10').merge().setFontColor('#555555').setFontSize(9);

  // 発行元(E7:F12)
  for (let r = 7; r <= 12; r++) sh.getRange(r, 5, 1, 2).merge().setHorizontalAlignment('left');
  sh.getRange('E7').setFontWeight('bold').setFontSize(11);
  sh.getRange('E8:E12').setFontSize(9);

  // 明細
  const head = sh.getRange(ITEM_FIRST_ROW - 1, 1, 1, 6);
  head.setValues([['No', '品目', '数量', '単位', '単価', '金額']])
    .setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  const body = sh.getRange(ITEM_FIRST_ROW, 1, ITEM_ROWS, 6);
  body.setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(ITEM_FIRST_ROW, 1, ITEM_ROWS, 1).setHorizontalAlignment('center');
  sh.getRange(ITEM_FIRST_ROW, 3, ITEM_ROWS, 2).setHorizontalAlignment('center');
  sh.getRange(ITEM_FIRST_ROW, 5, ITEM_ROWS, 2).setNumberFormat('#,##0').setHorizontalAlignment('right');

  // 合計欄
  const t = lastItemRow + 2; // 31
  sh.getRange(t, 5, 3, 1).setValues([['小計'], ['消費税(10%)'], ['合計']]).setFontWeight('bold');
  sh.getRange(t, 6, 3, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
  sh.getRange(t, 5, 3, 2).setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(t + 2, 5, 1, 2).setBackground('#EEF2F8');

  // 備考
  sh.getRange(t + 4, 1).setValue('備考').setFontWeight('bold');
  sh.getRange(t + 5, 1, 5, 6).merge().setVerticalAlignment('top').setWrap(true)
    .setBorder(true, true, true, true, null, null, '#999999', SpreadsheetApp.BorderStyle.SOLID);
}

// ---------- アプリから呼ぶ ----------

function getInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const st = readSettings_(ss);
  const ledger = ss.getSheetByName(SHEET_LEDGER);
  // 宛名の候補 = 見積台帳の宛名(新しい順) + 案件管理シートの「商談先」(下の行=新しい順)
  let names = [];
  if (ledger && ledger.getLastRow() > 1) {
    names = names.concat(ledger.getRange(2, 3, ledger.getLastRow() - 1, 1).getValues().map(r => r[0]).reverse());
  }
  const anken = ss.getSheetByName(SHEET_ANKEN);
  if (anken && anken.getLastRow() > 1) {
    const head = anken.getRange(1, 1, 1, anken.getLastColumn()).getValues()[0];
    const col = head.indexOf('商談先') + 1;
    if (col > 0) names = names.concat(anken.getRange(2, col, anken.getLastRow() - 1, 1).getValues().map(r => r[0]).reverse());
  }
  const recipients = [...new Set(names.map(v => String(v).trim()).filter(String))];
  const choices = ss.getSheetByName(SHEET_CHOICES);
  let items = [];
  if (choices && choices.getLastRow() > 1) {
    items = choices.getRange(2, 6, choices.getLastRow() - 1, 1).getValues().map(r => r[0]).filter(String);
  }
  return { company: st['社名'] || '', validDays: Number(st['有効期限(日)']) || 30, recipients, items, ready: !!ss.getSheetByName(SHEET_TEMPLATE) };
}

function createQuote(data) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const st = readSettings_(ss);
    const tpl = ss.getSheetByName(SHEET_TEMPLATE);
    const ledger = ss.getSheetByName(SHEET_LEDGER);
    if (!tpl || !ledger) throw new Error('初期設定がまだです。スプレッドシートの「見積書」メニューから「初期設定」を実行してください。');

    const items = (data.items || []).filter(it => it.name && Number(it.qty) && it.price !== '' && it.price !== null);
    if (!data.to) throw new Error('宛名が空です。');
    if (!items.length) throw new Error('品目が1つもありません。');
    if (items.length > ITEM_ROWS) throw new Error('品目は' + ITEM_ROWS + '行までです。');

    const issue = data.date ? new Date(data.date + 'T00:00:00') : new Date();
    const validDays = Number(st['有効期限(日)']) || 30;
    const valid = new Date(issue.getTime() + validDays * 86400000);
    const no = nextQuoteNo_(ledger, st['見積番号の頭'] || 'Q', issue);

    const sum = items.reduce((a, it) => a + Math.round(Number(it.qty) * Number(it.price)), 0);
    let subtotal, tax, total;
    if (data.taxMode === '内税') {
      total = sum; tax = Math.floor(sum * TAX_RATE / (1 + TAX_RATE)); subtotal = total - tax;
    } else {
      subtotal = sum; tax = Math.floor(sum * TAX_RATE); total = subtotal + tax;
    }

    // テンプレに流し込む
    const lastItemRow = ITEM_FIRST_ROW + ITEM_ROWS - 1;
    const t = lastItemRow + 2;
    tpl.getRange('A3').setValue(data.to + ' ' + (data.honorific || '御中'));
    tpl.getRange('F3').setValue(no);
    tpl.getRange('F4').setValue(issue);
    tpl.getRange('F5').setValue(valid);
    tpl.getRange('B7').setValue(data.subject || '');
    tpl.getRange('B9').setValue(total);
    tpl.getRange('A10').setValue(data.taxMode === '内税' ? '(税込・うち消費税 ¥' + tax.toLocaleString('ja-JP') + ')' : '(税込)');

    const issuer = [st['社名'] || '', [st['郵便番号'], st['住所']].filter(String).join(' '),
      st['電話'] ? 'TEL ' + st['電話'] : '', st['メール'] || '',
      st['登録番号'] ? '登録番号 ' + st['登録番号'] : ''].filter(String);
    const issuerCells = tpl.getRange('E7:E12');
    issuerCells.clearContent();
    issuer.forEach((v, i) => tpl.getRange(7 + i, 5).setValue(v));

    const body = tpl.getRange(ITEM_FIRST_ROW, 1, ITEM_ROWS, 6);
    body.clearContent();
    const rows = items.map((it, i) => [i + 1, it.name, Number(it.qty), it.unit || '式', Number(it.price), Math.round(Number(it.qty) * Number(it.price))]);
    tpl.getRange(ITEM_FIRST_ROW, 1, rows.length, 6).setValues(rows);

    tpl.getRange(t, 5).setValue(data.taxMode === '内税' ? '小計(税抜)' : '小計');
    tpl.getRange(t + 1, 5).setValue(data.taxMode === '内税' ? 'うち消費税(10%)' : '消費税(10%)');
    tpl.getRange(t, 6, 3, 1).setValues([[subtotal], [tax], [total]]);

    const note = [data.note || '', st['振込先'] ? 'お振込先: ' + st['振込先'] : ''].filter(String).join('\n');
    tpl.getRange(t + 5, 1).setValue(note);
    SpreadsheetApp.flush();

    // PDF化して保存
    const folder = getFolder_(ss, st);
    const fileName = no + '_' + safeName_(data.to) + '_見積書.pdf';
    const pdf = exportSheetPdf_(ss, tpl).setName(fileName);
    const file = folder.createFile(pdf);

    ledger.appendRow([no, issue, data.to, data.subject || '', data.taxMode === '内税' ? '内税' : '外税', subtotal, tax, total, valid, file.getUrl()]);
    return { no, url: file.getUrl(), total };
  } finally {
    lock.releaseLock();
  }
}

// ---------- 内部 ----------

function readSettings_(ss) {
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) out[r[0]] = r[1]; });
  return out;
}

function nextQuoteNo_(ledger, head, date) {
  const ymd = Utilities.formatDate(date, 'Asia/Tokyo', 'yyMMdd');
  const prefix = head + '-' + ymd + '-';
  let n = 0;
  if (ledger.getLastRow() > 1) {
    ledger.getRange(2, 1, ledger.getLastRow() - 1, 1).getValues().forEach(r => {
      const v = String(r[0]);
      if (v.indexOf(prefix) === 0) n = Math.max(n, Number(v.slice(prefix.length)) || 0);
    });
  }
  return prefix + String(n + 1).padStart(2, '0');
}

function getFolder_(ss, st) {
  const id = st['保存フォルダID'];
  if (id) return DriveApp.getFolderById(id);
  const folder = DriveApp.createFolder('Canup見積書');
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  const keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => r[0]);
  const row = keys.indexOf('保存フォルダID');
  if (row >= 0) sh.getRange(row + 2, 2).setValue(folder.getId());
  return folder;
}

function exportSheetPdf_(ss, sheet) {
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export'
    + '?format=pdf&gid=' + sheet.getSheetId()
    + '&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false'
    + '&pagenum=UNDEFINED&fzr=false&top_margin=0.6&bottom_margin=0.6&left_margin=0.6&right_margin=0.6'
    + '&r1=0&c1=0&r2=' + (ITEM_FIRST_ROW + ITEM_ROWS + 11) + '&c2=6';
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  return res.getBlob();
}

function safeName_(s) {
  return String(s).replace(/[\\\/:*?"<>|\s]+/g, '_').slice(0, 40);
}
