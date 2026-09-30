// 株式会社Canup 見積書アプリ
// スプレッドシート「株式会社Canup_物件管理リスト」に組み込むコンテナバインドスクリプト。
// シート: 設定 / 見積書テンプレ / 見積台帳 / 選択肢(F列=見積品目) を使う。

const SHEET_SETTINGS = '設定';
const SHEET_TEMPLATE = '見積書テンプレ';
const SHEET_LEDGER = '見積台帳';
const SHEET_CHOICES = '選択肢';
const SHEET_ANKEN = '案件管理';
const SHEET_INVOICE = '請求台帳';
const SHEET_DETAIL = '発行明細';

// 書類の種類ごとの台帳・番号・期限
const DOCS = {
  quote: { label: '見積書', sheet: SHEET_LEDGER, noHead: '見積番号', prefixKey: '見積番号の頭', prefixDef: 'Q',
    daysKey: '有効期限(日)', limitHead: '有効期限', extra: ['敬称', '備考', '発行元'] },
  invoice: { label: '請求書', sheet: SHEET_INVOICE, noHead: '請求番号', prefixKey: '請求番号の頭', prefixDef: 'INV',
    daysKey: '支払期限(日)', limitHead: '支払期限', extra: ['敬称', '備考', '元の見積番号', '入金日', '発行元'] },
};
// 発行元(送り主)。設定シートの項目名は「頭 + 社名」など。Canupは頭なし、志は「志_」
const ISSUERS = {
  canup: { name: '株式会社Canup', head: '', quoteDef: 'Q', invoiceDef: 'INV', sealName: 'canup_印影_見積用.png' },
  kokorozashi: { name: '株式会社志', head: '志_', quoteDef: 'KQ', invoiceDef: 'KINV', sealName: '志_印影_見積用.png' },
};
const DETAIL_HEAD = ['書類番号', '種別', '発行日', '宛名', '行', '品目', '数量', '単位', '単価', '金額'];
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
    else if (body.action === 'load') out = apiLoad_(ss, body.no);
    else throw new Error('不明な操作です: ' + body.action);
    out.ok = true;
  } catch (err) {
    out = { ok: false, error: err.message || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function apiInit_(ss, st) {
  ensureSettingRows_(ss);
  st = readSettings_(ss);
  const init = getInitData();
  const issuers = {};
  Object.keys(ISSUERS).forEach(k => issuers[k] = issuerInfo_(st, k));
  return {
    issuer: issuers.canup,
    issuers: issuers,
    validDays: init.validDays,
    payDays: Number(st['支払期限(日)']) || 30,
    recipients: init.recipients,
    items: init.items,
    history: history_(ss, 40),
    ready: init.ready && !!ss.getSheetByName(SHEET_LEDGER),
  };
}

function apiIssue_(ss, st, data) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const type = data.docType === 'invoice' ? 'invoice' : 'quote';
    const doc = DOCS[type];
    const ledger = ensureLedger_(ss, type);
    const detail = ensureDetail_(ss);
    const q = calcQuote_(data, st, type);
    const ik = ISSUERS[data.issuerKey] ? data.issuerKey : 'canup';
    const I = ISSUERS[ik];
    const noHead = st[I.head + doc.prefixKey] || (type === 'invoice' ? I.invoiceDef : I.quoteDef);
    const no = nextQuoteNo_(ledger, noHead, q.issue);
    const rec = {
      '発行日': q.issue, '宛名': data.to, '件名': data.subject || '', '税区分': q.taxMode,
      '小計': q.subtotal, '消費税': q.tax, '合計': q.total, 'PDF': '(PDF保存待ち)',
      '敬称': data.honorific || '御中', '備考': data.note || '', '元の見積番号': data.fromNo || '',
      '発行元': issuerInfo_(st, ik).company,
    };
    rec[doc.noHead] = no;
    rec[doc.limitHead] = q.valid;
    const head = headerOf_(ledger);
    ledger.appendRow(head.map(h => rec[h] !== undefined ? rec[h] : ''));
    // 明細は1行ずつ「発行明細」にたまる(集計・呼び出し用)
    const rows = q.items.map((it, i) => [no, doc.label, q.issue, data.to, i + 1, it.name, Number(it.qty), it.unit || '式',
      Number(it.price), Math.round(Number(it.qty) * Number(it.price))]);
    detail.getRange(detail.getLastRow() + 1, 1, rows.length, DETAIL_HEAD.length).setValues(rows);
    return {
      no: no, docType: type,
      issueDate: Utilities.formatDate(q.issue, 'Asia/Tokyo', 'yyyy-MM-dd'),
      validDate: Utilities.formatDate(q.valid, 'Asia/Tokyo', 'yyyy-MM-dd'),
      subtotal: q.subtotal, tax: q.tax, total: q.total,
      issuerKey: ik, issuer: issuerInfo_(st, ik),
    };
  } finally {
    lock.releaseLock();
  }
}

function apiAttachPdf_(ss, st, body) {
  if (!body.no || !body.pdf) throw new Error('書類番号かPDFがありません。');
  const hit = findDoc_(ss, body.no);
  if (!hit) throw new Error('台帳に ' + body.no + ' が見つかりません。');
  const blob = Utilities.newBlob(Utilities.base64Decode(body.pdf), 'application/pdf',
    body.no + '_' + safeName_(hit.rec['宛名']) + '_' + DOCS[hit.type].label + '.pdf');
  const file = getFolder_(ss, st).createFile(blob);
  hit.sheet.getRange(hit.row, hit.head.indexOf('PDF') + 1).setValue(file.getUrl());
  return { url: file.getUrl() };
}

// 過去の書類を画面に呼び出す(明細は「発行明細」から)
function apiLoad_(ss, no) {
  const hit = findDoc_(ss, no);
  if (!hit) throw new Error(no + ' が台帳に見つかりません。');
  const r = hit.rec;
  let items = [];
  const detail = ss.getSheetByName(SHEET_DETAIL);
  if (detail && detail.getLastRow() > 1) {
    items = detail.getRange(2, 1, detail.getLastRow() - 1, DETAIL_HEAD.length).getValues()
      .filter(v => String(v[0]) === String(no))
      .sort((a, b) => a[4] - b[4])
      .map(v => ({ name: String(v[5]), qty: v[6], unit: String(v[7]), price: v[8] }));
  }
  return {
    no: String(no), docType: hit.type, issuerKey: issuerKeyOf_(r['発行元']), to: String(r['宛名'] || ''), honorific: String(r['敬称'] || '御中'),
    subject: String(r['件名'] || ''), taxMode: r['税区分'] === '内税' ? '内税' : '外税', note: String(r['備考'] || ''), items,
  };
}

function history_(ss, limit) {
  const out = [];
  Object.keys(DOCS).forEach(type => {
    const sh = ss.getSheetByName(DOCS[type].sheet);
    if (!sh || sh.getLastRow() < 2) return;
    const head = headerOf_(sh);
    const c = h => head.indexOf(h);
    sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues().forEach(v => {
      if (!v[0]) return;
      const d = v[c('発行日')];
      out.push({ no: String(v[0]), type, date: d instanceof Date ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') : String(d),
        to: String(v[c('宛名')] || ''), subject: String(v[c('件名')] || ''), total: Number(v[c('合計')]) || 0,
        issuerKey: c('発行元') >= 0 ? issuerKeyOf_(v[c('発行元')]) : 'canup' });
    });
  });
  out.sort((a, b) => (b.date + b.no).localeCompare(a.date + a.no));
  return out.slice(0, limit);
}

function findDoc_(ss, no) {
  for (const type of Object.keys(DOCS)) {
    const sh = ss.getSheetByName(DOCS[type].sheet);
    if (!sh || sh.getLastRow() < 2) continue;
    const head = headerOf_(sh);
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]) === String(no)) {
        const rec = {};
        head.forEach((h, j) => rec[h] = vals[i][j]);
        return { type, sheet: sh, row: i + 2, head, rec };
      }
    }
  }
  return null;
}

function headerOf_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
}

// 台帳が無ければ作り、足りない列があれば右に足す(既存の行は動かさない)
function ensureLedger_(ss, type) {
  const doc = DOCS[type];
  let sh = ss.getSheetByName(doc.sheet);
  const want = [doc.noHead, '発行日', '宛名', '件名', '税区分', '小計', '消費税', '合計', doc.limitHead, 'PDF'].concat(doc.extra);
  if (!sh) {
    sh = ss.insertSheet(doc.sheet);
    sh.getRange(1, 1, 1, want.length).setValues([want]);
    sh.setFrozenRows(1);
    [130, 100, 220, 260, 70, 100, 90, 110, 100, 320].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.getRange('B:B').setNumberFormat('yyyy/mm/dd');
    sh.getRange('I:I').setNumberFormat('yyyy/mm/dd');
    sh.getRange('F:H').setNumberFormat('#,##0');
  } else {
    const head = headerOf_(sh);
    const miss = want.filter(h => head.indexOf(h) < 0);
    if (miss.length) sh.getRange(1, head.length + 1, 1, miss.length).setValues([miss]);
  }
  sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
  return sh;
}

function ensureDetail_(ss) {
  let sh = ss.getSheetByName(SHEET_DETAIL);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_DETAIL);
  sh.getRange(1, 1, 1, DETAIL_HEAD.length).setValues([DETAIL_HEAD])
    .setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  [130, 70, 100, 220, 40, 220, 60, 50, 90, 100].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange('C:C').setNumberFormat('yyyy/mm/dd');
  sh.getRange('I:J').setNumberFormat('#,##0');
  return sh;
}

// あとから増えた設定項目を「設定」シートの末尾に足す
function ensureSettingRows_(ss) {
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) return;
  const keys = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => r[0]);
  const add = [
    ['請求番号の頭', 'INV', '請求番号 = 頭-YYMMDD-連番'],
    ['支払期限(日)', 30, '請求書の支払期限の初期値(発行日からの日数)。画面で変えられる'],
    ['志_社名', '株式会社志', '発行元を「株式会社志」にしたときの社名。以下「志_」の行は志で発行するときだけ使う'],
    ['志_郵便番号', '', '空欄なら出ない'],
    ['志_住所', '', '空欄なら出ない(豊中へ移転予定のため未記入)'],
    ['志_電話', '', '空欄なら出ない'],
    ['志_メール', '', '空欄なら出ない'],
    ['志_登録番号', '', 'インボイス登録番号(T+13桁)。空欄なら出ない'],
    ['志_振込先', '', '備考欄の末尾に載る。空欄なら出ない'],
    ['志_印影ファイルID', '', '角印(背景透過PNG)のDriveファイルID。空欄でも「志_印影_見積用.png」という名前の画像がDriveにあれば自動で押す'],
    ['志_見積番号の頭', 'KQ', '志の見積番号 = 頭-YYMMDD-連番(Canupとは別の連番)'],
    ['志_請求番号の頭', 'KINV', '志の請求番号 = 頭-YYMMDD-連番(Canupとは別の連番)'],
  ].filter(r => keys.indexOf(r[0]) < 0);
  if (!add.length) return;
  const at = sh.getLastRow() + 1;
  sh.getRange(at, 1, add.length, 3).setValues(add);
  sh.getRange(at, 1, add.length, 1).setFontWeight('bold');
  sh.getRange(at, 2, add.length, 1).setBackground('#FFF9E6');
}

function issuerInfo_(st, key) {
  const I = ISSUERS[key] || ISSUERS.canup;
  const g = k => st[I.head + k] || '';
  return {
    key: ISSUERS[key] ? key : 'canup',
    company: g('社名') || I.name, zip: g('郵便番号'), address: g('住所'),
    tel: g('電話'), mail: g('メール'), regNo: g('登録番号'), bank: g('振込先'),
    seal: sealDataUrl_(g('印影ファイルID'), I.sealName),
  };
}

function issuerKeyOf_(company) {
  const c = String(company || '');
  return Object.keys(ISSUERS).find(k => c && c === ISSUERS[k].name) || (c.indexOf('志') >= 0 ? 'kokorozashi' : 'canup');
}

// 印影はリポジトリに置かず、Driveの画像を合言葉つきのAPIでだけ渡す。
// ファイルIDが空欄なら、決まった名前の画像がDriveにあればそれを使う(角印ができたら置くだけでよい)
function sealDataUrl_(id, fallbackName) {
  let file = null;
  id = String(id || '').trim();
  if (id) file = DriveApp.getFileById(id);
  else if (fallbackName) {
    const it = DriveApp.getFilesByName(fallbackName);
    if (it.hasNext()) file = it.next();
  }
  if (!file) return '';
  const blob = file.getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

function calcQuote_(data, st, type) {
  const items = (data.items || []).filter(it => it.name && Number(it.qty) && it.price !== '' && it.price !== null);
  if (!data.to) throw new Error('宛名が空です。');
  if (!items.length) throw new Error('品目・数量・単価がそろった行が1つもありません。');
  if (items.length > ITEM_ROWS) throw new Error('品目は' + ITEM_ROWS + '行までです。');
  const issue = data.date ? new Date(data.date + 'T00:00:00+09:00') : new Date();
  let valid;
  if (type === 'invoice' && data.due) valid = new Date(data.due + 'T00:00:00+09:00');
  else valid = new Date(issue.getTime() + (Number(st[DOCS[type || 'quote'].daysKey]) || 30) * 86400000);
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
  ensureSettingRows_(ss);
  ensureLedger_(ss, 'quote');
  ensureLedger_(ss, 'invoice');
  ensureDetail_(ss);
  setupTemplate_(ss);
  setupItemChoices_(ss);
  try {
    SpreadsheetApp.getUi().alert('初期設定が終わりました。「設定」シートの空欄(電話・メール・登録番号・振込先)を埋めてください。');
  } catch (e) {
    // エディタから実行したときは画面が無いので記録だけ残す
    Logger.log('初期設定が終わりました。');
  }
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
    ['印影ファイルID', '', '見積書に押す印影(背景透過PNG)のDriveファイルID。空欄なら押さない'],
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
  [ledger, ss.getSheetByName(SHEET_INVOICE)].forEach(sh => {
    if (sh && sh.getLastRow() > 1) names = names.concat(sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues().map(r => r[0]).reverse());
  });
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
