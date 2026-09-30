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
    .addItem('閲覧用ファイルに反映', 'syncViewMenu')
    .addSeparator()
    .addItem('初期設定(最初に1回)', 'setup')
    .addToUi();
}

// 物件管理: 受注状況が「受注」の行だけ支払確認のチェックボックスを出す(それ以外は消す)
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== '物件管理') return;
  const head = sh.getRange(3, 1, 1, sh.getLastColumn()).getValues()[0];
  const cs = head.indexOf('受注状況') + 1, cc = head.indexOf('支払確認') + 1;
  if (!cs || !cc || e.range.getColumn() > cs || e.range.getLastColumn() < cs) return;
  for (let r = Math.max(e.range.getRow(), 4); r <= e.range.getLastRow(); r++) {
    const box = sh.getRange(r, cc);
    if (sh.getRange(r, cs).getValue() === '受注') {
      if (!box.getDataValidation()) box.insertCheckboxes();
    } else {
      box.clearDataValidations().clearContent();
    }
  }
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
    // 合言葉は2種類。管理者=案件管理の全列 / 担当=案件管理の「公開する列」だけ。見積書・請求書はどちらでも出せる(普段使うのは担当)
    const adminKey = String(st['アプリの合言葉'] || '');
    const staffKey = String(st['案件_担当の合言葉'] || '');
    let role = '';
    if (adminKey && body.key === adminKey) role = 'admin';
    else if (staffKey && body.key === staffKey) role = 'staff';
    if (!role) throw new Error('合言葉が違います。スプレッドシートの「設定」シートにある合言葉を入れ直してください。');
    if (body.action === 'init') out = apiInit_(ss, st);
    else if (body.action === 'issue') out = apiIssue_(ss, st, body.data || {});
    else if (body.action === 'attachPdf') out = apiAttachPdf_(ss, st, body);
    else if (body.action === 'load') out = apiLoad_(ss, body.no);
    else if (body.action === 'ankenList') out = apiAnkenList_(ss, st, role);
    else if (body.action === 'ankenSave') out = apiAnkenSave_(ss, st, role, body);
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
  const file = docFolder_(ss, hit.type, issuerKeyOf_(hit.rec['発行元'])).createFile(blob);
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
    ['案件_担当の合言葉', Utilities.getUuid().replace(/-/g, ''), '担当者用。見積書・請求書は出せる。案件管理は「公開する列」だけ見る・直せる。変えると担当者の端末で入れ直し'],
    ['公開する列', PUBLIC_COLS_DEFAULT.join(','), '閲覧用ファイルと担当者用の画面に出す「物件管理」の列(カンマ区切り)。ここに無い列は原本にだけ残り、管理者の画面でだけ見える'],
    ['閲覧用ファイルID', '', '社内共有用の閲覧専用ファイル。空欄なら初回にこのファイルと同じフォルダに作る。共有はこのファイルだけにする'],
    ['閲覧用の最終反映', '', '自動で入る。日付が変わって最初に案件管理を開いたときにも反映し直す(支払アラートの日付計算のため)'],
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
    const folder = docFolder_(ss, 'quote', 'canup');
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

// 発行したPDFの置き場: スプレッドシートと同じフォルダ(Canup&志)の下に 会社名/見積書・請求書。無ければ作る
// (Canupと志は別の会社なので、会社ごと・書類ごとに分ける)
function docFolder_(ss, type, issuerKey) {
  const parents = DriveApp.getFileById(ss.getId()).getParents();
  const base = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const I = ISSUERS[issuerKey] || ISSUERS.canup;
  const sub = (folder, name) => { const it = folder.getFoldersByName(name); return it.hasNext() ? it.next() : folder.createFolder(name); };
  return sub(sub(base, I.name), DOCS[type] ? DOCS[type].label : '見積書');
}

// 1回だけ使う: 発行済みのPDFを 会社名/見積書・請求書 のフォルダへ移す(台帳のリンクはファイルIDなので変わらない)
function moveIssuedPdfs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const moved = [];
  Object.keys(DOCS).forEach(type => {
    const sh = ss.getSheetByName(DOCS[type].sheet);
    if (!sh || sh.getLastRow() < 2) return;
    const head = headerOf_(sh);
    const cPdf = head.indexOf('PDF'), cIss = head.indexOf('発行元');
    sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues().forEach(r => {
      const m = String(r[cPdf] || '').match(/\/d\/([\w-]+)/);
      if (!m) return;
      const to = docFolder_(ss, type, issuerKeyOf_(cIss >= 0 ? r[cIss] : ''));
      DriveApp.getFileById(m[1]).moveTo(to);
      moved.push(r[0] + ' -> ' + to.getName());
    });
  });
  Logger.log(moved.length + '件を移動しました\n' + moved.join('\n'));
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

// ---------- 案件管理(anken.html)から呼ぶ ----------
// 原本=このファイルの「物件管理」。アプリで読み書きし、公開する列だけを別ファイル(閲覧用)に書き写す。
// 閲覧用ファイルにはこのファイルの他のシート(設定・台帳)が入らないので、社内にはそちらだけを共有する。

const SHEET_BUKKEN = '物件管理';
const BUKKEN_HEAD_ROW = 3;
const PUBLIC_COLS_DEFAULT = ['No', '物件名', '支払い予定日', '受注状況', '支払確認', '支払アラート', '支払方法', '売上', '手数料(Canup)', '業務委託料', '担当者', '備考'];
const STATUS_COLORS = { '受注': '#DCE8F7', '提案中': '#FFF6D6' };
const ALERT_COLORS = {
  '期限超過・未確認': ['#9C1C1C', '#FFFFFF'], '7日以内': ['#F8C9A0', null], '14日以内': ['#FFF2B3', null], '確認済': ['#D5EDDA', null],
};

function publicCols_(st) {
  const v = String(st['公開する列'] || '').trim();
  return v ? v.split(/[,、，]/).map(x => x.trim()).filter(String) : PUBLIC_COLS_DEFAULT;
}

// 物件管理の列の並びと、データのある範囲を調べる。列の型は入力規則・数式・表示形式から決める(列を足しても画面が追従する)
function bukkenLayout_(ss) {
  const sh = ss.getSheetByName(SHEET_BUKKEN);
  if (!sh) throw new Error('「物件管理」シートが見つかりません。');
  const head = sh.getRange(BUKKEN_HEAD_ROW, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v).trim());
  let width = head.indexOf('');
  if (width < 0) width = head.length;
  const nameIdx = head.indexOf('物件名');
  if (nameIdx < 0) throw new Error('物件管理の' + BUKKEN_HEAD_ROW + '行目に「物件名」の見出しがありません。');
  // データ範囲 = 見出しの次の行から、A列に数字以外の文字が出る行(下の仕様メモ)の手前まで
  const start = BUKKEN_HEAD_ROW + 1;
  const lastRow = Math.max(sh.getLastRow(), start);
  const colA = sh.getRange(start, 1, lastRow - start + 1, 1).getValues();
  let end = start - 1;
  for (let i = 0; i < colA.length; i++) {
    const v = colA[i][0];
    if (v !== '' && typeof v !== 'number') break;
    end = start + i;
  }
  const n = Math.max(end - start + 1, 0);
  const rng = n ? sh.getRange(start, 1, n, width) : null;
  const vals = rng ? rng.getValues() : [];
  const fmls = rng ? rng.getFormulas() : [];
  const dvs = rng ? rng.getDataValidations() : [];
  const nfs = rng ? rng.getNumberFormats() : [];
  const fields = head.slice(0, width).map((h, c) => {
    const f = { name: h, type: 'text' };
    if (fmls.some(r => r[c])) { f.type = 'auto'; }
    else {
      for (let i = 0; i < dvs.length; i++) {
        const dv = dvs[i][c];
        if (!dv) continue;
        const t = dv.getCriteriaType();
        if (t === SpreadsheetApp.DataValidationCriteria.CHECKBOX) { f.type = 'check'; break; }
        if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
          f.type = 'select';
          f.options = dv.getCriteriaValues()[0].getValues().map(r => String(r[0]).trim()).filter(String);
          break;
        }
        if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          f.type = 'select'; f.options = dv.getCriteriaValues()[0].map(String); break;
        }
      }
      if (f.type === 'text') {
        const nf = nfs.map(r => String(r[c])).find(x => x && x !== 'General' && x !== '@') || '';
        if (vals.some(r => r[c] instanceof Date) || /y{2,4}/i.test(nf)) f.type = 'date';
        else if (/[¥￥]|#,##0/.test(nf) || vals.some(r => typeof r[c] === 'number' && r[c] >= 1000)) f.type = 'money';
        else if (h === '備考') f.type = 'note';
      }
    }
    return f;
  });
  return { sh, head: head.slice(0, width), width, nameIdx, start, end, vals, fields };
}

function cellOut_(v, type) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  if (type === 'check') return v === true;
  return v === null || v === undefined ? '' : v;
}

function apiAnkenList_(ss, st, role) {
  ensureSettingRows_(ss);
  st = readSettings_(ss);
  // 支払アラートは日付で変わるので、日付が変わって最初に開いたときに閲覧用も反映し直す
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  let viewUrl = '';
  const lastSync = st['閲覧用の最終反映'] instanceof Date
    ? Utilities.formatDate(st['閲覧用の最終反映'], 'Asia/Tokyo', 'yyyy-MM-dd') : String(st['閲覧用の最終反映'] || '').slice(0, 10);
  if (lastSync !== today) {
    try { viewUrl = syncView_(ss, st); } catch (e) { /* 反映の失敗で一覧を止めない */ }
  }
  const L = bukkenLayout_(ss);
  const pub = publicCols_(st);
  const show = L.fields.map((f, c) => role === 'admin' || pub.indexOf(f.name) >= 0 ? c : -1).filter(c => c >= 0);
  const rows = [];
  L.vals.forEach((r, i) => {
    if (String(r[L.nameIdx]).trim() === '') return;
    const v = {};
    show.forEach(c => v[L.fields[c].name] = cellOut_(r[c], L.fields[c].type));
    rows.push({ row: L.start + i, v });
  });
  const id = st['閲覧用ファイルID'];
  return {
    role, fields: show.map(c => Object.assign({ public: pub.indexOf(L.fields[c].name) >= 0 }, L.fields[c])),
    rows, today,
    viewUrl: viewUrl || (id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : ''), // 閲覧用は社内共有用なので担当にも渡す
  };
}

// body: { row: 行番号(新規はnull), name: 読み込んだ時の物件名(取り違え防止), values: {列名: 値} }
function apiAnkenSave_(ss, st, role, body) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const L = bukkenLayout_(ss);
    const sh = L.sh;
    const pub = publicCols_(st);
    const values = body.values || {};
    let row = Number(body.row) || 0;
    if (row) {
      if (row < L.start || row > L.end) throw new Error('行番号がデータの範囲外です。一覧を読み直してください。');
      const now = String(L.vals[row - L.start][L.nameIdx]).trim();
      if (now !== String(body.name || '').trim()) throw new Error('この行は別の場所で書き換えられています(今は「' + now + '」)。一覧を読み直してから直してください。');
    } else {
      if (!String(values['物件名'] || '').trim()) throw new Error('物件名が空です。');
      row = newBukkenRow_(L);
    }
    const statusIdx = L.head.indexOf('受注状況');
    const status = statusIdx >= 0 && values['受注状況'] !== undefined ? String(values['受注状況']) : (statusIdx >= 0 ? String(sh.getRange(row, statusIdx + 1).getValue()) : '');
    L.fields.forEach((f, c) => {
      if (f.type === 'auto' || values[f.name] === undefined) return;
      if (role !== 'admin' && pub.indexOf(f.name) < 0) return;
      const cell = sh.getRange(row, c + 1);
      let v = values[f.name];
      if (f.type === 'check') return; // 受注状況に合わせて下でまとめて扱う
      if (f.type === 'date') v = v ? new Date(String(v) + 'T00:00:00+09:00') : '';
      else if (f.type === 'money') v = v === '' || v === null ? '' : Number(String(v).replace(/[^\d.-]/g, ''));
      else if (f.type === 'select') { v = String(v || ''); if (v && f.options && f.options.indexOf(v) < 0) throw new Error(f.name + 'に「' + v + '」は選べません。'); }
      else v = String(v == null ? '' : v);
      if (f.type === 'money' && v !== '' && isNaN(v)) throw new Error(f.name + 'は数字で入れてください。');
      cell.setValue(v);
    });
    // 支払確認: 受注の行だけチェックボックスを出す(シートのonEditと同じ動き。スクリプトからの書き込みではonEditが動かないため)
    const checkIdx = L.head.indexOf('支払確認');
    if (checkIdx >= 0 && statusIdx >= 0 && (role === 'admin' || pub.indexOf('支払確認') >= 0)) {
      const box = sh.getRange(row, checkIdx + 1);
      if (status === '受注') {
        if (!box.getDataValidation()) box.insertCheckboxes();
        if (values['支払確認'] !== undefined) box.setValue(values['支払確認'] === true);
      } else {
        box.clearDataValidations().clearContent();
      }
    }
    SpreadsheetApp.flush();
    let viewErr = '';
    try { syncView_(ss, readSettings_(ss)); } catch (e) { viewErr = e.message || String(e); }
    const out = apiAnkenList_(ss, readSettings_(ss), role);
    out.saved = row;
    out.viewErr = viewErr;
    return out;
  } finally {
    lock.releaseLock();
  }
}

// 新しい行: データ範囲で物件名が空の最初の行。無ければ最後の行の下に1行足す。
// その行に数式(No・支払アラート等)が無ければ、上の行から数式と書式を写す
function newBukkenRow_(L) {
  const sh = L.sh;
  let row = 0;
  for (let i = 0; i < L.vals.length; i++) {
    if (String(L.vals[i][L.nameIdx]).trim() === '') { row = L.start + i; break; }
  }
  if (!row) {
    sh.insertRowAfter(L.end);
    row = L.end + 1;
  }
  const src = row > L.start ? row - 1 : 0;
  const fAuto = L.fields.map(f => f.type === 'auto');
  if (src && fAuto.some(Boolean)) {
    const tf = sh.getRange(row, 1, 1, L.width).getFormulas()[0];
    if (fAuto.some((a, c) => a && !tf[c])) {
      sh.getRange(src, 1, 1, L.width).copyTo(sh.getRange(row, 1, 1, L.width));
      // 写したのは数式と書式だけにしたいので、入力する列の値と入力規則(チェックボックス)を空にする
      L.fields.forEach((f, c) => {
        if (f.type === 'auto') return;
        const cell = sh.getRange(row, c + 1);
        if (f.type === 'check') cell.clearDataValidations();
        cell.clearContent();
      });
    }
  }
  return row;
}

function syncViewMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSettingRows_(ss);
  const url = syncView_(ss, readSettings_(ss));
  try {
    SpreadsheetApp.getUi().alert('閲覧用ファイルに反映しました。\n' + url);
  } catch (e) {
    Logger.log('閲覧用ファイルに反映しました。' + url); // エディタから実行したとき
  }
}

// 閲覧用ファイルに「公開する列」だけを書き写す。値と表示形式だけを写し、数式・入力規則は持ち込まない
function syncView_(ss, st) {
  const L = bukkenLayout_(ss);
  const pub = publicCols_(st);
  const cols = L.fields.map((f, c) => pub.indexOf(f.name) >= 0 ? c : -1).filter(c => c >= 0);
  let vs;
  const id = String(st['閲覧用ファイルID'] || '').trim();
  if (id) vs = SpreadsheetApp.openById(id);
  else {
    vs = SpreadsheetApp.create('株式会社Canup_物件管理リスト(閲覧用)');
    vs.setSpreadsheetTimeZone('Asia/Tokyo');
    const parents = DriveApp.getFileById(ss.getId()).getParents();
    if (parents.hasNext()) DriveApp.getFileById(vs.getId()).moveTo(parents.next());
    writeSetting_(ss, '閲覧用ファイルID', vs.getId());
  }
  let sh = vs.getSheets()[0];
  sh.setName('物件管理');
  vs.getSheets().slice(1).forEach(x => vs.deleteSheet(x));
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getDataRange().clearDataValidations();

  const srcRows = [];
  const srcFmts = [];
  const srcGroups = [];
  if (L.end >= L.start) {
    const nfs = L.sh.getRange(L.start, 1, L.end - L.start + 1, L.width).getNumberFormats();
    // 並び: 対応中 → 入金済み → 失注(同じ組の中は原本の順)。原本の行は動かさない(アプリは行番号で案件を見分けるため)
    const payC = L.head.indexOf('支払確認'), stC = L.head.indexOf('受注状況');
    const group = r => stC >= 0 && r[stC] === '失注' ? 2 : payC >= 0 && r[payC] === true ? 1 : 0;
    const picked = [];
    L.vals.forEach((r, i) => { if (String(r[L.nameIdx]).trim() !== '') picked.push({ r, i, g: group(r) }); });
    picked.sort((a, b) => a.g - b.g || a.i - b.i);
    picked.forEach(({ r, i, g }) => {
      srcRows.push(cols.map(c => L.fields[c].type === 'check' ? (r[c] === true ? '済' : r[c] === false ? '未' : '') : r[c]));
      srcFmts.push(cols.map(c => L.fields[c].type === 'check' ? '@' : nfs[i][c]));
      srcGroups.push(g);
    });
  }
  const w = cols.length;
  if (sh.getMaxColumns() < w) sh.insertColumnsAfter(sh.getMaxColumns(), w - sh.getMaxColumns());
  const name = c => L.fields[c].name, type = c => L.fields[c].type;
  const idx = n => cols.findIndex(c => name(c) === n);
  // 短い項目は中央ぞろえ、金額は右ぞろえ、物件名と備考は左ぞろえ
  const align = cols.map(c => type(c) === 'money' || (type(c) === 'auto' && name(c) === '業務委託料') ? 'right'
    : name(c) === '物件名' || name(c) === '備考' ? 'left' : 'center');

  // 見出し
  const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  sh.getRange(1, 1).setValue('株式会社Canup 物件管理リスト(閲覧用)').setFontSize(16).setFontWeight('bold');
  sh.getRange(2, 1).setValue('最終更新 ' + stamp + '　案件管理アプリから自動で書き写しています。このファイルは直さないでください。').setFontColor('#5B6470');
  sh.getRange(3, 1, 1, w).setValues([cols.map(name)])
    .setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 34); sh.setRowHeight(2, 22); sh.setRowHeight(3, 32);

  // 本体: 組ごとに見出し行(対応中 / 入金済み / 失注)を挟む
  const GROUPS = [['対応中', '#E6EBF3'], ['入金済み', '#E4F1E7'], ['失注', '#ECEEF1']];
  let r0 = 4;
  const aIdx = idx('支払アラート'), sIdx = idx('受注状況'), pIdx = idx('支払確認');
  GROUPS.forEach((G, g) => {
    const members = srcRows.map((row, i) => ({ row, fmt: srcFmts[i], g: srcGroups[i] })).filter(x => x.g === g);
    if (!members.length) return;
    // 見出し行の文字は物件名の列に置く(左端のNo列は狭く、文字が切れるため)
    sh.getRange(r0, 1, 1, w).setBackground(G[1]).setFontWeight('bold').setFontSize(12).setFontColor('#1F3864').setVerticalAlignment('middle');
    sh.getRange(r0, Math.max(idx('物件名'), 0) + 1).setValue('■ ' + G[0] + '　' + members.length + '件');
    sh.setRowHeight(r0, 30);
    r0++;
    const n = members.length;
    const body = sh.getRange(r0, 1, n, w);
    body.setNumberFormats(members.map(x => x.fmt)).setValues(members.map(x => x.row))
      .setVerticalAlignment('middle').setFontSize(11)
      .setHorizontalAlignments(members.map(() => align))
      .setBorder(true, true, true, true, true, true, '#D3D9E0', SpreadsheetApp.BorderStyle.SOLID);
    const bg = members.map((x, i) => x.row.map(() => i % 2 ? '#F7F9FC' : '#FFFFFF'));
    const fc = members.map(x => x.row.map(() => g === 2 ? '#9AA1A9' : g === 1 ? '#5B6470' : '#1B1F24'));
    const fw = members.map(x => x.row.map(() => 'normal'));
    members.forEach((x, i) => {
      if (g === 2) return;
      const st = sIdx >= 0 ? String(x.row[sIdx]) : '';
      if (sIdx >= 0 && STATUS_COLORS[st]) { bg[i][sIdx] = STATUS_COLORS[st]; fw[i][sIdx] = 'bold'; }
      if (pIdx >= 0 && x.row[pIdx] === '済') { bg[i][pIdx] = '#D5EDDA'; fc[i][pIdx] = '#1D6B3A'; fw[i][pIdx] = 'bold'; }
      if (pIdx >= 0 && x.row[pIdx] === '未') { fc[i][pIdx] = '#A3261E'; fw[i][pIdx] = 'bold'; }
      const col = aIdx >= 0 ? ALERT_COLORS[String(x.row[aIdx])] : null;
      if (col) { bg[i][aIdx] = col[0]; fw[i][aIdx] = 'bold'; if (col[1]) fc[i][aIdx] = col[1]; }
      const nIdx = idx('物件名');
      if (nIdx >= 0) fw[i][nIdx] = 'bold';
    });
    body.setBackgrounds(bg).setFontColors(fc).setFontWeights(fw);
    for (let k = 0; k < n; k++) sh.setRowHeight(r0 + k, 28);
    const noteIdx = idx('備考');
    if (noteIdx >= 0) sh.getRange(r0, noteIdx + 1, n, 1).setWrap(true).setFontSize(10).setFontColor(g === 2 ? '#9AA1A9' : '#5B6470');
    r0 += n;
  });
  sh.setFrozenRows(3);
  sh.setFrozenColumns(0);
  cols.forEach((c, i) => {
    const t = type(c), n = name(c);
    sh.setColumnWidth(i + 1, n === '物件名' ? 230 : n === '備考' ? 380 : n === 'No' ? 44 : n === '支払アラート' ? 136 : n === '支払確認' ? 72 : n === '担当者' ? 72 : t === 'money' || n === '業務委託料' ? 112 : t === 'date' ? 108 : 96);
  });
  const extra = sh.getMaxColumns() - w;
  if (extra > 0) sh.deleteColumns(w + 1, extra);
  writeSetting_(ss, '閲覧用の最終反映', Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'));
  return 'https://docs.google.com/spreadsheets/d/' + vs.getId() + '/edit';
}

function writeSetting_(ss, key, value) {
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  const keys = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => r[0]);
  const i = keys.indexOf(key);
  if (i >= 0) sh.getRange(i + 1, 2).setValue(value);
  else sh.appendRow([key, value, '']);
}
