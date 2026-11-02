/**
 * KIKCAFE アンケート＋クーポン ― Apps Script（本番用の受け口）
 *
 * 役割：
 *   doPost … 回答の登録（actionなし）／クーポンの使用状態の更新（action: "useCoupon"）
 *   doGet  … 確認アプリ（app.html）へ、回答＋クーポン状態の一覧をJSONで返す
 *
 * データは1つのシート「responses」に持ちます（クーポン番号と使用状態も同じ行）。
 * 項目は、依頼者の紙アンケート（Q1〜Q6）に合わせています。設置手順は README.md を参照。
 */
/* 対象スプレッドシート（スタンドアロン形式で設置するため、IDで開く） */
const SPREADSHEET_ID = '1JoyZ9kDN8inpkLcuZ_d_2gcJwJA6QjYAsoGaDTUR6YM';
function book_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

const SHEET_NAME = 'responses';
const HEADER = ['id', 'ts', 'age', 'gender', 'companion', 'source', 'sourceNote',
                'atmosphere', 'comment', 'couponCode', 'couponTitle', 'couponExpires', 'couponUsed'];

function getSheet_() {
  const ss = book_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
  }
  return sh;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

const CREDIT_SHEET = 'credit';
const CREDIT_DEFAULT_TOTAL = 15000;

function getCreditSheet_() {
  const ss = book_();
  let sh = ss.getSheetByName(CREDIT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CREDIT_SHEET);
    sh.appendRow(['kind', 'useId', 'ts', 'amount', 'memo']);
  }
  return sh;
}

function readCredit_() {
  const sh = getCreditSheet_();
  const values = sh.getDataRange().getValues();
  let total = CREDIT_DEFAULT_TOTAL;
  const uses = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === 'total') total = Number(values[i][3]) || 0;
    else if (values[i][0] === 'use') {
      uses.push({
        useId: String(values[i][1]),
        ts: values[i][2] instanceof Date ? values[i][2].toISOString() : String(values[i][2]),
        amount: Number(values[i][3]) || 0,
        memo: String(values[i][4] || '')
      });
    }
  }
  return { total: total, uses: uses };
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const sh = getSheet_();

  if (data.action === 'creditUse') {
    const csh = getCreditSheet_();
    csh.appendRow(['use', data.useId || ('u' + Date.now()),
                   data.ts || new Date().toISOString(),
                   Number(data.amount) || 0, data.memo || '']);
    return jsonOut_({ ok: true, credit: readCredit_() });
  }
  if (data.action === 'creditUndo') {
    const csh = getCreditSheet_();
    const values = csh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][1]) === String(data.useId)) { csh.deleteRow(i + 1); break; }
    }
    return jsonOut_({ ok: true, credit: readCredit_() });
  }
  if (data.action === 'creditTotal') {
    const csh = getCreditSheet_();
    const values = csh.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === 'total') { csh.getRange(i + 1, 4).setValue(Number(data.total) || 0); found = true; break; }
    }
    if (!found) csh.appendRow(['total', '', '', Number(data.total) || 0, '']);
    return jsonOut_({ ok: true, credit: readCredit_() });
  }

  if (data.action === 'useCoupon') {
    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][9]) === String(data.code)) {
        sh.getRange(i + 1, 13).setValue(data.used ? 'used' : '');
        break;
      }
    }
    return jsonOut_({ ok: true });
  }

  const c = data.coupon || {};
  if (c.code) {
    /* 発行済みの番号と照合し、かぶっていたらサーバー側で振り直す（前に出た番号は出さない） */
    const values = sh.getDataRange().getValues();
    const usedCodes = new Set(values.slice(1).map(r => String(r[9])).filter(Boolean));
    while (usedCodes.has(String(c.code))) c.code = makeCode_();
  }
  sh.appendRow([
    data.id || '',
    data.ts || new Date().toISOString(),
    data.age || '',
    data.gender || '',
    data.companion || '',
    data.source || '',
    data.sourceNote || '',
    data.atmosphere || '',
    data.comment || '',
    c.code || '',
    c.title || '',
    c.expires || '',
    c.used ? 'used' : ''
  ]);
  return jsonOut_({ ok: true, code: c.code || '' });
}

function makeCode_() {
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 6; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return 'KF-' + s;
}

function doGet() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  values.shift(); // 見出し行を除く
  const list = values.filter(r => r[0]).map(r => ({
    id: String(r[0]),
    ts: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
    age: String(r[2] || ''),
    gender: String(r[3] || ''),
    companion: String(r[4] || ''),
    source: String(r[5] || ''),
    sourceNote: String(r[6] || ''),
    atmosphere: String(r[7] || ''),
    comment: String(r[8] || ''),
    coupon: r[9] ? {
      code: String(r[9]),
      title: String(r[10] || ''),
      expires: String(r[11] || ''),
      used: String(r[12]) === 'used'
    } : null
  }));
  return jsonOut_({ responses: list, credit: readCredit_() });
}
