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
                'atmosphere', 'comment', 'couponCode', 'couponTitle', 'couponExpires', 'couponUsed',
                'wish', 'entry', 'atmoReason', 'entryNote', 'wishOther'];

function getSheet_() {
  const ss = book_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
  } else if (sh.getLastRow() >= 1) {
    /* 既存シートに新しい列（wish / entry）の見出しが無ければ補う */
    const head = sh.getRange(1, 1, 1, HEADER.length).getValues()[0];
    for (let i = 0; i < HEADER.length; i++) {
      if (!head[i]) sh.getRange(1, i + 1).setValue(HEADER[i]);
    }
  }
  return sh;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

const CONFIG_SHEET = 'config';
const CONFIG_DEFAULTS = { couponTitle: '次回ドリンク 200円引き', couponDays: 30 };
const WISH_DEFAULTS = ['甘いもの・ケーキ', 'ドリンクの種類', 'お食事メニュー', '占い', 'イベント・講座', '静かに過ごせる席', 'その他'];
const ENTRY_DEFAULTS = ['特になかった', 'お店の場所が分かりにくかった', '入口の様子で迷った', '中が見えなくて様子が分からなかった', 'その他'];
function parseItems_(v, defaults) {
  const items = String(v || '').split('\n').map(x => x.trim()).filter(Boolean);
  return items.length ? items : defaults.slice();
}

function getConfigSheet_() {
  const ss = book_();
  let sh = ss.getSheetByName(CONFIG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG_SHEET);
    sh.appendRow(['key', 'value']);
  }
  return sh;
}

function readConfig_() {
  const sh = getConfigSheet_();
  const values = sh.getDataRange().getValues();
  const cfg = { couponTitle: CONFIG_DEFAULTS.couponTitle, couponDays: CONFIG_DEFAULTS.couponDays,
                wishItems: WISH_DEFAULTS.slice(), entryItems: ENTRY_DEFAULTS.slice() };
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === 'couponTitle' && values[i][1]) cfg.couponTitle = String(values[i][1]);
    else if (values[i][0] === 'couponDays' && Number(values[i][1])) cfg.couponDays = Number(values[i][1]);
    else if (values[i][0] === 'wishItems') cfg.wishItems = parseItems_(values[i][1], WISH_DEFAULTS);
    else if (values[i][0] === 'entryItems') cfg.entryItems = parseItems_(values[i][1], ENTRY_DEFAULTS);
  }
  return cfg;
}
function cleanItems_(text) {
  return String(text || '').split('\n').map(x => x.trim().slice(0, 40)).filter(Boolean).slice(0, 12).join('\n');
}

function setConfig_(key, value) {
  const sh = getConfigSheet_();
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
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

  if (data.action === 'configCoupon') {
    const title = String(data.title || '').slice(0, 60);
    if (title) setConfig_('couponTitle', title);
    const days = Math.round(Number(data.days) || 0);
    if (days >= 1 && days <= 730) setConfig_('couponDays', days);
    return jsonOut_({ ok: true, config: readConfig_() });
  }

  if (data.action === 'configSurvey') {
    const w = cleanItems_(data.wishItems);
    if (w) setConfig_('wishItems', w);
    const e2 = cleanItems_(data.entryItems);
    if (e2) setConfig_('entryItems', e2);
    return jsonOut_({ ok: true, config: readConfig_() });
  }

  if (data.action === 'chargeEntry') {
    const charge = chargeApply_(data);
    return jsonOut_({ ok: true, charge: charge });
  }
  if (data.action === 'chargeUndo') {
    const charge = chargeUndo_(data);
    return jsonOut_({ ok: true, charge: charge });
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
    c.used ? 'used' : '',
    String(data.wish || ''),
    String(data.entry || ''),
    String(data.atmoReason || ''),
    String(data.entryNote || ''),
    String(data.wishOther || '')
  ]);
  return jsonOut_({ ok: true, code: c.code || '' });
}

/* ===== チャージ（お預かり残高） =====
   お客様ごとに、預かったお金（in）とお会計での利用（out）を記録します。
   残高はアプリ側で記録から計算します。正本はシート「charge」です。 */
const CHARGE_SHEET = 'charge';

function getChargeSheet_() {
  const ss = book_();
  let sh = ss.getSheetByName(CHARGE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CHARGE_SHEET);
    sh.appendRow(['entryId', 'accountId', 'name', 'kind', 'amount', 'ts', 'memo']);
  }
  return sh;
}

function readCharge_() {
  const sh = getChargeSheet_();
  const values = sh.getDataRange().getValues();
  const entries = [];
  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    entries.push({
      entryId: String(values[i][0]),
      accountId: String(values[i][1]),
      name: String(values[i][2] || ''),
      kind: String(values[i][3] || ''),
      amount: Number(values[i][4]) || 0,
      ts: values[i][5] instanceof Date ? values[i][5].toISOString() : String(values[i][5]),
      memo: String(values[i][6] || '')
    });
  }
  return { entries: entries };
}

function chargeApply_(p) {
  const kind = String(p.kind || '');
  if (['open', 'in', 'out'].indexOf(kind) < 0) throw new Error('bad kind');
  const sh = getChargeSheet_();
  const entryId = 'c' + new Date().getTime().toString(36) + Math.floor(Math.random() * 1e6).toString(36);
  sh.appendRow([
    entryId,
    String(p.accountId || ''),
    String(p.name || '').slice(0, 60),
    kind,
    Number(p.amount) || 0,
    p.ts || new Date().toISOString(),
    String(p.memo || '').slice(0, 200)
  ]);
  return readCharge_();
}

function chargeUndo_(p) {
  const sh = getChargeSheet_();
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(p.entryId)) { sh.deleteRow(i + 1); break; }
  }
  return readCharge_();
}

function makeCode_() {
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 6; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return 'KF-' + s;
}

function doGet(e) {
  /* フォーム用：設定だけを返す（回答データは渡さない） */
  if (e && e.parameter && e.parameter.mode === 'config') {
    return jsonOut_({ config: readConfig_() });
  }
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
    } : null,
    wish: String(r[13] || ''),
    entry: String(r[14] || ''),
    atmoReason: String(r[15] || ''),
    entryNote: String(r[16] || ''),
    wishOther: String(r[17] || '')
  }));
  return jsonOut_({ responses: list, credit: readCredit_(), config: readConfig_(), charge: readCharge_() });
}
