/**
 * 快速收銀 · Google 試算表同步程式
 *
 * 安裝方式：
 *  1. 開一個新的 Google 試算表 →「擴充功能」→「Apps Script」
 *  2. 把編輯器裡原本的內容全部刪掉，貼上這整份程式
 *  3. 改下面兩行密碼（老闆密碼和員工密碼不要一樣）
 *  4. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *     執行身分：我　　誰可以存取：所有人
 *  5. 按「部署」並授權，複製「網頁應用程式網址」貼到收銀 App
 */

const ADMIN_KEY = '改成老闆密碼';   // 老闆用：改商品、看全部人的統計
const STAFF_KEY = '改成員工密碼';   // 員工用：上傳自己的成交紀錄

// ---------------------------------------------------------------

function doGet() {
  return out({ ok: true, app: 'quick-pos', msg: '同步程式運作中' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return out({ ok: false, error: '資料格式錯誤' }); }

  if (ADMIN_KEY.indexOf('改成') === 0 || STAFF_KEY.indexOf('改成') === 0) {
    return out({ ok: false, error: '請先在程式裡設定老闆和員工密碼' });
  }
  const role = req.key === ADMIN_KEY ? 'admin' : (req.key === STAFF_KEY ? 'staff' : '');
  if (!role) return out({ ok: false, error: '密碼錯誤' });

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    switch (req.action) {
      case 'hello':
        return out({ ok: true, role: role, staffKey: role === 'admin' ? STAFF_KEY : undefined });
      case 'push':
        pushOps(req.ops || [], role, String(req.dev || ''));
        return out({ ok: true });
      case 'getConfig':
        return out({ ok: true, config: getConfig() });
      case 'setConfig':
        if (role !== 'admin') return out({ ok: false, error: '只有老闆可以修改商品' });
        setConfig(req.config);
        return out({ ok: true });
      case 'day':
        if (role !== 'admin') return out({ ok: false, error: '只有老闆可以看全部統計' });
        return out({ ok: true, orders: dayOrders(String(req.date || '')) });
      default:
        return out({ ok: false, error: '未知的動作' });
    }
  } finally {
    lock.releaseLock();
  }
}

// ---------- sheets ----------

function sheet(name, headers, hidden, plainText) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (plainText) sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    if (headers) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    if (hidden) sh.hideSheet();
  }
  return sh;
}
const RAW = () => sheet('_raw', ['單號', '日期', '資料', '已刪除', '裝置'], true, true);
const DETAIL = () => sheet('成交明細', ['日期', '時間', '班別', '人員', '品項', '單價', '數量', '小計', '整單合計', '單號', '狀態']);
const CONF = () => sheet('_設定', null, true);

function findRow(sh, col, value) {
  if (sh.getLastRow() < 2) return 0;
  const hit = sh.getRange(2, col, sh.getLastRow() - 1, 1)
    .createTextFinder(String(value)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

// ---------- orders ----------

function pushOps(ops, role, dev) {
  const raw = RAW(), det = DETAIL();
  const newRaw = [], newDet = [];
  const seen = {};
  ops.forEach(function (op) {
    const o = op.order || {};
    if (!o.id) return;
    if (op.op === 'add') {
      if (seen[o.id] || findRow(raw, 1, o.id)) return;   // 已經上傳過
      seen[o.id] = true;
      newRaw.push([o.id, o.date, JSON.stringify(o), '', dev]);
      const t = new Date(o.ts);
      const hm = Utilities.formatDate(t, 'Asia/Taipei', 'HH:mm');
      (o.items || []).forEach(function (i) {
        newDet.push([o.date, hm, o.shiftName || '', o.by || '', i.name, i.price, i.qty, i.price * i.qty, o.total, o.id, '']);
      });
    } else if (op.op === 'del') {
      const r = findRow(raw, 1, o.id);
      if (!r) return;
      if (role !== 'admin' && String(raw.getRange(r, 5).getValue()) !== dev) return; // 員工只能刪自己的
      raw.getRange(r, 4).setValue('Y');
      if (det.getLastRow() >= 2) {
        det.getRange(2, 10, det.getLastRow() - 1, 1)
          .createTextFinder(String(o.id)).matchEntireCell(true).findAll()
          .forEach(function (c) { det.getRange(c.getRow(), 11).setValue('已刪除'); });
      }
    }
  });
  if (newRaw.length) { ensureRows(raw, newRaw.length); raw.getRange(raw.getLastRow() + 1, 1, newRaw.length, 5).setValues(newRaw); }
  if (newDet.length) { ensureRows(det, newDet.length); det.getRange(det.getLastRow() + 1, 1, newDet.length, 11).setValues(newDet); }
}

function ensureRows(sh, n) {
  const free = sh.getMaxRows() - sh.getLastRow();
  if (free < n) sh.insertRowsAfter(sh.getMaxRows(), n - free + 500);   // 新列會沿用上一列的格式
}

function dayOrders(date) {
  const raw = RAW();
  if (raw.getLastRow() < 2) return [];
  return raw.getRange(2, 1, raw.getLastRow() - 1, 4).getValues()
    .filter(function (r) { return String(r[1]) === date && String(r[3]) !== 'Y'; })
    .map(function (r) { try { return JSON.parse(r[2]); } catch (e) { return null; } })
    .filter(function (o) { return o; });
}

// ---------- product settings ----------

function getConfig() {
  const v = CONF().getRange(1, 1).getValue();
  if (!v) return null;
  try { return JSON.parse(v); } catch (e) { return null; }
}
function setConfig(c) {
  CONF().getRange(1, 1).setValue(JSON.stringify(c || {}));
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
