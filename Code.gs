/*************************************************************************
 * PDF 智慧解析比賽報名系統  Code.gs（Phase 1 後端：GAS + Sheets + Drive）
 * 技術：Google Apps Script + Google Sheets + Google Drive + 單檔 HTML
 * 產出：Claude Code，可直接貼到 Apps Script 使用
 *
 * ── 部署步驟（重要，照做才會「連動」）─────────────────────────
 * 1. 用「會放正式資料」的 Google 帳號，新建一份 Google Sheets。
 * 2. 在該試算表開「擴充功能 → Apps Script」（這樣才是 container-bound，
 *    SpreadsheetApp.getActiveSpreadsheet() 會抓到這份表）。
 * 3. 把本檔內容貼成 Code.gs；再新增一個 HTML 檔，檔名打「Index」
 *    （不含 .html），貼入 index.html 的完整內容。
 * 4. 先在編輯器選函式 setupSheets 執行一次（建立資料表＋示範活動）。
 *    第一次會跳授權，按「允許」。
 * 5. 「部署 → 新增部署 → 類型選『網頁應用程式』」，
 *    執行身分＝我，存取權限＝「任何人」。取得 /exec 網址。
 * 6. 把 /exec 網址貼給家長/教練即可，資料全部進這份試算表。
 *
 * ── 改版重新部署（務必照做，否則 /exec 不會更新或會接到空白表）──
 *    部署 → 管理部署作業 → ✏️ → 版本選「新版本」→ 部署。
 *    千萬不要「新增部署」或另開新專案（會綁到另一份空白試算表）。
 *
 * ── 角色 ──────────────────────────────────────────────────
 *   家長：填報名表（免登入）。 教練：密碼登入後台（預設 1234，可改）。
 *************************************************************************/

const SHEETS = {
  EVENTS: 'Events',
  REGS: 'Registrations',
  SETTINGS: 'Settings'
};

/** 各表欄位（第一列標題）。⚠ 新增欄位一律 append 在最後，不可插中間（會整批錯位） */
const HEADERS = {
  Events: ['id', 'name', 'date', 'location', 'deadline', 'json', 'createdAt'],
  Registrations: ['id', 'registrationNumber', 'eventId', 'eventName', 'studentName',
                  'gender', 'birthday', 'school', 'grade', 'belt', 'items', 'group',
                  'weight', 'parentName', 'parentPhone', 'suggestedWeightClass', 'fee',
                  'paid', 'reviewStatus', 'uploadedAttachments', 'missingAttachments',
                  'consent', 'createdAt', 'updatedAt'],
  Settings: ['key', 'value']
};

const ATTACH_FOLDER_NAME = '比賽報名附件';

/* ===================== Web App 入口 ===================== */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('PDF 智慧解析比賽報名系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ===================== 共用工具 ===================== */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function uid_(prefix) {
  return (prefix || 'id') + '-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);
}

function readAll_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0];
  return values.slice(1).map(function (row) {
    const obj = {};
    head.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function appendRow_(name, obj) {
  const sh = sheet_(name);
  const head = HEADERS[name];
  sh.appendRow(head.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; }));
  return obj;
}

function findRowIndex_(name, idValue) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  const col = values[0].indexOf('id');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][col]) === String(idValue)) return i + 1;
  }
  return -1;
}

/** 依 id upsert 一整列（不存在就 append） */
function upsertRow_(name, obj) {
  const sh = sheet_(name);
  const head = HEADERS[name];
  const rowIndex = findRowIndex_(name, obj.id);
  const rowArr = head.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  if (rowIndex < 0) sh.appendRow(rowArr);
  else sh.getRange(rowIndex, 1, 1, head.length).setValues([rowArr]);
  return obj;
}

function patchRow_(name, idValue, patch) {
  const sh = sheet_(name);
  const rowIndex = findRowIndex_(name, idValue);
  if (rowIndex < 0) return false;
  const head = HEADERS[name];
  const current = sh.getRange(rowIndex, 1, 1, head.length).getValues()[0];
  head.forEach(function (h, i) { if (patch[h] !== undefined) current[i] = patch[h]; });
  sh.getRange(rowIndex, 1, 1, head.length).setValues([current]);
  return true;
}

function deleteRow_(name, idValue) {
  const rowIndex = findRowIndex_(name, idValue);
  if (rowIndex < 0) return false;
  sheet_(name).deleteRow(rowIndex);
  return true;
}

function json_(v, fallback) {
  try { return JSON.parse(v); } catch (e) { return fallback === undefined ? v : fallback; }
}

/* ===================== Settings ===================== */
function getSetting_(key, def) {
  const s = readAll_(SHEETS.SETTINGS).filter(function (r) { return r.key === key; })[0];
  return s ? String(s.value) : (def === undefined ? '' : def);
}
function setSetting_(key, value) {
  const sh = sheet_(SHEETS.SETTINGS);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(key)) { sh.getRange(i + 1, 2).setValue(value); return; }
  }
  sh.appendRow([key, value]);
}

/* ===================== 物件 <-> 列 轉換 ===================== */
function eventToRow_(ev) {
  return {
    id: ev.id, name: ev.name || '', date: ev.date || '', location: ev.location || '',
    deadline: ev.registrationDeadline || '', json: JSON.stringify(ev),
    createdAt: ev.createdAt || new Date().toISOString()
  };
}
function rowToEvent_(row) {
  const ev = json_(row.json, {}) || {};
  ev.id = ev.id || row.id;
  return ev;
}

function regToRow_(r) {
  return {
    id: r.id, registrationNumber: r.registrationNumber || '', eventId: r.eventId || '',
    eventName: r.eventName || '', studentName: r.studentName || '', gender: r.gender || '',
    birthday: r.birthday || '', school: r.school || '', grade: r.grade || '', belt: r.belt || '',
    items: (r.items || []).join('、'), group: r.group || '', weight: r.weight || 0,
    parentName: r.parentName || '', parentPhone: r.parentPhone || '',
    suggestedWeightClass: r.suggestedWeightClass || '', fee: Number(r.fee || 0),
    paid: r.paid ? 'TRUE' : 'FALSE', reviewStatus: r.reviewStatus || '待教練審核',
    uploadedAttachments: JSON.stringify(r.uploadedAttachments || []),
    missingAttachments: JSON.stringify(r.missingAttachments || []),
    consent: r.consent ? 'TRUE' : 'FALSE',
    createdAt: r.createdAt || new Date().toISOString(),
    updatedAt: r.updatedAt || new Date().toISOString()
  };
}
function rowToReg_(row) {
  return {
    id: row.id, registrationNumber: row.registrationNumber, eventId: row.eventId,
    eventName: row.eventName, studentName: row.studentName, gender: row.gender,
    birthday: row.birthday, school: row.school, grade: row.grade, belt: row.belt,
    items: row.items ? String(row.items).split('、').filter(Boolean) : [],
    group: row.group, weight: Number(row.weight || 0),
    parentName: row.parentName, parentPhone: row.parentPhone,
    suggestedWeightClass: row.suggestedWeightClass, fee: Number(row.fee || 0),
    paid: String(row.paid).toUpperCase() === 'TRUE',
    reviewStatus: row.reviewStatus || '待教練審核',
    uploadedAttachments: json_(row.uploadedAttachments, []) || [],
    missingAttachments: json_(row.missingAttachments, []) || [],
    consent: String(row.consent).toUpperCase() === 'TRUE',
    createdAt: row.createdAt, updatedAt: row.updatedAt
  };
}

/* ===================== Drive 附件 ===================== */
function getAttachFolder_() {
  const it = DriveApp.getFoldersByName(ATTACH_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(ATTACH_FOLDER_NAME);
}
/** 把前端傳來的 base64 附件存到 Drive，回傳含公開檢視連結的精簡物件 */
function saveAttachmentsToDrive_(files) {
  if (!files || !files.length) return [];
  const folder = getAttachFolder_();
  return files.map(function (f) {
    if (f.url) return { label: f.label, name: f.name, url: f.url, size: f.size || 0 };
    if (!f.dataUrl) return { label: f.label, name: f.name, url: '', size: f.size || 0 };
    const comma = f.dataUrl.indexOf(',');
    const meta = f.dataUrl.slice(0, comma);
    const b64 = f.dataUrl.slice(comma + 1);
    const ct = (meta.match(/data:(.*?);base64/) || [])[1] || f.type || 'application/octet-stream';
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), ct, f.name || 'attachment');
    const file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return { label: f.label, name: f.name, url: file.getUrl(), size: f.size || blob.getBytes().length };
  });
}

/* ===================== 前端 API（google.script.run 呼叫）===================== */

/** 一次撈回前端要的所有共用資料 */
function bootstrap() {
  const events = readAll_(SHEETS.EVENTS).map(rowToEvent_);
  const registrations = readAll_(SHEETS.REGS).map(rowToReg_)
    .sort(function (a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); });
  return {
    ok: true,
    events: events,
    registrations: registrations,
    paymentSettings: json_(getSetting_('paymentSettings', '{}'), {}) || {}
  };
}

/** 建立 / 更新活動（教練端）。回傳存好的活動清單 */
function saveEventRemote(ev) {
  if (!ev || !ev.id) return { ok: false, msg: '活動資料缺少 id' };
  if (!ev.createdAt) ev.createdAt = new Date().toISOString();
  upsertRow_(SHEETS.EVENTS, eventToRow_(ev));
  return { ok: true, events: readAll_(SHEETS.EVENTS).map(rowToEvent_) };
}

function deleteEventRemote(id) {
  deleteRow_(SHEETS.EVENTS, id);
  return { ok: true, events: readAll_(SHEETS.EVENTS).map(rowToEvent_) };
}

/** 家長送出報名：附件上 Drive，逐筆 upsert（多人同時送不互相覆蓋） */
function submitRegistrationRemote(row) {
  if (!row || !row.id) return { ok: false, msg: '報名資料缺少 id' };
  row.uploadedAttachments = saveAttachmentsToDrive_(row.uploadedAttachments || []);
  row.updatedAt = new Date().toISOString();
  if (!row.createdAt) row.createdAt = row.updatedAt;
  upsertRow_(SHEETS.REGS, regToRow_(row));
  return { ok: true, registration: row };
}

/** 教練改繳費 / 審核狀態等欄位 */
function patchRegistrationRemote(id, patch) {
  const sheetPatch = {};
  if (patch.paid !== undefined) sheetPatch.paid = patch.paid ? 'TRUE' : 'FALSE';
  if (patch.reviewStatus !== undefined) sheetPatch.reviewStatus = patch.reviewStatus;
  sheetPatch.updatedAt = new Date().toISOString();
  const ok = patchRow_(SHEETS.REGS, id, sheetPatch);
  return { ok: ok };
}

function deleteRegistrationRemote(id) {
  deleteRow_(SHEETS.REGS, id);
  return { ok: true };
}

function savePaymentSettingsRemote(settings) {
  setSetting_('paymentSettings', JSON.stringify(settings || {}));
  return { ok: true };
}

function coachLoginRemote(password) {
  const pw = getSetting_('coachPassword', '1234');
  return { ok: String(password) === String(pw) };
}

function changeCoachPasswordRemote(password) {
  if (!password || String(password).length < 4) return { ok: false, msg: '新密碼至少 4 碼' };
  setSetting_('coachPassword', String(password));
  return { ok: true };
}

/* ===================== 初始化（手動執行一次）===================== */
function setupSheets() {
  const ss = ss_();
  Object.keys(HEADERS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    // 只強制重寫第 1 列標題，不清空既有資料
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold')
      .setBackground('#12353f').setFontColor('#ffffff');
  });
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('工作表1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  // 預設值
  if (!getSetting_('coachPassword', '')) setSetting_('coachPassword', '1234');
  if (!getSetting_('paymentSettings', '')) setSetting_('paymentSettings', '{}');

  // 示範活動（若還沒有任何活動才建立）
  if (readAll_(SHEETS.EVENTS).length === 0) {
    const demo = demoEvent_();
    upsertRow_(SHEETS.EVENTS, eventToRow_(demo));
  }
  return '✅ 初始化完成：已建立 Events / Registrations / Settings 三張表，教練密碼預設 1234。';
}

/** 與前端 demoParsedData 對齊的示範活動 */
function demoEvent_() {
  return {
    id: uid_('event'),
    name: '2026 年新北議長盃全國跆拳道錦標賽',
    date: '115 年 7 月 16 日至 7 月 19 日',
    location: '新北市新莊體育館',
    address: '新北市新莊區中華路一段 75 號',
    registrationDeadline: '115 年 6 月 26 日 23:59',
    items: ['對練', '品勢', '競速踢擊'],
    weightClasses: [
      { group: '國小低年級男子組', gender: '男', label: '23 公斤以下', min: 0, max: 23 },
      { group: '國小低年級男子組', gender: '男', label: '23-25 公斤', min: 23.01, max: 25 },
      { group: '國小低年級女子組', gender: '女', label: '21 公斤以下', min: 0, max: 21 },
      { group: '國小高年級男子組', gender: '男', label: '34 公斤以下', min: 0, max: 34 },
      { group: '青少年男子電子護具組', gender: '男', label: '45 公斤以下', min: 0, max: 45 },
      { group: '青少年女子電子護具組', gender: '女', label: '42 公斤以下', min: 0, max: 42 }
    ],
    poomsaeGroups: ['國小低年級個人組', '國小中年級個人組', '國小高年級個人組', '青少年個人組', '混合雙人', '團體三人組'],
    speedKickGroups: ['國小低年級男子組', '國小低年級女子組', '國小中年級男子組', '國小中年級女子組'],
    feeRules: [
      { id: 'sparring_traditional', item: '對練', label: '國小組、青少年及青年色帶組傳統護具', amount: 600, match: ['國小', '色帶', '傳統'] },
      { id: 'sparring_electronic', item: '對練', label: '青少年及青年電子護具組', amount: 1000, match: ['電子護具'] },
      { id: 'poomsae_single', item: '品勢', label: '品勢個人組', amount: 500, match: ['個人'] },
      { id: 'poomsae_pair', item: '品勢', label: '混合雙人', amount: 800, match: ['混合雙人'] },
      { id: 'poomsae_team', item: '品勢', label: '團體三人組', amount: 1000, match: ['團體三人'] },
      { id: 'speed_kick', item: '競速踢擊', label: '競速踢擊', amount: 500, match: ['競速踢擊'] }
    ],
    attachments: ['選手切結書', '家長或監護人簽名', '段級證影本', '繳費證明', '指導教練確認單，選拔組需要'],
    createdAt: new Date().toISOString()
  };
}
