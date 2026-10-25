/* ============================================================
   一般內科 PGY 學員月評核（護理長 / CR 總醫師 → 一般內科 PGY1、PGY2）
   Google Apps Script 後端：給 index.html（問卷）與 dashboard.html（儀表板）使用

   安裝步驟
   1. 新建一份 Google 試算表 → 擴充功能 → Apps Script，把本檔全部貼上、儲存
   2. 在編輯器上方選函式「setup」→ 執行（第一次會要求授權）
      會自動建立分頁：「回覆」「PGY名單」「設定」
   3. 到「設定」分頁填入：
        B1 評核開放日（預設 20：每月 20 日起評「本月」，之前評「上個月」）
        B2 強制評核月份（平常留空；要固定某月時填 2026-10）
        B3 保留欄位（不再使用；留空即可）
        B4 管理者密碼（問卷右上角「管理者登入」用來設定每月 PGY 名單）
   4. 部署 → 新增部署作業 → 類型「網頁應用程式」
        執行身分：我　／　誰可以存取：所有人
      把產生的網址（…/exec）貼到 index.html 與 dashboard.html 的 CONFIG.GAS_URL
   之後修改程式需「管理部署作業 → 編輯 → 版本：新版本」才會生效。
   ============================================================ */

const SHEET_RESP = "回覆";
const SHEET_ROSTER = "PGY名單";
const SHEET_SET = "設定";
const TZ = "Asia/Taipei";

// 欄名與 index.html 送出的欄位一致；新增欄位只要加在這裡，舊資料不受影響
const ITEMS = [
  "該醫師對病人與家屬能耐心解釋",
  "該醫師會盡其所能減輕病人的痛苦",
  "該醫師會主動參與並關心醫院的公共事務",
  "該醫師會出席並積極參與醫院內部的會議",
  "該醫師願意和院內同仁討論關於病人病情的知識",
  "該醫師會提出建設性的建議來改進醫院的服務",
  "對於業務上需配合的事情，該醫師會盡力完成",
  "該醫師與其他醫療團隊能維持良好的溝通及互動",
  "該醫師會盡可能提供同仁所需的協助",
  "我很樂意和該醫師共事",
];
const HEADERS = ["時間戳記", "批次編號", "評核月份", "受評學員", "學員職別", "評估者姓名", "評估者身分", "狀態",
  ...ITEMS, "表現優良之處", "建議加強之處"];
const ROSTER_HEADERS = ["月份", "姓名", "職別"];

/* ---------------- 初始化 ---------------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const resp = ss.getSheetByName(SHEET_RESP) || ss.insertSheet(SHEET_RESP);
  if (resp.getLastRow() === 0) resp.appendRow(HEADERS);
  resp.setFrozenRows(1);
  resp.getRange("C:C").setNumberFormat("@");          // 評核月份保持文字，避免被轉成日期

  const ro = ss.getSheetByName(SHEET_ROSTER) || ss.insertSheet(SHEET_ROSTER);
  if (ro.getLastRow() === 0) ro.appendRow(ROSTER_HEADERS);
  ro.setFrozenRows(1);
  ro.getRange("A:A").setNumberFormat("@");

  const st = ss.getSheetByName(SHEET_SET) || ss.insertSheet(SHEET_SET);
  if (st.getLastRow() === 0) {
    st.getRange("A1:A4").setValues([["評核開放日"], ["強制評核月份"], ["保留欄位（不使用）"], ["管理者密碼"]]);
    st.getRange("B1").setValue(20);
    st.getRange("B2").setNumberFormat("@");
  }
}

/* ---------------- 共用 ---------------- */
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function setting(row) {
  const st = SpreadsheetApp.getActive().getSheetByName(SHEET_SET);
  return st ? st.getRange(row, 2).getDisplayValue().trim() : "";
}
function normMonth(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "yyyy-MM");
  const m = String(v || "").trim().match(/^(\d{4})[-\/.](\d{1,2})/);
  return m ? m[1] + "-" + ("0" + m[2]).slice(-2) : "";
}
function normLevel(v) {
  const s = String(v || "").toUpperCase().replace(/\s/g, "");
  return s === "PGY2" ? "PGY2" : "PGY1";
}

function readRoster() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ROSTER);
  const out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(r => {
    const month = normMonth(r[0]), name = String(r[1] || "").trim();
    if (month && name) out.push({ month: month, name: name, level: normLevel(r[2]) });
  });
  return out.sort((a, b) => a.month < b.month ? -1 : a.month > b.month ? 1 : 0);
}
function rosterMap(list) {
  const m = {};
  list.forEach(r => (m[r.month] = m[r.month] || []).push({ name: r.name, level: r.level }));
  return m;
}

/* ---------------- GET：問卷設定 / 儀表板資料 ---------------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === "config") {
      return json({ ok: true, roster: rosterMap(readRoster()), openDay: +setting(1) || 20, forceMonth: normMonth(setting(2)) });
    }
    if (p.action === "data") {
      // 儀表板免密碼讀取；管理名單仍由 doPost 的 B4 管理者密碼保護。
      const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RESP);
      const vals = sh && sh.getLastRow() > 0 ? sh.getDataRange().getDisplayValues() : [HEADERS];
      return json({ ok: true, headers: vals[0], rows: vals.slice(1), roster: rosterMap(readRoster()) });
    }
    return json({ ok: false, error: "unknown action" });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/* ---------------- POST：送出評核 / 管理 PGY 名單 ---------------- */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: "bad json" }); }

  if (body.action === "adminLogin" || body.action === "adminSave") {
    const pw = setting(4);
    if (!pw) return json({ ok: false, error: "disabled" });
    if (String(body.key || "") !== pw) return json({ ok: false, error: "auth" });
    if (body.action === "adminSave") saveRoster(body.roster || []);
    return json({ ok: true, roster: readRoster() });
  }

  // 送出評核：一次可包含多位學員（每位學員一列）
  const rows = Array.isArray(body.rows) ? body.rows : [body];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_RESP);
    let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    // 若送來的欄位試算表還沒有，就補在最右邊
    rows.forEach(r => Object.keys(r).forEach(k => {
      if (headers.indexOf(k) < 0) { headers.push(k); sh.getRange(1, headers.length).setValue(k); }
    }));
    const ts = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss");
    const batch = Utilities.getUuid().slice(0, 8);
    const data = rows.map(r => headers.map(h =>
      h === "時間戳記" ? ts : h === "批次編號" ? batch : (r[h] == null ? "" : String(r[h]))));
    if (data.length) {
      const start = sh.getLastRow() + 1;
      sh.getRange(start, 1, data.length, headers.length).setNumberFormat("@").setValues(data);
    }
    return json({ ok: true, n: data.length });
  } finally {
    lock.releaseLock();
  }
}

function saveRoster(list) {
  const seen = {};
  const rows = [];
  list.forEach(r => {
    const month = normMonth(r.month), name = String(r.name || "").trim();
    if (!month || !name || seen[month + "|" + name]) return;
    seen[month + "|" + name] = 1;
    rows.push([month, name, normLevel(r.level)]);
  });
  rows.sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0);
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_ROSTER);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setNumberFormat("@").setValues(rows);
}
