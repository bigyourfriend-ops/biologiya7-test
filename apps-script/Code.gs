/**
 * Биология 7 · нәтижелерді автоматты жинау (Google Таблица)
 *
 * Орнату:
 * 1. Жаңа Google Таблица ашыңыз → Кеңейтімдер (Extensions) → Apps Script.
 * 2. Осы кодты толығымен қойыңыз. Төмендегі SECRET-ті өз құпия сөзіңізге ауыстырыңыз.
 * 3. Deploy → New deployment → түрі: Web app.
 *    Execute as: Me · Who has access: Anyone → Deploy → рұқсат беріңіз.
 * 4. Шыққан Web app URL-ды (…/exec) және құпия сөзді мұғалім кабинетіндегі
 *    «Автоматты жинау» бөліміне қойыңыз да, «Оқушыларға жариялау» басыңыз.
 */
const SECRET = "ӨЗ-ҚҰПИЯ-СӨЗІҢІЗ";   // ← міндетті түрде өзгертіңіз
const SHEET = "Нәтижелер";
const STAGE = {pre: "Кіріс тест", post: "Қорытынды тест", late: "Кейінгі тексеру", prac: "Жаттығу"};

function doPost(e) {
  const code = String((e && e.postData && e.postData.contents) || "").trim();
  if (code.length > 20000 || !/^BIO7-[A-Za-z0-9_-]+\.[a-z0-9]+$/.test(code)) return out({ok: false});
  const o = decode(code);
  if (!o || o.s === "prac") return out({ok: false});
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet();
    const last = sh.getLastRow();
    if (last > 1) {
      const codes = sh.getRange(2, 8, last - 1, 1).getValues();
      if (codes.some(r => r[0] === code)) return out({ok: true, dup: true});
    }
    let ok = 0, n = 0;
    Object.keys(o.r || {}).forEach(k => { ok += o.r[k][0]; n += o.r[k][1]; });
    sh.appendRow([new Date(), o.n, o.c, STAGE[o.s] || o.s, o.m === "s" ? "қарапайым" : "қателермен жұмыс",
                  n ? Math.round(100 * ok / n) : "", o.x || 0, code]);
  } finally { lock.releaseLock(); }
  return out({ok: true});
}

function doGet(e) {
  if (!e || String(e.parameter.key || "") !== SECRET) return out({ok: false, error: "key"});
  const sh = sheet(), last = sh.getLastRow();
  const codes = last > 1 ? sh.getRange(2, 8, last - 1, 1).getValues().map(r => r[0]).filter(String) : [];
  return out({ok: true, codes: codes});
}

function decode(code) {
  try {
    let b = code.slice(5, code.lastIndexOf("."));
    while (b.length % 4) b += "=";
    return JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(b)).getDataAsString("UTF-8"));
  } catch (err) { return null; }
}

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(["Уақыты", "Аты-жөні", "Сынып", "Кезең", "Тест түрі", "Балл, %", "Қайта тексеру", "Код"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
