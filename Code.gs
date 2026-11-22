/** @OnlyCurrentDoc */

/**
 * מעקב התנסויות צוות — צד שרת (Google Apps Script)
 * הנתונים נשמרים בגיליון שאליו הסקריפט מחובר.
 */

const SHEET_PEOPLE = 'אנשים';
const SHEET_STAGES = 'שלבים';
const SHEET_EXP = 'התנסויות';
const TYPES = ['הדרכתית', 'פיקודית'];

/**
 * שלבי מעגל החניכה (לפי ה.ה קה"ד צוותי, סעיף 4(2)), זהים לשני הסוגים.
 * אחרי ההתקנה נערכים בלשונית "שלבים".
 */
const DEFAULT_STAGES = [
  ['הורדת התנסות',
   'קיבלת את ההתנסות "{התנסות}".\nעכשיו צריך למלא את טופס אישור התוכניות ולהעלות אותו לקמפוס תוך 24 שעות (או בזמן שנקבע להתנסות).\nאת הפורמט אפשר להוריד כאן למטה.'],
  ['אישור תוכניות',
   'הטופס הועלה ונמצא בבדיקה.\nעכשיו צריך לעבור על ההערות, לחדד ולדייק את התוכניות, ולוודא שהגרסה הסופית בקמפוס ותקינה עד שהמפק"ץ מאשר.'],
  ['ביצוע ההתנסות',
   'התוכניות אושרו.\nעכשיו צריך להתכונן ולבצע את ההתנסות לפי התוכנית שאושרה.'],
  ['משוב חם של הצוות',
   'מיד בסוף ההתנסות צריך לקיים משוב חם עם הצוות.\nלהקפיד שהמשוב יהיה איכותי, ולשים לב לנקודות שיעזרו בכתיבת התחקור.'],
  ['כתיבת דף תחקור עצמי',
   'עכשיו צריך לכתוב דף תחקור עצמי ולהגיש אותו בקמפוס עד 24 שעות מסיום ההתנסות.\nכדאי להעמיק בו, כי הוא הבסיס למשוב עם המפק"ץ.\nאת הפורמט אפשר להוריד כאן למטה.'],
  ['משוב מהמפק"ץ',
   'עכשיו צריך לתאם משוב עם המפק"ץ ולהגיע אליו עם דף התחקור.\nאחרי המשוב אפשר לסמן את ההתנסות כהושלמה.']
];

/** הרצה ידנית: מחליף את לשונית "שלבים" בשלבי מעגל החניכה. לא נוגע בהתנסויות. */
function resetStages() {
  ensureSetup_();
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STAGES);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 4).clearContent();
  const rows = defaultStageRows_();
  s.getRange(2, 1, rows.length, 4).setValues(rows);
  s.getRange(2, 4, rows.length, 1).setWrap(true);
}

function defaultStageRows_() {
  const rows = [];
  TYPES.forEach(function (type) {
    DEFAULT_STAGES.forEach(function (s, i) { rows.push([type, i + 1, s[0], s[1]]); });
  });
  return rows;
}

// ---------- API (האתר עצמו יושב ב-GitHub Pages וקורא לכאן) ----------

const ACTIONS = {
  addExperience: addExperience,
  moveStage: moveStage,
  setStage: setStage,
  deleteExperience: deleteExperience
};

/** GET → כל הנתונים */
function doGet() {
  return respond_(function () { return getData(); });
}

/** POST {action, args} → ביצוע פעולה, מחזיר את כל הנתונים המעודכנים */
function doPost(e) {
  return respond_(function () {
    const body = JSON.parse(e.postData.contents);
    const fn = ACTIONS[body.action];
    if (!fn) throw new Error('פעולה לא מוכרת');
    return fn.apply(null, body.args || []);
  });
}

function respond_(fn) {
  let out;
  try {
    out = fn();
  } catch (err) {
    out = { error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Setup ----------

/** הרצה ידנית פעם אחת (או אוטומטית בכניסה הראשונה): יוצר את הלשוניות עם ממלאי מקום. */
function setup() {
  ensureSetup_();
}

function ensureSetup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let people = ss.getSheetByName(SHEET_PEOPLE);
  if (!people) {
    people = ss.insertSheet(SHEET_PEOPLE);
    const rows = [['שם', 'צוות']];
    for (let i = 1; i <= 5; i++) rows.push(['חניך ' + i, '']);
    people.getRange(1, 1, rows.length, 2).setValues(rows);
    styleHeader_(people, 2);
  }

  let stages = ss.getSheetByName(SHEET_STAGES);
  if (!stages) {
    stages = ss.insertSheet(SHEET_STAGES);
    const rows = [['סוג', 'מספר', 'שם שלב', 'הודעה']].concat(defaultStageRows_());
    stages.getRange(1, 1, rows.length, 4).setValues(rows);
    stages.setColumnWidth(3, 180);
    stages.setColumnWidth(4, 480);
    stages.getRange('D:D').setWrap(true);
    styleHeader_(stages, 4);
  }

  let exp = ss.getSheetByName(SHEET_EXP);
  if (!exp) {
    exp = ss.insertSheet(SHEET_EXP);
    exp.getRange(1, 1, 1, 7).setValues([['id', 'שם', 'סוג', 'תיאור', 'שלב', 'נוצר', 'עודכן']]);
    exp.getRange('A:A').setNumberFormat('@');
    exp.getRange('D:D').setNumberFormat('@');
    styleHeader_(exp, 7);
  }

  // מחיקת הגיליון הריק שנוצר כברירת מחדל, אם קיים
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('גיליון1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
}

function styleHeader_(sheet, cols) {
  sheet.getRange(1, 1, 1, cols).setFontWeight('bold').setBackground('#eef2ff');
  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);
}

// ---------- Read ----------

function getData() {
  ensureSetup_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const people = rows_(ss.getSheetByName(SHEET_PEOPLE))
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (n) { return n; });

  const stages = {};
  TYPES.forEach(function (t) { stages[t] = []; });
  rows_(ss.getSheetByName(SHEET_STAGES)).forEach(function (r) {
    const type = String(r[0]).trim();
    if (!stages[type] || r[2] === '') return;
    stages[type].push({ num: Number(r[1]) || 0, name: String(r[2]), message: String(r[3] || '') });
  });
  TYPES.forEach(function (t) { stages[t].sort(function (a, b) { return a.num - b.num; }); });

  const experiences = rows_(ss.getSheetByName(SHEET_EXP))
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      return {
        id: String(r[0]),
        person: String(r[1]),
        type: String(r[2]),
        title: String(r[3]),
        stage: Number(r[4]) || 1,
        created: iso_(r[5]),
        updated: iso_(r[6])
      };
    });

  return { people: people, types: TYPES, stages: stages, experiences: experiences };
}

function rows_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
}

function iso_(v) {
  return v instanceof Date ? v.toISOString() : String(v || '');
}

// ---------- Write ----------

function addExperience(person, type, title) {
  person = String(person || '').trim();
  title = String(title || '').trim().slice(0, 200);
  if (TYPES.indexOf(type) === -1) throw new Error('סוג לא תקין');
  if (!title) throw new Error('חסר תיאור');

  return withLock_(function () {
    const data = getData();
    if (data.people.indexOf(person) === -1) throw new Error('שם לא קיים');
    const now = new Date();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EXP)
      .appendRow([Utilities.getUuid(), safe_(person), type, safe_(title), 1, now, now]);
  });
}

/** delta = +1 / -1. השלב המקסימלי הוא (מספר השלבים + 1) = "הושלמה". */
function moveStage(id, delta) {
  delta = delta > 0 ? 1 : -1;
  return withLock_(function () {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EXP);
    const row = findRow_(sheet, id);
    const type = String(sheet.getRange(row, 3).getValue());
    const max = getData().stages[type].length + 1;
    const cur = Number(sheet.getRange(row, 5).getValue()) || 1;
    const next = Math.min(max, Math.max(1, cur + delta));
    sheet.getRange(row, 5).setValue(next);
    sheet.getRange(row, 7).setValue(new Date());
  });
}

/** קפיצה ישירה לשלב (stage = מספר השלבים + 1 → הושלמה) */
function setStage(id, stage) {
  return withLock_(function () {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EXP);
    const row = findRow_(sheet, id);
    const type = String(sheet.getRange(row, 3).getValue());
    const max = getData().stages[type].length + 1;
    const next = Math.min(max, Math.max(1, Math.round(Number(stage)) || 1));
    sheet.getRange(row, 5).setValue(next);
    sheet.getRange(row, 7).setValue(new Date());
  });
}

function deleteExperience(id) {
  return withLock_(function () {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EXP);
    sheet.deleteRow(findRow_(sheet, id));
  });
}

function findRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last >= 2) {
    const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) return i + 2;
    }
  }
  throw new Error('ההתנסות לא נמצאה (אולי נמחקה)');
}

/** מבצע כתיבה תחת נעילה ומחזיר את כל הנתונים המעודכנים. */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    fn();
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return getData();
}

// ---------- ייבוא נתונים ----------

/**
 * מחליף את כל הגיליון בנתונים שהועברו: אנשים, שלבים (ברירת מחדל) והתנסויות.
 * people: [[שם, צוות], ...]
 * experiences: [[שם, סוג, תיאור, מספר שלב], ...]
 * נקרא מ-importInitialData() שבקובץ Import.gs (לא נשמר ב-GitHub).
 */
function replaceAllData_(people, experiences) {
  ensureSetup_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const p = ss.getSheetByName(SHEET_PEOPLE);
  if (p.getLastRow() > 1) p.getRange(2, 1, p.getLastRow() - 1, Math.max(2, p.getLastColumn())).clearContent();
  p.getRange(1, 1, 1, 2).setValues([['שם', 'צוות']]);
  p.getRange(2, 1, people.length, 2).setValues(people);

  const s = ss.getSheetByName(SHEET_STAGES);
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 4).clearContent();
  const stageRows = defaultStageRows_();
  s.getRange(2, 1, stageRows.length, 4).setValues(stageRows);

  const e = ss.getSheetByName(SHEET_EXP);
  if (e.getLastRow() > 1) e.getRange(2, 1, e.getLastRow() - 1, 7).clearContent();
  const now = new Date();
  const rows = experiences.map(function (x) {
    return [Utilities.getUuid(), x[0], x[1], safe_(x[2]), x[3] || 1, now, now];
  });
  if (rows.length) e.getRange(2, 1, rows.length, 7).setValues(rows);
}

/** מונע מטקסט שמתחיל ב-= להפוך לנוסחה. */
function safe_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
