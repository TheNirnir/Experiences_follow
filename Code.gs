/**
 * מעקב התנסויות צוות — צד שרת (Google Apps Script)
 * הנתונים נשמרים בגיליון שאליו הסקריפט מחובר.
 */

const SHEET_PEOPLE = 'אנשים';
const SHEET_STAGES = 'שלבים';
const SHEET_EXP = 'התנסויות';
const TYPES = ['הדרכתית', 'פיקודית'];
const SHEET_TASKS = 'מטלות';
const SHEET_SUBS = 'הגשות';
const PLACES = ['קמפוס', 'דרייב', 'אקסל', 'כתב יד'];
const DRIVE = 'דרייב';
/** אקסל: כל מגיש ממלא שדות, והם נשמרים בשורה שלו בגיליון של המטלה */
const EXCEL = 'אקסל';
const NAME_COL = 'שם';
const MAX_FIELDS = 20;
/** מטלות התחלתיות שנכנסות כשנוצרת לשונית "מטלות": [שם, תג"ב, מקומות, אחר] */
const DEFAULT_TASKS = [
  ['ל״ע 2 ומבנה צה״ל', '2026-11-23', 'קמפוס,דרייב', ''],
  ['למה קצונה?', '2026-11-26', 'כתב יד', ''],
  ['סכמ״ש', '2026-11-27', '', 'במחברת (עד כניסת שבת)'],
  ['ספרא וסייפא', '2026-12-14', '', '']
];
const SEED_KEY = 'מטלות התחלתיות נוספו';
const SHEET_SETTINGS = 'הגדרות';
/**
 * תיקיית ההגשות בדרייב (תת-תיקייה לכל מטלה).
 * הסקריפט יוצר אותה בעצמו, כי יש לו הרשאה רק לקבצים שהוא יצר (drive.file).
 * המזהה נשמר בלשונית "הגדרות".
 */
const ROOT_FOLDER_NAME = 'הגשות מטלות';
const ROOT_KEY = 'תיקיית הגשות';
const FOLDER_MIME = 'application/vnd.google-apps.folder';
const SHEET_MIME = 'application/vnd.google-apps.spreadsheet';
const TASK_COLS = 9;

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
  deleteExperience: deleteExperience,
  addAssignment: addAssignment,
  updateAssignment: updateAssignment,
  deleteAssignment: deleteAssignment,
  setSubmitted: setSubmitted,
  uploadSubmission: uploadSubmission,
  submitExcel: submitExcel
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
  const root = rootFolder_();
  Logger.log('תיקיית ההגשות: ' + root.url);
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

  let tasks = ss.getSheetByName(SHEET_TASKS);
  if (!tasks) {
    tasks = ss.insertSheet(SHEET_TASKS);
    tasks.getRange(1, 1, 1, 7).setValues([['id', 'שם', 'תג"ב', 'מקומות', 'אחר', 'תיקייה', 'נוצר']]);
    tasks.getRange('A:F').setNumberFormat('@');
    styleHeader_(tasks, 7);
  }
  // עמודות שנוספו בהמשך: שדות למילוי, ומזהה גיליון האקסל של המטלה
  if (tasks.getRange(1, 8).getValue() === '') {
    tasks.getRange(1, 8, 1, 2).setValues([['שדות', 'אקסל']]);
    tasks.getRange('H:I').setNumberFormat('@');
    styleHeader_(tasks, TASK_COLS);
  }

  let subs = ss.getSheetByName(SHEET_SUBS);
  if (!subs) {
    subs = ss.insertSheet(SHEET_SUBS);
    subs.getRange(1, 1, 1, 6).setValues([['מטלה', 'שם', 'אופן', 'fileId', 'קישור', 'זמן']]);
    subs.getRange('A:E').setNumberFormat('@');
    styleHeader_(subs, 6);
  }

  let settings = ss.getSheetByName(SHEET_SETTINGS);
  if (!settings) {
    settings = ss.insertSheet(SHEET_SETTINGS);
    settings.getRange(1, 1, 1, 3).setValues([['מפתח', 'ערך', 'קישור']]);
    settings.getRange('A:B').setNumberFormat('@');
    styleHeader_(settings, 3);
  }

  // מטלות התחלתיות: נכנסות פעם אחת בלבד (גם אם הלשונית כבר הייתה קיימת)
  if (!getSetting_(SEED_KEY).value) {
    const existing = rows_(tasks).map(function (r) { return String(r[1]); });
    const now = new Date();
    const seed = DEFAULT_TASKS
      .filter(function (t) { return existing.indexOf(t[0]) === -1; })
      .map(function (t) { return [Utilities.getUuid(), t[0], t[1], t[2], t[3], '', now]; });
    if (seed.length) tasks.getRange(tasks.getLastRow() + 1, 1, seed.length, 7).setValues(seed);
    setSetting_(SEED_KEY, 'כן', '');
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

  const people = peopleList_();

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

  const assignments = rows_(ss.getSheetByName(SHEET_TASKS))
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      return {
        id: String(r[0]),
        title: String(r[1]),
        due: dateStr_(r[2]),
        places: String(r[3] || '').split(',').map(function (p) { return p.trim(); })
          .filter(function (p) { return PLACES.indexOf(p) !== -1; }),
        other: String(r[4] || ''),
        created: iso_(r[6]),
        fields: fieldsOf_(r[7]),
        sheetUrl: r[8] ? sheetUrl_(r[8]) : ''
      };
    });

  const submissions = rows_(ss.getSheetByName(SHEET_SUBS))
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      const s = { aid: String(r[0]), person: String(r[1]), mode: String(r[2]), url: String(r[4] || ''), time: iso_(r[5]) };
      // באקסל: עמודת fileId שומרת עותק של מה שמולא, כדי שאפשר יהיה לערוך מהאתר
      if (s.mode === 'excel') { try { s.values = JSON.parse(String(r[3] || '{}')); } catch (e) { s.values = {}; } }
      return s;
    });

  return {
    people: people, types: TYPES, stages: stages, experiences: experiences,
    places: PLACES, assignments: assignments, submissions: submissions,
    driveFolder: getSetting_(ROOT_KEY).url
  };
}

function peopleList_() {
  return rows_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEOPLE))
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (n) { return n; });
}

function fieldsOf_(v) {
  return String(v || '').split('\n').map(function (f) { return f.trim(); }).filter(function (f) { return f; });
}

function sheetUrl_(id) {
  return 'https://docs.google.com/spreadsheets/d/' + id + '/edit';
}

function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
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

// ---------- מטלות ----------

function cleanTask_(title, due, places, other, fields) {
  title = String(title || '').trim().slice(0, 200);
  due = String(due || '').trim();
  other = String(other || '').trim().slice(0, 200);
  places = (Array.isArray(places) ? places : []).filter(function (p, i, a) {
    return PLACES.indexOf(p) !== -1 && a.indexOf(p) === i;
  });
  if (!title) throw new Error('חסר שם מטלה');
  if (due && !/^\d{4}-\d{2}-\d{2}$/.test(due)) throw new Error('תאריך לא תקין');
  if (!places.length && !other) throw new Error('צריך לבחור לפחות מקום הגשה אחד');
  fields = (Array.isArray(fields) ? fields : []).map(function (f) {
    return String(f || '').replace(/\s+/g, ' ').trim().slice(0, 100);
  }).filter(function (f, i, a) { return f && a.indexOf(f) === i; });
  if (places.indexOf(EXCEL) === -1) fields = [];
  else {
    if (!fields.length) throw new Error('צריך לפחות שדה אחד למילוי באקסל');
    if (fields.length > MAX_FIELDS) throw new Error('אפשר עד ' + MAX_FIELDS + ' שדות');
    if (fields.indexOf(NAME_COL) !== -1) throw new Error('העמודה "' + NAME_COL + '" נוספת לבד, אין צורך בשדה כזה');
  }
  return { title: title, due: due, places: places, other: other, fields: fields };
}

function addAssignment(title, due, places, other, fields) {
  const t = cleanTask_(title, due, places, other, fields);
  return withLock_(function () {
    const id = Utilities.getUuid();
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS)
      .appendRow([id, safe_(t.title), t.due, t.places.join(','), safe_(t.other), '', new Date(), t.fields.join('\n'), '']);
    if (t.fields.length) taskSheet_(id, t.title, t.fields);
  });
}

function updateAssignment(id, title, due, places, other, fields) {
  const t = cleanTask_(title, due, places, other, fields);
  return withLock_(function () {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
    const row = findRowIn_(sheet, function (r) { return String(r[0]) === String(id); }, 'המטלה לא נמצאה (אולי נמחקה)');
    const old = sheet.getRange(row, 1, 1, TASK_COLS).getValues()[0];
    sheet.getRange(row, 2, 1, 4).setValues([[safe_(t.title), t.due, t.places.join(','), safe_(t.other)]]);
    // שדות שהוסרו לא נמחקים מהגיליון, רק לא מופיעים יותר בטופס
    if (t.fields.length) sheet.getRange(row, 8).setValue(t.fields.join('\n'));
    // שינוי שם → גם שם התיקייה והגיליון בדרייב
    if (String(old[1]) !== t.title) {
      [old[5], old[8]].forEach(function (fid) {
        if (fid) { try { Drive.Files.update({ name: t.title }, String(fid)); } catch (e) {} }
      });
    }
    if (t.fields.length) syncHeader_(taskSheet_(String(id), t.title, t.fields), t.fields);
  });
}

/** מוחק את המטלה ואת רשומות ההגשה שלה. הקבצים בדרייב נשארים. */
function deleteAssignment(id) {
  return withLock_(function () {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TASKS);
    sheet.deleteRow(findRowIn_(sheet, function (r) { return String(r[0]) === String(id); }, 'המטלה לא נמצאה (אולי נמחקה)'));
    const subs = ss.getSheetByName(SHEET_SUBS);
    const rows = rows_(subs);
    for (let i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][0]) === String(id)) subs.deleteRow(i + 2);
    }
  });
}

/** סימון ידני שהוגש (on=true) או ביטול סימון (on=false). גם במטלות דרייב (הגשה בלי העלאה). */
function setSubmitted(aid, person, on) {
  person = String(person || '').trim();
  return withLock_(function () {
    const data = getData();
    if (data.people.indexOf(person) === -1) throw new Error('שם לא קיים');
    const task = data.assignments.filter(function (a) { return a.id === String(aid); })[0];
    if (!task) throw new Error('המטלה לא נמצאה (אולי נמחקה)');
    const subs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SUBS);
    const row = findSubRow_(subs, aid, person);
    if (on && !row) subs.appendRow([String(aid), safe_(person), 'manual', '', '', new Date()]);
    if (!on && row) {
      // הגשה שהועלתה לדרייב לא מבטלים מכאן, כדי לא לאבד את הקישור לקובץ
      if (String(subs.getRange(row, 3).getValue()) === 'upload') throw new Error('כבר הועלה קובץ להגשה הזאת');
      subs.deleteRow(row);
    }
  });
}

/** העלאת PDF לתיקיית המטלה בדרייב. העלאה חוזרת מחליפה את הקובץ הקודם (הישן עובר לפח). */
function uploadSubmission(aid, person, base64) {
  person = String(person || '').trim();
  const bytes = Utilities.base64Decode(String(base64 || ''));
  if (!bytes.length) throw new Error('הקובץ ריק');
  // בדיקה שזה באמת PDF (מתחיל ב-%PDF)
  if (bytes[0] !== 0x25 || bytes[1] !== 0x50 || bytes[2] !== 0x44 || bytes[3] !== 0x46) throw new Error('אפשר להעלות רק קובץ PDF');

  // 1. בדיקות + תיקיית המטלה (תחת נעילה, כדי לא ליצור שתי תיקיות באותו שם)
  let task, folderId;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const data = getData();
    if (data.people.indexOf(person) === -1) throw new Error('שם לא קיים');
    task = data.assignments.filter(function (a) { return a.id === String(aid); })[0];
    if (!task) throw new Error('המטלה לא נמצאה (אולי נמחקה)');
    if (task.places.indexOf(DRIVE) === -1) throw new Error('במטלה הזאת לא מגישים בדרייב');
    folderId = taskFolder_(task);
  } finally {
    lock.releaseLock();
  }

  // 2. יצירת הקובץ (מחוץ לנעילה, כי זה יכול לקחת זמן)
  const name = (person + ' - ' + task.title).replace(/[\\/]/g, '-') + '.pdf';
  const file = Drive.Files.create(
    { name: name, parents: [folderId], mimeType: 'application/pdf' },
    Utilities.newBlob(bytes, 'application/pdf', name),
    { fields: 'id,webViewLink' });

  // 3. רישום ההגשה
  return withLock_(function () {
    const subs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SUBS);
    const row = findSubRow_(subs, aid, person);
    const values = [String(aid), safe_(person), 'upload', file.id, file.webViewLink, new Date()];
    if (row) {
      const oldId = String(subs.getRange(row, 4).getValue());
      if (oldId && oldId !== file.id) {
        try { Drive.Files.update({ trashed: true }, oldId); } catch (e) {}
      }
      subs.getRange(row, 1, 1, 6).setValues([values]);
    } else {
      subs.appendRow(values);
    }
  });
}

/** הגשה באקסל: כותב את מה שמולא לשורה של המגיש בגיליון המטלה (שורה חדשה אם אין לו). */
function submitExcel(aid, person, values) {
  person = String(person || '').trim();
  values = values && typeof values === 'object' ? values : {};
  return withLock_(function () {
    const data = getData();
    if (data.people.indexOf(person) === -1) throw new Error('שם לא קיים');
    const task = data.assignments.filter(function (a) { return a.id === String(aid); })[0];
    if (!task) throw new Error('המטלה לא נמצאה (אולי נמחקה)');
    if (task.places.indexOf(EXCEL) === -1 || !task.fields.length) throw new Error('במטלה הזאת לא ממלאים באקסל');

    const clean = {};
    task.fields.forEach(function (f) { clean[f] = String(values[f] == null ? '' : values[f]).trim().slice(0, 5000); });

    const sid = taskSheet_(task.id, task.title, task.fields);
    syncHeader_(sid, task.fields);
    const all = Sheets.Spreadsheets.Values.get(sid, 'A:ZZ').values || [];
    const header = all[0] || [];
    let r = 1;
    while (r < all.length && String((all[r] || [])[0] || '').trim() !== person) r++;
    const existing = all[r] || [];
    // עמודות שאינן בטופס (למשל שדה שהוסר) נשארות כמו שהן
    const line = header.map(function (h, i) {
      if (i === 0) return person;
      return clean.hasOwnProperty(h) ? clean[h] : (existing[i] == null ? '' : existing[i]);
    });
    Sheets.Spreadsheets.Values.update({ values: [line] }, sid, 'A' + (r + 1), { valueInputOption: 'RAW' });

    const subs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SUBS);
    const row = findSubRow_(subs, aid, person);
    const rec = [String(aid), safe_(person), 'excel', JSON.stringify(clean), sheetUrl_(sid), new Date()];
    if (!row) subs.appendRow(rec);
    // הגשה שהועלתה לדרייב נשארת, כדי לא לאבד את הקישור לקובץ
    else if (String(subs.getRange(row, 3).getValue()) !== 'upload') subs.getRange(row, 1, 1, 6).setValues([rec]);
  });
}

/**
 * מחזיר את מזהה גיליון האקסל של המטלה, ויוצר אותו אם צריך (בתיקיית המטלה בדרייב),
 * עם עמודת שם, עמודה לכל שדה ושורה לכל אחד מהאנשים. נקרא תחת נעילה.
 */
function taskSheet_(id, title, fields) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const row = findRowIn_(sheet, function (r) { return String(r[0]) === String(id); }, 'המטלה לא נמצאה');
  const cached = String(sheet.getRange(row, 9).getValue() || '');
  if (cached && driveAlive_(cached)) return cached;

  const folderId = taskFolder_({ id: String(id), title: title });
  const f = Drive.Files.create({ name: title, mimeType: SHEET_MIME, parents: [folderId] }, null, { fields: 'id' });
  const tab = Sheets.Spreadsheets.get(f.id, { fields: 'sheets.properties.sheetId' }).sheets[0].properties.sheetId;
  Sheets.Spreadsheets.batchUpdate({ requests: [
    { updateSheetProperties: {
      properties: { sheetId: tab, rightToLeft: true, gridProperties: { frozenRowCount: 1 } },
      fields: 'rightToLeft,gridProperties.frozenRowCount' } },
    { repeatCell: {
      range: { sheetId: tab, startRowIndex: 0, endRowIndex: 1 },
      cell: { userEnteredFormat: { textFormat: { bold: true }, backgroundColor: { red: 0.933, green: 0.949, blue: 1 } } },
      fields: 'userEnteredFormat(textFormat,backgroundColor)' } }
  ] }, f.id);
  const values = [[NAME_COL].concat(fields)].concat(peopleList_().map(function (p) {
    return [p].concat(fields.map(function () { return ''; }));
  }));
  Sheets.Spreadsheets.Values.update({ values: values }, f.id, 'A1', { valueInputOption: 'RAW' });
  sheet.getRange(row, 9).setValue(f.id);
  return f.id;
}

/** מוסיף לשורת הכותרת בגיליון שדות חדשים שעוד אין להם עמודה. */
function syncHeader_(sid, fields) {
  const header = (Sheets.Spreadsheets.Values.get(sid, '1:1').values || [[]])[0] || [];
  const next = header.length ? header.slice() : [NAME_COL];
  fields.forEach(function (f) { if (next.indexOf(f) === -1) next.push(f); });
  if (next.length !== header.length) Sheets.Spreadsheets.Values.update({ values: [next] }, sid, 'A1', { valueInputOption: 'RAW' });
}

/** מחזיר את מזהה תת-התיקייה של המטלה, ויוצר אותה אם צריך. נקרא תחת נעילה. */
function taskFolder_(task) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TASKS);
  const row = findRowIn_(sheet, function (r) { return String(r[0]) === task.id; }, 'המטלה לא נמצאה');
  const cached = String(sheet.getRange(row, 6).getValue() || '');
  if (cached && driveAlive_(cached)) return cached;
  const rootId = rootFolder_().id;
  const found = Drive.Files.list({
    q: "name = '" + driveQ_(task.title) + "' and '" + rootId + "' in parents and mimeType = '" + FOLDER_MIME + "' and trashed = false",
    fields: 'files(id)', pageSize: 1
  }).files || [];
  const id = found.length ? found[0].id : createFolder_(task.title, rootId).id;
  sheet.getRange(row, 6).setValue(id);
  return id;
}

/** תיקיית ההגשות הראשית: נוצרת בפעם הראשונה ונשמרת בלשונית "הגדרות". */
function rootFolder_() {
  const cur = getSetting_(ROOT_KEY);
  if (cur.value && driveAlive_(cur.value)) return { id: cur.value, url: cur.url };
  const f = createFolder_(ROOT_FOLDER_NAME, null);
  setSetting_(ROOT_KEY, f.id, f.webViewLink);
  return { id: f.id, url: f.webViewLink };
}

function createFolder_(name, parentId) {
  const res = { name: name, mimeType: FOLDER_MIME };
  if (parentId) res.parents = [parentId];
  return Drive.Files.create(res, null, { fields: 'id,webViewLink' });
}

/** האם הקובץ/התיקייה קיימים, נגישים ולא בפח */
function driveAlive_(id) {
  try {
    return !Drive.Files.get(id, { fields: 'trashed' }).trashed;
  } catch (e) {
    return false;
  }
}

/** בריחה של תווים בשאילתת חיפוש של Drive */
function driveQ_(s) {
  return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function getSetting_(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sheet) return { value: '', url: '' };
  const r = rows_(sheet).filter(function (r) { return String(r[0]) === key; })[0];
  return r ? { value: String(r[1] || ''), url: String(r[2] || '') } : { value: '', url: '' };
}

function setSetting_(key, value, url) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  const rows = rows_(sheet);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === key) { sheet.getRange(i + 2, 2, 1, 2).setValues([[value, url]]); return; }
  }
  sheet.appendRow([key, value, url]);
}

function findSubRow_(subs, aid, person) {
  const rows = rows_(subs);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(aid) && String(rows[i][1]).replace(/^'/, '') === person) return i + 2;
  }
  return 0;
}

function findRowIn_(sheet, pred, msg) {
  const rows = rows_(sheet);
  for (let i = 0; i < rows.length; i++) if (pred(rows[i])) return i + 2;
  throw new Error(msg || 'השורה לא נמצאה');
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
