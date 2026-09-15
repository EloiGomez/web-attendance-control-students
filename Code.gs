/**
 * Attendance Control — backend (Google Apps Script)
 *
 * Sheets used by this script:
 *  - Config   : course settings (school days, weight per half-day, alert threshold)
 *  - Students : student roster
 *  - Teachers : emails authorized to access the web app (allow-list)
 *  - Records  : one row per student and day with any mark (absence or late arrival)
 *
 * "Justified morning"/"Justified afternoon" are shared between an absence and
 * a late arrival in that same half-day: since "Morning" (absence) and "Late
 * morning" (late arrival) can never both be true at once (checking one
 * disables the other in the UI), one justified flag per half-day is enough
 * to cover whichever of the two actually happened — no need for separate
 * "justified late arrival" columns.
 *
 * Run setup() ONCE from the Apps Script editor to create these sheets with
 * their headers and default values.
 */

const SHEET_CONFIG = 'Config';
const SHEET_STUDENTS = 'Students';
const SHEET_TEACHERS = 'Teachers';
const SHEET_RECORDS = 'Records';
const SHEET_PROMOTION = 'Promotion';

/**
 * Asking Google for the active Spreadsheet has a real network cost every time
 * it's done. This function requests it once per execution and reuses it,
 * instead of every function fetching it again on its own.
 */
let _ss = null;
function spreadsheet_() {
  if (!_ss) _ss = SpreadsheetApp.getActiveSpreadsheet();
  return _ss;
}

/**
 * A small counter, stored permanently (PropertiesService), incremented every
 * time a change is saved to "Records". getSummary() includes it in its cache
 * key, so bumping it automatically invalidates any cached Summary result
 * without having to delete it.
 */
function recordsVersion_() {
  return PropertiesService.getScriptProperties().getProperty('RECORDS_VERSION') || '0';
}
function incrementRecordsVersion_() {
  const props = PropertiesService.getScriptProperties();
  const current = Number(props.getProperty('RECORDS_VERSION') || '0');
  props.setProperty('RECORDS_VERSION', String(current + 1));
}

function doGet() {
  if (!isAuthorizedUser_()) {
    const detectedEmail = currentEmail_() || '(no email detected — you may need to share the Sheet with this user)';
    return HtmlService.createHtmlOutput(
      '<p style="font-family:sans-serif;padding:24px;">' +
      '🔒 Access denied. This application is for school staff only.<br>' +
      'If you think you should have access, ask someone to add your email to the "Teachers" tab.<br><br>' +
      '<b>Email detected by the system:</b> ' + detectedEmail +
      '</p>'
    ).setTitle('Attendance Control');
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Attendance Control')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function currentEmail_() {
  try {
    return (Session.getActiveUser().getEmail() || '').toLowerCase();
  } catch (e) {
    return '';
  }
}

/**
 * If the "Teachers" tab has no emails at all, everyone is let in (so you
 * don't lock yourself out before configuring it). As soon as there is at
 * least one email in the list, only those can get in.
 */
function isAuthorizedUser_() {
  const sh = spreadsheet_().getSheetByName(SHEET_TEACHERS);
  if (!sh) return true;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return true;

  const emails = sh.getRange(2, 1, lastRow - 1, 1).getValues()
    .map((r) => String(r[0] || '').toLowerCase().trim())
    .filter(String);

  if (emails.length === 0) return true;

  return emails.indexOf(currentEmail_()) !== -1;
}

/**
 * Creates (if missing) the Config, Students, Teachers, Records and Promotion
 * sheets. Safe to run again: it never deletes data that's already there.
 */
function setup() {
  const ss = spreadsheet_();

  let cfg = ss.getSheetByName(SHEET_CONFIG);
  if (!cfg) {
    cfg = ss.insertSheet(SHEET_CONFIG);
    cfg.getRange(1, 1, 5, 3).setValues([
      ['Key', 'Value', 'Description'],
      ['TOTAL_SCHOOL_DAYS', 177, 'Total school days in the year (used to calculate the %)'],
      ['WEIGHT_MORNING', 2 / 3, 'How much of a full day a morning absence counts as'],
      ['WEIGHT_AFTERNOON', 1 / 3, 'How much of a full day an afternoon absence counts as (morning+afternoon = full day)'],
      ['ALERT_THRESHOLD_PCT', 10, 'The % at which a student is flagged as an alert in the Summary'],
    ]);
    cfg.setFrozenRows(1);
    cfg.autoResizeColumns(1, 3);
  }

  let students = ss.getSheetByName(SHEET_STUDENTS);
  if (!students) {
    students = ss.insertSheet(SHEET_STUDENTS);
    students.getRange(1, 1, 4, 3).setValues([
      ['Class', 'Student', 'Do not promote (repeating)'],
      ['K3A', 'Example Student 1', false],
      ['K3A', 'Example Student 2', false],
      ['K4B', 'Example Student 3', false],
    ]);
    students.setFrozenRows(1);
    students.autoResizeColumns(1, 3);
  } else if (students.getRange(1, 3).getValue() === '') {
    // Migrating an older sheet that doesn't have the "repeating" column yet.
    students.getRange(1, 3).setValue('Do not promote (repeating)');
  }

  let promotion = ss.getSheetByName(SHEET_PROMOTION);
  if (!promotion) {
    promotion = ss.insertSheet(SHEET_PROMOTION);
    promotion.getRange(1, 1, 10, 2).setValues([
      ['Current class', 'New class (blank = leaves the school, e.g. Grade 6 graduating)'],
      ['K3A', 'K4A'], ['K3B', 'K4B'],
      ['K4A', 'K5A'], ['K4B', 'K5B'],
      ['K5A', 'G1A'], ['K5B', 'G1B'],
      ['...', '...'],
      ['G6A', ''], ['G6B', ''],
    ]);
    promotion.setFrozenRows(1);
    promotion.autoResizeColumns(1, 2);
  }

  let teachers = ss.getSheetByName(SHEET_TEACHERS);
  if (!teachers) {
    teachers = ss.insertSheet(SHEET_TEACHERS);
    teachers.getRange(1, 1, 1, 2).setValues([['Email', 'Name (optional)']]);
    const myEmail = currentEmail_();
    if (myEmail) {
      teachers.getRange(2, 1, 1, 2).setValues([[myEmail, 'Added automatically by setup()']]);
    }
    teachers.setFrozenRows(1);
    teachers.autoResizeColumns(1, 2);
  }

  let records = ss.getSheetByName(SHEET_RECORDS);
  if (!records) {
    records = ss.insertSheet(SHEET_RECORDS);
    records.getRange(1, 1, 1, 11).setValues([[
      'Date', 'Class', 'Student', 'Morning', 'Afternoon',
      'Justified Morning', 'Justified Afternoon',
      'Late Morning', 'Late Afternoon',
      'Updated By', 'Last Updated',
    ]]);
    records.setFrozenRows(1);
    records.autoResizeColumns(1, 11);
  }

  // Protects the admin tabs: even if a teacher is an Editor on the whole
  // Sheet (needed for the web app to work), they won't be able to edit or
  // unprotect these — only the Sheet's owner can (protected ranges/sheets
  // always stay editable by the file owner, regardless of the editors list).
  // "Students" is included here because the web app only ever reads it,
  // never writes to it — "Records" is deliberately left unprotected, since
  // teachers need to write there every day through the app.
  protectSheet_(cfg);
  protectSheet_(teachers);
  protectSheet_(promotion);
  protectSheet_(students);

  SpreadsheetApp.getUi().alert(
    'Sheets created successfully.\n\n' +
    'The Config, Teachers, Promotion and Students tabs are now protected: ' +
    'only the Sheet\'s owner can edit them, even if other teachers are ' +
    'Editors on the Sheet in general.\n\n' +
    'Remember: add the emails of authorized teachers to the "Teachers" tab ' +
    'before sharing the link with everyone.\n\n' +
    'You can now go to Deploy > New deployment.'
  );
}

/**
 * Protects an entire sheet so that only whoever runs this function (normally
 * the Sheet's owner) can edit it. If it was already protected by an earlier
 * setup() run, reuses the existing protection instead of duplicating it.
 */
function protectSheet_(sheet) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  const protection = protections.length ? protections[0] : sheet.protect();
  protection.setDescription('Admin only');
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
  protection.removeEditors(protection.getEditors());
}

/**
 * FOR PERFORMANCE TESTING ONLY: replaces the content of the "Students" tab
 * with fake data (a few hundred students spread across many classes), to
 * check that the web app is just as fast with a whole school as with 3
 * example students.
 *
 * Run it manually from the editor when you want to test, and run setup()
 * again (or clear the "Students" tab by hand) to go back to real data.
 */
function generateTestStudents() {
  const grades = ['K3', 'K4', 'K5', 'G1', 'G2', 'G3', 'G4', 'G5', 'G6'];
  const groups = ['A', 'B'];
  const studentsPerClass = 25;

  const rows = [];
  grades.forEach((grade) => {
    groups.forEach((group) => {
      const className = grade + group;
      for (let i = 1; i <= studentsPerClass; i++) {
        rows.push([className, `Test Student ${className}-${i}`]);
      }
    });
  });

  const sh = spreadsheet_().getSheetByName(SHEET_STUDENTS);
  const existingRows = Math.max(sh.getLastRow() - 1, 0);
  if (existingRows > 0) sh.getRange(2, 1, existingRows, 2).clearContent();
  sh.getRange(2, 1, rows.length, 2).setValues(rows);

  SpreadsheetApp.getUi().alert(
    `Generated ${rows.length} test students across ${grades.length * groups.length} classes.`
  );
}

/**
 * FOR TESTING ONLY: fills the "Records" tab with random absences and late
 * arrivals for every student, spread across the last `daysBack` school days
 * (weekends excluded). Useful for seeing the Summary, percentages, alerts
 * and search box with realistic-looking data, without taking attendance
 * day by day by hand.
 *
 * Replaces the ENTIRE current content of "Records". Run it manually from
 * the editor. To go back to empty, clear the content of "Records" by hand.
 */
function generateTestRecords(daysBack) {
  daysBack = daysBack || 60;

  const studentsSheet = spreadsheet_().getSheetByName(SHEET_STUDENTS);
  const studentRows = studentsSheet.getRange(2, 1, Math.max(studentsSheet.getLastRow() - 1, 0), 2).getValues()
    .filter((r) => r[0] && r[1]);

  if (!studentRows.length) {
    SpreadsheetApp.getUi().alert('There are no students in the "Students" tab. Generate them first with generateTestStudents().');
    return;
  }

  const dates = [];
  const cursor = new Date();
  cursor.setHours(0, 0, 0, 0);
  while (dates.length < daysBack) {
    const weekday = cursor.getDay();
    if (weekday !== 0 && weekday !== 6) {
      dates.push(Utilities.formatDate(cursor, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    }
    cursor.setDate(cursor.getDate() - 1);
  }

  let email;
  try {
    email = Session.getActiveUser().getEmail() || 'test';
  } catch (e) {
    email = 'test';
  }
  const now = new Date();

  const rows = [];
  studentRows.forEach(([className, student]) => {
    dates.forEach((date) => {
      const rand = Math.random();
      let morning = false, afternoon = false, lateMorning = false, lateAfternoon = false;
      let justifiedMorning = false, justifiedAfternoon = false;

      if (rand < 0.03) {
        morning = true; afternoon = true; // full day
      } else if (rand < 0.07) {
        morning = true;
      } else if (rand < 0.09) {
        afternoon = true;
      } else if (rand < 0.12) {
        lateMorning = true;
      } else if (rand < 0.13) {
        lateAfternoon = true;
      }

      // "Justified" is shared between an absence and a late arrival in the
      // same half-day (they're mutually exclusive, so this is unambiguous).
      if (morning || lateMorning) justifiedMorning = Math.random() < 0.5;
      if (afternoon || lateAfternoon) justifiedAfternoon = Math.random() < 0.5;

      if (morning || afternoon || lateMorning || lateAfternoon) {
        rows.push([date, className, student, morning, afternoon, justifiedMorning, justifiedAfternoon, lateMorning, lateAfternoon, email, now]);
      }
    });
  });

  const sh = spreadsheet_().getSheetByName(SHEET_RECORDS);
  const existingRows = Math.max(sh.getLastRow() - 1, 0);
  if (existingRows > 0) sh.getRange(2, 1, existingRows, 11).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 11).setValues(rows);
  }

  SpreadsheetApp.getUi().alert(
    `Generated ${rows.length} test absence/late-arrival records, ` +
    `spread across ${dates.length} school days (weekends excluded) and ${studentRows.length} students.`
  );
}

/**
 * End-of-year grade promotion: moves every student from their current class
 * to the next one, following the map written in the "Promotion" tab
 * (Current class -> New class; leave "New class" blank for grades that
 * graduate and leave the school, e.g. Grade 6).
 *
 * Students with the "Do not promote (repeating)" box checked in the
 * "Students" tab stay in the same class and are left untouched.
 */
function promoteToNextGrade() {
  const promotionSheet = spreadsheet_().getSheetByName(SHEET_PROMOTION);
  if (!promotionSheet) {
    SpreadsheetApp.getUi().alert('Can\'t find the "Promotion" tab. Run setup() first.');
    return;
  }

  const mapping = {};
  promotionSheet.getRange(2, 1, Math.max(promotionSheet.getLastRow() - 1, 0), 2).getValues().forEach(([current, next]) => {
    if (current) mapping[String(current).trim()] = String(next || '').trim();
  });

  const studentsSheet = spreadsheet_().getSheetByName(SHEET_STUDENTS);
  const numRows = Math.max(studentsSheet.getLastRow() - 1, 0);
  if (!numRows) {
    SpreadsheetApp.getUi().alert('There are no students in the "Students" tab.');
    return;
  }

  const data = studentsSheet.getRange(2, 1, numRows, 3).getValues();
  let promoted = 0;
  let repeating = 0;
  let graduated = 0;
  let unmapped = 0;
  const rowsToDelete = [];

  for (let i = 0; i < data.length; i++) {
    const [className, student, doNotPromote] = data[i];
    if (!student) continue;

    if (doNotPromote === true) {
      repeating++;
      continue;
    }

    if (!(className in mapping)) {
      unmapped++;
      continue;
    }

    const newClassName = mapping[className];
    if (!newClassName) {
      rowsToDelete.push(i + 2); // graduating / leaving the school
      graduated++;
    } else {
      studentsSheet.getRange(i + 2, 1).setValue(newClassName);
      promoted++;
    }
  }

  rowsToDelete
    .sort((a, b) => b - a)
    .forEach((row) => studentsSheet.deleteRow(row));

  // Reset the "repeating" checkboxes for next year.
  const finalRowCount = Math.max(studentsSheet.getLastRow() - 1, 0);
  if (finalRowCount) {
    studentsSheet.getRange(2, 3, finalRowCount, 1).setValue(false);
  }

  SpreadsheetApp.getUi().alert(
    'Promotion complete:\n\n' +
    `• ${promoted} students promoted to a new class\n` +
    `• ${repeating} repeating students kept in the same class\n` +
    `• ${graduated} students graduated/removed (blank new class in the mapping)\n` +
    (unmapped ? `• ⚠️ ${unmapped} students had a class not found in the "Promotion" tab (left untouched)\n` : '') +
    '\nCheck the "Students" tab to confirm everything looks right.'
  );
}

/**
 * FOR PERFORMANCE TESTING ONLY: calls getSummary() (with all classes, the
 * heaviest case) and shows how long it took in milliseconds.
 */
function measurePerformance() {
  const start = new Date().getTime();
  const result = getSummary('');
  const end = new Date().getTime();

  SpreadsheetApp.getUi().alert(
    `getSummary() took ${end - start} ms\n` +
    `(${result.length} students processed)`
  );
}

function getConfig_() {
  const sh = spreadsheet_().getSheetByName(SHEET_CONFIG);
  const values = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 0), 2).getValues();
  const config = {};
  values.forEach(([key, value]) => {
    if (key) config[key] = value;
  });
  return config;
}

function getClasses() {
  const sh = spreadsheet_().getSheetByName(SHEET_STUDENTS);
  const values = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 0), 1).getValues();
  const set = new Set(values.map((r) => r[0]).filter(String));
  return Array.from(set).sort();
}

function dayKey_(date, student) {
  return date + '||' + student;
}

function weekdayInfo_(isoDate) {
  const names = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const parts = isoDate.split('-').map(Number);
  const d = new Date(parts[0], parts[1] - 1, parts[2]);
  return { name: names[d.getDay()], isWeekend: d.getDay() === 0 || d.getDay() === 6 };
}

/**
 * Returns every student in a class along with whatever marks already exist
 * for `date` (if attendance hasn't been taken yet that day, all marks come
 * back false).
 */
function getAttendanceGrid(className, date) {
  const studentsSheet = spreadsheet_().getSheetByName(SHEET_STUDENTS);
  const studentValues = studentsSheet.getRange(2, 1, Math.max(studentsSheet.getLastRow() - 1, 0), 2).getValues();
  const students = studentValues
    .filter((r) => r[0] === className && r[1])
    .map((r) => r[1])
    .sort((a, b) => a.localeCompare(b, 'en'));

  const marks = {};
  const recordsSheet = spreadsheet_().getSheetByName(SHEET_RECORDS);
  const lastRow = recordsSheet.getLastRow();
  if (lastRow >= 2) {
    const recordValues = recordsSheet.getRange(2, 1, lastRow - 1, 9).getValues();
    recordValues.forEach(([rDate, rClass, rStudent, morning, afternoon, justifiedMorning, justifiedAfternoon, lateMorning, lateAfternoon]) => {
      if (rClass === className && formatDateISO_(rDate) === date) {
        marks[rStudent] = {
          morning: morning === true,
          afternoon: afternoon === true,
          justifiedMorning: justifiedMorning === true,
          justifiedAfternoon: justifiedAfternoon === true,
          lateMorning: lateMorning === true,
          lateAfternoon: lateAfternoon === true,
        };
      }
    });
  }

  const dayInfo = weekdayInfo_(date);

  return {
    weekday: dayInfo.name,
    isWeekend: dayInfo.isWeekend,
    students: students.map((student) => {
      const m = marks[student] || { morning: false, afternoon: false, justifiedMorning: false, justifiedAfternoon: false, lateMorning: false, lateAfternoon: false };
      return {
        student,
        morning: m.morning,
        afternoon: m.afternoon,
        justifiedMorning: m.justifiedMorning,
        justifiedAfternoon: m.justifiedAfternoon,
        lateMorning: m.lateMorning,
        lateAfternoon: m.lateAfternoon,
      };
    }),
  };
}

/**
 * rows = [{ student, morning, afternoon, justifiedMorning, justifiedAfternoon, lateMorning, lateAfternoon }, ...]
 * for a specific class and date. Updates, creates or deletes the matching row for each student as needed.
 * "justifiedMorning"/"justifiedAfternoon" apply to whichever of the absence
 * (morning/afternoon) or the late arrival (lateMorning/lateAfternoon) is set.
 */
function saveAttendanceGrid(className, date, rows) {
  if (!className || !date || !rows) throw new Error('Missing data.');

  const sh = spreadsheet_().getSheetByName(SHEET_RECORDS);
  const lastRow = sh.getLastRow();
  const numRows = Math.max(lastRow - 1, 0);
  const values = numRows ? sh.getRange(2, 1, numRows, 11).getValues() : [];

  const rowIndexByKey = {};
  values.forEach((row, i) => {
    if (row[1] === className && formatDateISO_(row[0]) === date) {
      rowIndexByKey[dayKey_(date, row[2])] = i + 2; // actual row on the sheet
    }
  });

  let email;
  try {
    email = Session.getActiveUser().getEmail() || 'unknown';
  } catch (e) {
    email = 'unknown';
  }
  const now = new Date();

  const rowsToDelete = [];

  rows.forEach((r) => {
    const hasAnyMark = r.morning || r.afternoon || r.lateMorning || r.lateAfternoon;
    const key = dayKey_(date, r.student);
    const existingRow = rowIndexByKey[key];

    if (!hasAnyMark) {
      if (existingRow) rowsToDelete.push(existingRow);
      return;
    }

    const newRow = [
      date, className, r.student,
      !!r.morning, !!r.afternoon,
      !!r.justifiedMorning, !!r.justifiedAfternoon,
      !!r.lateMorning, !!r.lateAfternoon,
      email, now,
    ];

    if (existingRow) {
      sh.getRange(existingRow, 1, 1, 11).setValues([newRow]);
    } else {
      sh.appendRow(newRow);
    }
  });

  rowsToDelete
    .sort((a, b) => b - a)
    .forEach((row) => sh.deleteRow(row));

  incrementRecordsVersion_();

  return true;
}

/**
 * Apps Script's cache only accepts ~100KB per key. With many students, the
 * Summary's JSON can exceed that, so we compress it (gzip + base64) before
 * storing it, and reverse that when reading it back.
 */
function cacheGet_(cache, key) {
  const base64 = cache.get(key);
  if (!base64) return null;
  try {
    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, 'application/x-gzip');
    return Utilities.ungzip(blob).getDataAsString();
  } catch (e) {
    return null;
  }
}
function cachePut_(cache, key, str, ttlSec) {
  try {
    const gzipBlob = Utilities.gzip(Utilities.newBlob(str));
    const base64 = Utilities.base64Encode(gzipBlob.getBytes());
    if (base64.length <= 100000) cache.put(key, base64, ttlSec);
  } catch (e) {
    // If it's still too large even after compression, we just don't cache it.
  }
}

function getSummary(classFilter) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'summary_v' + recordsVersion_() + '_' + (classFilter || '__all__');
  const cachedValue = cacheGet_(cache, cacheKey);
  if (cachedValue) return JSON.parse(cachedValue);

  const config = getConfig_();
  const weightMorning = Number(config.WEIGHT_MORNING) || 0;
  const weightAfternoon = Number(config.WEIGHT_AFTERNOON) || 0;
  const totalSchoolDays = Number(config.TOTAL_SCHOOL_DAYS) || 0;
  const threshold = Number(config.ALERT_THRESHOLD_PCT) || 10;

  const studentsSheet = spreadsheet_().getSheetByName(SHEET_STUDENTS);
  const studentValues = studentsSheet.getRange(2, 1, Math.max(studentsSheet.getLastRow() - 1, 0), 2).getValues();

  const emptyStats_ = () => ({
    unjustifiedDays: 0, justifiedDays: 0,
    lateMorningsJustified: 0, lateMorningsUnjustified: 0,
    lateAfternoonsJustified: 0, lateAfternoonsUnjustified: 0,
    byWeekday: [0, 0, 0, 0, 0, 0, 0],
  });
  const stats = {};
  const key_ = (className, student) => className + ' ||| ' + student;

  studentValues.forEach(([className, student]) => {
    if (!student) return;
    if (classFilter && className !== classFilter) return;
    stats[key_(className, student)] = Object.assign({ className, student }, emptyStats_());
  });

  const recordsSheet = spreadsheet_().getSheetByName(SHEET_RECORDS);
  const lastRow = recordsSheet.getLastRow();
  if (lastRow >= 2) {
    const recordValues = recordsSheet.getRange(2, 1, lastRow - 1, 9).getValues();
    recordValues.forEach(([date, className, student, morning, afternoon, justifiedMorning, justifiedAfternoon, lateMorning, lateAfternoon]) => {
      if (!student) return;
      if (classFilter && className !== classFilter) return;

      const k = key_(className, student);
      if (!stats[k]) {
        stats[k] = Object.assign({ className, student }, emptyStats_());
      }

      // "Morning" (absence) and "Late morning" (late arrival) are mutually
      // exclusive for the same half-day, so whichever one is set decides
      // what "justifiedMorning" actually refers to. Same for the afternoon.
      if (morning === true) {
        if (justifiedMorning === true) stats[k].justifiedDays += weightMorning;
        else stats[k].unjustifiedDays += weightMorning;
      } else if (lateMorning === true) {
        if (justifiedMorning === true) stats[k].lateMorningsJustified += 1;
        else stats[k].lateMorningsUnjustified += 1;
      }

      if (afternoon === true) {
        if (justifiedAfternoon === true) stats[k].justifiedDays += weightAfternoon;
        else stats[k].unjustifiedDays += weightAfternoon;
      } else if (lateAfternoon === true) {
        if (justifiedAfternoon === true) stats[k].lateAfternoonsJustified += 1;
        else stats[k].lateAfternoonsUnjustified += 1;
      }

      if (morning === true || afternoon === true) {
        const weekday = new Date(formatDateISO_(date)).getUTCDay();
        stats[k].byWeekday[weekday] += 1;
      }
    });
  }

  const WEEKDAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

  const result = Object.values(stats)
    .map((s) => {
      const total = s.unjustifiedDays + s.justifiedDays;
      const unjustifiedPct = totalSchoolDays ? (s.unjustifiedDays / totalSchoolDays) * 100 : 0;
      const justifiedPct = totalSchoolDays ? (s.justifiedDays / totalSchoolDays) * 100 : 0;
      const totalPct = unjustifiedPct + justifiedPct;

      // Justification rate: what % of THEIR OWN absences are justified
      // (different from justifiedPct, which is over the whole school year).
      // Kept as a number|null for sorting, plus a ready-to-display string
      // (computed here, not in the frontend) so the web page never has to
      // guess how to render a missing value.
      const justificationRate = total > 0 ? round_((s.justifiedDays / total) * 100) : null;
      const justificationRateLabel = justificationRate === null ? '—' : justificationRate + '%';

      // Breakdown of every weekday with at least one absence, sorted from most to least frequent.
      const mostFrequentDay = s.byWeekday
        .map((n, i) => ({ name: WEEKDAY_NAMES[i], n }))
        .filter((d) => d.n > 0)
        .sort((a, b) => b.n - a.n)
        .map((d) => `${d.name} (${d.n})`)
        .join(', ') || '—';

      // Progressive 3-step scale, relative to the configured alert threshold.
      let level = 'ok';
      if (totalPct >= threshold) level = 'alert';
      else if (totalPct >= threshold / 2) level = 'warning';

      return {
        className: s.className,
        student: s.student,
        unjustifiedDays: round_(s.unjustifiedDays),
        justifiedDays: round_(s.justifiedDays),
        totalDays: round_(total),
        unjustifiedPct: round_(unjustifiedPct),
        justifiedPct: round_(justifiedPct),
        totalPct: round_(totalPct),
        justificationRate,
        justificationRateLabel,
        isAlert: totalPct >= threshold,
        level,
        lateMorningsJustified: s.lateMorningsJustified,
        lateMorningsUnjustified: s.lateMorningsUnjustified,
        lateAfternoonsJustified: s.lateAfternoonsJustified,
        lateAfternoonsUnjustified: s.lateAfternoonsUnjustified,
        mostFrequentDay,
      };
    })
    .sort((a, b) => b.totalPct - a.totalPct);

  cachePut_(cache, cacheKey, JSON.stringify(result), 120); // 2 minutes

  return result;
}

function formatDateISO_(d) {
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return d;
}

function round_(n) {
  return Math.round(n * 100) / 100;
}
