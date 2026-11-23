// ============================================================
// MORNING ARENA — Google Apps Script Backend
// ============================================================

const SHEET_NAME   = 'Logs_V2';
const LOCK_MINUTES = 5;
const USERS        = ['אלון', 'איתן'];

const TASKS_CONFIG = [
  { id:'t1',  title:'התעוררתי באנרגיה טובה',       icon:'⚡' },
  { id:'t2',  title:'שטפתי פנים והתרעננתי',         icon:'💧' },
  { id:'t3',  title:'צחצחתי שיניים יסודי',           icon:'🦷' },
  { id:'t4',  title:'התארגנתי בשירותים',             icon:'🚽' },
  { id:'t5',  title:'התלבשתי',                       icon:'👕' },
  { id:'t6',  title:'נעלתי נעליים',                  icon:'👟' },
  { id:'t7',  title:'סידרתי את השיער',               icon:'💇‍♂️' },
  { id:'t8',  title:'סידרתי את המיטה',         icon:'🛏️' },
  { id:'t9',  title:'ארגנתי אוכל ומים לתיק',         icon:'🎒' },
  { id:'t10', title:'אכלתי ארוחת בוקר',              icon:'🥪' }
];

// ── Entry Point ──────────────────────────────────────────
function doGet() {
  initSheetIfNeeded();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('MORNING ARENA | אלון ואיתן')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── Sheet helpers ────────────────────────────────────────
function getSheet_() {
  initSheetIfNeeded();
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function initSheetIfNeeded() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) return;

  sheet = ss.insertSheet(SHEET_NAME);
  sheet.appendRow(['Date', 'User', 'TaskID', 'TaskTitle', 'Status', 'CheckedAtTimestamp']);
  sheet.getRange('A1:F1')
    .setFontWeight('bold')
    .setBackground('#111a2d')
    .setFontColor('#38bdf8')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, 6, 160);
}

function validateUser_(u) {
  if (!USERS.includes(u)) throw new Error('משתמש לא מורשה');
}

function validateTask_(id) {
  const t = TASKS_CONFIG.find(t => t.id === id);
  if (!t) throw new Error('משימה לא חוקית');
  return t;
}

function getTodayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatRowDate(value) {
  if (!value) return '';
  if (value instanceof Date)
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(value).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

// ── Dashboard ────────────────────────────────────────────
function getUserDashboardData(userName) {
  validateUser_(userName);

  const sheet   = getSheet_();
  const data    = sheet.getDataRange().getValues();
  const today   = getTodayStr();
  const now     = Date.now();
  const lockMs  = LOCK_MINUTES * 60 * 1000;

  const todayState = {};
  TASKS_CONFIG.forEach(t => {
    todayState[t.id] = { status:false, checkedAt:null, locked:false, remainingSeconds:0 };
  });

  const completedDates = new Set();
  const dailyCounts    = {};

  for (let i = 1; i < data.length; i++) {
    const row     = data[i];
    const date    = formatRowDate(row[0]);
    const user    = String(row[1]).trim();
    const taskId  = String(row[2]).trim();
    const status  = row[4] === true || row[4] === 'TRUE';
    const ts      = Number(row[5]) || 0;

    if (user !== userName || !date) continue;

    if (status) dailyCounts[date] = (dailyCounts[date] || 0) + 1;

    if (date === today && todayState[taskId]) {
      const locked    = status && ts > 0 && (now - ts) >= lockMs;
      const remaining = status && ts > 0 && !locked
        ? Math.ceil((lockMs - (now - ts)) / 1000) : 0;
      todayState[taskId] = { status, checkedAt: ts || null, locked, remainingSeconds: remaining };
    }
  }

  Object.keys(dailyCounts).forEach(d => { if (dailyCounts[d] >= 7) completedDates.add(d); });

  return {
    user:       userName,
    tasks:      TASKS_CONFIG,
    todayState,
    todayDate:  today,
    streak:     calculateStreak_(completedDates, today)
  };
}

function calculateStreak_(completedDates, today) {
  let streak = 0;
  const [y, m, d] = today.split('-').map(Number);
  const date = new Date(y, m - 1, d);

  if (!completedDates.has(today)) date.setDate(date.getDate() - 1);

  while (true) {
    const str = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!completedDates.has(str)) break;
    streak++;
    date.setDate(date.getDate() - 1);
  }
  return streak;
}

// ── Toggle Task ──────────────────────────────────────────
function toggleTask(userName, taskId, newStatus) {
  validateUser_(userName);
  const task = validateTask_(taskId);
  if (typeof newStatus !== 'boolean') throw new Error('סטטוס משימה לא חוקי');

  const lock  = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet  = getSheet_();
    const data   = sheet.getDataRange().getValues();
    const today  = getTodayStr();
    const now    = Date.now();
    const lockMs = LOCK_MINUTES * 60 * 1000;

    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (formatRowDate(data[i][0]) === today &&
          String(data[i][1]).trim() === userName &&
          String(data[i][2]).trim() === taskId) {
        targetRow = i + 1; break;
      }
    }

    if (targetRow > 0) {
      const curStatus  = data[targetRow-1][4] === true || data[targetRow-1][4] === 'TRUE';
      const checkedAt  = Number(data[targetRow-1][5]) || 0;

      if (curStatus && checkedAt > 0 && (now - checkedAt) >= lockMs)
        return { success:false, error:'🔒 עברו 5 דקות. המשימה ננעלה להיום.' };

      if (curStatus === newStatus)
        return {
          success:true, status:curStatus, locked:false,
          remainingSeconds: curStatus ? Math.max(0, Math.ceil((lockMs-(now-checkedAt))/1000)) : 0
        };

      sheet.getRange(targetRow, 5, 1, 2).setValues([[newStatus, newStatus ? now : '']]);

    } else {
      if (!newStatus) return { success:true, status:false, locked:false, remainingSeconds:0 };
      sheet.appendRow([today, userName, taskId, task.title, true, now]);
    }

    return { success:true, status:newStatus, locked:false, remainingSeconds: newStatus ? LOCK_MINUTES*60 : 0 };

  } finally {
    lock.releaseLock();
  }
}

// ── History Stats ────────────────────────────────────────
function getHistoryStats(userName, period) {
  validateUser_(userName);
  if (!['day','week','month'].includes(period)) throw new Error('תקופה לא חוקית');

  const sheet = getSheet_();
  const data  = sheet.getDataRange().getValues();
  const dailyMap = {};

  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const date   = formatRowDate(row[0]);
    const user   = String(row[1]).trim();
    const title  = String(row[3]);
    const status = row[4] === true || row[4] === 'TRUE';

    if (!date || user !== userName) continue;
    if (!dailyMap[date]) dailyMap[date] = { date, completedCount:0, completedTasks:[] };
    if (status) {
      dailyMap[date].completedCount++;
      if (!dailyMap[date].completedTasks.includes(title)) dailyMap[date].completedTasks.push(title);
    }
  }

  const todayParts = getTodayStr().split('-').map(Number);
  const today      = new Date(todayParts[0], todayParts[1]-1, todayParts[2]);
  const maxDays    = { day:2, week:7, month:30 }[period];

  const filtered = Object.values(dailyMap)
    .filter(item => {
      const [y,m,d] = item.date.split('-').map(Number);
      const diff    = Math.round((today - new Date(y,m-1,d)) / 86400000);
      return diff >= 0 && diff <= maxDays;
    })
    .sort((a,b) => b.date.localeCompare(a.date));

  const totalCompleted = filtered.reduce((s, i) => s + i.completedCount, 0);
  const totalPossible  = filtered.length * TASKS_CONFIG.length;

  return {
    days:          filtered,
    totalTasksPerDay: TASKS_CONFIG.length,
    avgPercent:    totalPossible ? Math.round(totalCompleted/totalPossible*100) : 0,
    totalCompleted
  };
}

// ============================================================
// ATTENDANCE MODULE (Alon – school lessons)
// ============================================================
const ATTENDANCE_SHEET = 'Attendance_V1';

function initAttendanceSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(ATTENDANCE_SHEET)) return;
  const sh = ss.insertSheet(ATTENDANCE_SHEET);
  sh.appendRow(['Date','User','Slot','LessonTitle','Present','UpdatedAt']);
  sh.getRange('A1:F1').setFontWeight('bold').setBackground('#111a2d').setFontColor('#38bdf8').setHorizontalAlignment('center');
  sh.setFrozenRows(1); sh.setColumnWidths(1,6,160);
}

function getAttendSheet_() {
  initAttendanceSheet_();
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET);
}

function getAttendanceData(userName, dateStr) {
  validateUser_(userName);
  if (!dateStr) dateStr = getTodayStr();
  const sheet = getAttendSheet_();
  const data  = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const date = formatRowDate(row[0]);
    const user = String(row[1]).trim();
    const slot = Number(row[2]);
    if (date === dateStr && user === userName)
      result[slot] = { present: row[4] === true || row[4] === 'TRUE', title: String(row[3]) };
  }
  return { date: dateStr, attendance: result };
}

function setAttendance(userName, slot, lessonTitle, present, dateStr) {
  validateUser_(userName);
  if (!dateStr) dateStr = getTodayStr();
  slot = Number(slot);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sheet = getAttendSheet_();
    const data  = sheet.getDataRange().getValues();
    const now   = Date.now();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (formatRowDate(data[i][0]) === dateStr && String(data[i][1]).trim() === userName && Number(data[i][2]) === slot) {
        targetRow = i + 1; break;
      }
    }
    if (targetRow > 0) sheet.getRange(targetRow,5,1,2).setValues([[present, now]]);
    else sheet.appendRow([dateStr, userName, slot, lessonTitle, present, now]);
    return { success: true };
  } finally { lock.releaseLock(); }
}

function getAttendanceHistory(userName, period) {
  validateUser_(userName);
  if (!['week','month','term'].includes(period)) period = 'week';
  const sheet = getAttendSheet_();
  const data  = sheet.getDataRange().getValues();
  const dailyMap = {};
  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const date = formatRowDate(row[0]);
    const user = String(row[1]).trim();
    const pres = row[4] === true || row[4] === 'TRUE';
    if (!date || user !== userName) continue;
    if (!dailyMap[date]) dailyMap[date] = { date, present:0, absent:0 };
    if (pres) dailyMap[date].present++; else dailyMap[date].absent++;
  }
  const todayParts = getTodayStr().split('-').map(Number);
  const today = new Date(todayParts[0], todayParts[1]-1, todayParts[2]);
  const maxDays = { week:7, month:30, term:90 }[period];
  const filtered = Object.values(dailyMap)
    .filter(item => { const [y,m,d] = item.date.split('-').map(Number); const diff = Math.round((today - new Date(y,m-1,d))/86400000); return diff >= 0 && diff <= maxDays; })
    .sort((a,b) => b.date.localeCompare(a.date));
  const totalPresent  = filtered.reduce((s,i) => s + i.present, 0);
  const totalLessons  = filtered.reduce((s,i) => s + i.present + i.absent, 0);
  return { days: filtered, totalPresent, totalLessons, avgPercent: totalLessons ? Math.round(totalPresent/totalLessons*100) : 0 };
}

// ============================================================
// EVENING ROUTINE MODULE (Alon + Eitan)
// ============================================================
const EVENING_SHEET = 'Evening_V1';
const EVENING_TASKS = [
  { id:'ev1', title:'שיעורי בית',       icon:'📚' },
  { id:'ev2', title:'מקלחת',            icon:'🚿' },
  { id:'ev3', title:'ארוחת ערב',        icon:'🍽️' },
  { id:'ev4', title:'הכנת בגדים למחר',  icon:'👕' },
];

function initEveningSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(EVENING_SHEET)) return;
  const sh = ss.insertSheet(EVENING_SHEET);
  sh.appendRow(['Date','User','TaskID','TaskTitle','Done','DoneAt']);
  sh.getRange('A1:F1').setFontWeight('bold').setBackground('#111a2d').setFontColor('#38bdf8').setHorizontalAlignment('center');
  sh.setFrozenRows(1); sh.setColumnWidths(1,6,160);
}

function getEveningSheet_() {
  initEveningSheet_();
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EVENING_SHEET);
}

function getEveningData(userName) {
  validateUser_(userName);
  const sheet = getEveningSheet_();
  const data  = sheet.getDataRange().getValues();
  const today = getTodayStr();
  const doneMap = {};
  const allDaysDone = {};
  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const date = formatRowDate(row[0]);
    const user = String(row[1]).trim();
    const tid  = String(row[2]).trim();
    const done = row[4] === true || row[4] === 'TRUE';
    if (!date || user !== userName) continue;
    if (date === today) doneMap[tid] = done;
    if (done && EVENING_TASKS.find(t => t.id === tid)) { allDaysDone[date] = (allDaysDone[date]||0)+1; }
  }
  const completedDates = new Set(Object.keys(allDaysDone).filter(d => allDaysDone[d] >= EVENING_TASKS.length));
  return { tasks: EVENING_TASKS, doneMap, today, streak: calculateStreak_(completedDates, today) };
}

function toggleEveningTask(userName, taskId, done) {
  validateUser_(userName);
  const task = EVENING_TASKS.find(t => t.id === taskId);
  if (!task) throw new Error('משימת ערב לא חוקית');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sheet = getEveningSheet_();
    const data  = sheet.getDataRange().getValues();
    const today = getTodayStr();
    const now   = Date.now();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (formatRowDate(data[i][0]) === today && String(data[i][1]).trim() === userName && String(data[i][2]).trim() === taskId) {
        targetRow = i + 1; break;
      }
    }
    if (targetRow > 0) sheet.getRange(targetRow,5,1,2).setValues([[done, done ? now : '']]);
    else if (done) sheet.appendRow([today, userName, taskId, task.title, true, now]);
    return { success: true };
  } finally { lock.releaseLock(); }
}

function getEveningHistory(userName, period) {
  validateUser_(userName);
  if (!['week','month'].includes(period)) period = 'week';
  const sheet = getEveningSheet_();
  const data  = sheet.getDataRange().getValues();
  const dailyMap = {};
  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const date = formatRowDate(row[0]);
    const user = String(row[1]).trim();
    const tid  = String(row[2]).trim();
    const done = row[4] === true || row[4] === 'TRUE';
    if (!date || user !== userName) continue;
    if (!dailyMap[date]) dailyMap[date] = { date, done:0 };
    if (done && EVENING_TASKS.find(t => t.id === tid)) dailyMap[date].done++;
  }
  const todayParts = getTodayStr().split('-').map(Number);
  const today = new Date(todayParts[0], todayParts[1]-1, todayParts[2]);
  const maxDays = { week:7, month:30 }[period];
  const filtered = Object.values(dailyMap)
    .filter(item => { const [y,m,d] = item.date.split('-').map(Number); const diff = Math.round((today - new Date(y,m-1,d))/86400000); return diff >= 0 && diff <= maxDays; })
    .sort((a,b) => b.date.localeCompare(a.date));
  return { days: filtered, totalTasks: EVENING_TASKS.length, totalFullDays: filtered.filter(d => d.done >= EVENING_TASKS.length).length };
}

// ============================================================
// FULL HISTORY — combined morning + schedule + evening per day
// ============================================================
function getFullHistory(userName, period) {
  validateUser_(userName);
  if (!['day','week','month'].includes(period)) period = 'week';

  const todayParts = getTodayStr().split('-').map(Number);
  const today = new Date(todayParts[0], todayParts[1]-1, todayParts[2]);
  const maxDays = {day:1, week:7, month:30}[period];

  const HEB = ['ראשון','שני','שלישי','רביעי','חמישי','שישי','שבת'];

  function inRange(dateStr) {
    const [y,m,d] = dateStr.split('-').map(Number);
    const diff = Math.round((today - new Date(y,m-1,d)) / 86400000);
    return diff >= 0 && diff < maxDays;
  }

  const dayMap = {};
  function ensureDay(dateStr) {
    if (!dayMap[dateStr]) {
      const [y,m,d] = dateStr.split('-').map(Number);
      dayMap[dateStr] = {date:dateStr, dayName:HEB[new Date(y,m-1,d).getDay()], morningMap:{}, attendMap:{}, eveningMap:{}};
    }
    return dayMap[dateStr];
  }

  // Morning (Logs_V2): Date, User, TaskID, TaskTitle, Status, ...
  const logData = getSheet_().getDataRange().getValues();
  for (var i = 1; i < logData.length; i++) {
    var r = logData[i];
    var date = formatRowDate(r[0]);
    if (!date || String(r[1]).trim() !== userName || !inRange(date)) continue;
    var dy = ensureDay(date);
    dy.morningMap[String(r[2]).trim()] = {done: r[4]===true||r[4]==='TRUE'};
  }

  // Attendance (Attendance_V1): Date, User, Slot, LessonTitle, Present, ...
  var attData = getAttendSheet_().getDataRange().getValues();
  for (var i = 1; i < attData.length; i++) {
    var r = attData[i];
    var date = formatRowDate(r[0]);
    if (!date || String(r[1]).trim() !== userName || !inRange(date)) continue;
    var dy = ensureDay(date);
    dy.attendMap[Number(r[2])] = {present: r[4]===true||r[4]==='TRUE'};
  }

  // Evening (Evening_V1): Date, User, TaskID, TaskTitle, Done, ...
  var evData = getEveningSheet_().getDataRange().getValues();
  for (var i = 1; i < evData.length; i++) {
    var r = evData[i];
    var date = formatRowDate(r[0]);
    if (!date || String(r[1]).trim() !== userName || !inRange(date)) continue;
    var dy = ensureDay(date);
    dy.eveningMap[String(r[2]).trim()] = {done: r[4]===true||r[4]==='TRUE'};
  }

  var days = Object.values(dayMap)
    .sort(function(a,b){return b.date.localeCompare(a.date)})
    .map(function(d) {
      var morningDone  = Object.values(d.morningMap).filter(function(t){return t.done}).length;
      var attendArr    = Object.values(d.attendMap);
      var attPresent   = attendArr.filter(function(a){return a.present}).length;
      var eveningDone  = Object.values(d.eveningMap).filter(function(t){return t.done}).length;
      return {
        date:        d.date,
        dayName:     d.dayName,
        morning:     {done: morningDone,  total: TASKS_CONFIG.length},
        attendance:  {present: attPresent, total: attendArr.length},
        evening:     {done: eveningDone,  total: EVENING_TASKS.length}
      };
    });

  return {days:days, tasksPerDay:TASKS_CONFIG.length, eveningTotal:EVENING_TASKS.length};
}
