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
  { id:'t8',  title:'סידרתי את המיטה והחדר',         icon:'🛏️' },
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
