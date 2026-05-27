// ============================================================
// TASK MANAGER — Code.gs
// ============================================================
// SETUP INSTRUCTIONS:
//
// 1. Replace YOUR_SPREADSHEET_ID_HERE with your Google Sheets ID
//    (found in the sheet URL between /d/ and /edit)
//
// 2. If using SMS/WhatsApp notifications, fill in:
//    - NOTIFY_MODE: 'sms_fast2sms' | 'sms_twilio' | 'whatsapp_twilio'
//    - FAST2SMS_API_KEY: your Fast2SMS API key
//    - TWILIO_ACCOUNT_SID / TWILIO_AUTH_TOKEN: from Twilio console
//    - TWILIO_SMS_FROM / TWILIO_WA_FROM: your Twilio numbers
//    Leave NOTIFY_MODE as 'none' if not using SMS/WhatsApp.
//
// 3. In doGet(), set:
//    const page = 'director'  → for the Director deployment
//    const page = 'employee'  → for the Employee deployment
//
// 4. Deploy as Web App:
//    Execute as: Me
//    Who has access: Anyone with Google Account
//
// 5. Run setupDailyTrigger() once to enable 7 AM auto-refresh
// 6. Run setupDailySummaryTrigger() once to enable 5 PM email summary
// ============================================================

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // ← Replace with your Google Sheets ID
const TASKS_SHEET    = 'Tasks';
const DATA_SHEET     = 'Data';
const DASH_SHEET     = 'Dashboard';

// ── Tasks sheet column indices (0-based) ─────────────────────
const COL = {
  DATE        : 0,   // A
  TASK        : 1,   // B
  DEPARTMENT  : 2,   // C
  ASSIGNED_TO : 3,   // D
  EMAIL       : 4,   // E
  PLANNED_WHEN: 5,   // F
  ACTUAL_DATE : 6,   // G
  STATUS      : 7,   // H
  REMARKS     : 8,   // I
  ASSIGNED_BY : 9,   // J
  ROW_ID      : 10,  // K
  DELAY       : 11,  // L  ← auto-calculated by scoring engine
  SCORE       : 12,  // M  ← GREEN / YELLOW / RED / PENDING
  WEEK        : 13   // N  ← Week Number based on Planned When
};

// ============================================================
// NOTIFICATION CONFIG
// 'sms_fast2sms' | 'sms_twilio' | 'whatsapp_twilio' | 'none'
// ============================================================
const NOTIFY_MODE = 'none';

const FAST2SMS_API_KEY   = 'YOUR_FAST2SMS_API_KEY_HERE';
const TWILIO_ACCOUNT_SID = 'ACxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
const TWILIO_AUTH_TOKEN  = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
const TWILIO_SMS_FROM    = '+1XXXXXXXXXX';
const TWILIO_WA_FROM     = 'whatsapp:+14155238886';

// ============================================================
// WEB APP ENTRY POINT
// Director deployment → const page = 'director'
// Employee deployment → const page = 'employee'
// ============================================================
function doGet(e) {
  const page = 'director'; // CHANGE TO 'employee' in employee deployment
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(page === 'director' ? 'Task Manager – Director' : 'Task Manager – Employee')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// AUTH — TWO-STEP USER IDENTIFICATION
// Step 1: getViewerEmail() → gets actual logged-in user email
// Step 2: getUserByEmail(email) → looks up full record from Data sheet
// ============================================================
function getViewerEmail() {
  try {
    const email = Session.getEffectiveUser().getEmail();
    if (!email) return { error: 'Could not detect logged-in user.' };
    return { email };
  } catch(err) { return { error: err.message }; }
}

function getUserByEmail(email) {
  try {
    if (!email) return { error: 'No email provided' };
    const data = SpreadsheetApp.openById(SPREADSHEET_ID)
                   .getSheetByName(DATA_SHEET).getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((data[i][2]||'').toString().trim().toLowerCase() === email.toLowerCase()) {
        return {
          email,
          name      : (data[i][1]||email).toString().trim(),
          department: (data[i][0]||'').toString().trim(),
          role      : (data[i][3]||'employee').toString().trim().toLowerCase(),
          phone     : (data[i][4]||'').toString().trim()
        };
      }
    }
    return { email, name: email, department: '', role: 'employee', phone: '' };
  } catch(err) { return { error: err.message }; }
}

function getCurrentUser() {
  const r = getViewerEmail();
  if (r.error) return r;
  return getUserByEmail(r.email);
}

// ============================================================
// DROPDOWN DATA
// A2:A=Departments | B2:E=Name,Email,Role,Phone | F2:F=Statuses
// ============================================================
function getDropdownData() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET);
    const last  = sheet.getLastRow();
    if (last < 2) return { departments:[], employees:[], statuses:[] };

    const departments = sheet.getRange('A2:A'+last).getValues()
      .map(r => r[0].toString().trim()).filter(Boolean);

    const employees = sheet.getRange('B2:E'+last).getValues()
      .filter(r => r[0].toString().trim())
      .map(r => ({
        name : r[0].toString().trim(),
        email: (r[1]||'').toString().trim(),
        role : (r[2]||'employee').toString().trim().toLowerCase(),
        phone: (r[3]||'').toString().trim()
      }));

    const statuses = sheet.getRange('F2:F'+last).getValues()
      .map(r => r[0].toString().trim()).filter(Boolean);

    return { departments, employees, statuses };
  } catch(err) { return { error: err.message }; }
}

// ============================================================
// ENSURE Tasks sheet exists with headers
// ============================================================
function ensureTasksSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(TASKS_SHEET);
  if (!sheet) sheet = ss.insertSheet(TASKS_SHEET);
  if (sheet.getRange(1,1).getValue() !== 'Date') {
    const h = ['Date','Task','Department','Assigned To','Email',
               'Planned When','Actual Date','Status','Remarks','Assigned By','Row ID'];
    sheet.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold');
  }
  return sheet;
}

// ============================================================
// SEND TASK — creates new row, sends email + SMS/WhatsApp
// Also triggers scoring so Dashboard stays current
// ============================================================
function sendTask(taskData) {
  try {
    const sheet = ensureTasksSheet();
    const rowId = Utilities.getUuid();

    // Use assignedByName sent from the frontend (the actual logged-in user)
    // This fixes the bug where Execute-as-Me returned the script owner instead
    const assignedBy = (taskData.assignedByName || '').toString().trim()
                    || (taskData.assignedByEmail || '').toString().trim()
                    || 'Unknown';

    sheet.appendRow([
      taskData.date, taskData.task, taskData.department,
      taskData.assignedTo, taskData.email, taskData.plannedWhen,
      '', 'Pending', '', assignedBy, rowId
    ]);

    if (taskData.email) _emailNew(taskData, assignedBy);
    if (taskData.phone) {
      _sendNotification(taskData.phone,
        'New Task Assigned!\nTask: ' + taskData.task +
        '\nDept: ' + taskData.department +
        '\nBy: ' + assignedBy +
        '\nDeadline: ' + taskData.plannedWhen +
        '\nLogin to Task Manager to update.');
    }

    // Auto-refresh scoring & dashboard after new task
    _runScoringEngine();

    return { success: true, rowId };
  } catch(err) { return { error: err.message }; }
}

// ============================================================
// UPDATE TASK — writes Status + Remarks, auto-sets Actual Date
// Also triggers scoring so Dashboard stays current
// ============================================================
function updateTask(taskData, callerInfo) {
  try {
    const sheet   = ensureTasksSheet();
    const allData = sheet.getDataRange().getValues();

    // Use updatedByName from taskData (passed by frontend) or callerInfo
    const updatedBy = (taskData.updatedByName || '')
                   || (callerInfo && callerInfo.name ? callerInfo.name : '')
                   || (taskData.updatedByEmail || '')
                   || 'Unknown';

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][COL.ROW_ID].toString() !== taskData.rowId.toString()) continue;

      const rowNum       = i + 1;
      const isCompleting = (taskData.status||'').toLowerCase() === 'completed';
      let   actualDate   = allData[i][COL.ACTUAL_DATE];
      if (isCompleting && !actualDate) actualDate = _fmt(new Date());

      sheet.getRange(rowNum, COL.STATUS+1).setValue(taskData.status||allData[i][COL.STATUS]);
      sheet.getRange(rowNum, COL.REMARKS+1).setValue(
        taskData.remarks !== undefined ? taskData.remarks : allData[i][COL.REMARKS]
      );
      if (actualDate) sheet.getRange(rowNum, COL.ACTUAL_DATE+1).setValue(actualDate);

      // Notify assigner
      const assignerData = _getPersonData(allData[i][COL.ASSIGNED_BY]);
      const updaterEmail = (taskData.updatedByEmail||'').toLowerCase();
      if (assignerData.email && assignerData.email.toLowerCase() !== updaterEmail) {
        _emailUpdate(allData[i], taskData, updatedBy, actualDate);
      }
      if (assignerData.phone) {
        const updaterPhone = _getPersonData(updatedBy).phone;
        if (assignerData.phone !== updaterPhone) {
          _sendNotification(assignerData.phone,
            'Task Updated!\nTask: ' + allData[i][COL.TASK] +
            '\nBy: ' + updatedBy +
            '\nStatus: ' + taskData.status +
            '\nRemarks: ' + (taskData.remarks||'') +
            (actualDate ? '\nCompleted: ' + actualDate : ''));
        }
      }

      // Auto-refresh scoring & dashboard after status change
      _runScoringEngine();

      return { success: true, isCompleted: isCompleting };
    }
    return { error: 'Task not found' };
  } catch(err) { return { error: err.message }; }
}

// ============================================================
// GET TASKS
// filter: 'pending' | 'all'
// Director/Owner → all rows; Employee → only their email
// ============================================================
function getTasks(filter, callerInfo) {
  try {
    const sheet   = ensureTasksSheet();

    // callerInfo is passed from frontend after login — contains name, email, role
    // This avoids relying on Session.getEffectiveUser() which always returns
    // the script owner when deployed as "Execute as: Me"
    const user = callerInfo && callerInfo.name
      ? callerInfo
      : getCurrentUser();
    if (!user || user.error) return { error: (user && user.error) || 'Auth error' };

    const allData = sheet.getDataRange().getValues();
    if (allData.length <= 1) return { tasks:[], user };

    const isLeader      = (user.role||'').toLowerCase() === 'director'
                       || (user.role||'').toLowerCase() === 'owner';
    const userEmail     = (user.email||'').toString().trim().toLowerCase();
    const userNameLower = (user.name||'').toString().trim().toLowerCase();
    const tasks         = [];

    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      if (!row[COL.TASK]) continue;
      const taskEmail  = (row[COL.EMAIL]      ||'').toString().trim().toLowerCase();
      const assignedBy = (row[COL.ASSIGNED_BY]||'').toString().trim().toLowerCase();

      // Directors/owners see ALL tasks
      // Employees:
      //   Pending tab → only tasks assigned TO them (they need to act on these)
      //   All Tasks tab → tasks assigned TO them OR BY them (full history)
      if (!isLeader) {
        const isAssignee = userEmail && taskEmail === userEmail;
        const isAssigner = userNameLower && assignedBy === userNameLower;
        if (filter === 'pending') {
          // Pending: ONLY show tasks where THIS user is the assignee
          if (!isAssignee) continue;
        } else {
          // All Tasks: show tasks assigned to OR assigned by this user
          if (!isAssignee && !isAssigner) continue;
        }
      }
      const status = (row[COL.STATUS]||'').toString().trim();
      if (filter === 'pending' && status.toLowerCase() === 'completed') continue;
      tasks.push({
        rowId      : row[COL.ROW_ID],
        date       : _fmt(row[COL.DATE]),
        task       : row[COL.TASK],
        department : row[COL.DEPARTMENT],
        assignedTo : row[COL.ASSIGNED_TO],
        email      : row[COL.EMAIL],
        plannedWhen: _fmt(row[COL.PLANNED_WHEN]),
        actualDate : _fmt(row[COL.ACTUAL_DATE]),
        status,
        remarks    : row[COL.REMARKS],
        assignedBy : row[COL.ASSIGNED_BY]
      });
    }
    return { tasks, user };
  } catch(err) { return { error: err.message }; }
}

// ============================================================
// DASHBOARD DATA — for web app dashboard tab
// ============================================================
function getDashboardData(callerInfo) {
  try {
    const sheet   = ensureTasksSheet();

    const user = callerInfo && callerInfo.name
      ? callerInfo
      : getCurrentUser();
    if (!user || user.error) return { error: (user && user.error) || 'Auth error' };

    const allData = sheet.getDataRange().getValues();
    if (allData.length <= 1) return { byEmployee:[], byDept:[], totals:{onTime:0,late:0,overdue:0,pending:0}, user };

    const isLeader  = (user.role||'').toLowerCase() === 'director'
                   || (user.role||'').toLowerCase() === 'owner';
    const userEmail = (user.email||'').toString().trim().toLowerCase();
    const today    = new Date(); today.setHours(0,0,0,0);
    const empMap   = {}, deptMap = {};

    function blank(name) { return { name, onTime:0, late:0, overdue:0, pending:0 }; }

    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      if (!row[COL.TASK]) continue;
      const taskEmail = (row[COL.EMAIL]||'').toString().trim().toLowerCase();
      if (!isLeader && taskEmail !== userEmail) continue;

      const empName  = (row[COL.ASSIGNED_TO]||'').toString().trim() || 'Unknown';
      const deptName = (row[COL.DEPARTMENT] ||'').toString().trim() || 'Unknown';
      const status   = (row[COL.STATUS]     ||'').toString().trim().toLowerCase();
      const planned  = _parseDate(row[COL.PLANNED_WHEN]);
      const actual   = _parseDate(row[COL.ACTUAL_DATE]);

      if (!empMap[empName])   empMap[empName]   = blank(empName);
      if (!deptMap[deptName]) deptMap[deptName] = blank(deptName);

      const isCompleted = status === 'completed';
      let bucket;
      if (isCompleted) {
        bucket = (actual && planned && actual > planned) ? 'late' : 'onTime';
      } else {
        bucket = (planned && today > planned) ? 'overdue' : 'pending';
      }
      empMap[empName][bucket]++;
      deptMap[deptName][bucket]++;
    }

    const byEmployee = Object.values(empMap).sort((a,b) => a.name.localeCompare(b.name));
    const byDept     = Object.values(deptMap).sort((a,b) => a.name.localeCompare(b.name));
    const totals     = byEmployee.reduce((acc,e) => {
      acc.onTime+=e.onTime; acc.late+=e.late; acc.overdue+=e.overdue; acc.pending+=e.pending;
      return acc;
    }, {onTime:0,late:0,overdue:0,pending:0});

    return { byEmployee, byDept, totals, user };
  } catch(err) { return { error: err.message }; }
}

// ============================================================
// ═══════════════════════════════════════════════════════════
//   SCORING ENGINE & DASHBOARD BUILDER
//   Runs automatically after every sendTask / updateTask
//   Also runs on the daily time-trigger at 7 AM
// ═══════════════════════════════════════════════════════════
// ============================================================

// ── Scoring column indices in Tasks sheet ────────────────────
// L=11 (Delay), M=12 (Score), N=13 (Week) — already in COL above

// ── Called after sendTask and updateTask automatically ───────
function _runScoringEngine() {
  try {
    runScoringAndDashboard();
  } catch(e) {
    Logger.log('Scoring engine error: ' + e.message);
  }
}

// ── Main public function — also called by time trigger ───────
function runScoringAndDashboard() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TASKS_SHEET);
  if (!sheet) { Logger.log('Tasks sheet not found'); return; }

  // Ensure L, M, N headers
  const hdrRange = sheet.getRange(1, COL.DELAY+1, 1, 3);
  if (hdrRange.getValues()[0][0] !== 'Delay (Days)') {
    hdrRange.setValues([['Delay (Days)', 'Performance Score', 'Week Number']]);
    hdrRange.setFontWeight('bold').setBackground('#e8eaf6');
  }

  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const today  = new Date(); today.setHours(0,0,0,0);
  const delays = [], scores = [], weeks = [];

  for (let i = 1; i < data.length; i++) {
    const row     = data[i];
    const planned = _parseDate(row[COL.PLANNED_WHEN]);
    const actual  = _parseDate(row[COL.ACTUAL_DATE]);
    const status  = (row[COL.STATUS]||'').toString().trim().toLowerCase();
    const hasTask = (row[COL.TASK]  ||'').toString().trim() !== '';

    let delay = '', score = '', week = '';

    if (hasTask && planned) {
      week = _weekLabel(planned);
      if (status === 'completed') {
        if (actual) {
          const diff = Math.round((actual - planned) / 86400000);
          delay = diff > 0 ? diff : 0;
          score = diff <= 0 ? 'GREEN' : 'YELLOW';
        } else {
          score = 'GREEN';
        }
      } else if (today > planned) {
        delay = Math.round((today - planned) / 86400000);
        score = 'RED';
      } else {
        score = 'PENDING';
      }
    }

    delays.push([delay]);
    scores.push([score]);
    weeks.push([week]);
  }

  const n = delays.length;
  sheet.getRange(2, COL.DELAY+1, n, 1).setValues(delays);
  sheet.getRange(2, COL.SCORE+1, n, 1).setValues(scores);
  sheet.getRange(2, COL.WEEK +1, n, 1).setValues(weeks);

  _applyScoreFormatting(sheet, COL.SCORE+1, n+1);
  _buildDashboard(ss, data, delays, scores, weeks);

  Logger.log('Scoring & Dashboard updated at ' + new Date());
}

// ── Conditional formatting for Score column ──────────────────
function _applyScoreFormatting(sheet, colIndex, lastRow) {
  if (lastRow < 2) return;
  const range    = sheet.getRange(2, colIndex, lastRow-1, 1);
  const existing = sheet.getConditionalFormatRules()
    .filter(r => r.getRanges()[0].getColumn() !== colIndex);

  function mk(text, bg, fc) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(text).setBackground(bg).setFontColor(fc)
      .setRanges([range]).build();
  }
  existing.push(mk('GREEN',   '#b7e1cd', '#0E3221'));
  existing.push(mk('YELLOW',  '#fce8b2', '#7A4800'));
  existing.push(mk('RED',     '#f4cccc', '#8A1A1A'));
  existing.push(mk('PENDING', '#e8eaf6', '#3c4043'));
  sheet.setConditionalFormatRules(existing);
}

// ── Week number label ────────────────────────────────────────
function _weekLabel(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
  const yr = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return 'Week ' + Math.ceil((((d-yr)/86400000)+1)/7);
}

// ── Build the Dashboard sheet ────────────────────────────────
function _buildDashboard(ss, rawData, delays, scores, weeks) {
  let dash = ss.getSheetByName(DASH_SHEET);
  if (!dash) {
    dash = ss.insertSheet(DASH_SHEET);
    ss.setActiveSheet(dash);
    ss.moveActiveSheet(2);
  }
  dash.clearContents();
  dash.clearFormats();
  dash.clearConditionalFormatRules();

  // Aggregate stats
  let tGreen=0, tYellow=0, tRed=0, tPending=0, totalDelay=0, delayCount=0;
  const empMap={}, deptMap={}, weekMap={};

  function bk() { return {green:0,yellow:0,red:0,pending:0,delay:0,dc:0}; }

  for (let i = 1; i < rawData.length; i++) {
    const row  = rawData[i];
    if (!(row[COL.TASK]||'').toString().trim()) continue;

    const emp   = (row[COL.ASSIGNED_TO]||'Unknown').toString().trim();
    const dept  = (row[COL.DEPARTMENT] ||'Unknown').toString().trim();
    const sc    = (scores[i-1]?scores[i-1][0]:'').toString();
    const dl    = Number(delays[i-1]?delays[i-1][0]:0)||0;
    const wk    = (weeks[i-1] ?weeks[i-1][0] :'').toString();

    if (!empMap[emp])  empMap[emp]  = bk();
    if (!deptMap[dept]) deptMap[dept] = bk();
    if (wk && !weekMap[wk]) weekMap[wk] = bk();

    function add(m) {
      if (sc==='GREEN') m.green++;
      else if (sc==='YELLOW') m.yellow++;
      else if (sc==='RED')    m.red++;
      else if (sc==='PENDING') m.pending++;
      if (dl>0){ m.delay+=dl; m.dc++; }
    }
    add(empMap[emp]); add(deptMap[dept]);
    if (wk) add(weekMap[wk]);

    if (sc==='GREEN') tGreen++;
    else if (sc==='YELLOW') tYellow++;
    else if (sc==='RED')    tRed++;
    else if (sc==='PENDING') tPending++;
    if (dl>0){ totalDelay+=dl; delayCount++; }
  }

  const grand    = tGreen+tYellow+tRed+tPending;
  const avgDelay = delayCount>0 ? Math.round(totalDelay/delayCount) : 0;

  // Column widths
  dash.setColumnWidth(1,170); dash.setColumnWidth(2,72);
  dash.setColumnWidth(3,72);  dash.setColumnWidth(4,72);
  dash.setColumnWidth(5,72);  dash.setColumnWidth(6,60);
  dash.setColumnWidth(7,90);  dash.setColumnWidth(8,75);

  const BDR  = SpreadsheetApp.BorderStyle.SOLID;
  const BDRC = '#cccccc';

  function sR(r,c,val,o){
    o=o||{};
    const cell=dash.getRange(r,c);
    if(val!==null&&val!==undefined) cell.setValue(val);
    if(o.bg)    cell.setBackground(o.bg);
    if(o.fc)    cell.setFontColor(o.fc);
    if(o.bold)  cell.setFontWeight('bold');
    if(o.sz)    cell.setFontSize(o.sz);
    if(o.align) cell.setHorizontalAlignment(o.align);
    if(o.wrap)  cell.setWrap(true);
    if(o.bdr)   cell.setBorder(true,true,true,true,false,false,BDRC,BDR);
    if(o.italic) cell.setFontStyle('italic');
    return cell;
  }
  function mg(r,c1,c2,val,o){
    dash.getRange(r,c1,1,c2-c1+1).merge(); sR(r,c1,val,o);
    if(o&&o.bg) for(let c=c1+1;c<=c2;c++) dash.getRange(r,c).setBackground(o.bg);
  }
  function secH(r,txt){
    mg(r,1,8,'  '+txt,{bg:'#1B4B35',fc:'#ffffff',bold:true,sz:11});
    dash.setRowHeight(r,26);
  }
  function tblH(r,cols){
    cols.forEach((h,i)=>sR(r,i+1,h,{bg:'#E4EFEA',fc:'#0E3221',bold:true,sz:10,align:'center',bdr:true}));
    dash.setRowHeight(r,22);
  }
  function nC(r,c,val,bg){
    const disp=(val===0||val==='')?'—':val;
    sR(r,c,disp,{bg:bg||'#ffffff',align:'center',bdr:true,
      fc:(val===0||val==='')?'#aaaaaa':'#202124'});
  }
  function rBg(r){ return r%2===0?'#f8f9fa':'#ffffff'; }

  // ── Title ──
  mg(1,1,8,'TASK PERFORMANCE DASHBOARD',
    {bg:'#1B4B35',fc:'#ffffff',bold:true,sz:18,align:'center'});
  dash.setRowHeight(1,44);

  mg(2,1,8,
    'Last updated: '+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy  HH:mm'),
    {bg:'#f1f3f4',fc:'#5f6368',sz:9,align:'right',italic:true});
  dash.setRowHeight(2,16);
  dash.setRowHeight(3,10);

  // ── Summary cards ──
  secH(4,'OVERALL SUMMARY');
  const cds=[
    {lbl:'Total Tasks',       val:grand,    bg:'#E8F5E9',fc:'#1B4B35'},
    {lbl:'On Time (GREEN)',   val:tGreen,   bg:'#b7e1cd',fc:'#0E3221'},
    {lbl:'Late (YELLOW)',     val:tYellow,  bg:'#fce8b2',fc:'#7A4800'},
    {lbl:'Overdue (RED)',     val:tRed,     bg:'#f4cccc',fc:'#8A1A1A'},
    {lbl:'Pending',           val:tPending, bg:'#e8eaf6',fc:'#3c4043'},
    {lbl:'Avg Delay (Days)',  val:avgDelay||'—', bg:'#fff3e0',fc:'#7A4800'},
  ];
  cds.forEach((c,i)=>sR(5,i+1,c.lbl,{bg:c.bg,fc:c.fc,bold:true,sz:9,align:'center',wrap:true,bdr:true}));
  dash.setRowHeight(5,30);
  cds.forEach((c,i)=>sR(6,i+1,c.val,{bg:c.bg,fc:c.fc,bold:true,sz:22,align:'center',bdr:true}));
  dash.setRowHeight(6,40);
  cds.forEach((c,i)=>{
    const pct=(i>=1&&i<=4&&grand>0)?Math.round((c.val/grand)*100)+'%':'';
    sR(7,i+1,pct,{bg:c.bg,fc:c.fc,sz:9,align:'center',bdr:true,italic:true});
  });
  // % Score card (overall on-time %)
  const overallPct = grand>0 ? Math.round(((tGreen)/(tGreen+tYellow+tRed))*100||0) : 0;
  sR(5,7,'% On-Time Score',{bg:'#E8F5E9',fc:'#1B4B35',bold:true,sz:9,align:'center',wrap:true,bdr:true});
  sR(6,7,overallPct+'%',{bg:'#E8F5E9',fc:'#1B4B35',bold:true,sz:22,align:'center',bdr:true});
  sR(7,7,'of completed tasks',{bg:'#E8F5E9',fc:'#5a5f6a',sz:9,align:'center',bdr:true,italic:true});
  dash.setRowHeight(7,18);
  dash.setRowHeight(8,10);

  // ── By Employee ──
  let r=9;
  secH(r,'PERFORMANCE BY EMPLOYEE'); r++;
  tblH(r,['Employee','GREEN','YELLOW','RED','PENDING','Total','Avg Delay','% Score']); r++;

  Object.entries(empMap).sort((a,b)=>a[0].localeCompare(b[0])).forEach(([name,m])=>{
    const tot=m.green+m.yellow+m.red+m.pending;
    const avg=m.dc>0?Math.round(m.delay/m.dc):0;
    const bg=rBg(r);
    const completed=m.green+m.yellow;
    const pctScore=completed>0?Math.round((m.green/completed)*100):0;
    const pctBg=pctScore>=80?'#b7e1cd':pctScore>=50?'#fce8b2':'#f4cccc';
    const pctFc=pctScore>=80?'#0E3221':pctScore>=50?'#7A4800':'#8A1A1A';
    sR(r,1,name,{bg,bold:true,wrap:true,bdr:true});
    nC(r,2,m.green,   m.green  >0?'#b7e1cd':bg);
    nC(r,3,m.yellow,  m.yellow >0?'#fce8b2':bg);
    nC(r,4,m.red,     m.red    >0?'#f4cccc':bg);
    nC(r,5,m.pending, m.pending>0?'#e8eaf6':bg);
    sR(r,6,tot,{bg,bold:true,align:'center',bdr:true});
    sR(r,7,avg>0?avg+'d':'—',{bg:avg>0?'#fce8b2':bg,fc:avg>0?'#7A4800':'#aaaaaa',align:'center',bdr:true});
    sR(r,8,completed>0?pctScore+'%':'—',{bg:completed>0?pctBg:bg,fc:completed>0?pctFc:'#aaaaaa',bold:completed>0,align:'center',bdr:true});
    dash.setRowHeight(r,22); r++;
  });
  // Totals row
  const totalCompleted=tGreen+tYellow;
  const totalPct=totalCompleted>0?Math.round((tGreen/totalCompleted)*100):0;
  sR(r,1,'TOTAL',{bg:'#E4EFEA',fc:'#0E3221',bold:true,bdr:true});
  sR(r,2,tGreen,  {bg:'#b7e1cd',fc:'#0E3221',bold:true,align:'center',bdr:true});
  sR(r,3,tYellow, {bg:'#fce8b2',fc:'#7A4800',bold:true,align:'center',bdr:true});
  sR(r,4,tRed,    {bg:'#f4cccc',fc:'#8A1A1A',bold:true,align:'center',bdr:true});
  sR(r,5,tPending,{bg:'#e8eaf6',fc:'#3c4043',bold:true,align:'center',bdr:true});
  sR(r,6,grand,   {bg:'#E4EFEA',fc:'#0E3221',bold:true,align:'center',bdr:true});
  sR(r,7,avgDelay>0?avgDelay+'d':'—',{bg:'#E4EFEA',fc:'#0E3221',bold:true,align:'center',bdr:true});
  sR(r,8,totalCompleted>0?totalPct+'%':'—',{bg:'#E4EFEA',fc:'#0E3221',bold:true,align:'center',bdr:true});
  dash.setRowHeight(r,24); r++;

  // ── By Department ──
  dash.setRowHeight(r,10); r++;
  secH(r,'PERFORMANCE BY DEPARTMENT'); r++;
  tblH(r,['Department','GREEN','YELLOW','RED','PENDING','Total','Avg Delay','% Score']); r++;

  Object.entries(deptMap).sort((a,b)=>a[0].localeCompare(b[0])).forEach(([name,m])=>{
    const tot=m.green+m.yellow+m.red+m.pending;
    const avg=m.dc>0?Math.round(m.delay/m.dc):0;
    const bg=rBg(r);
    const dc=m.green+m.yellow;
    const dp=dc>0?Math.round((m.green/dc)*100):0;
    const dpBg=dp>=80?'#b7e1cd':dp>=50?'#fce8b2':'#f4cccc';
    const dpFc=dp>=80?'#0E3221':dp>=50?'#7A4800':'#8A1A1A';
    sR(r,1,name,{bg,bold:true,wrap:true,bdr:true});
    nC(r,2,m.green,   m.green  >0?'#b7e1cd':bg);
    nC(r,3,m.yellow,  m.yellow >0?'#fce8b2':bg);
    nC(r,4,m.red,     m.red    >0?'#f4cccc':bg);
    nC(r,5,m.pending, m.pending>0?'#e8eaf6':bg);
    sR(r,6,tot,{bg,bold:true,align:'center',bdr:true});
    sR(r,7,avg>0?avg+'d':'—',{bg:avg>0?'#fce8b2':bg,fc:avg>0?'#7A4800':'#aaaaaa',align:'center',bdr:true});
    sR(r,8,dc>0?dp+'%':'—',{bg:dc>0?dpBg:bg,fc:dc>0?dpFc:'#aaaaaa',bold:dc>0,align:'center',bdr:true});
    dash.setRowHeight(r,22); r++;
  });

  // ── By Week ──
  dash.setRowHeight(r,10); r++;
  secH(r,'PERFORMANCE BY WEEK (based on Planned When date)'); r++;
  tblH(r,['Week','GREEN','YELLOW','RED','PENDING','Total','Avg Delay','% Score']); r++;

  Object.entries(weekMap)
    .sort((a,b)=>parseInt(a[0].replace('Week ',''))-parseInt(b[0].replace('Week ','')))
    .forEach(([name,m])=>{
      const tot=m.green+m.yellow+m.red+m.pending;
      const avg=m.dc>0?Math.round(m.delay/m.dc):0;
      const bg=rBg(r);
      const wc=m.green+m.yellow;
      const wp=wc>0?Math.round((m.green/wc)*100):0;
      const wpBg=wp>=80?'#b7e1cd':wp>=50?'#fce8b2':'#f4cccc';
      const wpFc=wp>=80?'#0E3221':wp>=50?'#7A4800':'#8A1A1A';
      sR(r,1,name,{bg,wrap:true,bdr:true});
      nC(r,2,m.green,   m.green  >0?'#b7e1cd':bg);
      nC(r,3,m.yellow,  m.yellow >0?'#fce8b2':bg);
      nC(r,4,m.red,     m.red    >0?'#f4cccc':bg);
      nC(r,5,m.pending, m.pending>0?'#e8eaf6':bg);
      sR(r,6,tot,{bg,bold:true,align:'center',bdr:true});
      sR(r,7,avg>0?avg+'d':'—',{bg:avg>0?'#fce8b2':bg,fc:avg>0?'#7A4800':'#aaaaaa',align:'center',bdr:true});
      sR(r,8,wc>0?wp+'%':'—',{bg:wc>0?wpBg:bg,fc:wc>0?wpFc:'#aaaaaa',bold:wc>0,align:'center',bdr:true});
      dash.setRowHeight(r,22); r++;
    });

  // ── Legend ──
  dash.setRowHeight(r,10); r++;
  mg(r,1,8,'LEGEND',{bg:'#1B4B35',fc:'#ffffff',bold:true,sz:10});
  dash.setRowHeight(r,22); r++;
  [
    ['GREEN',  '#b7e1cd','#0E3221','Completed on or before the deadline'],
    ['YELLOW', '#fce8b2','#7A4800','Completed after the deadline (task done but late)'],
    ['RED',    '#f4cccc','#8A1A1A','Not completed and deadline has already passed (overdue)'],
    ['PENDING','#e8eaf6','#3c4043','Not completed but deadline has not passed yet'],
  ].forEach(([lbl,bg,fc,desc])=>{
    sR(r,1,lbl,{bg,fc,bold:true,sz:10,align:'center',bdr:true});
    mg(r,2,8,desc,{bg,fc,sz:10,bdr:true});
    dash.setRowHeight(r,20); r++;
  });

  // Freeze rows only (no column freeze — conflicts with merged cells)
  dash.setFrozenRows(11);
}

// ============================================================
// TIME-BASED TRIGGER SETUP
// Runs runScoringAndDashboard automatically every day at 7 AM
// Call setupDailyTrigger() ONCE from Apps Script to activate
// ============================================================
function setupDailyTrigger() {
  // Remove existing triggers first to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runScoringAndDashboard') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('runScoringAndDashboard')
    .timeBased().everyDays(1).atHour(7).create();
  Logger.log('Daily trigger set — runScoringAndDashboard will run every day at 7 AM.');
}

function removeDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runScoringAndDashboard') {
      ScriptApp.deleteTrigger(t);
    }
  });
  Logger.log('Daily trigger removed.');
}

// ============================================================
// NOTIFICATION DISPATCHER
// ============================================================
function _sendNotification(phone, message) {
  if (!phone || NOTIFY_MODE === 'none') return;
  try {
    if      (NOTIFY_MODE === 'sms_fast2sms')   _sendFast2SMS(phone, message);
    else if (NOTIFY_MODE === 'sms_twilio')      _sendTwilioSMS(phone, message);
    else if (NOTIFY_MODE === 'whatsapp_twilio') _sendTwilioWhatsApp(phone, message);
  } catch(e) { Logger.log('Notification error: ' + e.message); }
}

function _sendFast2SMS(phone, message) {
  const cleaned = phone.toString().replace(/\D/g,'').slice(-10);
  if (cleaned.length !== 10) { Logger.log('Fast2SMS: invalid number ' + phone); return; }
  const resp = UrlFetchApp.fetch('https://www.fast2sms.com/dev/bulkV2', {
    method:'post', headers:{authorization: FAST2SMS_API_KEY},
    payload:{route:'q', message, language:'english', flash:'0', numbers:cleaned},
    muteHttpExceptions:true
  });
  const r = JSON.parse(resp.getContentText());
  Logger.log(r.return === true ? 'Fast2SMS sent to '+cleaned : 'Fast2SMS error: '+JSON.stringify(r));
}

function _sendTwilioSMS(phone, message) {
  const to  = phone.startsWith('+') ? phone : '+91'+phone;
  const url = 'https://api.twilio.com/2010-04-01/Accounts/'+TWILIO_ACCOUNT_SID+'/Messages.json';
  const opts = {
    method:'post', contentType:'application/x-www-form-urlencoded',
    headers:{Authorization:'Basic '+Utilities.base64Encode(TWILIO_ACCOUNT_SID+':'+TWILIO_AUTH_TOKEN)},
    payload:'From='+encodeURIComponent(TWILIO_SMS_FROM)+'&To='+encodeURIComponent(to)+'&Body='+encodeURIComponent(message),
    muteHttpExceptions:true
  };
  const r = JSON.parse(UrlFetchApp.fetch(url, opts).getContentText());
  Logger.log(r.sid ? 'Twilio SMS sent '+to : 'Twilio SMS error: '+r.message);
}

function _sendTwilioWhatsApp(phone, message) {
  const to  = 'whatsapp:'+(phone.startsWith('+') ? phone : '+91'+phone);
  const url = 'https://api.twilio.com/2010-04-01/Accounts/'+TWILIO_ACCOUNT_SID+'/Messages.json';
  const opts = {
    method:'post', contentType:'application/x-www-form-urlencoded',
    headers:{Authorization:'Basic '+Utilities.base64Encode(TWILIO_ACCOUNT_SID+':'+TWILIO_AUTH_TOKEN)},
    payload:'From='+encodeURIComponent(TWILIO_WA_FROM)+'&To='+encodeURIComponent(to)+'&Body='+encodeURIComponent(message),
    muteHttpExceptions:true
  };
  const r = JSON.parse(UrlFetchApp.fetch(url, opts).getContentText());
  Logger.log(r.sid ? 'WhatsApp sent '+to : 'WhatsApp error: '+r.message);
}

// ============================================================
// PRIVATE HELPERS
// ============================================================
function _fmt(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return v.toString();
}

function _parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    const d = new Date(val); d.setHours(0,0,0,0);
    return isNaN(d.getTime()) ? null : d;
  }
  const s = val.toString().trim();
  const parts = s.split('/');
  if (parts.length === 3) {
    const d = new Date(+parts[2], +parts[1]-1, +parts[0]);
    d.setHours(0,0,0,0);
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(s); d.setHours(0,0,0,0);
  return isNaN(d.getTime()) ? null : d;
}

function _getPersonData(name) {
  try {
    const data = SpreadsheetApp.openById(SPREADSHEET_ID)
                   .getSheetByName(DATA_SHEET).getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((data[i][1]||'').toString().trim() === name.toString().trim()) {
        return { email:(data[i][2]||'').toString().trim(), phone:(data[i][4]||'').toString().trim() };
      }
    }
  } catch(e) {}
  return { email:'', phone:'' };
}

function _emailOfName(name) { return _getPersonData(name).email; }

function _emailNew(t, by) {
  try {
    const subject = 'New Task Assigned: ' + t.task.substring(0, 60);
    const plain   = 'Hello ' + t.assignedTo + ',\n\nA new task has been assigned to you.\n\n'
      + 'Task: ' + t.task + '\nDepartment: ' + t.department + '\nAssigned By: ' + by
      + '\nDate: ' + t.date + '\nDeadline: ' + t.plannedWhen
      + '\n\nLog in to Task Manager to view and update this task.\n\nRegards,\nTask Manager';
    const html = '<div style="font-family:Arial,sans-serif;max-width:600px;padding:20px">'
      + '<div style="background:#1B4B35;padding:16px 20px;border-radius:8px 8px 0 0">'
      + '<h2 style="color:#fff;margin:0;font-size:18px">Task Manager</h2></div>'
      + '<div style="background:#fff;border:1px solid #e2e5ea;border-top:none;padding:24px;border-radius:0 0 8px 8px">'
      + '<h3 style="color:#1a1c20;margin:0 0 16px">New Task Assigned</h3>'
      + '<p style="color:#5a5f6a;margin:0 0 16px">Hello <strong>' + t.assignedTo + '</strong>,</p>'
      + '<table style="width:100%;border-collapse:collapse;margin-bottom:20px">'
      + '<tr><td style="padding:8px 12px;background:#f6f7f9;border:1px solid #e2e5ea;font-weight:700;width:35%;color:#5a5f6a">Task</td>'
      + '<td style="padding:8px 12px;border:1px solid #e2e5ea">' + t.task + '</td></tr>'
      + '<tr><td style="padding:8px 12px;background:#f6f7f9;border:1px solid #e2e5ea;font-weight:700;color:#5a5f6a">Department</td>'
      + '<td style="padding:8px 12px;border:1px solid #e2e5ea">' + t.department + '</td></tr>'
      + '<tr><td style="padding:8px 12px;background:#f6f7f9;border:1px solid #e2e5ea;font-weight:700;color:#5a5f6a">Assigned By</td>'
      + '<td style="padding:8px 12px;border:1px solid #e2e5ea">' + by + '</td></tr>'
      + '<tr><td style="padding:8px 12px;background:#f6f7f9;border:1px solid #e2e5ea;font-weight:700;color:#5a5f6a">Date Assigned</td>'
      + '<td style="padding:8px 12px;border:1px solid #e2e5ea">' + t.date + '</td></tr>'
      + '<tr><td style="padding:8px 12px;background:#f6f7f9;border:1px solid #e2e5ea;font-weight:700;color:#5a5f6a">Deadline</td>'
      + '<td style="padding:8px 12px;border:1px solid #e2e5ea;color:#8A1A1A;font-weight:700">' + t.plannedWhen + '</td></tr>'
      + '</table>'
      + '<p style="color:#5a5f6a;margin:0">Please log in to Task Manager to view and update this task.</p>'
      + '</div></div>';
    GmailApp.sendEmail(t.email, subject, plain, { htmlBody: html });
  } catch(e) { Logger.log('Email error: ' + e.message); }
}

function _emailUpdate(orig, upd, by, actual) {
  try {
    const to = _emailOfName(orig[COL.ASSIGNED_BY]);
    if (!to) return;
    GmailApp.sendEmail(to,
      'Task Updated: '+orig[COL.TASK].toString().substring(0,60),
      'Hello,\n\nA task has been updated.\n\n'
      +'Task: '+orig[COL.TASK]+'\nUpdated By: '+by+'\nNew Status: '+upd.status
      +'\nRemarks: '+(upd.remarks||'—')+'\n'+(actual?'Completed On: '+actual+'\n':'')
      +'\nLog in to Task Manager to view full details.\n\nRegards,\nTask Manager');
  } catch(e) { Logger.log('Update email error: '+e.message); }
}


// ============================================================
// END-OF-DAY SUMMARY EMAIL
// Sends each employee a list of tasks assigned to them today
// Called automatically by time-trigger at 5 PM daily
// Also sends a summary to all directors
// ============================================================
function sendEndOfDaySummary() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TASKS_SHEET);
  if (!sheet) return;

  const data  = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const tz    = Session.getScriptTimeZone();

  // Collect tasks added today grouped by assignee email
  const byEmployee = {};
  for (let i = 1; i < data.length; i++) {
    const row      = data[i];
    if (!row[COL.TASK]) continue;
    const taskDate = _fmt(row[COL.DATE]);
    if (taskDate !== today) continue;

    const email = (row[COL.EMAIL]||'').toString().trim();
    if (!email) continue;
    if (!byEmployee[email]) byEmployee[email] = { name: row[COL.ASSIGNED_TO], tasks: [] };
    byEmployee[email].tasks.push({
      task      : row[COL.TASK],
      dept      : row[COL.DEPARTMENT],
      deadline  : _fmt(row[COL.PLANNED_WHEN]),
      status    : row[COL.STATUS],
      assignedBy: row[COL.ASSIGNED_BY]
    });
  }

  if (Object.keys(byEmployee).length === 0) {
    Logger.log('End-of-day summary: no tasks assigned today.');
    return;
  }

  // Send individual summary to each employee
  Object.entries(byEmployee).forEach(function([email, empData]) {
    const rows = empData.tasks.map(t =>
      '<tr><td style="padding:8px 12px;border:1px solid #e2e5ea">' + t.task + '</td>'
      + '<td style="padding:8px 12px;border:1px solid #e2e5ea">' + t.dept + '</td>'
      + '<td style="padding:8px 12px;border:1px solid #e2e5ea;color:#8A1A1A;font-weight:700">' + t.deadline + '</td>'
      + '<td style="padding:8px 12px;border:1px solid #e2e5ea">' + t.assignedBy + '</td>'
      + '<td style="padding:8px 12px;border:1px solid #e2e5ea">' + t.status + '</td></tr>'
    ).join('');

    const html = '<div style="font-family:Arial,sans-serif;max-width:620px;padding:20px">'
      + '<div style="background:#1B4B35;padding:16px 20px;border-radius:8px 8px 0 0">'
      + '<h2 style="color:#fff;margin:0;font-size:18px">Task Manager — Daily Summary</h2>'
      + '<p style="color:#a8d5c2;margin:4px 0 0;font-size:13px">' + today + '</p></div>'
      + '<div style="background:#fff;border:1px solid #e2e5ea;border-top:none;padding:24px;border-radius:0 0 8px 8px">'
      + '<p style="color:#1a1c20;margin:0 0 16px">Hello <strong>' + empData.name + '</strong>,</p>'
      + '<p style="color:#5a5f6a;margin:0 0 16px">Here are the tasks assigned to you today (' + today + '):</p>'
      + '<table style="width:100%;border-collapse:collapse;font-size:13px">'
      + '<tr style="background:#E4EFEA"><th style="padding:8px 12px;border:1px solid #e2e5ea;text-align:left;color:#0E3221">Task</th>'
      + '<th style="padding:8px 12px;border:1px solid #e2e5ea;text-align:left;color:#0E3221">Dept</th>'
      + '<th style="padding:8px 12px;border:1px solid #e2e5ea;text-align:left;color:#0E3221">Deadline</th>'
      + '<th style="padding:8px 12px;border:1px solid #e2e5ea;text-align:left;color:#0E3221">Assigned By</th>'
      + '<th style="padding:8px 12px;border:1px solid #e2e5ea;text-align:left;color:#0E3221">Status</th></tr>'
      + rows
      + '</table>'
      + '<p style="color:#5a5f6a;margin:16px 0 0;font-size:13px">Total tasks assigned today: <strong>' + empData.tasks.length + '</strong></p>'
      + '</div></div>';

    const plain = 'Daily Task Summary for ' + empData.name + ' (' + today + '):\n\n'
      + empData.tasks.map((t,i) => (i+1)+'. '+t.task+' | Deadline: '+t.deadline+' | By: '+t.assignedBy).join('\n');

    try {
      GmailApp.sendEmail(email, 'Your Task Summary for ' + today + ' — Task Manager', plain, { htmlBody: html });
    } catch(e) { Logger.log('EOD email error for ' + email + ': ' + e.message); }
  });

  // Send full summary to directors
  const dataSheet = ss.getSheetByName(DATA_SHEET);
  const dData     = dataSheet.getDataRange().getValues();
  const directors = dData.slice(1).filter(r => {
    const role = (r[3]||'').toString().trim().toLowerCase();
    return (role === 'director' || role === 'owner') && r[2];
  });

  if (directors.length === 0) return;

  // Build full summary table for directors
  let allRows = '';
  Object.entries(byEmployee).forEach(function([email, empData]) {
    empData.tasks.forEach(function(t) {
      allRows += '<tr><td style="padding:7px 10px;border:1px solid #e2e5ea">' + empData.name + '</td>'
        + '<td style="padding:7px 10px;border:1px solid #e2e5ea">' + t.task + '</td>'
        + '<td style="padding:7px 10px;border:1px solid #e2e5ea">' + t.dept + '</td>'
        + '<td style="padding:7px 10px;border:1px solid #e2e5ea;color:#8A1A1A;font-weight:700">' + t.deadline + '</td>'
        + '<td style="padding:7px 10px;border:1px solid #e2e5ea">' + t.assignedBy + '</td>'
        + '<td style="padding:7px 10px;border:1px solid #e2e5ea">' + t.status + '</td></tr>';
    });
  });

  const totalTasks = Object.values(byEmployee).reduce((s, e) => s + e.tasks.length, 0);

  const dirHtml = '<div style="font-family:Arial,sans-serif;max-width:700px;padding:20px">'
    + '<div style="background:#1D3F6E;padding:16px 20px;border-radius:8px 8px 0 0">'
    + '<h2 style="color:#fff;margin:0;font-size:18px">Team Task Summary — ' + today + '</h2>'
    + '<p style="color:#a8c4e8;margin:4px 0 0;font-size:13px">' + totalTasks + ' tasks assigned today</p></div>'
    + '<div style="background:#fff;border:1px solid #e2e5ea;border-top:none;padding:24px;border-radius:0 0 8px 8px">'
    + '<table style="width:100%;border-collapse:collapse;font-size:13px">'
    + '<tr style="background:#D6E4F0"><th style="padding:8px 10px;border:1px solid #e2e5ea;text-align:left;color:#1D3F6E">Employee</th>'
    + '<th style="padding:8px 10px;border:1px solid #e2e5ea;text-align:left;color:#1D3F6E">Task</th>'
    + '<th style="padding:8px 10px;border:1px solid #e2e5ea;text-align:left;color:#1D3F6E">Dept</th>'
    + '<th style="padding:8px 10px;border:1px solid #e2e5ea;text-align:left;color:#1D3F6E">Deadline</th>'
    + '<th style="padding:8px 10px;border:1px solid #e2e5ea;text-align:left;color:#1D3F6E">Assigned By</th>'
    + '<th style="padding:8px 10px;border:1px solid #e2e5ea;text-align:left;color:#1D3F6E">Status</th></tr>'
    + allRows + '</table></div></div>';

  const dirPlain = 'Team Task Summary for ' + today + ':\nTotal tasks assigned today: ' + totalTasks;

  directors.forEach(function(dir) {
    try {
      GmailApp.sendEmail(dir[2], 'Team Task Summary for ' + today + ' — Task Manager', dirPlain, { htmlBody: dirHtml });
    } catch(e) { Logger.log('Director EOD email error: ' + e.message); }
  });

  Logger.log('End-of-day summary sent to ' + Object.keys(byEmployee).length + ' employees and ' + directors.length + ' directors.');
}

// Set up end-of-day trigger at 5 PM
function setupEndOfDayTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendEndOfDaySummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendEndOfDaySummary').timeBased().everyDays(1).atHour(17).create();
  Logger.log('End-of-day trigger set for 5 PM daily.');
}


// ============================================================
// USER ID / PASSWORD LOGIN (Change 5)
// Reads from "Users" sheet: A=UserID, B=Password, C=Name, D=Role
// Called from the login screen before the main app loads
// ============================================================
const USERS_SHEET = 'Users';

function loginWithCredentials(userId, password) {
  try {
    if (!userId || !password) return { error: 'Please enter both User ID and Password.' };

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(USERS_SHEET);
    if (!sheet) return { error: 'Users sheet not found. Please set it up.' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowId   = (data[i][0]||'').toString().trim();
      const rowPass = (data[i][1]||'').toString().trim();
      const rowName = (data[i][2]||'').toString().trim();
      const rowRole = (data[i][3]||'employee').toString().trim().toLowerCase();

      if (rowId.toLowerCase() === userId.toLowerCase() && rowPass === password) {
        // Match found — now get full profile from Data sheet by name
        const dataSheet = ss.getSheetByName(DATA_SHEET);
        const dData     = dataSheet.getDataRange().getValues();
        for (let j = 1; j < dData.length; j++) {
          if ((dData[j][1]||'').toString().trim().toLowerCase() === rowName.toLowerCase()) {
            return {
              success   : true,
              name      : dData[j][1].toString().trim(),
              email     : (dData[j][2]||'').toString().trim(),
              department: (dData[j][0]||'').toString().trim(),
              role      : rowRole,
              phone     : (dData[j][4]||'').toString().trim()
            };
          }
        }
        // Name not in Data sheet — return basic info from Users sheet
        return {
          success: true,
          name   : rowName,
          email  : '',
          department: '',
          role   : rowRole,
          phone  : ''
        };
      }
    }
    return { error: 'Invalid User ID or Password. Please try again.' };
  } catch(err) { return { error: err.message }; }
}



// ============================================================
// FIX 1: DAILY SUMMARY EMAIL
// Sends each employee a list of ALL their tasks assigned today
// Schedule: run setupDailyEmailTrigger() once to activate
// ============================================================
function sendDailySummaryEmails() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(TASKS_SHEET);
    if (!sheet) return;

    const data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    const today     = new Date();
    const todayStr  = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const todayFull = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM dd, yyyy');

    // Group today's tasks by employee email
    const byEmail = {};
    for (let i = 1; i < data.length; i++) {
      const row    = data[i];
      const task   = (row[COL.TASK]  ||'').toString().trim();
      const email  = (row[COL.EMAIL] ||'').toString().trim();
      const name   = (row[COL.ASSIGNED_TO]||'').toString().trim();
      const taskDate = _fmt(row[COL.DATE]);

      if (!task || !email) continue;
      // Match today's date
      if (taskDate !== todayStr) continue;

      if (!byEmail[email]) byEmail[email] = { name, tasks: [] };
      byEmail[email].tasks.push({
        task     : task,
        dept     : (row[COL.DEPARTMENT] ||'').toString().trim(),
        deadline : _fmt(row[COL.PLANNED_WHEN]),
        status   : (row[COL.STATUS]||'Pending').toString().trim(),
        assignedBy: (row[COL.ASSIGNED_BY]||'').toString().trim()
      });
    }

    if (Object.keys(byEmail).length === 0) {
      Logger.log('Daily summary: no tasks assigned today.');
      return;
    }

    // Send one email per employee
    Object.entries(byEmail).forEach(([email, info]) => {
      let taskRows = '';
      info.tasks.forEach((t, idx) => {
        const bg = idx % 2 === 0 ? '#f8f9fa' : '#ffffff';
        taskRows += `<tr style="background:${bg}">
          <td style="padding:10px 12px;border-bottom:1px solid #e2e5ea">${idx+1}. ${t.task}</td>
          <td style="padding:10px 12px;border-bottom:1px solid #e2e5ea;white-space:nowrap">${t.dept}</td>
          <td style="padding:10px 12px;border-bottom:1px solid #e2e5ea;white-space:nowrap">${t.deadline}</td>
          <td style="padding:10px 12px;border-bottom:1px solid #e2e5ea">
            <span style="background:#e8eaf6;color:#3c4043;padding:2px 8px;border-radius:12px;font-size:12px;font-weight:600">${t.status}</span>
          </td>
          <td style="padding:10px 12px;border-bottom:1px solid #e2e5ea;color:#5a5f6a;font-size:12px">${t.assignedBy}</td>
        </tr>`;
      });

      const subject = `📋 Your Task Summary for ${todayFull} — ${info.tasks.length} task(s)`;
      const html = `
<div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto">
  <div style="background:#1B4B35;padding:18px 24px;border-radius:10px 10px 0 0">
    <h2 style="color:#fff;margin:0;font-size:20px">📋 Daily Task Summary</h2>
    <p style="color:#9FE1CB;margin:4px 0 0;font-size:13px">${todayFull}</p>
  </div>
  <div style="background:#fff;border:1px solid #e2e5ea;border-top:none;padding:24px;border-radius:0 0 10px 10px">
    <p style="color:#1A1C20;font-size:15px;margin:0 0 18px">Hello <strong>${info.name}</strong>,</p>
    <p style="color:#5a5f6a;font-size:14px;margin:0 0 18px">
      Here is a summary of <strong>${info.tasks.length} task(s)</strong> assigned to you today:
    </p>
    <table style="width:100%;border-collapse:collapse;font-size:14px">
      <thead>
        <tr style="background:#E4EFEA">
          <th style="padding:10px 12px;text-align:left;color:#0E3221;font-size:11px;text-transform:uppercase;letter-spacing:.05em">Task</th>
          <th style="padding:10px 12px;text-align:left;color:#0E3221;font-size:11px;text-transform:uppercase;letter-spacing:.05em">Dept</th>
          <th style="padding:10px 12px;text-align:left;color:#0E3221;font-size:11px;text-transform:uppercase;letter-spacing:.05em">Deadline</th>
          <th style="padding:10px 12px;text-align:left;color:#0E3221;font-size:11px;text-transform:uppercase;letter-spacing:.05em">Status</th>
          <th style="padding:10px 12px;text-align:left;color:#0E3221;font-size:11px;text-transform:uppercase;letter-spacing:.05em">Assigned By</th>
        </tr>
      </thead>
      <tbody>${taskRows}</tbody>
    </table>
    <p style="color:#5a5f6a;font-size:13px;margin:20px 0 0;border-top:1px solid #e2e5ea;padding-top:16px">
      Please log in to Task Manager to view and update your tasks.<br>
      <em>This is an automated daily summary email.</em>
    </p>
  </div>
</div>`;

      GmailApp.sendEmail(email, subject, `Daily summary: ${info.tasks.length} task(s) assigned to you today. Log in to Task Manager.`, { htmlBody: html, name: 'Task Manager' });
      Logger.log('Daily summary sent to ' + email);
    });

  } catch(e) {
    Logger.log('Daily summary error: ' + e.message);
  }
}

// Run this ONCE to set up the daily 5 PM summary trigger
function setupDailySummaryTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailySummaryEmails') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailySummaryEmails')
    .timeBased().everyDays(1).atHour(17).create();
  Logger.log('Daily summary trigger set for 5 PM every day.');
}

// ============================================================
// TEST FUNCTIONS
// ============================================================
function testSMS() {
  _sendNotification('9876543210', 'Task Manager SMS test. Setup working!');
}
function testWhatsApp() {
  _sendTwilioWhatsApp('+919876543210', 'Task Manager WhatsApp test. Setup working!');
}
