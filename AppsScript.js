// ═══════════════════════════════════════════════════════════════
// Daily Tracker — Google Apps Script Backend
// Deploy as Web App: Execute as "Me", Access "Anyone"
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME   = 'TrackerData';
const PHOTO_SHEET  = 'Photos';
const NOTES_SHEET  = 'Notes';
const CONFIG_SHEET = 'UserConfig';
const SPORTS_SHEET = 'SportsPlan';
const FOODDEF_SHEET = 'FoodDefs';
const FOODLOG_SHEET = 'FoodLog';
const DIETTARGET_SHEET = 'DietTargets';
const FOCUSLOG_SHEET = 'FocusLog';

// ─── credentials ────────────────────────────────────────────────
// TESTER backend:  Ibrahim→"tester"   Sameeha→"tester1"
// LIVE backend:    Ibrahim→(live pwd) Sameeha→(live pwd)
const USERS = {
  "Ibrahim": "YOUR_IBRAHIM_PASSWORD",
  "Sameeha": "YOUR_SAMEEHA_PASSWORD"
};

// ─── sheet helpers ───────────────────────────────────────────────
function getOrCreateSheet(name, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if(!sh){
    sh = ss.insertSheet(name);
    if(headers) sh.appendRow(headers);
  }
  return sh;
}

function getTrackerSheet(){  return getOrCreateSheet(SHEET_NAME,  ['user','date','activities','anotes']); }
function getPhotoSheet(){    return getOrCreateSheet(PHOTO_SHEET,  ['user','date','type','value','file_id','photo_id']); }
function getNotesSheet(){    return getOrCreateSheet(NOTES_SHEET,  ['user','date','note']); }
function getConfigSheet(){   return getOrCreateSheet(CONFIG_SHEET, ['user','activities','categories']); }
function getSportsSheet(){   return getOrCreateSheet(SPORTS_SHEET, ['user','weekStart','plan','done']); }
function getFoodDefSheet(){  return getOrCreateSheet(FOODDEF_SHEET,['user','foods']); }
// One row per logged food item (not a JSON blob) so the sheet itself is a
// readable table you can scroll through and review at the end of the week.
function getFoodLogSheet(){  return getOrCreateSheet(FOODLOG_SHEET,['user','date','meal','food','qty','kcal','protein','fat','fiber','carbs','entryId']); }
function getDietTargetSheet(){ return getOrCreateSheet(DIETTARGET_SHEET,['user','kcal','protein']); }
// One row per completed/stopped Focus Timer session, same row-per-entry
// shape as FoodLog, so the sheet itself stays a plain readable table.
function getFocusLogSheet(){ return getOrCreateSheet(FOCUSLOG_SHEET,['user','date','activityId','activityLabel','cat','seconds','startedAt','entryId']); }

function formatSheetDate(val){
  if(!val) return '';
  if(val instanceof Date){
    const y = val.getFullYear();
    const m = String(val.getMonth()+1).padStart(2,'0');
    const d = String(val.getDate()).padStart(2,'0');
    return `${y}-${m}-${d}`;
  }
  return String(val);
}

function cors(output){
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── token store (per-execution cache) ──────────────────────────
// We use CacheService (script cache) so tokens survive across requests
function makeToken(user){
  const tok = Utilities.getUuid();
  CacheService.getScriptCache().put('tok:'+tok, user, 21600); // 6 hrs
  return tok;
}
function tokenUser(tok){
  return CacheService.getScriptCache().get('tok:'+tok);
}

// ═══════════════════════════════════════════════════════════════
// MAIN ROUTER
// ═══════════════════════════════════════════════════════════════
// Photo chunk uploads arrive as POST (payload in the body, not the URL,
// so mobile carrier proxies can't mangle long GET query strings). Apps
// Script merges POST body params into e.parameter same as GET, so both
// entry points can share the same router.
function doPost(e){
  return doGet(e);
}

function doGet(e){
  const p = e.parameter || {};
  const action = p.action || '';

  // LOGIN doesn't need a token
  if(action === 'login'){
    const name = p.user || '';
    const pass = p.password || '';
    if(USERS[name] && USERS[name] === pass){
      const tok = makeToken(name);
      return cors(JSON.stringify({ ok:true, user:name, token:tok }));
    }
    return cors(JSON.stringify({ ok:false, error:'Invalid credentials' }));
  }

  // All other actions require a valid token
  const tok = p.token || '';
  const authedUser = tokenUser(tok);
  if(!authedUser){
    return cors(JSON.stringify({ ok:false, error:'Unauthorized' }));
  }

  try{
    if(action === 'getWeekBoth')  return cors(getWeekBoth(p));
    if(action === 'saveDay')      return cors(saveDay(p, authedUser));
    if(action === 'resetDay')     return cors(resetDay(p, authedUser));
    if(action === 'getNote')      return cors(getNote(p));
    if(action === 'saveNote')     return cors(saveNote(p, authedUser));
    if(action === 'getPhotos')    return cors(getPhotos(p));
    if(action === 'uploadChunk')  return cors(uploadChunk(p, authedUser));
    if(action === 'deletePhoto')  return cors(deletePhoto(p, authedUser));
    if(action === 'getUserConfig')return cors(getUserConfig(p));
    if(action === 'saveUserConfig')return cors(saveUserConfig(p, authedUser));
    if(action === 'saveAnote')    return cors(saveAnote(p, authedUser));
    if(action === 'getSportsPlan')     return cors(getSportsPlan(p));
    if(action === 'saveSportsPlan')    return cors(saveSportsPlan(p, authedUser));
    if(action === 'toggleSportsSession')return cors(toggleSportsSession(p, authedUser));
    if(action === 'getFoodDefs')  return cors(getFoodDefs(p));
    if(action === 'saveFoodDefs') return cors(saveFoodDefs(p, authedUser));
    if(action === 'getFoodLog')   return cors(getFoodLog(p));
    if(action === 'saveFoodLog')  return cors(saveFoodLog(p, authedUser));
    if(action === 'getDietTarget') return cors(getDietTarget(p));
    if(action === 'saveDietTarget')return cors(saveDietTarget(p, authedUser));
    if(action === 'getFocusLog')  return cors(getFocusLog(p));
    if(action === 'saveFocusLog') return cors(saveFocusLog(p, authedUser));
    return cors(JSON.stringify({ ok:false, error:'Unknown action' }));
  }catch(err){
    return cors(JSON.stringify({ ok:false, error: err.message }));
  }
}

// ═══════════════════════════════════════════════════════════════
// getWeekBoth — returns toggle data + anotes for both users
// ═══════════════════════════════════════════════════════════════
function getWeekBoth(p){
  const start = p.start;
  const end   = p.end;
  if(!start || !end) return JSON.stringify({ ok:false, error:'Missing start/end' });

  const sh   = getTrackerSheet();
  const rows = sh.getDataRange().getValues();

  const result   = {};
  const anotes   = {};

  // skip header row (row 0)
  for(let i = 1; i < rows.length; i++){
    const user = String(rows[i][0]);
    const date = formatSheetDate(rows[i][1]);
    const acts = rows[i][2];
    const ants = rows[i][3];

    if(date < start || date > end) continue;

    if(!result[user])  result[user]  = {};
    if(!anotes[user])  anotes[user]  = {};

    try{ result[user][date] = JSON.parse(acts || '[]'); }catch{ result[user][date] = []; }
    try{ anotes[user][date] = JSON.parse(ants || '{}'); }catch{ anotes[user][date] = {}; }
  }

  return JSON.stringify({ ok:true, data:result, anotes });
}

// ═══════════════════════════════════════════════════════════════
// saveDay — merges a set of {activityId: checked} toggles into the
// existing stored activities array for one user+date, instead of
// overwriting it wholesale with the caller's local snapshot.
//
// The old version replaced the whole day's array with whatever the
// client had in memory. When the same profile was open on two devices
// at once, whichever device's save happened to land last on the server
// silently erased the other device's more recent toggle — the two
// devices would then permanently disagree even after a full reload,
// since the server's own stored data was wrong, not just a local cache.
// A script lock serializes concurrent saves so two near-simultaneous
// requests always merge onto each other instead of racing.
// ═══════════════════════════════════════════════════════════════
function saveDay(p, authedUser){
  const user = p.user;
  const date = p.date;
  let diffs;
  try{ diffs = JSON.parse(p.diffs || '{}'); }catch(e){ diffs = {}; }

  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const sh   = getTrackerSheet();
    const rows = sh.getDataRange().getValues();

    for(let i = 1; i < rows.length; i++){
      if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date){
        let acts = [];
        try{ acts = JSON.parse(rows[i][2] || '[]'); }catch(e){}
        const set = new Set(acts);
        Object.keys(diffs).forEach(actId => { if(diffs[actId]) set.add(actId); else set.delete(actId); });
        const updated = [...set];
        sh.getRange(i+1, 3).setValue(JSON.stringify(updated));
        return JSON.stringify({ ok:true, activities: updated });
      }
    }
    // new row — preserve anotes column as empty object
    const set = new Set();
    Object.keys(diffs).forEach(actId => { if(diffs[actId]) set.add(actId); });
    const updated = [...set];
    sh.appendRow([user, date, JSON.stringify(updated), '{}']);
    return JSON.stringify({ ok:true, activities: updated });
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════
// resetDay — explicit full clear of a user+date's activities, for the
// "Reset this day" button. Unlike saveDay this intentionally wipes
// everything rather than merging, since that's exactly what the user
// asked for by clicking it.
// ═══════════════════════════════════════════════════════════════
function resetDay(p, authedUser){
  const user = p.user;
  const date = p.date;

  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const sh   = getTrackerSheet();
    const rows = sh.getDataRange().getValues();

    for(let i = 1; i < rows.length; i++){
      if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date){
        sh.getRange(i+1, 3).setValue('[]');
        return JSON.stringify({ ok:true });
      }
    }
    sh.appendRow([user, date, '[]', '{}']);
    return JSON.stringify({ ok:true });
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════
// saveAnote — saves a per-activity note into column 4 of TrackerData
// ═══════════════════════════════════════════════════════════════
function saveAnote(p, authedUser){
  const user  = p.user;
  const date  = p.date;
  const actId = p.actId;
  const text  = p.text || '';

  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });

  const sh   = getTrackerSheet();
  const rows = sh.getDataRange().getValues();

  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date){
      let ants = {};
      try{ ants = JSON.parse(rows[i][3] || '{}'); }catch{}
      if(text) ants[actId] = text;
      else     delete ants[actId];
      sh.getRange(i+1, 4).setValue(JSON.stringify(ants));
      return JSON.stringify({ ok:true });
    }
  }
  // row doesn't exist yet — create it
  const ants = text ? JSON.stringify({ [actId]: text }) : '{}';
  sh.appendRow([user, date, '[]', ants]);
  return JSON.stringify({ ok:true });
}

// ═══════════════════════════════════════════════════════════════
// getNote / saveNote — per-day sticky notes
// ═══════════════════════════════════════════════════════════════
function getNote(p){
  const user = p.user;
  const date = p.date;
  const sh   = getNotesSheet();
  const rows = sh.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date){
      return JSON.stringify({ ok:true, note: rows[i][2] || '' });
    }
  }
  return JSON.stringify({ ok:true, note:'' });
}

function saveNote(p, authedUser){
  const user = p.user;
  const date = p.date;
  const note = p.note || '';

  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });

  const sh   = getNotesSheet();
  const rows = sh.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date){
      sh.getRange(i+1, 3).setValue(note);
      return JSON.stringify({ ok:true });
    }
  }
  sh.appendRow([user, date, note]);
  return JSON.stringify({ ok:true });
}

// ═══════════════════════════════════════════════════════════════
// getUserConfig / saveUserConfig — device-independent activities + categories
// UserConfig sheet columns: user | activities (JSON) | categories (JSON)
// ═══════════════════════════════════════════════════════════════
function getUserConfig(p){
  const user = p.user;
  const sh   = getConfigSheet();
  const rows = sh.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user){
      let activities = null, categories = null;
      try{ activities  = JSON.parse(rows[i][1] || 'null'); }catch{}
      try{ categories  = JSON.parse(rows[i][2] || 'null'); }catch{}
      return JSON.stringify({ ok:true, activities, categories });
    }
  }
  return JSON.stringify({ ok:true, activities:null, categories:null });
}

function saveUserConfig(p, authedUser){
  const user = p.user;
  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });

  const sh   = getConfigSheet();
  const rows = sh.getDataRange().getValues();

  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user){
      if(p.activities) sh.getRange(i+1, 2).setValue(p.activities);
      if(p.categories) sh.getRange(i+1, 3).setValue(p.categories);
      return JSON.stringify({ ok:true });
    }
  }
  // new row
  sh.appendRow([user, p.activities || 'null', p.categories || 'null']);
  return JSON.stringify({ ok:true });
}

// ═══════════════════════════════════════════════════════════════
// Sports plan — a per-user, per-week checklist of training sessions.
// SportsPlan sheet cols: user | weekStart | plan (JSON) | done (JSON)
//   plan: [{ day, intensity, sessions:[{id,label,detail,icon}], notes }, ...]
//   done: { sessionId: true }
// Plan structure is edited/uploaded as a whole (like resetDay, an
// intentional full replace); completion ticks are merged per-session
// (like saveDay) so two devices ticking around the same time don't
// clobber each other.
// ═══════════════════════════════════════════════════════════════
function getSportsPlan(p){
  const user = p.user, weekStart = p.weekStart;
  const sh   = getSportsSheet();
  const rows = sh.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === weekStart){
      let plan = [], done = {};
      try{ plan = JSON.parse(rows[i][2] || '[]'); }catch(e){}
      try{ done = JSON.parse(rows[i][3] || '{}'); }catch(e){}
      return JSON.stringify({ ok:true, plan, done });
    }
  }
  return JSON.stringify({ ok:true, plan:null, done:{} });
}

function saveSportsPlan(p, authedUser){
  const user = p.user, weekStart = p.weekStart;
  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });
  let plan;
  try{ plan = JSON.parse(p.plan || '[]'); }catch(e){ plan = []; }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const sh   = getSportsSheet();
    const rows = sh.getDataRange().getValues();
    for(let i = 1; i < rows.length; i++){
      if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === weekStart){
        sh.getRange(i+1, 3).setValue(JSON.stringify(plan));
        return JSON.stringify({ ok:true });
      }
    }
    sh.appendRow([user, weekStart, JSON.stringify(plan), '{}']);
    return JSON.stringify({ ok:true });
  } finally {
    lock.releaseLock();
  }
}

function toggleSportsSession(p, authedUser){
  const user = p.user, weekStart = p.weekStart, sessionId = p.sessionId;
  const checked = p.checked === 'true' || p.checked === true;
  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });
  if(!sessionId) return JSON.stringify({ ok:false, error:'Missing sessionId' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const sh   = getSportsSheet();
    const rows = sh.getDataRange().getValues();
    for(let i = 1; i < rows.length; i++){
      if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === weekStart){
        let done = {};
        try{ done = JSON.parse(rows[i][3] || '{}'); }catch(e){}
        if(checked) done[sessionId] = true; else delete done[sessionId];
        sh.getRange(i+1, 4).setValue(JSON.stringify(done));
        return JSON.stringify({ ok:true, done });
      }
    }
    const done = checked ? { [sessionId]: true } : {};
    sh.appendRow([user, weekStart, '[]', JSON.stringify(done)]);
    return JSON.stringify({ ok:true, done });
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════
// Predefined foods + daily food log.
// FoodDefs sheet cols: user | foods (JSON array of {id,name,kcal,protein,unit})
// FoodLog  sheet cols: user | date | entries (JSON array of
//   {id,foodId,name,qty,kcal,protein,meal}) — name/kcal/protein are
//   snapshotted at log time so editing a food definition later doesn't
//   silently rewrite past days' totals.
// ═══════════════════════════════════════════════════════════════
function getFoodDefs(p){
  const user = p.user;
  const sh   = getFoodDefSheet();
  const rows = sh.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user){
      let foods = [];
      try{ foods = JSON.parse(rows[i][1] || '[]'); }catch(e){}
      return JSON.stringify({ ok:true, foods });
    }
  }
  return JSON.stringify({ ok:true, foods:[] });
}

function saveFoodDefs(p, authedUser){
  const user = p.user;
  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });
  const foods = p.foods || '[]';

  const sh   = getFoodDefSheet();
  const rows = sh.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user){
      sh.getRange(i+1, 2).setValue(foods);
      return JSON.stringify({ ok:true });
    }
  }
  sh.appendRow([user, foods]);
  return JSON.stringify({ ok:true });
}

function getFoodLog(p){
  const user = p.user, date = p.date;
  const sh   = getFoodLogSheet();
  const rows = sh.getDataRange().getValues();
  const entries = [];
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date){
      entries.push({
        id: rows[i][10] || ('row_' + i),
        meal: rows[i][2] || '',
        name: rows[i][3] || '',
        qty: Number(rows[i][4]) || 0,
        kcal: Number(rows[i][5]) || 0,
        protein: Number(rows[i][6]) || 0,
        fat: Number(rows[i][7]) || 0,
        fiber: Number(rows[i][8]) || 0,
        carbs: Number(rows[i][9]) || 0,
      });
    }
  }
  return JSON.stringify({ ok:true, entries });
}

// Full day replace: clears every row for this user+date, then appends one
// row per entry — keeps the sheet itself a plain readable table (date |
// meal | food | qty | kcal | protein | fat | fiber | carbs) you can scroll
// and review, instead of one JSON blob per day.
function saveFoodLog(p, authedUser){
  const user = p.user, date = p.date;
  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });
  let entries = [];
  try{ entries = JSON.parse(p.entries || '[]'); }catch(e){}

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const sh   = getFoodLogSheet();
    const rows = sh.getDataRange().getValues();
    const toDelete = [];
    for(let i = 1; i < rows.length; i++){
      if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date) toDelete.push(i + 1);
    }
    toDelete.sort((a,b) => b-a).forEach(r => sh.deleteRow(r));

    entries.forEach(e => {
      sh.appendRow([user, date, e.meal || '', e.name || '', e.qty || 0, e.kcal || 0, e.protein || 0, e.fat || 0, e.fiber || 0, e.carbs || 0, e.id || '']);
    });
    return JSON.stringify({ ok:true });
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════
// Diet targets — one editable kcal/protein goal per user.
// ═══════════════════════════════════════════════════════════════
function getDietTarget(p){
  const user = p.user;
  const sh   = getDietTargetSheet();
  const rows = sh.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user){
      return JSON.stringify({ ok:true, kcal: Number(rows[i][1]) || 0, protein: Number(rows[i][2]) || 0 });
    }
  }
  return JSON.stringify({ ok:true, kcal:0, protein:0 });
}

function saveDietTarget(p, authedUser){
  const user = p.user;
  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });
  const kcal    = Number(p.kcal) || 0;
  const protein = Number(p.protein) || 0;

  const sh   = getDietTargetSheet();
  const rows = sh.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user){
      sh.getRange(i+1, 2, 1, 2).setValues([[kcal, protein]]);
      return JSON.stringify({ ok:true });
    }
  }
  sh.appendRow([user, kcal, protein]);
  return JSON.stringify({ ok:true });
}

// ═══════════════════════════════════════════════════════════════
// Focus Timer log — one row per completed/stopped session.
// Full day replace on save (like FoodLog), keeping the sheet a plain
// readable table of date | activity | category | seconds | startedAt.
// ═══════════════════════════════════════════════════════════════
function getFocusLog(p){
  const user = p.user, date = p.date;
  const sh   = getFocusLogSheet();
  const rows = sh.getDataRange().getValues();
  const entries = [];
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date){
      entries.push({
        id: rows[i][7] || ('row_' + i),
        activityId: rows[i][2] || '',
        activityLabel: rows[i][3] || '',
        cat: rows[i][4] || '',
        seconds: Number(rows[i][5]) || 0,
        startedAt: rows[i][6] || '',
      });
    }
  }
  return JSON.stringify({ ok:true, entries });
}

function saveFocusLog(p, authedUser){
  const user = p.user, date = p.date;
  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });
  let entries = [];
  try{ entries = JSON.parse(p.entries || '[]'); }catch(e){}

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const sh   = getFocusLogSheet();
    const rows = sh.getDataRange().getValues();
    const toDelete = [];
    for(let i = 1; i < rows.length; i++){
      if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date) toDelete.push(i + 1);
    }
    toDelete.sort((a,b) => b-a).forEach(r => sh.deleteRow(r));

    entries.forEach(e => {
      sh.appendRow([user, date, e.activityId || '', e.activityLabel || '', e.cat || '', e.seconds || 0, e.startedAt || '', e.id || '']);
    });
    return JSON.stringify({ ok:true });
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════
// Photo upload / get / delete — multiple photos per user/day
// Photos sheet cols: user | date | type | value | file_id | photo_id
//   type='session:N' rows are temporary chunks during upload
//   type='done'       row stores one photo's Drive URL + file ID + photo_id
// Drive folder structure: Daily Tracker Photos / {user} / {year} / {month} / {date}
// (older photos may still sit directly in the {month} folder as "{date}.jpg" —
//  run migratePhotosToDateFolders() once from the editor to move them in)
// ═══════════════════════════════════════════════════════════════
const MONTH_NAMES = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December'];

function getDriveFolder(parent, name){
  const iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parent.createFolder(name);
}

function getDayFolder(user, date){
  const parts       = date.split('-');
  const year        = parts[0];
  const month       = MONTH_NAMES[parseInt(parts[1]) - 1];

  const root        = DriveApp.getRootFolder();
  const appFolder   = getDriveFolder(root,        'Daily Tracker Photos');
  const userFolder  = getDriveFolder(appFolder,   user);
  const yearFolder  = getDriveFolder(userFolder,  year);
  const monthFolder = getDriveFolder(yearFolder,  month);
  return getDriveFolder(monthFolder, date);
}

function saveToDrive(user, date, base64, photoId){
  const dayFolder = getDayFolder(user, date);

  const bytes = Utilities.base64Decode(base64);
  const blob  = Utilities.newBlob(bytes, 'image/jpeg', photoId + '.jpg');
  const file  = dayFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    url:    'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w720',
    fileId: file.getId()
  };
}

function getPhotos(p){
  const user = p.user, date = p.date;
  const sh   = getPhotoSheet();
  const rows = sh.getDataRange().getValues();
  const photos = [];
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date && rows[i][2] === 'done'){
      photos.push({ id: rows[i][5] || ('legacy_' + i), url: rows[i][3] || null });
    }
  }
  return JSON.stringify({ ok:true, photos: photos });
}

function uploadChunk(p, authedUser){
  const user    = p.user;
  const date    = p.date;
  const photoId = (p.photoId || Utilities.getUuid()).replace(/[^a-zA-Z0-9_-]/g,'').slice(0,40);
  const session = (p.session || '').replace(/[^a-zA-Z0-9]/g,'').slice(-12);
  const index   = parseInt(p.index || '0');
  const data    = p.chunk || '';
  const total   = parseInt(p.total || '1');

  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });

  // Reject a chunk immediately if it arrived truncated, instead of silently
  // storing bad data and only discovering it at final assembly. Only the
  // last chunk is allowed to be shorter than the fixed chunk size — this
  // caught a real case where one middle chunk arrived as 7 chars instead
  // of 3500, corrupting the whole upload on a flaky mobile connection.
  const EXPECTED_CHUNK_SIZE = 3500;
  const isFinalChunk = index === total - 1;
  if(!isFinalChunk && data.length !== EXPECTED_CHUNK_SIZE){
    return JSON.stringify({ ok:false, error: `Chunk ${index} arrived truncated (${data.length}/${EXPECTED_CHUNK_SIZE} chars) — retry` });
  }

  const sh = getPhotoSheet();

  // Append this chunk with a session-scoped key — no delete here to avoid race conditions
  // when parallel requests run. Cleanup happens only on the last chunk during assembly.
  sh.appendRow([user, date, session + ':' + index, data, '', '']);

  if(index === total - 1){
    const rows     = sh.getDataRange().getValues();
    const byIndex  = {};
    const toDelete = [];
    const prefix   = session + ':';

    for(let i = 1; i < rows.length; i++){
      if(String(rows[i][0]) !== user || formatSheetDate(rows[i][1]) !== date) continue;
      const t = String(rows[i][2]);
      if(t.startsWith(prefix)){
        byIndex[parseInt(t.slice(prefix.length))] = rows[i][3];
        toDelete.push(i + 1);
      }
    }

    // Fail loudly on a dropped chunk instead of silently base64-decoding a
    // corrupted string later (surfaced to users as a confusing "Could not
    // decode string" error with no indication what actually went wrong).
    const received = Object.keys(byIndex).length;
    if(received !== total){
      return JSON.stringify({ ok:false, error: `Upload incomplete: received ${received}/${total} chunks — please retry` });
    }

    const sortedIdx  = Object.keys(byIndex).map(Number).sort((a,b) => a-b);
    const fullBase64 = sortedIdx.map(k => byIndex[k]).join('');
    let result;
    try{
      result = saveToDrive(user, date, fullBase64, photoId);
    }catch(err){
      // Surface exactly what was received so we can see which chunk (if any)
      // came back the wrong size, instead of a bare "Could not decode string".
      const lens = sortedIdx.map(k => byIndex[k].length);
      return JSON.stringify({ ok:false, error: `Decode failed: ${err.message} | total=${total} chunkLens=[${lens.join(',')}] finalLen=${fullBase64.length}` });
    }

    toDelete.sort((a,b) => b-a).forEach(r => sh.deleteRow(r));
    sh.appendRow([user, date, 'done', result.url, result.fileId, photoId]);

    return JSON.stringify({ ok:true, done:true, url:result.url, photoId:photoId });
  }

  return JSON.stringify({ ok:true, done:false });
}

function deletePhoto(p, authedUser){
  const user = p.user, date = p.date, photoId = p.photoId || '';
  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });
  if(!photoId) return JSON.stringify({ ok:false, error:'Missing photoId' });

  const sh   = getPhotoSheet();
  const rows = sh.getDataRange().getValues();
  for(let i = rows.length - 1; i >= 1; i--){
    if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date && rows[i][2] === 'done'){
      const rowPhotoId = rows[i][5] || ('legacy_' + i);
      if(String(rowPhotoId) !== String(photoId)) continue;
      const fileId = String(rows[i][4] || '');
      if(fileId) try{ DriveApp.getFileById(fileId).setTrashed(true); }catch(e){}
      sh.deleteRow(i+1);
    }
  }
  return JSON.stringify({ ok:true });
}

// ═══════════════════════════════════════════════════════════════
// One-time migration: move older photos that were saved directly as
// ".../{year}/{month}/{date}.jpg" into their own "{date}/" subfolder,
// and backfill a photo_id for their Photos-sheet row so they behave
// like any other gallery photo. Run manually from the Apps Script
// editor once after deploying the date-folder version — safe to
// run more than once (already-migrated files are skipped).
// ═══════════════════════════════════════════════════════════════
function migratePhotosToDateFolders(){
  const sh   = getPhotoSheet();
  const rows = sh.getDataRange().getValues();
  let moved  = 0;

  for(let i = 1; i < rows.length; i++){
    const user = String(rows[i][0]);
    const date = formatSheetDate(rows[i][1]);
    const type = rows[i][2];
    const fileId = String(rows[i][4] || '');
    if(type !== 'done' || !fileId || rows[i][5]) continue; // skip chunks and already-tagged rows

    let file;
    try{ file = DriveApp.getFileById(fileId); }catch(e){ continue; }

    const parts       = date.split('-');
    const year        = parts[0];
    const month       = MONTH_NAMES[parseInt(parts[1]) - 1];
    const root        = DriveApp.getRootFolder();
    const appFolder   = getDriveFolder(root,       'Daily Tracker Photos');
    const userFolder  = getDriveFolder(appFolder,  user);
    const yearFolder  = getDriveFolder(userFolder, year);
    const monthFolder = getDriveFolder(yearFolder, month);
    const dayFolder    = getDriveFolder(monthFolder, date);

    // Move the file into the day folder if it isn't already there
    const alreadyThere = file.getParents().hasNext() &&
      (function(){ const ps = file.getParents(); while(ps.hasNext()){ if(ps.next().getId() === dayFolder.getId()) return true; } return false; })();
    if(!alreadyThere){
      dayFolder.addFile(file);
      const parents = file.getParents();
      while(parents.hasNext()){
        const parent = parents.next();
        if(parent.getId() !== dayFolder.getId()) parent.removeFile(file);
      }
    }

    const photoId = 'legacy_' + i;
    sh.getRange(i + 1, 6).setValue(photoId); // photo_id column
    moved++;
  }

  Logger.log('Migrated ' + moved + ' photo(s) into date folders.');
  return moved;
}

// ═══════════════════════════════════════════════════════════════
// Utility: clear all data (run manually in Apps Script editor to reset)
// ═══════════════════════════════════════════════════════════════
function clearAllData(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEET_NAME, PHOTO_SHEET, NOTES_SHEET, CONFIG_SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if(!sh) return;
    const lastRow = sh.getLastRow();
    if(lastRow > 1) sh.getRange(2, 1, lastRow-1, sh.getLastColumn()).clearContent();
  });
}
