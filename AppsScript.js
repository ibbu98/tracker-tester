// ═══════════════════════════════════════════════════════════════
// Daily Tracker — Google Apps Script Backend
// Deploy as Web App: Execute as "Me", Access "Anyone"
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME   = 'TrackerData';
const PHOTO_SHEET  = 'Photos';
const NOTES_SHEET  = 'Notes';
const CONFIG_SHEET = 'UserConfig';

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
function getPhotoSheet(){    return getOrCreateSheet(PHOTO_SHEET,  ['user','date','type','value','file_id']); }
function getNotesSheet(){    return getOrCreateSheet(NOTES_SHEET,  ['user','date','note']); }
function getConfigSheet(){   return getOrCreateSheet(CONFIG_SHEET, ['user','activities','categories']); }

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
    if(action === 'getNote')      return cors(getNote(p));
    if(action === 'saveNote')     return cors(saveNote(p, authedUser));
    if(action === 'getPhoto')     return cors(getPhoto(p));
    if(action === 'uploadChunk')  return cors(uploadChunk(p, authedUser));
    if(action === 'deletePhoto')  return cors(deletePhoto(p, authedUser));
    if(action === 'getUserConfig')return cors(getUserConfig(p));
    if(action === 'saveUserConfig')return cors(saveUserConfig(p, authedUser));
    if(action === 'saveAnote')    return cors(saveAnote(p, authedUser));
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
// saveDay — saves activity toggle array for one user+date
// ═══════════════════════════════════════════════════════════════
function saveDay(p, authedUser){
  const user = p.user;
  const date = p.date;
  const acts = p.activities; // JSON string

  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });

  const sh   = getTrackerSheet();
  const rows = sh.getDataRange().getValues();

  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date){
      sh.getRange(i+1, 3).setValue(acts);
      return JSON.stringify({ ok:true });
    }
  }
  // new row — preserve anotes column as empty object
  sh.appendRow([user, date, acts, '{}']);
  return JSON.stringify({ ok:true });
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
// Photo upload / get / delete
// Photos sheet cols: user | date | type | value | file_id
//   type='chunk_N' rows are temporary during upload
//   type='done'    row stores the final Drive URL + file ID
// Drive folder structure: Daily Tracker Photos / {user} / {year} / {month}
// ═══════════════════════════════════════════════════════════════
const MONTH_NAMES = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December'];

function getDriveFolder(parent, name){
  const iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parent.createFolder(name);
}

function saveToDrive(user, date, base64){
  const parts       = date.split('-');
  const year        = parts[0];
  const month       = MONTH_NAMES[parseInt(parts[1]) - 1];

  const root        = DriveApp.getRootFolder();
  const appFolder   = getDriveFolder(root,        'Daily Tracker Photos');
  const userFolder  = getDriveFolder(appFolder,   user);
  const yearFolder  = getDriveFolder(userFolder,  year);
  const monthFolder = getDriveFolder(yearFolder,  month);

  // Delete any existing file for this date (replacing a photo)
  const existing = monthFolder.getFilesByName(date + '.jpg');
  while(existing.hasNext()) existing.next().setTrashed(true);

  const bytes = Utilities.base64Decode(base64);
  const blob  = Utilities.newBlob(bytes, 'image/jpeg', date + '.jpg');
  const file  = monthFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    url:    'https://drive.google.com/uc?export=view&id=' + file.getId(),
    fileId: file.getId()
  };
}

function getPhoto(p){
  const user = p.user, date = p.date;
  const sh   = getPhotoSheet();
  const rows = sh.getDataRange().getValues();
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date && rows[i][2] === 'done'){
      return JSON.stringify({ ok:true, url: rows[i][3] || null });
    }
  }
  return JSON.stringify({ ok:true, url:null });
}

function uploadChunk(p, authedUser){
  const user    = p.user;
  const date    = p.date;
  const session = (p.session || '').replace(/[^a-zA-Z0-9]/g,'').slice(-12);
  const index   = parseInt(p.index || '0');
  const data    = p.chunk || '';
  const total   = parseInt(p.total || '1');

  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });

  const sh = getPhotoSheet();

  // Append this chunk with a session-scoped key — no delete here to avoid race conditions
  // when parallel requests run. Cleanup happens only on the last chunk during assembly.
  sh.appendRow([user, date, session + ':' + index, data, '']);

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
      } else {
        toDelete.push(i + 1); // delete orphan chunks and old 'done' rows
      }
    }

    const fullBase64 = Object.keys(byIndex).map(Number).sort((a,b) => a-b).map(k => byIndex[k]).join('');
    const result = saveToDrive(user, date, fullBase64);

    toDelete.sort((a,b) => b-a).forEach(r => sh.deleteRow(r));
    sh.appendRow([user, date, 'done', result.url, result.fileId]);

    return JSON.stringify({ ok:true, done:true, url:result.url });
  }

  return JSON.stringify({ ok:true, done:false });
}

function deletePhoto(p, authedUser){
  const user = p.user, date = p.date;
  if(user !== authedUser) return JSON.stringify({ ok:false, error:'Forbidden' });

  const sh   = getPhotoSheet();
  const rows = sh.getDataRange().getValues();
  for(let i = rows.length - 1; i >= 1; i--){
    if(String(rows[i][0]) === user && formatSheetDate(rows[i][1]) === date){
      const fileId = String(rows[i][4] || '');
      if(fileId) try{ DriveApp.getFileById(fileId).setTrashed(true); }catch(e){}
      sh.deleteRow(i+1);
    }
  }
  return JSON.stringify({ ok:true });
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
