const SHEET_SESSIONS   = 'Sessioni';
const SHEET_CATEGORIES = 'Categorie';
const SHEET_PROJECTS   = 'Progetti';
const SHEET_TASKS      = 'Tasks';
const SHEET_NOTES      = 'Note';

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const body = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents)
      : {};
    const action = e && e.parameter && e.parameter.action
      ? e.parameter.action
      : body.action;

    let result;

    if      (action === 'get')            result = getAllData();
    else if (action === 'add')            result = addSession(body.session);
    else if (action === 'delete')         result = deleteSession(body.id);
    else if (action === 'saveCategories') result = saveCategories(body.categories || []);
    else if (action === 'saveProjects')   result = saveProjects(body.projects || []);
    else if (action === 'saveTasks')      result = saveTasks(body.tasks || []);
    else if (action === 'saveNotes')      result = saveNotes(body.notes || []);
    else                                  result = { ok: false, error: 'Azione non riconosciuta' };

    output.setContent(JSON.stringify(result));
  } catch (err) {
    output.setContent(JSON.stringify({
      ok: false,
      error: err && err.message ? err.message : String(err)
    }));
  }

  return output;
}

// SESSIONI
function getSessionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_SESSIONS);
  if (!s) {
    s = ss.insertSheet(SHEET_SESSIONS);
    s.appendRow(['ID','Data','Categoria','Minuti','Pomodori','DurataPomodoro','Note','Umore','Progetti','Timestamp']);
    s.getRange(1,1,1,10).setFontWeight('bold');
  }
  return s;
}

function getSessions() {
  const rows = getSessionSheet().getDataRange().getValues();
  if (rows.length <= 1) return [];

  return rows.slice(1).map(r => {
    let dateVal = r[1];
    if (dateVal instanceof Date) {
      dateVal = dateVal.getFullYear()+'-'+String(dateVal.getMonth()+1).padStart(2,'0')+'-'+String(dateVal.getDate()).padStart(2,'0');
    } else {
      dateVal = dateVal ? dateVal.toString().slice(0,10) : '';
    }

    const projRaw = r[8] ? r[8].toString() : '';
    const projects = projRaw ? projRaw.split(',').map(p => p.trim()).filter(Boolean) : [];

    return {
      id:              r[0],
      date:            dateVal,
      category:        r[2],
      minutes:         r[3],
      pomodori:        r[4],
      durata_pomodoro: r[5],
      note:            r[6] || '',
      mood:            r[7] || '',
      projects:        projects
    };
  });
}

function addSession(s) {
  if (!s) throw new Error('Sessione mancante');

  const id = Date.now().toString();
  const dateStr = s.date ? s.date.toString().slice(0,10) : '';
  const projStr = Array.isArray(s.projects) ? s.projects.join(',') : (s.projects || '');

  getSessionSheet().appendRow([
    id,
    dateStr,
    s.category,
    s.minutes,
    s.pomodori || '',
    s.durata_pomodoro || '',
    s.note || '',
    s.mood || '',
    projStr,
    new Date().toISOString()
  ]);

  return { ok: true, id: id };
}

function deleteSession(id) {
  const sheet = getSessionSheet();
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }

  return { ok: false, error: 'Non trovata' };
}

// CATEGORIE
function getCategorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_CATEGORIES);
  if (!s) {
    s = ss.insertSheet(SHEET_CATEGORIES);
    s.appendRow(['Categoria']);
    s.getRange(1,1,1,1).setFontWeight('bold');
  }
  return s;
}

function getCategories() {
  const rows = getCategorySheet().getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(r => r[0]).filter(Boolean);
}

function saveCategories(cats) {
  const sheet = getCategorySheet();
  sheet.clearContents();
  sheet.appendRow(['Categoria']);
  if (cats.length) sheet.getRange(2,1,cats.length,1).setValues(cats.map(c => [c]));
  return { ok: true };
}

// PROGETTI
function getProjectSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_PROJECTS);
  if (!s) {
    s = ss.insertSheet(SHEET_PROJECTS);
    s.appendRow(['Progetto']);
    s.getRange(1,1,1,1).setFontWeight('bold');
  }
  return s;
}

function getProjects() {
  const rows = getProjectSheet().getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(r => r[0]).filter(Boolean);
}

function saveProjects(projs) {
  const sheet = getProjectSheet();
  sheet.clearContents();
  sheet.appendRow(['Progetto']);
  if (projs.length) sheet.getRange(2,1,projs.length,1).setValues(projs.map(p => [p]));
  return { ok: true };
}

// TASK CONDIVISE TRA PC E CELLULARE
function getTaskSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_TASKS);
  if (!s) {
    s = ss.insertSheet(SHEET_TASKS);
    s.appendRow(['ID','JSON']);
    s.getRange(1,1,1,2).setFontWeight('bold');
  }
  return s;
}

function getTasks() {
  const rows = getTaskSheet().getDataRange().getValues();
  if (rows.length <= 1) return [];

  return rows.slice(1).map(r => {
    try {
      const task = JSON.parse(r[1]);
      return task && task.id ? task : null;
    } catch (err) {
      return null;
    }
  }).filter(Boolean);
}

function saveTasks(tasks) {
  if (!Array.isArray(tasks)) throw new Error('Formato task non valido');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getTaskSheet();
    sheet.clearContents();
    sheet.getRange(1,1,1,2).setValues([['ID','JSON']]);
    sheet.getRange(1,1,1,2).setFontWeight('bold');

    if (tasks.length) {
      const rows = tasks.map(task => [
        String(task.id || ''),
        JSON.stringify(task)
      ]);
      sheet.getRange(2,1,rows.length,2).setValues(rows);
    }

    return { ok: true, count: tasks.length };
  } finally {
    lock.releaseLock();
  }
}

// NOTE RAPIDE CONDIVISE TRA PC E CELLULARE
function getNoteSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_NOTES);
  if (!s) {
    s = ss.insertSheet(SHEET_NOTES);
    s.appendRow(['ID','JSON']);
    s.getRange(1,1,1,2).setFontWeight('bold');
  }
  return s;
}

function getNotes() {
  const rows = getNoteSheet().getDataRange().getValues();
  if (rows.length <= 1) return [];

  return rows.slice(1).map(r => {
    try {
      const note = JSON.parse(r[1]);
      return note && note.text ? note : null;
    } catch (err) {
      return null;
    }
  }).filter(Boolean);
}

function saveNotes(notes) {
  if (!Array.isArray(notes)) throw new Error('Formato note non valido');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getNoteSheet();
    sheet.clearContents();
    sheet.getRange(1,1,1,2).setValues([['ID','JSON']]);
    sheet.getRange(1,1,1,2).setFontWeight('bold');

    if (notes.length) {
      const rows = notes.map(note => [
        String(note.id || note.at || ''),
        JSON.stringify(note)
      ]);
      sheet.getRange(2,1,rows.length,2).setValues(rows);
    }

    return { ok: true, count: notes.length };
  } finally {
    lock.releaseLock();
  }
}

// RISPOSTA COMPLETA AL TRACKER
function getAllData() {
  return {
    ok:         true,
    sessions:   getSessions(),
    categories: getCategories(),
    projects:   getProjects(),
    tasks:      getTasks(),
    notes:      getNotes()
  };
}
