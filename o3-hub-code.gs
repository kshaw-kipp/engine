/**
 * Kevin × Sophie O3 Hub — Google Apps Script Backend
 * ─────────────────────────────────────────────────────
 * SETUP:
 *   1. Open your Google Sheet → Extensions → Apps Script
 *   2. Paste this entire file into Code.gs
 *   3. Run initSheets() once to create the sheet structure
 *   4. Deploy → New deployment → Web app
 *      - Execute as: Me
 *      - Who has access: Anyone with the link
 *   5. Copy the web app URL and set it as GAS_URL env var in Vercel
 *
 * NOTE: No SHEET_ID needed — this script is bound to your Sheet.
 *
 * SHEETS CREATED:
 *   - Sessions    : one row per O3 session
 *   - Agenda      : one row per workstream per session
 *   - Actions     : running action item tracker
 *   - Categories  : workstream groupings + colors
 *   - Skills      : key-value store for skill tracking
 */

// ──────────────────────────────────────────────────────────────
// HELPERS
// ──────────────────────────────────────────────────────────────

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function zipObj(keys, vals) {
  const obj = {};
  keys.forEach((k, i) => { obj[k] = vals[i]; });
  return obj;
}

// ──────────────────────────────────────────────────────────────
// GET — read data
// ──────────────────────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  let result;

  try {
    if      (action === 'getAll')       result = getAll();
    else if (action === 'getSessions')  result = getSessions();
    else if (action === 'getAgenda')    result = getAgenda(e.parameter.sessionId);
    else if (action === 'getAllAgenda') result = getAllAgenda();
    else if (action === 'getActions')   result = getActions();
    else if (action === 'getCategories')result = getCategories();
    else if (action === 'getSkills')      result = getSkills();
    else if (action === 'getQuarterlies') result = getJsonBlob('Quarterlies');
    else if (action === 'getEvalEdits')   result = getJsonBlob('EvalEdits');
    else                                  result = { error: 'Unknown action: ' + action };
  } catch (err) {
    result = { error: err.message };
  }

  return jsonOut(result);
}

// ──────────────────────────────────────────────────────────────
// POST — write data
// ──────────────────────────────────────────────────────────────
function doPost(e) {
  let result;

  try {
    const p = JSON.parse(e.postData.contents);

    if      (p.action === 'createSession')     result = createSession(p);
    else if (p.action === 'updateSession')     result = updateSession(p);
    else if (p.action === 'deleteSession')     result = deleteSession(p.id);
    else if (p.action === 'updateAgendaItem')  result = updateAgendaItem(p.sessionId, p.workstream, p.data);
    else if (p.action === 'createAction')      result = createAction(p);
    else if (p.action === 'updateAction')      result = updateAction(p.id, p);
    else if (p.action === 'updateActionStatus')result = updateActionStatus(p.id, p.status);
    else if (p.action === 'deleteAction')      result = deleteAction(p.id);
    else if (p.action === 'syncCategories')    result = syncCategories(p.categories);
    else if (p.action === 'syncSkills')        result = syncSkills(p.skills);
    else if (p.action === 'syncQuarterlies')   result = setJsonBlob('Quarterlies', p.quarters);
    else if (p.action === 'syncEvalEdits')     result = setJsonBlob('EvalEdits',   p.edits);
    else if (p.action === 'syncAll')           result = syncAll(p.sessions, p.agenda, p.actions, p.categories, p.skills);
    else                                       result = { error: 'Unknown action: ' + p.action };
  } catch (err) {
    result = { error: err.message };
  }

  return jsonOut(result);
}

// ──────────────────────────────────────────────────────────────
// INIT — run once to create sheets
// ──────────────────────────────────────────────────────────────
function initSheets() {
  const ss = getSS();

  let sessions = ss.getSheetByName('Sessions');
  if (!sessions) {
    sessions = ss.insertSheet('Sessions');
    sessions.appendRow(['id', 'date', 'label', 'status', 'created']);
    sessions.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1E293B').setFontColor('white');
    sessions.setFrozenRows(1);
    sessions.setColumnWidths(1, 5, [160, 120, 200, 100, 160]);
  }

  let agenda = ss.getSheetByName('Agenda');
  if (!agenda) {
    agenda = ss.insertSheet('Agenda');
    const h = ['sessionId','date','workstream','category','questions','isNA','urgency','links','notes','kevinNext','sophieNext','status','updatedAt'];
    agenda.appendRow(h);
    agenda.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1E293B').setFontColor('white');
    agenda.setFrozenRows(1);
  }

  let actions = ss.getSheetByName('Actions');
  if (!actions) {
    actions = ss.insertSheet('Actions');
    const h = ['id','sessionId','sessionDate','workstream','text','owner','due','status','created'];
    actions.appendRow(h);
    actions.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1E293B').setFontColor('white');
    actions.setFrozenRows(1);
    actions.setColumnWidths(1, 9, [160, 160, 120, 200, 320, 100, 100, 80, 160]);
  }

  let cats = ss.getSheetByName('Categories');
  if (!cats) {
    cats = ss.insertSheet('Categories');
    cats.appendRow(['id', 'name', 'color', 'workstreams']);
    cats.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1E293B').setFontColor('white');
    cats.setFrozenRows(1);
    cats.setColumnWidths(1, 4, [140, 200, 90, 500]);
  }

  let skills = ss.getSheetByName('Skills');
  if (!skills) {
    skills = ss.insertSheet('Skills');
    skills.appendRow(['key', 'value']);
    skills.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1E293B').setFontColor('white');
    skills.setFrozenRows(1);
    skills.setColumnWidths(1, 2, [140, 600]);
  }

  ['Quarterlies', 'EvalEdits'].forEach(name => {
    if (!ss.getSheetByName(name)) {
      const sh = ss.insertSheet(name);
      sh.appendRow(['key', 'value']);
      sh.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1E293B').setFontColor('white');
      sh.setFrozenRows(1);
      sh.setColumnWidths(1, 2, [140, 600]);
    }
  });

  SpreadsheetApp.flush();
  Logger.log('✅ Sheets ready: Sessions, Agenda, Actions, Categories, Skills, Quarterlies, EvalEdits');
}

// ──────────────────────────────────────────────────────────────
// READ FUNCTIONS
// ──────────────────────────────────────────────────────────────

function getAll() {
  return {
    sessions:   getSessions(),
    agenda:     getAllAgenda(),
    actions:    getActions(),
    categories: getCategories(),
    skills:     getSkills(),
    quarterlies: getJsonBlob('Quarterlies'),
    evalEdits:   getJsonBlob('EvalEdits'),
    _ts:        new Date().toISOString()
  };
}

function getSessions() {
  const sh = getSheet('Sessions');
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const [headers, ...data] = rows;
  return data
    .filter(r => r[0])
    .map(r => {
      const obj = zipObj(headers, r);
      obj.date = normDate(obj.date);   // always return YYYY-MM-DD string
      return obj;
    })
    .sort((a, b) => new Date(b.date) - new Date(a.date));
}

function getAgenda(sessionId) {
  const sh = getSheet('Agenda');
  if (!sh) return {};
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return {};
  const [headers, ...data] = rows;
  const result = {};
  data.filter(r => r[0] === sessionId).map(r => zipObj(headers, r)).forEach(item => {
    result[item.workstream] = agendaShape(item);
  });
  return result;
}

function getAllAgenda() {
  const sh = getSheet('Agenda');
  if (!sh) return {};
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return {};
  const [headers, ...data] = rows;
  const result = {};
  data.filter(r => r[0]).map(r => zipObj(headers, r)).forEach(item => {
    if (!result[item.sessionId]) result[item.sessionId] = {};
    result[item.sessionId][item.workstream] = agendaShape(item);
  });
  return result;
}

function agendaShape(item) {
  return {
    questions:  item.questions  || '',
    na:         item.isNA === true || item.isNA === 'TRUE',
    urgency:    item.urgency    || '',
    links:      item.links      || '',
    notes:      item.notes      || '',
    kevinNext:  item.kevinNext  || '',
    sophieNext: item.sophieNext || '',
    status:     item.status     || 'pending'
  };
}

function getActions() {
  const sh = getSheet('Actions');
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const [headers, ...data] = rows;
  return data
    .filter(r => r[0])
    .map(r => zipObj(headers, r))
    .sort((a, b) => new Date(b.created) - new Date(a.created));
}

function getCategories() {
  const sh = getSheet('Categories');
  if (!sh) return null;
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return null;
  const [headers, ...data] = rows;
  return data.filter(r => r[0]).map(r => {
    const obj = zipObj(headers, r);
    let ws = [];
    try { ws = JSON.parse(obj.workstreams || '[]'); } catch { ws = []; }
    return { id: obj.id, name: obj.name, color: obj.color, workstreams: ws };
  });
}

function getSkills() {
  const sh = getSheet('Skills');
  if (!sh) return null;
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return null;
  const result = {};
  const [, ...data] = rows;
  data.filter(r => r[0]).forEach(r => {
    try { result[r[0]] = JSON.parse(r[1]); } catch { result[r[0]] = r[1]; }
  });
  return Object.keys(result).length ? result : null;
}

// ──────────────────────────────────────────────────────────────
// WRITE FUNCTIONS
// ──────────────────────────────────────────────────────────────

function createSession(p) {
  const sh = getSheet('Sessions');
  sh.appendRow([p.id, p.date, p.label || '', p.status || 'prep', p.created || Date.now()]);
  return { ok: true, id: p.id };
}

function updateSession(p) {
  const sh   = getSheet('Sessions');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === p.id) {
      if (p.date   !== undefined) sh.getRange(i + 1, 2).setValue(p.date);
      if (p.label  !== undefined) sh.getRange(i + 1, 3).setValue(p.label);
      if (p.status !== undefined) sh.getRange(i + 1, 4).setValue(p.status);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Session not found: ' + p.id };
}

function deleteSession(id) {
  const ss = getSheet('Sessions');
  deleteRowById(ss, id);

  const ag = getSheet('Agenda');
  if (ag) {
    const data = ag.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === id) ag.deleteRow(i + 1);
    }
  }
  return { ok: true };
}

function updateAgendaItem(sessionId, workstream, data) {
  // data is a partial patch object from the frontend
  if (!data || !sessionId || !workstream) return { ok: false, error: 'Missing params' };

  const sh    = getSheet('Agenda');
  const rows  = sh.getDataRange().getValues();
  const sessions = getSessions();
  const session  = sessions.find(s => s.id === sessionId) || {};

  // Find existing row index
  let rowIdx = -1;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === sessionId && rows[i][2] === workstream) {
      rowIdx = i + 1; break;
    }
  }

  // Build full row, merging patch over existing values
  const existing = rowIdx > 0 ? zipObj(rows[0], rows[rowIdx - 1]) : {};
  const merged   = Object.assign({}, existing, data);

  const rowData = [
    sessionId,
    session.date || existing.date || '',
    workstream,
    getCategoryForWorkstream(workstream),
    merged.questions  || '',
    merged.na         || false,
    merged.urgency    || '',
    merged.links      || '',
    merged.notes      || '',
    merged.kevinNext  || '',
    merged.sophieNext || '',
    merged.status     || 'pending',
    new Date().toISOString()
  ];

  if (rowIdx > 0) {
    sh.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sh.appendRow(rowData);
  }
  return { ok: true };
}

function createAction(p) {
  const sh = getSheet('Actions');
  const sessions = getSessions();
  const session  = sessions.find(s => s.id === p.sessionId) || {};
  sh.appendRow([
    p.id,
    p.sessionId   || '',
    session.date  || '',
    p.workstream  || '',
    p.text        || '',
    p.owner       || 'Kevin',
    p.due         || '',
    p.status      || 'open',
    p.created     || Date.now()
  ]);
  return { ok: true, id: p.id };
}

function updateAction(id, fields) {
  const sh   = getSheet('Actions');
  const data = sh.getDataRange().getValues();
  // Columns: id(1) sessionId(2) sessionDate(3) workstream(4) text(5) owner(6) due(7) status(8) created(9)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      if (fields.text      !== undefined) sh.getRange(i + 1, 5).setValue(fields.text);
      if (fields.owner     !== undefined) sh.getRange(i + 1, 6).setValue(fields.owner);
      if (fields.due       !== undefined) sh.getRange(i + 1, 7).setValue(fields.due);
      if (fields.status    !== undefined) sh.getRange(i + 1, 8).setValue(fields.status);
      if (fields.workstream!== undefined) sh.getRange(i + 1, 4).setValue(fields.workstream);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Action not found: ' + id };
}

function updateActionStatus(id, status) {
  const sh   = getSheet('Actions');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sh.getRange(i + 1, 8).setValue(status);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Action not found: ' + id };
}

function deleteAction(id) {
  const sh = getSheet('Actions');
  return deleteRowById(sh, id) ? { ok: true } : { ok: false, error: 'Action not found: ' + id };
}

function syncCategories(categories) {
  if (!Array.isArray(categories)) return { ok: false, error: 'categories must be an array' };
  // LockService prevents two concurrent requests from interleaving their
  // clearContent + appendRow operations, which is the root cause of duplicate rows.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // wait up to 15s for any concurrent write to finish
    let sh = getSheet('Categories');
    if (!sh) { initSheets(); sh = getSheet('Categories'); }
    // Delete existing data rows (not just clearContent — actually removes the rows
    // so appendRow always starts fresh from row 2 with no phantom empty rows).
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.deleteRows(2, lastRow - 1);
    categories.forEach(c => {
      sh.appendRow([c.id, c.name, c.color, JSON.stringify(c.workstreams || [])]);
    });
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function syncSkills(skills) {
  if (!skills || typeof skills !== 'object') return { ok: false, error: 'skills must be an object' };
  let sh = getSheet('Skills');
  if (!sh) { initSheets(); sh = getSheet('Skills'); }
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  Object.entries(skills).forEach(([k, v]) => {
    sh.appendRow([k, typeof v === 'object' ? JSON.stringify(v) : v]);
  });
  return { ok: true };
}

function syncAll(sessions, agenda, actions, categories, skills) {
  const ss = getSheet('Sessions');

  // Sessions
  if (ss && Array.isArray(sessions)) {
    if (ss.getLastRow() > 1) ss.getRange(2, 1, ss.getLastRow() - 1, 5).clearContent();
    sessions.forEach(s => {
      ss.appendRow([s.id, s.date, s.label || '', s.status || 'prep', s.created || '']);
    });
  }

  // Agenda
  const ag = getSheet('Agenda');
  if (ag && agenda && typeof agenda === 'object') {
    if (ag.getLastRow() > 1) ag.getRange(2, 1, ag.getLastRow() - 1, 13).clearContent();
    Object.entries(agenda).forEach(([sessionId, workstreams]) => {
      const session = (sessions || []).find(s => s.id === sessionId) || {};
      Object.entries(workstreams).forEach(([workstream, item]) => {
        ag.appendRow([
          sessionId, session.date || '', workstream,
          getCategoryForWorkstream(workstream),
          item.questions  || '', item.na || false, item.urgency || '',
          item.links      || '', item.notes || '',
          item.kevinNext  || '', item.sophieNext || '',
          item.status     || 'pending', new Date().toISOString()
        ]);
      });
    });
  }

  // Actions
  const ac = getSheet('Actions');
  if (ac && Array.isArray(actions)) {
    if (ac.getLastRow() > 1) ac.getRange(2, 1, ac.getLastRow() - 1, 9).clearContent();
    actions.forEach(a => {
      const session = (sessions || []).find(s => s.id === a.sessionId) || {};
      ac.appendRow([a.id, a.sessionId || '', session.date || '', a.workstream || '',
                    a.text || '', a.owner || '', a.due || '', a.status || 'open', a.created || '']);
    });
  }

  if (categories) syncCategories(categories);
  if (skills)     syncSkills(skills);

  SpreadsheetApp.flush();
  return { ok: true };
}

// ──────────────────────────────────────────────────────────────
// UTILITIES
// ──────────────────────────────────────────────────────────────

function getSheet(name) {
  return getSS().getSheetByName(name);
}

function deleteRowById(sh, id) {
  if (!sh) return false;
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === id) { sh.deleteRow(i + 1); return true; }
  }
  return false;
}

/** Normalize any date value (Date object, timestamp, string) → 'YYYY-MM-DD' */
function normDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (typeof val === 'number') {
    return Utilities.formatDate(new Date(val), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  // Already a string — strip any time component and return date part only
  return String(val).split('T')[0].substring(0, 10);
}

// ──────────────────────────────────────────────────────────────
// JSON BLOB HELPERS — used for Quarterlies and EvalEdits sheets
// Each sheet stores exactly one data row: [key, JSON string]
// ──────────────────────────────────────────────────────────────

function getJsonBlob(sheetName) {
  const sh = getSheet(sheetName);
  if (!sh || sh.getLastRow() <= 1) return null;
  try { return JSON.parse(sh.getRange(2, 2).getValue()); } catch { return null; }
}

function setJsonBlob(sheetName, data) {
  let sh = getSheet(sheetName);
  if (!sh) { initSheets(); sh = getSheet(sheetName); }
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  sh.appendRow([sheetName.toLowerCase(), JSON.stringify(data)]);
  return { ok: true };
}

function getCategoryForWorkstream(workstream) {
  const sh = getSheet('Categories');
  if (!sh) return '';
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    try {
      const ws = JSON.parse(rows[i][3] || '[]');
      if (ws.includes(workstream)) return rows[i][1];
    } catch { /* skip */ }
  }
  return '';
}
