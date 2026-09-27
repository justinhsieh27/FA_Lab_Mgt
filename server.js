const express = require('express');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const XLSX = require('xlsx');
const Database = require('better-sqlite3');
const bcrypt = require('bcryptjs');
const multer = require('multer');

const app = express();
const PORT = Number(process.env.PORT || 3000);
const ROOT = __dirname;
const DATA_DIR = process.env.FA_LAB_DATA_DIR || path.join(ROOT, 'data');
const UPLOAD_DIR = process.env.FA_LAB_UPLOAD_DIR || path.join(ROOT, 'uploads');
const BACKUP_DIR = process.env.FA_LAB_BACKUP_DIR || path.join(ROOT, 'backups');
const WORKBOOK_PATH = path.join(DATA_DIR, 'FA_Lab_Case_Management.xlsx');
const DB_PATH = path.join(DATA_DIR, 'system.sqlite');
const CASE_HEADERS = [
  'case_no', 'tech', 'tech_other', 'cut_count', 'department', 'department_other', 'type', 'type_other', 'layer',
  'priority', 'lot_id', 'wafer_id', 'case_owner', 'start_date', 'end_date', 'cycle_time_days',
  'committed_eta', 'override_status', 'requester_name', 'requester_contact', 'request_summary',
  'analysis_objective', 'remark', 'created_at', 'updated_at'
];
const JOB_HEADERS = [
  'job_id', 'case_no', 'sample_id', 'sample_location', 'cut_id', 'route_preparation',
  'route_process', 'route_analysis', 'job_owner', 'current_step', 'status', 'start_date',
  'end_date', 'cycle_time_days', 'redo_count', 'redo_reason', 'sequence_override', 'sequence_override_reason', 'steps_json', 'remark',
  'created_at', 'updated_at'
];
const ROLES = ['manager', 'engineer', 'guest'];
const DEPARTMENTS = ['ETC', 'LIT', 'PVD', 'CVD', 'WET', 'DIF', 'PIE', 'PDE', 'RA', 'Others'];
const CASE_TYPES = ['PRS', 'Reliability', 'WAT', 'Yield', 'Defect', 'Others'];
const TECHNOLOGIES = ['N40ULP', 'N90BCD', '130BCD', 'E-Flash', 'I-Proj', 'Others'];
const PRIORITIES = ['Super Hot Run', 'Hot Run', 'Normal'];

for (const dir of [DATA_DIR, UPLOAD_DIR, BACKUP_DIR]) fs.mkdirSync(dir, { recursive: true });

const db = new Database(DB_PATH);
db.pragma('journal_mode = WAL');
db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    display_name TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL,
    language TEXT NOT NULL DEFAULT 'zh-TW',
    active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS sessions (
    token TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL,
    expires_at TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at TEXT NOT NULL,
    username TEXT NOT NULL,
    action TEXT NOT NULL,
    entity_type TEXT NOT NULL,
    entity_id TEXT NOT NULL,
    detail TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS attachments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    case_no TEXT NOT NULL,
    job_id TEXT,
    original_name TEXT NOT NULL,
    stored_name TEXT NOT NULL,
    uploaded_by TEXT NOT NULL,
    created_at TEXT NOT NULL
  );
`);

function now() { return new Date().toISOString(); }
function dateOnly(value) { return value ? String(value).slice(0, 10) : ''; }
function escapeLike(value) { return String(value).replace(/[\\%_]/g, '\\$&'); }
function json(value, fallback) { try { return value ? JSON.parse(value) : fallback; } catch { return fallback; } }
function safeText(value) { return value === undefined || value === null ? '' : String(value).trim(); }
function dayDiff(start, end) {
  if (!start || !end) return '';
  const ms = Date.parse(`${dateOnly(end)}T00:00:00Z`) - Date.parse(`${dateOnly(start)}T00:00:00Z`);
  return Number.isFinite(ms) ? Math.max(0, Math.round(ms / 86400000)) : '';
}
function averageDate(dates) {
  if (!dates.length) return '';
  const valid = dates.map(d => Date.parse(`${dateOnly(d)}T00:00:00Z`)).filter(Number.isFinite);
  if (!valid.length) return '';
  return new Date(Math.round(valid.reduce((a, b) => a + b, 0) / valid.length)).toISOString().slice(0, 10);
}

function seedUsers() {
  if (db.prepare('SELECT COUNT(*) AS count FROM users').get().count) return;
  const insert = db.prepare('INSERT INTO users (username, display_name, password_hash, role, created_at) VALUES (?, ?, ?, ?, ?)');
  const stamp = now();
  insert.run('manager', 'FA Lab Manager', bcrypt.hashSync('ChangeMe!2026', 10), 'manager', stamp);
  for (let index = 1; index <= 8; index += 1) {
    insert.run(`engineer${String(index).padStart(2, '0')}`, `Engineer ${index}`, bcrypt.hashSync('ChangeMe!2026', 10), 'engineer', stamp);
  }
  insert.run('guest', 'Guest Viewer', bcrypt.hashSync('ChangeMe!2026', 10), 'guest', stamp);
}
seedUsers();

function emptyWorkbook() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([CASE_HEADERS]), 'FA Case');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([JOB_HEADERS]), 'Job');
  XLSX.writeFile(wb, WORKBOOK_PATH, { bookType: 'xlsx' });
}
if (!fs.existsSync(WORKBOOK_PATH)) emptyWorkbook();

function sheetRows(workbook, name, headers) {
  const sheet = workbook.Sheets[name];
  if (!sheet) return [];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
  return rows.map(row => Object.fromEntries(headers.map(header => [header, row[header] ?? ''])));
}
function loadData() {
  const workbook = XLSX.readFile(WORKBOOK_PATH, { cellDates: false });
  return { cases: sheetRows(workbook, 'FA Case', CASE_HEADERS), jobs: sheetRows(workbook, 'Job', JOB_HEADERS) };
}
function ensureWorkbookSchema() {
  const workbook = XLSX.readFile(WORKBOOK_PATH, { cellDates: false });
  const existingHeaders = XLSX.utils.sheet_to_json(workbook.Sheets['FA Case'], { header: 1, defval: '' })[0] || [];
  if (existingHeaders.includes('tech')) return;
  const data = { cases: sheetRows(workbook, 'FA Case', CASE_HEADERS), jobs: sheetRows(workbook, 'Job', JOB_HEADERS) };
  const updated = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(updated, XLSX.utils.json_to_sheet(data.cases, { header: CASE_HEADERS }), 'FA Case');
  XLSX.utils.book_append_sheet(updated, XLSX.utils.json_to_sheet(data.jobs, { header: JOB_HEADERS }), 'Job');
  XLSX.writeFile(updated, WORKBOOK_PATH, { bookType: 'xlsx' });
}
ensureWorkbookSchema();
let fileQueue = Promise.resolve();
function saveData(data) {
  fileQueue = fileQueue.then(() => {
    const wb = XLSX.utils.book_new();
    const caseSheet = XLSX.utils.json_to_sheet(data.cases, { header: CASE_HEADERS });
    const jobSheet = XLSX.utils.json_to_sheet(data.jobs, { header: JOB_HEADERS });
    XLSX.utils.book_append_sheet(wb, caseSheet, 'FA Case');
    XLSX.utils.book_append_sheet(wb, jobSheet, 'Job');
    const temporary = `${WORKBOOK_PATH}.tmp`;
    XLSX.writeFile(wb, temporary, { bookType: 'xlsx' });
    fs.renameSync(temporary, WORKBOOK_PATH);
  });
  return fileQueue;
}

function makeSteps(route = {}) {
  return [
    { id: 'prep', name: 'Sample preparation', owner: '', status: 'Not Started', start_date: '', end_date: '', required: true },
    { id: 'process', name: route.process || 'FIB', owner: '', status: 'Not Started', start_date: '', end_date: '', required: true },
    { id: 'analysis', name: route.analysis || 'TEM', owner: '', status: 'Not Started', start_date: '', end_date: '', required: true }
  ];
}
function standardizeSteps(job) {
  let steps = json(job.steps_json, []);
  if (!Array.isArray(steps) || !steps.length) steps = makeSteps({ process: job.route_process || 'FIB', analysis: job.route_analysis || 'TEM' });
  return steps.map((step, index) => ({
    id: safeText(step.id) || `custom-${index + 1}`,
    name: safeText(step.name) || `Step ${index + 1}`,
    owner: safeText(step.owner),
    status: ['Not Started', 'In Progress', 'Completed', 'Cancelled'].includes(step.status) ? step.status : 'Not Started',
    start_date: dateOnly(step.start_date),
    end_date: dateOnly(step.end_date),
    required: step.required !== false
  }));
}
function deriveJob(job) {
  const steps = standardizeSteps(job);
  const activeSteps = steps.filter(step => step.status !== 'Cancelled');
  const allFinished = activeSteps.length > 0 && activeSteps.every(step => step.status === 'Completed');
  const anyStarted = activeSteps.some(step => step.status === 'In Progress' || step.status === 'Completed');
  const allCancelled = steps.length > 0 && steps.every(step => step.status === 'Cancelled');
  const current = activeSteps.find(step => step.status === 'In Progress') || activeSteps.find(step => step.status === 'Not Started') || activeSteps.at(-1) || {};
  const starts = activeSteps.map(step => step.start_date).filter(Boolean).sort();
  const ends = activeSteps.map(step => step.end_date).filter(Boolean).sort();
  const status = allCancelled ? 'Cancelled' : allFinished ? 'Completed' : anyStarted ? 'In Progress' : 'Not Started';
  return {
    ...job,
    steps,
    steps_json: JSON.stringify(steps),
    current_step: current.name || '',
    job_owner: current.owner || '',
    status,
    start_date: starts[0] || '',
    end_date: allFinished ? ends.at(-1) || '' : '',
    cycle_time_days: allFinished ? dayDiff(starts[0], ends.at(-1)) : ''
  };
}
function deriveCase(caseRow, jobs) {
  const caseJobs = jobs.filter(job => job.case_no === caseRow.case_no).map(deriveJob);
  const active = caseJobs.filter(job => job.status !== 'Cancelled');
  let calculated = 'New';
  if (active.length && active.every(job => job.status === 'Completed')) calculated = 'Completed';
  else if (active.some(job => job.status === 'In Progress' || job.status === 'Completed')) calculated = 'In Progress';
  const override = ['On Hold', 'Cancelled'].includes(caseRow.override_status) ? caseRow.override_status : '';
  const isCompleted = calculated === 'Completed' && !override;
  const endDate = isCompleted ? averageDate(active.map(job => job.end_date).filter(Boolean)) : '';
  return { ...caseRow, status: override || calculated, end_date: endDate, cycle_time_days: endDate ? dayDiff(caseRow.start_date, endDate) : '', job_count: caseJobs.length, completed_jobs: caseJobs.filter(job => job.status === 'Completed').length };
}
function audit(user, action, entityType, entityId, detail) {
  db.prepare('INSERT INTO audit_log (created_at, username, action, entity_type, entity_id, detail) VALUES (?, ?, ?, ?, ?, ?)')
    .run(now(), user.username, action, entityType, entityId, JSON.stringify(detail));
}
function exportUser(user) { return { id: user.id, username: user.username, display_name: user.display_name, role: user.role, language: user.language, active: Boolean(user.active) }; }
function sessionUser(req) {
  const token = req.headers.authorization?.replace(/^Bearer\s+/i, '') || req.query?.token || req.cookies?.fa_session;
  if (!token) return null;
  return db.prepare(`SELECT u.* FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token=? AND s.expires_at>? AND u.active=1`).get(token, now()) || null;
}
app.use(express.json({ limit: '2mb' }));
app.use((req, res, next) => {
  req.user = sessionUser(req);
  next();
});
function requireAuth(...roles) {
  return (req, res, next) => {
    if (!req.user) return res.status(401).json({ error: 'Authentication required.' });
    if (roles.length && !roles.includes(req.user.role)) return res.status(403).json({ error: 'Permission denied.' });
    next();
  };
}
function cleanCase(input, existing = {}) {
  const row = { ...existing };
  for (const field of CASE_HEADERS) if (field in input && !['end_date', 'cycle_time_days', 'created_at', 'updated_at'].includes(field)) row[field] = safeText(input[field]);
  row.case_no = safeText(row.case_no);
  row.department = DEPARTMENTS.includes(row.department) ? row.department : 'Others';
  row.type = CASE_TYPES.includes(row.type) ? row.type : 'PRS';
  row.tech = TECHNOLOGIES.includes(row.tech) ? row.tech : 'N40ULP';
  row.priority = PRIORITIES.includes(row.priority) ? row.priority : 'Normal';
  row.start_date = dateOnly(row.start_date);
  row.committed_eta = dateOnly(row.committed_eta);
  row.override_status = ['On Hold', 'Cancelled'].includes(row.override_status) ? row.override_status : '';
  return row;
}
function cleanJob(input, existing = {}) {
  const row = { ...existing };
  for (const field of JOB_HEADERS) if (field in input && !['current_step', 'job_owner', 'status', 'start_date', 'end_date', 'cycle_time_days', 'created_at', 'updated_at'].includes(field)) row[field] = typeof input[field] === 'object' ? JSON.stringify(input[field]) : safeText(input[field]);
  row.job_id = safeText(row.job_id);
  row.case_no = safeText(row.case_no);
  row.route_preparation = 'Sample preparation';
  row.route_process = ['FIB', 'SELA'].includes(row.route_process) ? row.route_process : 'FIB';
  row.route_analysis = ['TEM', 'SEM'].includes(row.route_analysis) ? row.route_analysis : 'TEM';
  row.redo_count = Math.max(0, Number.parseInt(row.redo_count || 0, 10) || 0);
  row.sequence_override = String(row.sequence_override) === 'true' ? 'true' : '';
  row.sequence_override_reason = safeText(row.sequence_override_reason);
  if (Array.isArray(input.steps)) row.steps_json = JSON.stringify(input.steps);
  row.steps_json = JSON.stringify(standardizeSteps(row));
  return deriveJob(row);
}
function validateSequence(job) {
  if (job.sequence_override === 'true') return null;
  const steps = standardizeSteps(job);
  for (let index = 0; index < steps.length; index += 1) {
    const step = steps[index];
    if (step.status === 'In Progress' || step.status === 'Completed') {
      const priorIsComplete = steps.slice(0, index).filter(item => item.status !== 'Cancelled').every(item => item.status === 'Completed');
      if (!priorIsComplete) return `Complete the preceding step before starting ${step.name}, or ask a Manager to enable a sequence override.`;
    }
  }
  return null;
}
function createBackup() {
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const folder = path.join(BACKUP_DIR, stamp);
  fs.mkdirSync(folder, { recursive: true });
  for (const file of [WORKBOOK_PATH, DB_PATH, `${DB_PATH}-wal`, `${DB_PATH}-shm`]) if (fs.existsSync(file)) fs.copyFileSync(file, path.join(folder, path.basename(file)));
  fs.cpSync(UPLOAD_DIR, path.join(folder, 'uploads'), { recursive: true });
}
createBackup();
setInterval(createBackup, 24 * 60 * 60 * 1000).unref();

app.post('/api/auth/login', (req, res) => {
  const username = safeText(req.body.username).toLowerCase();
  const user = db.prepare('SELECT * FROM users WHERE username=? AND active=1').get(username);
  if (!user || !bcrypt.compareSync(String(req.body.password || ''), user.password_hash)) return res.status(401).json({ error: 'Invalid username or password.' });
  const token = crypto.randomBytes(32).toString('hex');
  db.prepare('INSERT INTO sessions (token, user_id, expires_at) VALUES (?, ?, ?)').run(token, user.id, new Date(Date.now() + 12 * 60 * 60 * 1000).toISOString());
  audit(user, 'LOGIN', 'user', user.username, {});
  res.json({ token, user: exportUser(user) });
});
app.post('/api/auth/logout', requireAuth(), (req, res) => {
  const token = req.headers.authorization?.replace(/^Bearer\s+/i, '');
  if (token) db.prepare('DELETE FROM sessions WHERE token=?').run(token);
  res.json({ ok: true });
});
app.get('/api/auth/me', requireAuth(), (req, res) => res.json({ user: exportUser(req.user) }));
app.put('/api/auth/language', requireAuth(), (req, res) => {
  const language = req.body.language === 'en' ? 'en' : 'zh-TW';
  db.prepare('UPDATE users SET language=? WHERE id=?').run(language, req.user.id);
  res.json({ language });
});

app.get('/api/meta', requireAuth(), (req, res) => res.json({ departments: DEPARTMENTS, types: CASE_TYPES, technologies: TECHNOLOGIES, priorities: PRIORITIES }));
app.get('/api/cases', requireAuth('manager', 'engineer'), (req, res) => {
  const data = loadData();
  res.json({ cases: data.cases.map(row => deriveCase(row, data.jobs)).sort((a, b) => String(b.created_at).localeCompare(String(a.created_at))) });
});
app.get('/api/cases/:caseNo', requireAuth('manager', 'engineer'), (req, res) => {
  const data = loadData();
  const caseRow = data.cases.find(row => row.case_no === req.params.caseNo);
  if (!caseRow) return res.status(404).json({ error: 'Case not found.' });
  const jobs = data.jobs.filter(row => row.case_no === caseRow.case_no).map(deriveJob);
  res.json({ case: deriveCase(caseRow, jobs), jobs, attachments: db.prepare('SELECT * FROM attachments WHERE case_no=? ORDER BY created_at DESC').all(caseRow.case_no) });
});
app.get('/api/jobs', requireAuth('manager', 'engineer'), (req, res) => {
  const data = loadData();
  const caseNo = safeText(req.query.case_no);
  const jobs = data.jobs.filter(row => !caseNo || row.case_no === caseNo).map(deriveJob);
  res.json({ jobs });
});
app.post('/api/cases', requireAuth('manager', 'engineer'), async (req, res) => {
  const data = loadData();
  const row = cleanCase(req.body);
  if (!row.case_no) return res.status(400).json({ error: 'Case No. is required.' });
  if (data.cases.some(item => item.case_no === row.case_no)) return res.status(409).json({ error: 'Case No. already exists.' });
  row.created_at = row.updated_at = now();
  data.cases.push(row);
  await saveData(data);
  audit(req.user, 'CREATE', 'case', row.case_no, row);
  res.status(201).json({ case: deriveCase(row, []) });
});
app.put('/api/cases/:caseNo', requireAuth('manager', 'engineer'), async (req, res) => {
  const data = loadData();
  const index = data.cases.findIndex(row => row.case_no === req.params.caseNo);
  if (index < 0) return res.status(404).json({ error: 'Case not found.' });
  if ('override_status' in req.body && !['', 'On Hold', 'Cancelled'].includes(req.body.override_status)) return res.status(400).json({ error: 'Invalid status override.' });
  if ('override_status' in req.body && req.user.role !== 'manager') return res.status(403).json({ error: 'Only a manager can override Case status.' });
  const updated = cleanCase({ ...req.body, case_no: data.cases[index].case_no }, data.cases[index]);
  updated.updated_at = now();
  data.cases[index] = updated;
  await saveData(data);
  audit(req.user, 'UPDATE', 'case', updated.case_no, req.body);
  res.json({ case: deriveCase(updated, data.jobs) });
});
app.post('/api/jobs', requireAuth('manager', 'engineer'), async (req, res) => {
  const data = loadData();
  if (!data.cases.some(row => row.case_no === safeText(req.body.case_no))) return res.status(400).json({ error: 'A valid Case No. is required.' });
  let row = cleanJob(req.body);
  if (row.sequence_override === 'true' && req.user.role !== 'manager') return res.status(403).json({ error: 'Only a manager can enable a sequence override.' });
  const sequenceError = validateSequence(row);
  if (sequenceError) return res.status(400).json({ error: sequenceError });
  if (!row.job_id) return res.status(400).json({ error: 'Job ID is required.' });
  if (data.jobs.some(item => item.job_id === row.job_id)) return res.status(409).json({ error: 'Job ID already exists.' });
  row.created_at = row.updated_at = now();
  row = deriveJob(row);
  data.jobs.push(row);
  await saveData(data);
  audit(req.user, 'CREATE', 'job', row.job_id, row);
  res.status(201).json({ job: row });
});
app.put('/api/jobs/:jobId', requireAuth('manager', 'engineer'), async (req, res) => {
  const data = loadData();
  const index = data.jobs.findIndex(row => row.job_id === req.params.jobId);
  if (index < 0) return res.status(404).json({ error: 'Job not found.' });
  let updated = cleanJob({ ...req.body, job_id: data.jobs[index].job_id, case_no: data.jobs[index].case_no }, data.jobs[index]);
  if (updated.sequence_override === 'true' && req.user.role !== 'manager') return res.status(403).json({ error: 'Only a manager can enable a sequence override.' });
  const sequenceError = validateSequence(updated);
  if (sequenceError) return res.status(400).json({ error: sequenceError });
  updated.updated_at = now();
  updated = deriveJob(updated);
  data.jobs[index] = updated;
  await saveData(data);
  audit(req.user, 'UPDATE', 'job', updated.job_id, req.body);
  res.json({ job: updated });
});

app.get('/api/dashboard', requireAuth('manager', 'engineer'), (req, res) => {
  const data = loadData();
  const jobs = data.jobs.map(deriveJob);
  const cases = data.cases.map(row => deriveCase(row, jobs));
  const open = cases.filter(row => !['Completed', 'Cancelled'].includes(row.status));
  const overdue = open.filter(row => row.committed_eta && row.committed_eta < dateOnly(now()));
  const by = (rows, key) => rows.reduce((out, row) => { const value = row[key] || 'Unassigned'; out[value] = (out[value] || 0) + 1; return out; }, {});
  const trend = (rows, field) => rows.reduce((out, row) => { const key = dateOnly(row[field]); if (key) out[key] = (out[key] || 0) + 1; return out; }, {});
  const cycle = rows => { const values = rows.map(row => Number(row.cycle_time_days)).filter(Number.isFinite); return values.length ? Math.round((values.reduce((a, b) => a + b, 0) / values.length) * 10) / 10 : 0; };
  res.json({
    kpis: { totalCases: cases.length, openCases: open.length, overdueCases: overdue.length, completedCases: cases.filter(row => row.status === 'Completed').length, openJobs: jobs.filter(row => !['Completed', 'Cancelled'].includes(row.status)).length, avgCaseCycle: cycle(cases.filter(row => row.status === 'Completed')), avgJobCycle: cycle(jobs.filter(row => row.status === 'Completed')), redoCount: jobs.reduce((total, row) => total + Number(row.redo_count || 0), 0) },
    distributions: { priority: by(open, 'priority'), department: by(open, 'department'), owner: by(jobs.filter(row => !['Completed', 'Cancelled'].includes(row.status)), 'job_owner'), step: by(jobs.filter(row => !['Completed', 'Cancelled'].includes(row.status)), 'current_step'), redoReason: by(jobs.filter(row => Number(row.redo_count) > 0), 'redo_reason') },
    trend: { casesCreated: trend(cases, 'created_at'), casesCompleted: trend(cases.filter(row => row.status === 'Completed'), 'end_date'), jobsCreated: trend(jobs, 'created_at'), jobsCompleted: trend(jobs.filter(row => row.status === 'Completed'), 'end_date') },
    queue: open.sort((a, b) => PRIORITIES.indexOf(a.priority) - PRIORITIES.indexOf(b.priority) || String(a.committed_eta).localeCompare(String(b.committed_eta)))
  });
});
app.get('/api/queue', requireAuth(), (req, res) => {
  const data = loadData();
  const cases = data.cases.map(row => deriveCase(row, data.jobs)).filter(row => !['Completed', 'Cancelled'].includes(row.status));
  const queue = cases.map(row => ({ case_no: row.case_no, department: row.department === 'Others' ? row.department_other : row.department, type: row.type === 'Others' ? row.type_other : row.type, priority: row.priority, lot_id: row.lot_id, wafer_id: row.wafer_id, status: row.status, created_at: row.created_at, committed_eta: row.committed_eta }))
    .sort((a, b) => PRIORITIES.indexOf(a.priority) - PRIORITIES.indexOf(b.priority) || String(a.committed_eta).localeCompare(String(b.committed_eta)));
  res.json({ queue });
});
app.get('/api/my-work', requireAuth('manager', 'engineer'), (req, res) => {
  const data = loadData();
  const jobs = data.jobs.map(deriveJob);
  const work = jobs.filter(job => job.steps.some(step => step.owner === req.user.username && !['Completed', 'Cancelled'].includes(step.status))).map(job => ({ ...job, case: data.cases.find(row => row.case_no === job.case_no) || {} }));
  res.json({ jobs: work });
});

app.get('/api/users', requireAuth('manager'), (req, res) => res.json({ users: db.prepare('SELECT * FROM users ORDER BY role, display_name').all().map(exportUser) }));
app.get('/api/users/assignees', requireAuth('manager', 'engineer'), (req, res) => res.json({ users: db.prepare("SELECT id, username, display_name, role FROM users WHERE active=1 AND role IN ('manager', 'engineer') ORDER BY display_name").all() }));
app.post('/api/users', requireAuth('manager'), (req, res) => {
  const username = safeText(req.body.username).toLowerCase();
  const displayName = safeText(req.body.display_name);
  const role = ROLES.includes(req.body.role) ? req.body.role : null;
  if (!/^[a-z0-9._-]{3,40}$/.test(username) || !displayName || !role || String(req.body.password || '').length < 8) return res.status(400).json({ error: 'Please enter a valid username, display name, role and 8+ character password.' });
  try {
    const result = db.prepare('INSERT INTO users (username, display_name, password_hash, role, created_at) VALUES (?, ?, ?, ?, ?)').run(username, displayName, bcrypt.hashSync(req.body.password, 10), role, now());
    const user = db.prepare('SELECT * FROM users WHERE id=?').get(result.lastInsertRowid);
    audit(req.user, 'CREATE', 'user', username, { displayName, role });
    res.status(201).json({ user: exportUser(user) });
  } catch { res.status(409).json({ error: 'Username already exists.' }); }
});
app.put('/api/users/:id', requireAuth('manager'), (req, res) => {
  const current = db.prepare('SELECT * FROM users WHERE id=?').get(req.params.id);
  if (!current) return res.status(404).json({ error: 'User not found.' });
  const displayName = safeText(req.body.display_name) || current.display_name;
  const role = ROLES.includes(req.body.role) ? req.body.role : current.role;
  const active = typeof req.body.active === 'boolean' ? Number(req.body.active) : current.active;
  const passwordHash = req.body.password ? bcrypt.hashSync(req.body.password, 10) : current.password_hash;
  db.prepare('UPDATE users SET display_name=?, role=?, active=?, password_hash=? WHERE id=?').run(displayName, role, active, passwordHash, current.id);
  audit(req.user, 'UPDATE', 'user', current.username, { displayName, role, active });
  res.json({ user: exportUser(db.prepare('SELECT * FROM users WHERE id=?').get(current.id)) });
});
app.get('/api/audit', requireAuth('manager'), (req, res) => res.json({ entries: db.prepare('SELECT * FROM audit_log ORDER BY id DESC LIMIT 500').all().map(row => ({ ...row, detail: json(row.detail, {}) })) }));

const storage = multer.diskStorage({ destination: (_req, _file, cb) => cb(null, UPLOAD_DIR), filename: (_req, file, cb) => cb(null, `${Date.now()}-${crypto.randomBytes(8).toString('hex')}${path.extname(file.originalname)}`) });
const upload = multer({ storage, limits: { fileSize: 25 * 1024 * 1024 } });
app.post('/api/attachments', requireAuth('manager', 'engineer'), upload.single('file'), (req, res) => {
  if (!req.file || !safeText(req.body.case_no)) return res.status(400).json({ error: 'A file and Case No. are required.' });
  const result = db.prepare('INSERT INTO attachments (case_no, job_id, original_name, stored_name, uploaded_by, created_at) VALUES (?, ?, ?, ?, ?, ?)').run(safeText(req.body.case_no), safeText(req.body.job_id) || null, req.file.originalname, req.file.filename, req.user.username, now());
  audit(req.user, 'UPLOAD', 'attachment', String(result.lastInsertRowid), { case_no: req.body.case_no, job_id: req.body.job_id, name: req.file.originalname });
  res.status(201).json({ attachment: db.prepare('SELECT * FROM attachments WHERE id=?').get(result.lastInsertRowid) });
});
app.get('/api/attachments/:id/download', requireAuth('manager', 'engineer'), (req, res) => {
  const attachment = db.prepare('SELECT * FROM attachments WHERE id=?').get(req.params.id);
  if (!attachment) return res.status(404).json({ error: 'Attachment not found.' });
  res.download(path.join(UPLOAD_DIR, attachment.stored_name), attachment.original_name);
});
app.get('/api/export/excel', requireAuth('manager', 'engineer'), (_req, res) => res.download(WORKBOOK_PATH, 'FA_Lab_Case_Management.xlsx'));

app.use(express.static(path.join(ROOT, 'public')));
app.get('*splat', (_req, res) => res.sendFile(path.join(ROOT, 'public', 'index.html')));
app.use((error, _req, res, _next) => { console.error(error); res.status(500).json({ error: error.message || 'Unexpected server error.' }); });

app.listen(PORT, () => console.log(`FA Lab Case Management is running at http://localhost:${PORT}`));
