const API_ACTIONS = ['login', 'saveRecord', 'listRecords', 'listEmployees', 'addEmployee', 'stats', 'listDutyOfficers', 'saveDutyOfficer', 'saveEmployeeCard'];
const DEFAULT_API_KEY = 'DEV_KEY';
const DEFAULT_SPREADSHEET_ID = '1K-XEE97ddfdZG6t3br76bzLSQB4R6VFZQT9SbAAJb2E';
const DEFAULT_PIN = '1234';
const DEFAULT_USER = { name: 'Ns. Ayu Lestari', role: 'Perawat', email: 'ayu.lestari@tpknm.id', initials: 'AL' };

const RECORD_HEADERS = ['id', 'timestamp', 'employeeId', 'employeeName', 'systolic', 'diastolic', 'temperature', 'spo2', 'heartRate', 'weightKg', 'heightCm', 'bmi', 'status', 'source', 'reason', 'examiner', 'examinerRole'];
const ANSWER_KEYS = ['q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10'];
const ANSWERS_HEADERS = ['id', 'timestamp', 'employeeId', 'employeeName'].concat(ANSWER_KEYS);
const EMP_HEADERS = ['id', 'Nomor kartu RF', 'Nama', 'Unit', 'Grup Tugas', 'Tanggal Lahir', 'Tinggi Badan'];
const EMP_KEYS = ['id', 'cardUid', 'name', 'unit', 'grup', 'dob', 'heightCm'];
const DEFAULT_EMPLOYEES = [
  ['K-1001', null, 'Budi Santoso', 'HSSE', 'A', '1992-01-01', 172],
  ['K-1002', null, 'Siti Rahma', 'Produksi', 'A', '1998-01-01', 160],
  ['K-1003', null, 'Andi Wijaya', 'Maintenance', 'B', '1985-01-01', 168],
  ['K-1004', null, 'Dewi Anggraini', 'Logistik', 'A', '1995-01-01', 158],
  ['K-1005', null, 'Rudi Hartono', 'Produksi', 'C', '1981-01-01', 175],
  ['K-1006', null, 'Lestari Putri', 'HRD', 'A', '2000-01-01', 162],
  ['K-1007', null, 'Hendra Gunawan', 'Maintenance', 'B', '1988-01-01', 170],
  ['K-1008', null, 'Maya Sari', 'Logistik', 'A', '1997-01-01', 155],
];

const DUTY_HEADERS = ['id', 'name', 'role', 'initials', 'defaultName', 'defaultRole'];
const DEFAULT_DUTY_OFFICERS = [
  ['DOC-01', 'dr. Pratama Wijaya, Sp.OK', 'Dokter', 'PW', 'dr. Pratama Wijaya, Sp.OK', 'Dokter'],
  ['DOC-02', 'dr. Ratna Kusuma, Sp.PD', 'Dokter', 'RK', 'dr. Ratna Kusuma, Sp.PD', 'Dokter'],
  ['NUR-01', 'Ns. Ayu Lestari, S.Kep', 'Perawat', 'AL', 'Ns. Ayu Lestari, S.Kep', 'Perawat'],
  ['NUR-02', 'Ns. Budi Hartono, S.Kep', 'Perawat', 'BH', 'Ns. Budi Hartono, S.Kep', 'Perawat'],
  ['NUR-03', 'Ns. Siti Rahma, S.Kep', 'Perawat', 'SR', 'Ns. Siti Rahma, S.Kep', 'Perawat'],
];

const STR_REQUIRED = [
  ['timestamp', 'Timestamp'],
  ['employeeId', 'ID karyawan'],
  ['employeeName', 'Nama karyawan'],
  ['source', 'Sumber data'],
];
const RANGE_FIELDS = [
  { key: 'systolic', label: 'Sistolik', min: 50, max: 300, required: true },
  { key: 'diastolic', label: 'Diastolik', min: 20, max: 200, required: true },
  { key: 'temperature', label: 'Suhu', min: 20, max: 50, required: true },
  { key: 'spo2', label: 'SPO2', min: 50, max: 100, required: true },
  { key: 'weightKg', label: 'Berat badan', min: 20, max: 300, required: true },
  { key: 'heartRate', label: 'Detak jantung', min: 20, max: 250, required: false },
  { key: 'heightCm', label: 'Tinggi badan', min: 50, max: 250, required: false },
  { key: 'bmi', label: 'BMI', min: 5, max: 100, required: false },
];
const WEEKDAY_ID = { '1': 'Sen', '2': 'Sel', '3': 'Rab', '4': 'Kam', '5': 'Jum', '6': 'Sab', '7': 'Min' };

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function prop_(key) {
  try {
    return PropertiesService.getScriptProperties().getProperty(key);
  } catch (e) {
    return null;
  }
}

function apiKeyOk_(provided) {
  const expected = prop_('API_KEY') || DEFAULT_API_KEY;
  return String(provided === undefined || provided === null ? '' : provided) === expected;
}

function pinMap_() {
  const raw = prop_('PIN_MAP');
  if (raw) {
    try {
      const map = JSON.parse(raw);
      if (map && typeof map === 'object') return map;
    } catch (e) {}
  }
  const def = {};
  def[DEFAULT_PIN] = DEFAULT_USER;
  return def;
}

function ensureHeaders_(sh, headers) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  const cur = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  let changed = false;
  const merged = headers.map(function (h, i) {
    const v = cur[i];
    if (v === '' || v === null || v === undefined) {
      changed = true;
      return h;
    }
    return v;
  });
  if (changed) sh.getRange(1, 1, 1, merged.length).setValues([merged]);
}

function migrateEmployees_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow === 0) return;
  const header = sh.getRange(1, 1, 1, 8).getValues()[0];
  if (String(header[1]) === 'Nomor kartu RF') return;

  const idxId = header.findIndex(function(h) { return String(h) === 'id'; });
  const idxName = header.findIndex(function(h) { return String(h) === 'name'; });
  const idxCard = header.findIndex(function(h) { return String(h) === 'cardUid'; });
  if (idxId === -1 || idxName === -1 || idxCard === -1) return; // Unrecognized layout or missing card

  const idxUnit = header.findIndex(function(h) { return String(h) === 'unit'; });
  const idxShift = header.findIndex(function(h) { return String(h) === 'shift'; });
  const idxAge = header.findIndex(function(h) { return String(h) === 'age'; });
  const idxHeight = header.findIndex(function(h) { return String(h) === 'heightCm'; });
  const idxDob = header.findIndex(function(h) { return String(h) === 'dob'; });

  const dataRange = sh.getRange(1, 1, lastRow, 8);
  const rows = dataRange.getValues();
  const currentYear = new Date().getFullYear();

  const newRows = rows.map(function(r, i) {
    if (i === 0) return EMP_HEADERS.concat(['']);
    const shift = idxShift >= 0 ? String(r[idxShift] || '').trim() : '';
    let grup = '';
    if (/^Pagi$/i.test(shift)) grup = 'A';
    else if (/^Siang$/i.test(shift)) grup = 'B';
    else if (/^Malam$/i.test(shift)) grup = 'C';
    else if (/^[A-D]$/i.test(shift)) grup = shift.toUpperCase();

    let dob = idxDob >= 0 ? r[idxDob] : '';
    if (dob instanceof Date) {
      dob = Utilities.formatDate(dob, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (typeof dob === 'string' && dob.trim() !== '') {
      dob = dob.trim();
    } else if (idxAge >= 0 && typeof r[idxAge] === 'number') {
      dob = (currentYear - r[idxAge]) + '-01-01';
    } else {
      dob = '';
    }

    const cId = idxId >= 0 ? r[idxId] : '';
    const cCard = idxCard >= 0 ? r[idxCard] : '';
    const cName = idxName >= 0 ? r[idxName] : '';
    const cUnit = idxUnit >= 0 ? r[idxUnit] : '';
    const cHeight = idxHeight >= 0 ? r[idxHeight] : '';

    return [cId, cCard, cName, cUnit, grup, dob, cHeight, ''];
  });

  dataRange.setValues(newRows);
}

function sheet_(name, headers) {
  const id = prop_('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID;
  if (!id) return null;
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (name === 'Employees') migrateEmployees_(sh);
  ensureHeaders_(sh, headers);
  if (name === 'Employees') {
    ensureCardColText_(sh);
    cleanEmptyEmployeeRows_(sh);
  }
  return sh;
}

function ensureCardColText_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const fmt = sh.getRange(2, 2).getNumberFormat();
  if (fmt !== '@') sh.getRange(2, 2, last - 1, 1).setNumberFormat('@');
}

function cleanEmptyEmployeeRows_(sh) {
  const last = sh.getLastRow();
  if (last < 3) return;
  const cols = sh.getLastColumn();
  if (cols < 1) return;
  const vals = sh.getRange(2, 1, last - 1, cols).getValues();
  let boundary = 1;
  for (let i = 0; i < vals.length; i++) {
    if (vals[i].some(function(v) { return v !== '' && v !== null && v !== undefined; })) {
      boundary = i + 2;
    }
  }
  if (last > boundary) sh.deleteRows(boundary + 1, last - boundary);
}

function coerce_(v) {
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'number') return v;
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'string' && v.trim() !== '') {
    const n = Number(v);
    if (!isNaN(n)) return n;
  }
  return v;
}

function readRows_(sh, headers) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, headers.length).getValues();
  return values.map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) {
      if (h === 'Nomor kartu RF' || h === 'cardUid') {
        const v = row[i];
        obj[h] = (v === '' || v === null || v === undefined) ? null : String(v);
      } else {
        obj[h] = coerce_(row[i]);
      }
    });
    return obj;
  });
}

function readEmpRows_(sh) {
  const rows = readRows_(sh, EMP_HEADERS);
  return rows.map(function(r) {
    const o = {};
    EMP_KEYS.forEach(function(k, i) {
      o[k] = r[EMP_HEADERS[i]];
    });
    return o;
  });
}

function numOrNull_(v) {
  if (v === '' || v === null || v === undefined || v === 0) return null;
  const n = Number(v);
  return isFinite(n) ? n : null;
}

function calcStatus_(systolic, temperature) {
  return Number(systolic) < 140 && Number(temperature) < 37.5 ? 'FIT' : 'UNFIT';
}

function dateKey_(ts, tz) {
  const d = new Date(ts);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

function timeMs_(ts) {
  const t = Date.parse(ts);
  return isNaN(t) ? 0 : t;
}

function validateRecord_(d) {
  if (!d || typeof d !== 'object') return 'Data tidak valid';
  const missing = [];
  STR_REQUIRED.forEach(function (pair) {
    const v = d[pair[0]];
    if (v === undefined || v === null || String(v).trim() === '') missing.push(pair[1]);
  });
  if (missing.length) return 'Field wajib belum diisi: ' + missing.join(', ');
  if (isNaN(Date.parse(String(d.timestamp)))) return 'Timestamp tidak valid';
  for (let i = 0; i < RANGE_FIELDS.length; i++) {
    const f = RANGE_FIELDS[i];
    const raw = d[f.key];
    const isEmpty = raw === undefined || raw === null || raw === '' || Number(raw) === 0;
    if (isEmpty) {
      if (f.required) return f.label + ' wajib diisi (tidak boleh nol)';
      continue;
    }
    const n = Number(raw);
    if (!isFinite(n)) return f.label + ' bukan angka yang valid';
    if (n < f.min || n > f.max) return f.label + ' di luar rentang ' + f.min + '-' + f.max;
  }
  return null;
}

function validateAnswers_(d) {
  if (!d || typeof d !== 'object') return 'Kuesioner wajib dijawab lengkap (10 pertanyaan)';
  for (let i = 0; i < ANSWER_KEYS.length; i++) {
    const v = String(d[ANSWER_KEYS[i]] === undefined || d[ANSWER_KEYS[i]] === null ? '' : d[ANSWER_KEYS[i]]).trim().toLowerCase();
    if (v !== 'ya' && v !== 'tidak') return 'Kuesioner wajib dijawab lengkap (10 pertanyaan)';
  }
  return null;
}

function normAns_(v) {
  const s = String(v).trim().toLowerCase();
  return s === 'ya' ? 'Ya' : 'Tidak';
}

function makeRecordId_(rows) {
  const tz = Session.getScriptTimeZone();
  const prefix = 'TRX-' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd') + '-';
  const used = {};
  let maxSeq = 0;
  rows.forEach(function (r) {
    const id = String(r.id === null || r.id === undefined ? '' : r.id);
    if (id.indexOf(prefix) === 0) {
      used[id] = true;
      const seq = parseInt(id.slice(prefix.length), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  });
  let seq = maxSeq + 1;
  let id = prefix + String(seq).padStart(3, '0');
  while (used[id]) {
    seq++;
    id = prefix + String(seq).padStart(3, '0');
  }
  return id;
}

function actionLogin_(data) {
  const pin = data && data.pin !== undefined && data.pin !== null ? String(data.pin).trim() : '';
  if (!pin) return { ok: false, message: 'PIN wajib diisi' };
  const map = pinMap_();
  const user = map[pin];
  if (!user) return { ok: false, message: 'PIN salah' };
  return {
    ok: true,
    user: {
      name: user.name || '',
      role: user.role || '',
      email: user.email || '',
      initials: user.initials || '',
    },
  };
}

function actionSaveRecord_(data) {
  const err = validateRecord_(data);
  if (err) return { ok: false, message: err };
  const ansErr = validateAnswers_(data.answers);
  if (ansErr) return { ok: false, message: ansErr };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { ok: false, message: 'Server sedang sibuk — coba lagi sebentar.' };
  }
  try {
    const sh = sheet_('Records', RECORD_HEADERS);
    if (!sh) return { ok: false, message: 'SPREADSHEET_ID belum dikonfigurasi' };
    const rows = readRows_(sh, RECORD_HEADERS);
    const ts = String(data.timestamp);
    const empId = String(data.employeeId);
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i].timestamp) === ts && String(rows[i].employeeId) === empId) {
        return { ok: true, id: rows[i].id, status: rows[i].status, duplicate: true };
      }
    }
    const id = makeRecordId_(rows);
    const status = calcStatus_(data.systolic, data.temperature);
    const rec = {
      id: id,
      timestamp: ts,
      employeeId: empId,
      employeeName: String(data.employeeName),
      systolic: Number(data.systolic),
      diastolic: Number(data.diastolic),
      temperature: Number(data.temperature),
      spo2: Number(data.spo2),
      heartRate: numOrNull_(data.heartRate),
      weightKg: Number(data.weightKg),
      heightCm: numOrNull_(data.heightCm),
      bmi: numOrNull_(data.bmi),
      status: status,
      source: String(data.source),
      reason: data.reason ? String(data.reason) : '',
      examiner: data.examiner ? String(data.examiner) : '',
      examinerRole: data.examinerRole ? String(data.examinerRole) : '',
    };
    sh.appendRow(RECORD_HEADERS.map(function (h) {
      return rec[h] === null || rec[h] === undefined ? '' : rec[h];
    }));
    const ash = sheet_('Answers', ANSWERS_HEADERS);
    if (!ash) return { ok: false, message: 'SPREADSHEET_ID belum dikonfigurasi' };
    const ansRow = [id, rec.timestamp, rec.employeeId, rec.employeeName];
    ANSWER_KEYS.forEach(function (k) {
      ansRow.push(normAns_(data.answers[k]));
    });
    ash.appendRow(ansRow);
    return { ok: true, id: id, status: status };
  } finally {
    lock.releaseLock();
  }
}

function actionListRecords_() {
  const sh = sheet_('Records', RECORD_HEADERS);
  if (!sh) return { ok: false, message: 'SPREADSHEET_ID belum dikonfigurasi' };
  const rows = readRows_(sh, RECORD_HEADERS);
  rows.sort(function (a, b) {
    const diff = timeMs_(b.timestamp) - timeMs_(a.timestamp);
    if (diff !== 0) return diff;
    return String(b.id === null || b.id === undefined ? '' : b.id).localeCompare(String(a.id === null || a.id === undefined ? '' : a.id));
  });
  return { ok: true, records: rows.slice(0, 500) };
}

function seedEmployees_(sh) {
  if (sh.getLastRow() < 2) {
    DEFAULT_EMPLOYEES.forEach(function (row) {
      sh.appendRow(row);
    });
  }
}

function ageFromDob_(dob, now) {
  const p = String(dob).split('-');
  const y = Number(p[0]);
  const m = Number(p[1]);
  const d = Number(p[2]);
  let age = now.getFullYear() - y;
  const nm = now.getMonth() + 1;
  if (nm < m || (nm === m && now.getDate() < d)) age--;
  return age;
}

function makeEmployeeId_(rows) {
  let maxNum = 1000;
  const used = {};
  rows.forEach(function (r) {
    const id = String(r.id === null || r.id === undefined ? '' : r.id);
    if (/^K-\d+$/.test(id)) {
      used[id] = true;
      const n = parseInt(id.slice(2), 10);
      if (n > maxNum) maxNum = n;
    }
  });
  let n = maxNum + 1;
  let id = 'K-' + String(n).padStart(4, '0');
  while (used[id]) {
    n++;
    id = 'K-' + String(n).padStart(4, '0');
  }
  return id;
}

function actionListEmployees_() {
  const sh = sheet_('Employees', EMP_HEADERS);
  if (!sh) return { ok: false, message: 'SPREADSHEET_ID belum dikonfigurasi' };
  seedEmployees_(sh);
  const rows = readEmpRows_(sh);
  const now = new Date();
    rows.forEach(function(r) {
      r.age = r.dob ? ageFromDob_(r.dob, now) : null;
      if (r.cardUid) r.cardUid = String(r.cardUid).replace(/[-:\s]/g, '').toUpperCase();
    });
    return { ok: true, employees: rows };
}

function actionAddEmployee_(data) {
  if (!data || typeof data !== 'object') return { ok: false, message: 'Data tidak valid' };
  const name = data.name === undefined || data.name === null ? '' : String(data.name).trim();
  const unit = data.unit === undefined || data.unit === null ? '' : String(data.unit).trim();
  let grupRaw = data.grup === undefined || data.grup === null ? '' : String(data.grup).trim();
  let grup = '';
  if (/^[a-d]$/i.test(grupRaw)) grup = grupRaw.toUpperCase();
  else if (grupRaw.toLowerCase() === 'non shift') grup = 'Non Shift';
  const dob = data.dob === undefined || data.dob === null ? '' : String(data.dob).trim();
  const heightRaw = data.heightCm === undefined || data.heightCm === null ? '' : String(data.heightCm).trim();
  if (!name) return { ok: false, message: 'Nama lengkap wajib diisi' };
  if (!unit) return { ok: false, message: 'Unit wajib diisi' };
  if (!grup || (!/^[A-D]$/.test(grup) && grup !== 'Non Shift')) return { ok: false, message: 'Grup tugas wajib dipilih (A, B, C, D, atau Non Shift)' };
  if (!dob) return { ok: false, message: 'Tanggal lahir wajib diisi' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dob) || isNaN(Date.parse(dob))) {
    return { ok: false, message: 'Tanggal lahir tidak valid' };
  }
  const now = new Date();
  if (Date.parse(dob) > now.getTime()) return { ok: false, message: 'Tanggal lahir tidak valid' };
  if (!heightRaw) return { ok: false, message: 'Tinggi badan wajib diisi' };
  const height = Number(heightRaw);
  if (!isFinite(height)) return { ok: false, message: 'Tinggi badan bukan angka yang valid' };
  if (height < 50 || height > 250) return { ok: false, message: 'Tinggi badan di luar rentang 50-250' };
  const age = ageFromDob_(dob, now);
  if (age < 0 || age > 120) return { ok: false, message: 'Tanggal lahir tidak valid' };
  const sh = sheet_('Employees', EMP_HEADERS);
  if (!sh) return { ok: false, message: 'SPREADSHEET_ID belum dikonfigurasi' };
  seedEmployees_(sh);
  const rows = readEmpRows_(sh);
  const id = makeEmployeeId_(rows);
  const emp = { id: id, cardUid: null, name: name, unit: unit, grup: grup, dob: dob, heightCm: height, age: age };
  sh.appendRow([emp.id, '', emp.name, emp.unit, emp.grup, emp.dob, emp.heightCm]);
  return { ok: true, employee: emp };
}

function seedDutyOfficers_(sh) {
  if (sh.getLastRow() < 2) {
    DEFAULT_DUTY_OFFICERS.forEach(function (row) {
      sh.appendRow(row);
    });
  }
}

function actionListDutyOfficers_() {
  const sh = sheet_('DutyOfficers', DUTY_HEADERS);
  if (!sh) return { ok: false, message: 'SPREADSHEET_ID belum dikonfigurasi' };
  seedDutyOfficers_(sh);
  return { ok: true, officers: readRows_(sh, DUTY_HEADERS) };
}

function actionSaveDutyOfficer_(data) {
  if (!data || typeof data !== 'object') return { ok: false, message: 'Data tidak valid' };
  const id = data.id === undefined || data.id === null ? '' : String(data.id).trim();
  if (!id) return { ok: false, message: 'ID petugas wajib diisi' };
  
  const sh = sheet_('DutyOfficers', DUTY_HEADERS);
  if (!sh) return { ok: false, message: 'SPREADSHEET_ID belum dikonfigurasi' };
  seedDutyOfficers_(sh);
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { ok: false, message: 'Server sedang sibuk — coba lagi sebentar.' };
  }
  try {
    const rows = readRows_(sh, DUTY_HEADERS);
    let rowIndex = -1;
    let officer = null;
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i].id) === id) {
        rowIndex = i + 2; // Data mulai di baris 2
        officer = rows[i];
        break;
      }
    }
    if (rowIndex === -1) return { ok: false, message: 'Petugas tidak ditemukan' };
    
    let newName = data.name === undefined || data.name === null ? '' : String(data.name).trim();
    let newRole = data.role === undefined || data.role === null ? '' : String(data.role).trim();
    
    if (!newName) newName = String(officer.defaultName || '');
    if (!newRole) newRole = String(officer.defaultRole || '');
    
    if (newName.length > 80) return { ok: false, message: 'Nama terlalu panjang (maks 80 karakter)' };
    if (newRole !== 'Dokter' && newRole !== 'Perawat') return { ok: false, message: 'Role harus Dokter atau Perawat' };
    
    officer.name = newName;
    officer.role = newRole;
    
    sh.getRange(rowIndex, 2).setValue(newName);
    sh.getRange(rowIndex, 3).setValue(newRole);
    
    return { ok: true, officer: officer };
  } finally {
    lock.releaseLock();
  }
}

function actionSaveEmployeeCard_(data) {
  if (!data || typeof data !== 'object') return { ok: false, message: 'Data tidak valid' };
  const employeeId = data.employeeId === undefined || data.employeeId === null ? '' : String(data.employeeId).trim();
  let cardUid = data.cardUid === undefined || data.cardUid === null ? '' : String(data.cardUid);
  
  if (!employeeId) return { ok: false, message: 'ID karyawan wajib diisi' };
  
  if (cardUid !== '') {
    cardUid = cardUid.replace(/[-:\s]/g, '').toUpperCase();
    if (!/^[A-Z0-9]{4,32}$/.test(cardUid)) {
      return { ok: false, message: 'Format UID tidak valid (4-32 karakter huruf/angka)' };
    }
  }

  const sh = sheet_('Employees', EMP_HEADERS);
  if (!sh) return { ok: false, message: 'SPREADSHEET_ID belum dikonfigurasi' };
  seedEmployees_(sh);

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { ok: false, message: 'Server sedang sibuk — coba lagi sebentar.' };
  }
  
  try {
    const rows = readEmpRows_(sh);
    let rowIndex = -1;
    let employee = null;
    
    // Check uniqueness and find target
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (String(r.id) === employeeId) {
        rowIndex = i + 2;
        employee = r;
      } else if (cardUid !== '' && String(r.cardUid || '').replace(/[-:\s]/g, '').toUpperCase() === cardUid) {
        return { ok: false, message: 'Kartu sudah terdaftar untuk ' + (r.name || 'karyawan lain') + ' (' + r.id + ')' };
      }
    }
    
    if (rowIndex === -1) return { ok: false, message: 'Karyawan tidak ditemukan' };
    
      employee.cardUid = cardUid === '' ? null : cardUid;
      employee.age = employee.dob ? ageFromDob_(employee.dob, new Date()) : null;
      sh.getRange(rowIndex, 2).setNumberFormat('@').setValue(cardUid);
    
    return { ok: true, employee: employee };
  } finally {
    lock.releaseLock();
  }
}

function actionStats_() {
  const sh = sheet_('Records', RECORD_HEADERS);
  if (!sh) return { ok: false, message: 'SPREADSHEET_ID belum dikonfigurasi' };
  const rows = readRows_(sh, RECORD_HEADERS);
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const isToday = rows.filter(function (r) {
    return dateKey_(r.timestamp, tz) === today;
  });
  let totalToday = 0;
  let fitToday = 0;
  let unfitToday = 0;
  isToday.forEach(function (r) {
    totalToday++;
    if (r.status === 'UNFIT') unfitToday++;
    else if (r.status === 'FIT') fitToday++;
  });
  let avgMinutes = 0;
  if (isToday.length >= 2) {
    const sorted = isToday.slice().sort(function (a, b) {
      return timeMs_(a.timestamp) - timeMs_(b.timestamp);
    });
    let sum = 0;
    for (let i = 1; i < sorted.length; i++) {
      sum += (timeMs_(sorted[i].timestamp) - timeMs_(sorted[i - 1].timestamp)) / 1000;
    }
    avgMinutes = Math.round(sum / (sorted.length - 1));
  }
  const weekly = [];
  for (let k = 6; k >= 0; k--) {
    const day = new Date(now.getTime() - k * 86400000);
    const key = Utilities.formatDate(day, tz, 'yyyy-MM-dd');
    const dow = Utilities.formatDate(day, tz, 'u');
    let fit = 0;
    let unfit = 0;
    rows.forEach(function (r) {
      if (dateKey_(r.timestamp, tz) !== key) return;
      if (r.status === 'UNFIT') unfit++;
      else if (r.status === 'FIT') fit++;
    });
    weekly.push({ label: WEEKDAY_ID[dow] || '', fit: fit, unfit: unfit });
  }
  return { ok: true, totalToday: totalToday, fitToday: fitToday, unfitToday: unfitToday, avgMinutes: avgMinutes, weekly: weekly };
}

function doGet(e) {
  return jsonResponse_({ ok: true, service: 'SMART FIT2WORK API', version: '1.0.0', actions: API_ACTIONS });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return jsonResponse_({ ok: false, message: 'Request tidak valid' });
  }
  if (!body || typeof body !== 'object') {
    return jsonResponse_({ ok: false, message: 'Request tidak valid' });
  }
  const action = typeof body.action === 'string' ? body.action : '';
  const data = body.data && typeof body.data === 'object' ? body.data : {};
  if (!action) return jsonResponse_({ ok: false, message: 'Action tidak dikenal' });
  if (!apiKeyOk_(body.apiKey)) return jsonResponse_({ ok: false, message: 'API key tidak valid' });
  try {
    let res;
    if (action === 'login') res = actionLogin_(data);
    else if (action === 'saveRecord') res = actionSaveRecord_(data);
    else if (action === 'listRecords') res = actionListRecords_();
    else if (action === 'listEmployees') res = actionListEmployees_();
    else if (action === 'addEmployee') res = actionAddEmployee_(data);
    else if (action === 'stats') res = actionStats_();
    else if (action === 'listDutyOfficers') res = actionListDutyOfficers_();
    else if (action === 'saveDutyOfficer') res = actionSaveDutyOfficer_(data);
    else if (action === 'saveEmployeeCard') res = actionSaveEmployeeCard_(data);
    else res = { ok: false, message: 'Action tidak dikenal: ' + action };
    return jsonResponse_(res);
  } catch (err) {
    return jsonResponse_({ ok: false, message: 'Server error: ' + (err && err.message ? err.message : String(err)) });
  }
}
