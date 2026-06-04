#!/usr/bin/env node
/**
 * P4P Excel Merge + SK03 Pipeline
 * Step 1  — 6-month window in BE format
 * Step 2  — Google Drive: find month folder, list Excel files
 * Step 3  — Supabase: fetch person list + metadata
 * Step 4  — Compare Drive vs Supabase lists
 * Step 5  — Merge individual Excel files → one workbook per person sheet
 * Step 6  — Create SK03 Google Spreadsheet via Sheets API (mirrors GAS doPost)
 */

require('dotenv').config();

const { google }        = require('googleapis');
const { createClient }  = require('@supabase/supabase-js');
const { Readable }      = require('stream');
const fs                = require('fs');
const os                = require('os');
const path              = require('path');
const { spawnSync }     = require('child_process');
const ExcelJS           = require('exceljs');

// ═══════════════════════════════════════════════════════════════════
//  CONFIG
// ═══════════════════════════════════════════════════════════════════
const CONFIG = {
  google: {
    clientId:     process.env.GOOGLE_CLIENT_ID,
    clientSecret: process.env.GOOGLE_CLIENT_SECRET,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
  },
  supabase: {
    url: process.env.SUPABASE_URL,
    key: process.env.SUPABASE_KEY,
  },
  rootFolderId:   process.env.GOOGLE_ROOT_FOLDER_ID,
  outputFolderId: process.env.GOOGLE_MERGE_FOLDER_ID,
  sk03FolderId:   process.env.GOOGLE_SK03_FOLDER_ID,
  sk03TemplateId: process.env.GOOGLE_SK03_TEMPLATE_ID, // spreadsheet ID of SK03 template
};

// ─── Thai month names ────────────────────────────────────────────
const THAI_MONTHS = {
  1:'มกราคม', 2:'กุมภาพันธ์', 3:'มีนาคม', 4:'เมษายน',
  5:'พฤษภาคม', 6:'มิถุนายน', 7:'กรกฎาคม', 8:'สิงหาคม',
  9:'กันยายน', 10:'ตุลาคม', 11:'พฤศจิกายน', 12:'ธันวาคม',
};

const THAI_MONTH_ABBR = {
  1:'ม.ค.', 2:'ก.พ.', 3:'มี.ค.', 4:'เม.ย.',
  5:'พ.ค.', 6:'มิ.ย.', 7:'ก.ค.', 8:'ส.ค.',
  9:'ก.ย.', 10:'ต.ค.', 11:'พ.ย.', 12:'ธ.ค.',
};

// ─── Dept colours & sort order ───────────────────────────────────
// ORDER of keys defines sheet order. INTERN is last.
const DEPT_COLORS = {
  'กุมารเวชกรรม':                          '#C8D3B8',
  'จักษุวิทยา':                            '#EEE7D3',
  'จิตเวชและยาเสพติด':                     '#D9B4E2',
  'เทคนิคการแพทย์และพยาธิวิทยาคลินิก':     '#F7D5C2',
  'นิติเวช':                               '#B1D1E3',
  'พยาธิวิทยากายวิภาค':                    '#BEAFE1',
  'ผู้ป่วยนอก':                             '#EAD7C3',
  'รังสีวิทยา':                             '#B0E0E6',
  'วิสัญญีวิทยา':                           '#FFFAE6',
  'เวชกรรมฟื้นฟู':                          '#F3BBDC',
  'เวชกรรมสังคม':                           '#DAEDF7',
  'เวชศาสตร์ฉุกเฉิน':                       '#9ACFC4',
  'ศัลยกรรม':                              '#FFE5EC',
  'ศัลยกรรมออร์โธปิดิกส์':                  '#CFDFEF',
  'สูติ-นรีเวชกรรม':                        '#F9ECD9',
  'โสต ศอ นาสิก':                           '#FADA5E',
  'อาชีวเวชกรรม':                           '#FFDDE4',
  'อายุรกรรม':                              '#A3DDBD',
  'INTERN':                               '#ECC1D1',
};
// ─── Management allowance data ───────────────────────────────────
// Key: firstname + " " + lastname (no prefix). Provides col M (staff) and col E + I (dept).
const MANAGEMENT_DATA = [
  { name: 'ศิริพันธ์ บุญโต',             remark: 'รองผู้อำนวยการ',      amount: 7000 },
  { name: 'นิธินันท์ สร้อยอากาศ',         remark: 'หัวหน้ากลุ่มงาน',     amount: 1000 },
  { name: 'อภิสรา กูลวงศ์ธนโรจน์',       remark: 'หัวหน้ากลุ่มงาน',     amount: 1000 },
  { name: 'ศิรดา แสงไพบูลย์',            remark: 'หัวหน้ากลุ่มงาน',     amount: 1000 },
  { name: 'ณัฏฐพัชร จันทร์หอม',          remark: 'ผู้ช่วยผู้อำนวยการ',   amount: 3000 },
  { name: 'ลักขณา จิราพงษ์',             remark: 'ผู้ช่วยผู้อำนวยการ',   amount: 3000 },
  { name: 'อรวรรณ อุตราวิสิทธิกุล',      remark: 'หัวหน้ากลุ่มงาน',     amount: 1000 },
  { name: 'ดวงพร เกื้อกูลเกียรติ',        remark: 'หัวหน้ากลุ่มงาน',     amount: 1000 },
  { name: 'พงศ์พจน์ ฉุยฉาย',             remark: 'ผู้ช่วยผู้อำนวยการ',   amount: 3000 },
  { name: 'ทรงพล โพธิ์สุวรรณ',           remark: 'ประธาน PCT มะเร็ง',   amount:  800 },
  { name: 'ฉัตรดาว สุจริต',              remark: 'ผู้ช่วยผู้อำนวยการ',   amount: 3000 },
  { name: 'พิสิษฐ์ เลิศเชาวพัฒน์',       remark: 'หัวหน้ากลุ่มงาน',     amount: 1000 },
  { name: 'วราวุธ เมธีศิริวัฒน์',         remark: 'รองผู้อำนวยการ',      amount: 7000 },
  { name: 'ศุภศรัณย์ ศุภพัฒนพงศ์',        remark: 'รองผู้อำนวยการ',      amount: 7000 },
  { name: 'ธิรัญฎา สุทธิพงศ์',            remark: 'ผู้ช่วยผู้อำนวยการ',   amount: 3000 },
  { name: 'ธญาภร ลิขิตธรรมากุล',          remark: 'หัวหน้ากลุ่มงาน',     amount: 1000 },
  { name: 'นฤวัต เกสรสุคนธ์',            remark: 'หัวหน้ากลุ่มงาน',     amount: 1000 },
  { name: 'พยุงศักดิ์ ศักดาภิพาณิชย์',   remark: 'ประธาน PCT ENT',     amount:  800 },
  { name: 'อัญชลี ชุ่มแจ่ม',             remark: 'ผู้ช่วยผู้อำนวยการ',   amount: 3000 },
  { name: 'สงกรานต์ ชุนหวัฒนา',          remark: 'หัวหน้ากลุ่มงาน',     amount: 1000 },
  { name: 'โอภาส ไชยมหาพฤกษ์',           remark: 'หัวหน้ากลุ่มงาน',     amount: 1000 },
  { name: 'เกษมศักดิ์ จึงจรูญ',           remark: 'หัวหน้ากลุ่มงาน',     amount: 1000 },
];

// O(1) lookup: "firstname lastname" → { remark, amount }
const MGMT_LOOKUP = Object.fromEntries(
  MANAGEMENT_DATA.map(d => [normaliseName(d.name), { remark: d.remark, amount: d.amount }])
);

// Dept-head physicians eligible for extra score of 1,320 in col I
const DEPT_HEAD_SET = new Set([
  'ฉัตรดาว สุจริต',
  'ศิริพันธ์ บุญโต',
  'อรวรรณ อุตราวิสิทธิกุล',
  'นฤวัต เกสรสุคนธ์',
  'ดวงพร เกื้อกูลเกียรติ',
  'พิสิษฐ์ เลิศเชาวพัฒน์',
  'นิธินันท์ สร้อยอากาศ',
  'ธญาภร ลิขิตธรรมากุล',
  'ธิรัญฎา สุทธิพงศ์',
  'ศิรดา แสงไพบูลย์',
  'อภิสรา กูลวงศ์ธนโรจน์',
  'สงกรานต์ ชุนหวัฒนา',
  'โอภาส ไชยมหาพฤกษ์',
  'เกษมศักดิ์ จึงจรูญ',
].map(normaliseName));

/** Return management entry for a person, or null */
function getMgmt(person) {
  const key = normaliseName(`${person.firstname} ${person.lastname}`);
  return MGMT_LOOKUP[key] ?? null;
}

const DEPT_ORDER     = Object.fromEntries(Object.keys(DEPT_COLORS).map((k, i) => [k, i]));
const DEPT_ORDER_MAX = Object.keys(DEPT_COLORS).length;

// ─── Supabase column name mapping ────────────────────────────────
const SB = {
  prefix:    'prefix',
  position:  'position',
  level:     'level',
  rank:      'rank',            // col F in staff sheet and col C suffix in dept sheets
  type:      'type',             // col G in staff/intern sheet, col D in dept sheets
  std:       'standard_score',
  boss:      'boss_score',
  perf:      'performance_score', // fallback col J
  score:     'score',               // col J override — used when not null
  index:     'index',           // for updateSupabaseRowNum
};

// ═══════════════════════════════════════════════════════════════════
//  STEP 1 — 6-month window
// ═══════════════════════════════════════════════════════════════════
function getTargetMonths() {
  const now = new Date();
  return Array.from({ length: 6 }, (_, i) => {
    const d      = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const beYear = d.getFullYear() + 543;
    const month  = d.getMonth() + 1;
    const key    = `${beYear}_${String(month).padStart(2, '0')}`;
    return { key, beYear, month };
  });
}

// ═══════════════════════════════════════════════════════════════════
//  Utility
// ═══════════════════════════════════════════════════════════════════
function stripExt(filename)    { return filename.replace(/\.(xlsx|xls)$/i, '').trim(); }
function normaliseName(name)   { return (name ?? '').trim().replace(/\s+/g, ' '); }
function toSheetName(name)     { return name.replace(/[\\/:?*[\]]/g, '').substring(0, 31).trim(); }

function hexToArgb(hex) {
  if (!hex) return undefined;
  const c = hex.replace('#', '').toUpperCase();
  return c.length === 6 ? `FF${c}` : undefined;
}

/** Convert #RRGGBB to Sheets API {red, green, blue} (0-1 range) */
function hexToRgb(hex) {
  const c = hex.replace('#', '');
  return {
    red:   parseInt(c.slice(0, 2), 16) / 255,
    green: parseInt(c.slice(2, 4), 16) / 255,
    blue:  parseInt(c.slice(4, 6), 16) / 255,
  };
}

/**
 * Convert a monthKey (e.g. "2568_12") to the Thai sheet name base
 * used by the GAS scripts (e.g. "ธ.ค. 68"), matching the merge step's key.
 */
function monthKeyToSheetBase(monthKey) {
  const [beYear, monthStr] = monthKey.split('_');
  const month = parseInt(monthStr, 10);
  return `${THAI_MONTH_ABBR[month]} ${beYear.slice(2)}`;  // e.g. "ธ.ค. 68"
}

function log(msg, level = 'info') {
  (level === 'warn' ? console.error : console.log)((level === 'warn' ? '⚠  ' : '') + msg);
}

/** ms delay */
const sleep = ms => new Promise(r => setTimeout(r, ms));

/**
 * Retry an async fn on Google quota (429) or server (5xx) errors.
 * Exponential backoff starting at 3 s, capped at 60 s.
 */
async function withRetry(fn, maxRetries = 7) {
  let delay = 3000;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return await fn();
    } catch (err) {
      const status  = err?.response?.status ?? err?.status ?? 0;
      const msg     = err?.message ?? '';
      const isQuota = status === 429 || msg.includes('Quota exceeded') || msg.includes('RESOURCE_EXHAUSTED');
      const isServer = status >= 500 && status < 600;
      if ((isQuota || isServer) && attempt < maxRetries) {
        const wait = delay + Math.random() * 1000;
        log(`  [Retry] ${isQuota ? 'Quota' : 'Server'} — waiting ${(wait/1000).toFixed(1)}s (attempt ${attempt+1}/${maxRetries})`, 'warn');
        await sleep(wait);
        delay = Math.min(delay * 2, 60000);
      } else {
        throw err;
      }
    }
  }
}

function sbVal(row, colName, defaultVal = null) {
  const v = row[colName];
  return (v !== null && v !== undefined && v !== '') ? v : defaultVal;
}

// ═══════════════════════════════════════════════════════════════════
//  Google API clients
// ═══════════════════════════════════════════════════════════════════
function createAuth() {
  const auth = new google.auth.OAuth2(CONFIG.google.clientId, CONFIG.google.clientSecret);
  auth.setCredentials({ refresh_token: CONFIG.google.refreshToken });
  return auth;
}

function createDriveClient() {
  return google.drive({ version: 'v3', auth: createAuth() });
}

function createSheetsClient() {
  return google.sheets({ version: 'v4', auth: createAuth() });
}

// ═══════════════════════════════════════════════════════════════════
//  Google Drive helpers
// ═══════════════════════════════════════════════════════════════════
async function driveListAll(drive, params) {
  const items = [];
  let pageToken;
  do {
    const res = await withRetry(() => drive.files.list({ ...params, pageToken }));
    items.push(...(res.data.files ?? []));
    pageToken = res.data.nextPageToken;
  } while (pageToken);
  return items;
}

async function listFolders(drive, parentId) {
  return driveListAll(drive, {
    q: `'${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
    fields: 'nextPageToken, files(id, name)', pageSize: 100,
  });
}

async function listExcelFiles(drive, folderId) {
  const mimes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
  ].map(m => `mimeType='${m}'`).join(' or ');
  return driveListAll(drive, {
    q: `'${folderId}' in parents and (${mimes}) and trashed=false`,
    fields: 'nextPageToken, files(id, name)', pageSize: 100,
  });
}

async function downloadFile(drive, fileId) {
  const res = await withRetry(() => drive.files.get({ fileId, alt: 'media' }, { responseType: 'arraybuffer' }));
  return Buffer.from(res.data);
}

async function uploadFileToDrive(drive, folderId, fileName, buffer) {
  const mime     = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const safeName = fileName.replace(/'/g, "\\'");
  const existing = await driveListAll(drive, {
    q: `'${folderId}' in parents and name='${safeName}' and trashed=false`,
    fields: 'nextPageToken, files(id, name)', pageSize: 10,
  });
  if (existing.length > 0) {
    const [first, ...dupes] = existing;
    for (const d of dupes) {
      await withRetry(() => drive.files.delete({ fileId: d.id }));
      log(`  [Drive] Deleted duplicate id: ${d.id}`, 'warn');
    }
    log(`  [Drive] Overwriting "${fileName}" (id: ${first.id})`);
    const res = await withRetry(() => drive.files.update({
      fileId: first.id,
      requestBody: { name: fileName },
      media: { mimeType: mime, body: Readable.from([buffer]) },
      fields: 'id, name, webViewLink',
    }));
    return res.data;
  }
  const res = await withRetry(() => drive.files.create({
    requestBody: { name: fileName, parents: [folderId], mimeType: mime },
    media: { mimeType: mime, body: Readable.from([buffer]) },
    fields: 'id, name, webViewLink',
  }));
  return res.data;
}

// ═══════════════════════════════════════════════════════════════════
//  STEP 2 — Google Drive: find month folder
// ═══════════════════════════════════════════════════════════════════
async function getDriveMonthData(drive, { beYear, month }) {
  const yearFolders = await listFolders(drive, CONFIG.rootFolderId);
  const yearFolder  = yearFolders.find(f => f.name === String(beYear));
  if (!yearFolder) { log(`  [Drive] Year folder "${beYear}" not found`, 'warn'); return null; }

  const thai       = THAI_MONTHS[month];
  const candidates = [`${month} - ${thai}`, `${String(month).padStart(2,'0')} - ${thai}`];
  const monthFolders = await listFolders(drive, yearFolder.id);
  const monthFolder  = monthFolders.find(f => candidates.includes(f.name));
  if (!monthFolder) { log(`  [Drive] Month folder not found (tried: ${candidates.join(' | ')})`, 'warn'); return null; }

  const files = await listExcelFiles(drive, monthFolder.id);
  if (files.length === 0) { log(`  [Drive] No Excel files in "${monthFolder.name}"`, 'warn'); return null; }

  const names = files.map(f => normaliseName(stripExt(f.name)));
  return { names, files, folderId: monthFolder.id };
}

// ═══════════════════════════════════════════════════════════════════
//  STEP 3 — Supabase
// ═══════════════════════════════════════════════════════════════════
async function getSupabaseMonthData(supabase, tableKey) {
  const { data, error } = await supabase.from(tableKey).select('*');
  if (error) {
    if (error.code === '42P01' || error.message?.includes('does not exist')) {
      log(`  [Supabase] Table "${tableKey}" not found`, 'warn'); return null;
    }
    throw new Error(`Supabase error (${tableKey}): ${error.message}`);
  }
  if (!data || data.length === 0) { log(`  [Supabase] Table "${tableKey}" is empty`, 'warn'); return null; }

  return data.map(r => ({
    fullname:   normaliseName(`${r.firstname ?? ''} ${r.lastname ?? ''}`),
    firstname:  r.firstname  ?? '',
    lastname:   r.lastname   ?? '',
    department: (r.department ?? '').trim(),
    prefix:     sbVal(r, SB.prefix,   ''),
    position:   sbVal(r, SB.position, 'นายแพทย์'),
    level:      sbVal(r, SB.level,    ''),
    rank:       sbVal(r, SB.rank,     ''),
    type:       sbVal(r, SB.type,     ''),
    std_score:  sbVal(r, SB.std,      2200),
    boss_score: sbVal(r, SB.boss,     0),
    perf_score: sbVal(r, SB.perf,     0),
    score:      sbVal(r, SB.score,    null), // overrides perf_score in col J when not null
    index:      sbVal(r, SB.index,    null),
  }));
}

// ═══════════════════════════════════════════════════════════════════
//  STEP 4 — Compare lists
// ═══════════════════════════════════════════════════════════════════
function listsMatch(driveNames, supabasePersons) {
  const normDrive = driveNames.map(normaliseName);
  const normSB    = supabasePersons.map(p => normaliseName(p.fullname));
  if (normDrive.length !== normSB.length) {
    log(`  [Compare] Length mismatch — Drive: ${normDrive.length}, Supabase: ${normSB.length}`, 'warn');
    const setSB    = new Set(normSB);
    const setDrive = new Set(normDrive);
    normDrive.forEach(n => { if (!setSB.has(n))    log(`    ✗ Drive but not SB: "${n}"`, 'warn'); });
    normSB.forEach(n   => { if (!setDrive.has(n)) log(`    ✗ SB but not Drive: "${n}"`, 'warn'); });
    return false;
  }
  const s1 = [...normDrive].sort(), s2 = [...normSB].sort();
  let ok = true;
  for (let i = 0; i < s1.length; i++) {
    if (s1[i] !== s2[i]) { log(`  [Compare] Mismatch: "${s1[i]}" ≠ "${s2[i]}"`, 'warn'); ok = false; }
  }
  return ok;
}

// ═══════════════════════════════════════════════════════════════════
//  STEP 4.5 — Fill missing scores from Excel (ported from claude-analyst.js)
// ═══════════════════════════════════════════════════════════════════

// ── Score extraction constants ────────────────────────────────────
const GRAND_TOTAL_LABELS = [
  'รวมแต้มทั้งหมด', 'รวมคะแนนทั้งหมด', 'รวมทั้งสิ้น', 'ยอดรวมทั้งหมด',
  'รวมทั้งหมด', 'คะแนนรวมทั้งหมด',
];
const SUBTOTAL_LABELS = [
  'รวมคะแนน', 'รวมแต้ม', 'คะแนนรวม', 'ผลรวม', 'รวม',
];
const TOTAL_LABELS = [...GRAND_TOTAL_LABELS, ...SUBTOTAL_LABELS];

function isYearLike(n) {
  if (n >= 1900 && n <= 2099) return true;
  if (n >= 2400 && n <= 2699 && Number.isInteger(n)) return true;
  return false;
}

function toNum(val) {
  if (val === null || val === undefined || val === '') return NaN;
  if (typeof val === 'number')  return val;
  if (typeof val === 'boolean') return NaN;
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return NaN;
  return parseFloat(s.replace(/,/g, ''));
}

function numsFromText(val) {
  const s = String(val ?? '').replace(/,/g, '');
  return [...s.matchAll(/\d+(?:\.\d+)?/g)]
    .map(m => parseFloat(m[0]))
    .filter(n => !isNaN(n) && n > 0 && !isYearLike(n));
}

function collectScoreCandidates(rows) {
  const results = [];
  for (const row of rows) {
    for (const val of Object.values(row)) {
      const n = toNum(val);
      if (isNaN(n) || n <= 0 || isYearLike(n)) continue;
      results.push(n);
    }
  }
  return results;
}

function extractScoreFromRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return { score: null, method: 'no rows' };

  // Pass 1: grand-total labels (search all columns — these never appear as column headers)
  const grandCandidates = [];
  for (const row of rows) {
    const allValues = Object.values(row).map(v => String(v ?? ''));
    const labelCells = allValues.filter(s => GRAND_TOTAL_LABELS.some(lbl => s.includes(lbl)));
    if (labelCells.length > 0) {
      grandCandidates.push(...collectScoreCandidates([row]));
      grandCandidates.push(...labelCells.flatMap(s => numsFromText(s)));
    }
  }
  if (grandCandidates.length > 0) {
    return { score: Math.max(...grandCandidates), method: 'grand-total label row' };
  }

  // Pass 2: sub-total labels (first 3 columns only — avoids col-header false matches)
  const subCandidates = [];
  for (const row of rows) {
    const firstThree = ['col_1', 'col_2', 'col_3'].map(k => String(row[k] ?? ''));
    if (firstThree.some(s => SUBTOTAL_LABELS.some(lbl => s.includes(lbl)))) {
      subCandidates.push(...collectScoreCandidates([row]));
    }
  }
  if (subCandidates.length > 0) {
    const subMax  = Math.max(...subCandidates);
    const allNums = collectScoreCandidates(rows);
    const sheetMax = allNums.length > 0 ? Math.max(...allNums) : subMax;
    if (sheetMax > subMax) return { score: sheetMax, method: 'largest in sheet (exceeds sub-total rows)' };
    return { score: subMax, method: 'sub-total label row' };
  }

  // Pass 3: largest valid number in the whole sheet
  const all = collectScoreCandidates(rows);
  if (all.length > 0) return { score: Math.max(...all), method: 'largest in sheet' };

  // Pass 4: weight × day-count (last resort for formula-only sheets)
  let computedTotal = 0;
  for (const row of rows) {
    const isLabel = ['col_1', 'col_2', 'col_3']
      .some(k => TOTAL_LABELS.some(lbl => String(row[k] ?? '').includes(lbl)));
    if (isLabel) continue;
    const weightRaw = row['col_3'];
    if (weightRaw == null) continue;
    let weight = typeof weightRaw === 'number'
      ? weightRaw
      : parseFloat(String(weightRaw).replace(/,/g, '').match(/^(\d+\.?\d*)/)?.[1] ?? 'NaN');
    if (isNaN(weight) || weight <= 0) continue;
    let daySum = 0;
    for (let d = 6; d <= 36; d++) {
      const v = toNum(row[`col_${d}`]);
      if (!isNaN(v) && v > 0) daySum += v;
    }
    if (daySum > 0) computedTotal += weight * daySum;
  }
  if (computedTotal > 0) return { score: computedTotal, method: 'weight × day-count' };

  return { score: null, method: 'no candidates found' };
}

function resolveScore(rows) {
  const { score: jsScore, method: jsMethod } = extractScoreFromRows(rows);

  const isGrandRow = row =>
    Object.values(row).some(v => GRAND_TOTAL_LABELS.some(lbl => String(v ?? '').includes(lbl)));
  const isSubRow = row =>
    ['col_1', 'col_2', 'col_3'].some(k => SUBTOTAL_LABELS.some(lbl => String(row[k] ?? '').includes(lbl)));
  const rowNums = row =>
    Object.values(row).map(toNum).filter(n => !isNaN(n) && n > 0 && !isYearLike(n));

  // Detect: grand-total label row present but holds no numbers (uncached formula)
  const grandRowEmpty = rows.some(row => isGrandRow(row) && rowNums(row).length === 0);
  if (!grandRowEmpty) return { score: jsScore, method: jsMethod };

  const populated = rows.filter(r => isSubRow(r) && rowNums(r).length > 0);
  const empty     = rows.filter(r => isSubRow(r) && rowNums(r).length === 0);

  // Tier 1: sum max from each populated sub-total row
  const subtotalSum = populated.reduce((s, r) => s + Math.max(...rowNums(r)), 0);

  // Tier 2: when some sub-totals also uncached, sum score column from data rows
  let dataRowSum = 0;
  if (populated.length > 0 && empty.length > 0) {
    let scoreColIndex = -1;
    for (const row of populated) {
      const indices = Object.keys(row)
        .filter(k => /^col_\d+$/.test(k) && !isNaN(toNum(row[k])) && toNum(row[k]) > 0)
        .map(k => parseInt(k.slice(4)));
      if (indices.length > 0) scoreColIndex = Math.max(scoreColIndex, Math.max(...indices));
    }
    if (scoreColIndex > 0) {
      const scoreColKey = `col_${scoreColIndex}`;
      for (const row of rows) {
        if (isSubRow(row) || isGrandRow(row)) continue;
        const n = toNum(row[scoreColKey]);
        if (!isNaN(n) && n > 0 && !isYearLike(n)) dataRowSum += n;
      }
    }
  }

  const best = Math.max(subtotalSum, dataRowSum);
  if (best > 0 && best > (jsScore ?? 0)) {
    const method = dataRowSum >= subtotalSum
      ? 'sum of score-column data rows (sub-totals partially uncached)'
      : 'sum of sub-total rows (grand-total formula uncached)';
    return { score: best, method };
  }
  return { score: jsScore, method: jsMethod };
}

/** Convert an Excel buffer to the { col_1, col_2, … } row format used by resolveScore */
async function excelToRows(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];
  if (!ws) return [];

  const rows = [];
  ws.eachRow({ includeEmpty: false }, row => {
    const obj = {};
    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      let val = cell.value;
      if (val !== null && typeof val === 'object') {
        if (val instanceof Date)           val = val.toISOString();
        else if (val.richText)             val = val.richText.map(r => r.text).join('');
        else if (val.result !== undefined) {
          val = val.result;
          // result itself can be a Date (date formula) or error object {error:'#REF!'}
          if (val instanceof Date)                          val = val.toISOString();
          else if (val !== null && typeof val === 'object') val = null;
        }
        else if (val.formula !== undefined) val = null;  // formula with no cached result
        else if (val.text !== undefined)    val = val.text;
        else                               val = String(val);
      }
      obj[`col_${colNum}`] = val;
    });
    if (Object.values(obj).some(v => v !== null && v !== undefined && v !== '')) {
      rows.push(obj);
    }
  });
  return rows;
}

/** Step 4.5 — for each person with score=null, extract sum score from their Drive Excel */
async function fillMissingScores(drive, driveFiles, sbData, monthKey, supabase) {
  const nullScorePersons = sbData.filter(p => p.score === null);
  if (nullScorePersons.length === 0) {
    log('  [Score] All persons already have scores');
    return;
  }
  log(`  [Score] ${nullScorePersons.length} person(s) with null score — extracting from Excel…`);

  const driveMap = {};
  for (const file of driveFiles) {
    driveMap[normaliseName(stripExt(file.name))] = file;
  }

  let updated = 0;
  for (const person of nullScorePersons) {
    const normName = normaliseName(person.fullname);
    const file = driveMap[normName];
    if (!file) {
      log(`  [Score] ⚠ No Drive file for "${normName}" — skipping`, 'warn');
      continue;
    }
    log(`  [Score] ↓ ${person.fullname}`);
    try {
      const buffer = await downloadFile(drive, file.id);
      const rows   = await excelToRows(buffer);
      const { score, method } = resolveScore(rows);
      if (score === null || score <= 0) {
        log(`  [Score] ⚠ "${normName}" — could not extract (${method})`, 'warn');
        continue;
      }
      log(`  [Score]   → ${score.toFixed(2)}  (${method})`);

      if (person.index !== null) {
        const { error } = await supabase
          .from(monthKey)
          .update({ [SB.score]: score })
          .eq(SB.index, person.index);
        if (error) {
          log(`  [Score] ⚠ Supabase update failed for "${normName}": ${error.message}`, 'warn');
        } else {
          person.score = score;  // keep sbData consistent for Step 5+
          updated++;
        }
      }
    } catch (err) {
      log(`  [Score] ⚠ Error for "${normName}": ${err.message}`, 'warn');
    }
  }
  log(`  [Score] ✓ ${updated}/${nullScorePersons.length} scores extracted and saved`);
}

// ═══════════════════════════════════════════════════════════════════
//  STEP 5 — Merge Excel files (via Python/openpyxl)
// ═══════════════════════════════════════════════════════════════════

function sortByDeptThenName(a, b, deptMap) {
  const dA   = deptMap[a.normName] ?? '';
  const dB   = deptMap[b.normName] ?? '';
  const idxA = DEPT_ORDER[dA] ?? DEPT_ORDER_MAX;
  const idxB = DEPT_ORDER[dB] ?? DEPT_ORDER_MAX;
  if (idxA !== idxB) return idxA - idxB;
  return a.normName.localeCompare(b.normName, 'th');
}

/** Merge all Excel files → one workbook with one sheet per person (via Python/openpyxl) */
async function mergeAndUpload(drive, driveFiles, supabasePersons, monthKey) {
  const deptMap = {};
  supabasePersons.forEach(p => { deptMap[normaliseName(p.fullname)] = p.department; });

  const filesWithName = driveFiles.map(f => ({
    ...f, origName: stripExt(f.name), normName: normaliseName(stripExt(f.name)),
  }));
  filesWithName.sort((a, b) => sortByDeptThenName(a, b, deptMap));

  log('\n  [Merge] Sorted order:');
  filesWithName.forEach((f, i) =>
    log(`    ${String(i+1).padStart(2)}. ${f.origName.padEnd(32)} [${deptMap[f.normName] ?? '?'}]`));

  // Create temp directory for this run
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'p4p_merge_'));
  log(`\n  [Merge] Temp dir: ${tmpDir}`);

  try {
    const manifestEntries = [];
    const usedSheetNames  = new Set();

    // Download each file into the temp directory
    for (const file of filesWithName) {
      log(`\n  [Merge] ↓ ${file.origName}`);
      const buffer   = await downloadFile(drive, file.id);
      const filePath = path.join(tmpDir, `${file.id}.xlsx`);
      fs.writeFileSync(filePath, buffer);

      let sheetName = toSheetName(file.origName);
      if (usedSheetNames.has(sheetName)) {
        let c = 2;
        while (usedSheetNames.has(`${sheetName}_${c}`)) c++;
        const suffix = `_${c}`;
        sheetName = `${sheetName.substring(0, 31 - suffix.length)}${suffix}`;
      }
      usedSheetNames.add(sheetName);

      const dept    = deptMap[file.normName] ?? '';
      const tabArgb = hexToArgb(DEPT_COLORS[dept]) ?? '';
      manifestEntries.push({ path: filePath, sheet_name: sheetName, tab_color: tabArgb });
    }

    if (manifestEntries.length === 0) {
      log('  [Merge] ⚠ No files — nothing to upload', 'warn');
      return null;
    }

    // Write manifest for Python script
    const manifestPath = path.join(tmpDir, 'manifest.json');
    const outputPath   = path.join(tmpDir, 'merged.xlsx');
    fs.writeFileSync(manifestPath, JSON.stringify({ files: manifestEntries }, null, 2), 'utf8');

    // Invoke Python merge script
    const scriptPath = path.join(__dirname, 'merge.py');
    log('\n  [Merge] Running Python/openpyxl merge…');
    const py = spawnSync(
      process.platform === 'win32' ? 'python' : 'python3',
      [scriptPath, manifestPath, outputPath],
      { encoding: 'utf8', stdio: 'pipe' }
    );
    if (py.stdout) log(py.stdout.trimEnd());
    if (py.stderr) log(py.stderr.trimEnd(), 'warn');
    if (py.status !== 0) throw new Error(`merge.py exited with code ${py.status}`);

    // Upload merged file to Drive
    const buf      = fs.readFileSync(outputPath);
    const name     = `merged_${monthKey}.xlsx`;
    const sheetCnt = manifestEntries.length;
    log(`\n  [Merge] Uploading "${name}" — ${sheetCnt} sheets…`);
    const uploaded = await uploadFileToDrive(drive, CONFIG.outputFolderId, name, buf);
    log(`  [Merge] ✓ id: ${uploaded.id}\n           🔗 ${uploaded.webViewLink}`);
    return uploaded;

  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
    log('  [Merge] Temp dir cleaned up');
  }
}

// ═══════════════════════════════════════════════════════════════════
//  STEP 6 — Create SK03 Google Spreadsheet (mirrors GAS doPost × 3)
//  Sheet order: staff → intern → one per dept (DEPT_COLORS order)
// ═══════════════════════════════════════════════════════════════════

// ─── Sheets API helpers ───────────────────────────────────────────
/** Grid range helper — all indices 0-based, r2/c2 inclusive */
function gridRange(sheetId, r1, c1, r2, c2) {
  return { sheetId, startRowIndex: r1, endRowIndex: r2 + 1, startColumnIndex: c1, endColumnIndex: c2 + 1 };
}

const SOLID = { style: 'SOLID', width: 1 };

function bordersReq(sheetId, r1, c1, r2, c2, inner = false) {
  return {
    updateBorders: {
      range: gridRange(sheetId, r1, c1, r2, c2),
      top: SOLID, bottom: SOLID, left: SOLID, right: SOLID,
      ...(inner ? { innerHorizontal: SOLID, innerVertical: SOLID } : {}),
    },
  };
}

// ─── Overall sheet helpers (staff + intern share the same template) ─

/**
 * Build the 22-column data rows for an overall (staff/intern) sheet.
 * Mirrors the appendRow + setFormula loop in both GAS scripts.
 */
function buildOverallRows(persons, S, isIntern) {
  const rateCell = isIntern ? 'Y13' : 'Y14';
  return persons.map((p, idx) => {
    const r = S + idx;
    return [
      idx + 1,                                               // A: sequence
      p.prefix    || '',                                     // B
      p.firstname || '',                                     // C
      p.lastname  || '',                                     // D
      p.position  || 'นายแพทย์',                            // E
      p.rank      || '',                                     // F: rank (supabase 'rank')
      p.type      || '',                                     // G: type (supabase 'type')
      p.std_score  ?? 2200,                                  // H
      DEPT_HEAD_SET.has(normaliseName(`${p.firstname} ${p.lastname}`)) ? 1320 : 0, // I: dept-head extra score
      (p.score ?? p.perf_score ?? 0) - (DEPT_HEAD_SET.has(normaliseName(`${p.firstname} ${p.lastname}`)) ? 1320 : 0), // J: score minus dept-head extra (already in col I)
      `=I${r}+J${r}`,                                       // K
      `=K${r}-H${r}`,                                       // L
      getMgmt(p)?.amount ?? p.boss_score ?? 0,              // M: mgmt amount or boss score (moved from I)
      `=ROUNDDOWN(V${r}*${rateCell},0)`,                    // N
      `=M${r}+N${r}`,                                       // O
      ' ',                                                   // P: remark
      `=IF(L${r}>0,IF(L${r}>Q7,0.90,0.00),0.00)`,         // Q
      `=IF(L${r}>R7,0.95,0.00)`,                           // R
      `=IF(L${r}>=S7,1.00,0.00)`,                          // S
      `=IF(L${r}>T7,1.05,0.00)`,                           // T
      `=IF(L${r}>U7,1.10,0.00)`,                           // U
      `=MAX(Q${r}:U${r})`,                                  // V
    ];
  });
}

/**
 * Write stat formulas (S5, U5) and side panel values.
 * Staff and intern have different side panel structure — mirrors both GAS scripts.
 */
async function writeOverallMeta(sheets, ssId, sheetName, lastRow, isIntern, internSheetName) {
  const S = 8;

  const panelData = isIntern
    ? [  // ── Intern side panel (rows 8, 11, 15) ────────────────
        { range: `'${sheetName}'!X8:Y8`,   values: [['งบจัดสรร', '']] },
        { range: `'${sheetName}'!X9:Y9`,   values: [['เฉพาะ intern', 0]] },
        { range: `'${sheetName}'!X11:Y11`, values: [['ค่าที่คำนวณได้', '']] },
        { range: `'${sheetName}'!X12:Y13`, values: [
          ['sum point', `=SUM(V${S}:V${lastRow})`],
          ['per P4P',   '=Y9/Y12'],
        ]},
        { range: `'${sheetName}'!X15:Y15`, values: [['ผลรวม', '']] },
        { range: `'${sheetName}'!X16:Y16`, values: [['P4P intern', `=SUM(O${S}:O${lastRow})`]] },
      ]
    : [  // ── Staff side panel (rows 8, 12, 16, 22) ─────────────
        { range: `'${sheetName}'!X8:Y8`,   values: [['งบจัดสรร', '']] },
        { range: `'${sheetName}'!X9:Y10`,  values: [
          ['งบทั้งหมด',   0],
          ['เฉพาะ staff', `=Y9-'${internSheetName}'!Y16`],
        ]},
        { range: `'${sheetName}'!X12:Y12`, values: [['ค่าที่คำนวณได้', '']] },
        { range: `'${sheetName}'!X13:Y14`, values: [
          ['sum point', `=SUM(V${S}:V${lastRow})`],
          ['per P4P',   '=Y10/Y13'],
        ]},
        { range: `'${sheetName}'!X16:Y16`, values: [['ผลรวม', '']] },
        { range: `'${sheetName}'!X17:Y20`, values: [
          ['ค่าตอบแทนตามสัดส่วนวิชาชีพ',     `=SUM(N${S}:N${lastRow})`],
          ['ค่าตอบแทนบริหาร',                 `=SUM(M${S}:M${lastRow})`],
          ['P4P staff & intern (ไม่รวมค่าบริหาร)', `=Y17+'${internSheetName}'!Y16`],
          ['P4P staff & intern (รวมค่าบริหาร)',     '=Y19+Y18'],
        ]},
        { range: `'${sheetName}'!X22:Y22`, values: [['คงคืน', '']] },
        { range: `'${sheetName}'!Y23`,     values: [['=Y9-Y19']] },
      ];

  await withRetry(() => sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: ssId,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [
        { range: `'${sheetName}'!S5`, values: [[`=AVERAGE(L${S}:L${lastRow})`]] },
        { range: `'${sheetName}'!U5`, values: [[`=STDEV.P(L${S}:L${lastRow})`]] },
        ...panelData,
      ],
    },
  }));
}

/**
 * Apply all formatting for an overall sheet:
 * - J column dept background per row
 * - Number formats (H:O = ###,##0.00 | Q:V = 0.00 | Y = ###,##0.00)
 * - Data area borders
 * - Side panel grey headers + borders + merges
 */
async function formatOverallSheet(sheets, ssId, sheetId, persons, lastRow, isIntern) {
  const S     = 8;
  const R8    = S - 1;          // 0-based row 8
  const RLast = lastRow - 1;    // 0-based last data row
  const GRAY  = { red: 0.831, green: 0.831, blue: 0.831 };
  const GRAY_CELL   = { userEnteredFormat: { backgroundColor: GRAY, textFormat: { bold: true }, horizontalAlignment: 'CENTER' } };
  const GRAY_FIELDS = 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)';

  const fmtReqs = [];

  const WHITE = { red: 1, green: 1, blue: 1 };

  // J column (idx 9): apply dept colour per data row only
  persons.forEach((p, idx) => {
    const hex = DEPT_COLORS[p.department];
    if (!hex) return;
    const ri = S - 1 + idx; // 0-based row
    fmtReqs.push({
      repeatCell: {
        range: gridRange(sheetId, ri, 9, ri, 9),
        cell: { userEnteredFormat: { backgroundColor: hexToRgb(hex) } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    });
  });

  // Extend the grid so the clear range below never exceeds the sheet boundary.
  // This is needed when persons fill the sheet to the template's row limit.
  fmtReqs.push({ appendDimension: { sheetId, dimension: 'ROWS', length: 20 } });

  // Clear J background for rows after last person (template may have residual colours)
  fmtReqs.push({
    repeatCell: {
      range: gridRange(sheetId, RLast + 1, 9, RLast + 20, 9),
      cell: { userEnteredFormat: { backgroundColor: WHITE } },
      fields: 'userEnteredFormat.backgroundColor',
    },
  });

  // Number formats
  fmtReqs.push({ repeatCell: { range: gridRange(sheetId, R8, 7,  RLast, 14), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '###,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });
  fmtReqs.push({ repeatCell: { range: gridRange(sheetId, R8, 16, RLast, 21), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '0.00' } } },        fields: 'userEnteredFormat.numberFormat' } });
  fmtReqs.push({ repeatCell: { range: gridRange(sheetId, R8, 24, RLast, 24), cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '###,##0.00' } } }, fields: 'userEnteredFormat.numberFormat' } });

  // Data area borders (mirrors GAS: A8:A, B8:D, E8:P with inner)
  fmtReqs.push(bordersReq(sheetId, R8, 0,  RLast, 0));
  fmtReqs.push(bordersReq(sheetId, R8, 1,  RLast, 3));
  fmtReqs.push(bordersReq(sheetId, R8, 4,  RLast, 15, true));

  // Side panel grey header rows (0-based)
  // Staff: rows 8,12,16,22 → 0-based 7,11,15,21
  // Intern: rows 8,11,15   → 0-based 7,10,14
  const hdrRows = isIntern ? [7, 10, 14] : [7, 11, 15, 21];
  hdrRows.forEach(r => {
    fmtReqs.push({ repeatCell: { range: gridRange(sheetId, r, 23, r, 24), cell: GRAY_CELL, fields: GRAY_FIELDS } });
    fmtReqs.push(bordersReq(sheetId, r, 23, r, 24));
    fmtReqs.push({ mergeCells: { range: gridRange(sheetId, r, 23, r, 24), mergeType: 'MERGE_ALL' } });
  });

  // Side panel data borders
  if (isIntern) {
    fmtReqs.push(bordersReq(sheetId, 8,  23, 8,  24));  // X9:Y9
    fmtReqs.push(bordersReq(sheetId, 11, 23, 12, 24));  // X12:Y13
    fmtReqs.push(bordersReq(sheetId, 14, 23, 15, 24));  // X15:Y16
  } else {
    fmtReqs.push(bordersReq(sheetId, 8,  23, 9,  24));  // X9:Y10
    fmtReqs.push(bordersReq(sheetId, 12, 23, 13, 24));  // X13:Y14
    fmtReqs.push(bordersReq(sheetId, 16, 23, 19, 24));  // X17:Y20
    fmtReqs.push(bordersReq(sheetId, 22, 23, 22, 24));  // X23:Y23
    fmtReqs.push({ mergeCells: { range: gridRange(sheetId, 22, 23, 22, 24), mergeType: 'MERGE_ALL' } });
  }

  await withRetry(() => sheets.spreadsheets.batchUpdate({ spreadsheetId: ssId, requestBody: { requests: fmtReqs } }));
}

/**
 * Append 6 blank rows then signature rows at lastRow+4 and lastRow+5.
 * Mirrors GAS: both staff and intern scripts.
 */
async function writeOverallSignature(sheets, ssId, sheetName, sheetId, lastRow) {
  const blankRows = Array.from({ length: 6 }, () => Array.from({ length: 26 }, () => ' '));
  await withRetry(() => sheets.spreadsheets.values.append({
    spreadsheetId: ssId,
    range: `'${sheetName}'!A${lastRow + 1}`,
    valueInputOption: 'RAW',
    requestBody: { values: blankRows },
  }));

  const sig1 = lastRow + 4;
  const sig2 = lastRow + 5;

  await withRetry(() => sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: ssId,
    requestBody: {
      valueInputOption: 'RAW',
      data: [
        { range: `'${sheetName}'!A${sig1}`, values: [['ลงชื่อ................................................................................................................หัวหน้างาน/กลุ่มงาน/งาน']] },
        { range: `'${sheetName}'!A${sig2}`, values: [['(.....................................................................................................................................)']] },
      ],
    },
  }));

  await withRetry(() => sheets.spreadsheets.batchUpdate({
    spreadsheetId: ssId,
    requestBody: {
      requests: [
        { mergeCells: { range: gridRange(sheetId, sig1 - 1, 0, sig1 - 1, 15), mergeType: 'MERGE_ALL' } },
        { mergeCells: { range: gridRange(sheetId, sig2 - 1, 0, sig2 - 1, 15), mergeType: 'MERGE_ALL' } },
        { repeatCell: {
            range: gridRange(sheetId, sig1 - 1, 0, sig2 - 1, 15),
            cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } },
            fields: 'userEnteredFormat(horizontalAlignment,textFormat)',
        }},
        { repeatCell: {
            range: gridRange(sheetId, sig1 - 2, 0, sig2 - 1, 25),
            cell: { userEnteredFormat: { textFormat: { fontFamily: 'TH SarabunPSK', fontSize: 16 } } },
            fields: 'userEnteredFormat.textFormat.fontFamily,userEnteredFormat.textFormat.fontSize',
        }},
      ],
    },
  }));
}

// ─── Department sheets ─────────────────────────────────────────────

/**
 * Build all department sheets from dep_template.
 * Mirrors GAS dep doPost: one sheet per dept in DEPT_COLORS order.
 * Uses rowNumMap to link back to staff/intern sheet row numbers.
 *
 * rowNumMap key: "prefix firstname  lastname" (double space before lastname)
 * rowNumMap value: 1-based row number in staff or intern sheet
 */
async function buildDeptSheets(sheets, ssId, depTmplSheetId, allPersons, beYear, month, staffSheetName, internSheetName, rowNumMap) {
  const D = 6; // first data row in dep_template (rows 1-5 are header)

  for (const dept of Object.keys(DEPT_COLORS)) {
    const persons = allPersons
      .filter(p => p.department === dept)
      .sort((a, b) => normaliseName(a.fullname).localeCompare(normaliseName(b.fullname), 'th'));

    if (persons.length === 0) continue;
    // Brief pause between dept sheets to avoid write-quota exhaustion
    await sleep(1000);
    log(`  [SK03] Dept sheet: "${dept}" (${persons.length})`);

    // Copy dep_template
    const copyRes = await withRetry(() => sheets.spreadsheets.sheets.copyTo({
      spreadsheetId: CONFIG.sk03TemplateId, sheetId: depTmplSheetId,
      requestBody: { destinationSpreadsheetId: ssId },
    }));
    const sheetId = copyRes.data.sheetId;

    // Rename + tab colour
    const rgb = DEPT_COLORS[dept] ? hexToRgb(DEPT_COLORS[dept]) : null;
    await withRetry(() => sheets.spreadsheets.batchUpdate({
      spreadsheetId: ssId,
      requestBody: { requests: [{
        updateSheetProperties: {
          properties: { sheetId, title: dept, ...(rgb ? { tabColorStyle: { rgbColor: rgb } } : {}) },
          fields: rgb ? 'title,tabColorStyle' : 'title',
        },
      }]},
    }));

    // Header: A2 = month/year, A3 = dept name (mirrors GAS)
    await withRetry(() => sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: ssId,
      requestBody: {
        valueInputOption: 'RAW',
        data: [
          { range: `'${dept}'!A2`, values: [[`ประจำเดือน  ${THAI_MONTHS[month]}  พ.ศ.  ${beYear}`]] },
          { range: `'${dept}'!A3`, values: [[`หน่วยงาน  ${dept} กลุ่มภารกิจ ด้าน${dept === 'เวชกรรมสังคม' ? 'ปฐมภูมิ' : 'ตติยภูมิ'}`]] },
        ],
      },
    }));

    // Data rows: [count, "prefix firstname  lastname", "position+rank", type, mgmt_amount, 0, 0, " ", remark]
    // Note: double space before lastname — this is the key for rowNumMap lookup
    // F and G are later overwritten by formulas; E stays as the direct MANAGEMENT_DATA amount.
    const rows = persons.map((p, idx) => {
      const mgmt = getMgmt(p);
      return [
        idx + 1,
        `${p.prefix} ${p.firstname}  ${p.lastname}`,  // double space before lastname
        `${p.position}${p.rank ? ' ' + p.rank : ''}`,    // position + rank (supabase 'rank')
        p.type || '',                                   // D: type (supabase 'type')
        mgmt?.amount ?? 0,   // E: management amount (or 0)
        0, 0, ' ',
        mgmt?.remark ?? ' ', // I: remark
      ];
    });
    await withRetry(() => sheets.spreadsheets.values.append({
      spreadsheetId: ssId,
      range: `'${dept}'!A${D}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: rows },
    }));

    const lastRow      = D + persons.length - 1;
    const isInternDept = dept === 'INTERN';
    const formulaData  = [];

    // Set E, F, G formulas referencing staff/intern sheet via row_num
    // Non-INTERN: E = staff!M, F = staff!N, G = E+F
    // INTERN:     F = intern!N, G = E+F (E left as 0)
    for (let i = 0; i < persons.length; i++) {
      const p       = persons[i];
      const nameKey = `${p.prefix} ${p.firstname}  ${p.lastname}`;
      const rowNum  = rowNumMap[nameKey];
      const r       = D + i;

      if (rowNum == null) { log(`  [SK03] ⚠ row_num missing for "${nameKey}"`, 'warn'); continue; }

      if (isInternDept) {
        formulaData.push({ range: `'${dept}'!F${r}`, values: [[`='${internSheetName}'!N${rowNum}`]] });
        formulaData.push({ range: `'${dept}'!G${r}`, values: [[`=E${r}+F${r}`]] });
      } else {
        // E: direct amount from MANAGEMENT_DATA (written in rows above); no formula override
        formulaData.push({ range: `'${dept}'!F${r}`, values: [[`='${staffSheetName}'!N${rowNum}`]] });
        formulaData.push({ range: `'${dept}'!G${r}`, values: [[`=E${r}+F${r}`]] });
      }
    }
    if (formulaData.length > 0) {
      await withRetry(() => sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: ssId,
        requestBody: { valueInputOption: 'USER_ENTERED', data: formulaData },
      }));
    }

    // Borders + number format (mirrors GAS dep: A6:I with inner, E6:G number format)
    const R_S = D - 1; // 0-based data start
    const R_E = lastRow - 1;
    await withRetry(() => sheets.spreadsheets.batchUpdate({
      spreadsheetId: ssId,
      requestBody: { requests: [
        bordersReq(sheetId, R_S, 0, R_E, 8, true),
        { repeatCell: {
            range: gridRange(sheetId, R_S, 4, R_E, 6),
            cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '###,##0.00' } } },
            fields: 'userEnteredFormat.numberFormat',
        }},
      ]},
    }));

    // 5 blank rows then signatures at lastRow+3 (two sections) and lastRow+5 (one)
    const blankRows = Array.from({ length: 5 }, () => Array.from({ length: 10 }, () => ' '));
    await withRetry(() => sheets.spreadsheets.values.append({
      spreadsheetId: ssId,
      range: `'${dept}'!A${lastRow + 1}`,
      valueInputOption: 'RAW',
      requestBody: { values: blankRows },
    }));

    const sig1 = lastRow + 3;
    const sig2 = lastRow + 5;
    await withRetry(() => sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: ssId,
      requestBody: {
        valueInputOption: 'RAW',
        data: [
          { range: `'${dept}'!A${sig1}`, values: [['................................................................................................................หัวหน้างาน/กลุ่มงาน/ฝ่าย']] },
          { range: `'${dept}'!E${sig1}`, values: [['................................................................................................................หัวหน้ากลุ่มภารกิจ']] },
          { range: `'${dept}'!A${sig2}`, values: [['................................................................................................................ประธานคณะกรรมการตรวจสอบค่าคะแนน']] },
        ],
      },
    }));
    await withRetry(() => sheets.spreadsheets.batchUpdate({
      spreadsheetId: ssId,
      requestBody: { requests: [
        { mergeCells: { range: gridRange(sheetId, sig1-1, 0, sig1-1, 3), mergeType: 'MERGE_ALL' } }, // A:D
        { mergeCells: { range: gridRange(sheetId, sig1-1, 4, sig1-1, 8), mergeType: 'MERGE_ALL' } }, // E:I
        { mergeCells: { range: gridRange(sheetId, sig2-1, 0, sig2-1, 8), mergeType: 'MERGE_ALL' } }, // A:I
        { repeatCell: { range: gridRange(sheetId, sig1-1, 0, sig1-1, 8), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } },
        { repeatCell: { range: gridRange(sheetId, sig2-1, 0, sig2-1, 8), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } },
        { repeatCell: { range: gridRange(sheetId, sig1-1, 0, sig2-1, 8), cell: { userEnteredFormat: { textFormat: { fontFamily: 'TH SarabunPSK', fontSize: 16 } } }, fields: 'userEnteredFormat.textFormat.fontFamily,userEnteredFormat.textFormat.fontSize' } },
      ]},
    }));
  }

  // Delete "ชีต1" if still present (created with the spreadsheet, normally deleted in staff step)
  const meta   = await withRetry(() => sheets.spreadsheets.get({ spreadsheetId: ssId, fields: 'sheets.properties' }));
  const sheet1 = meta.data.sheets.find(s => s.properties.title === 'ชีต1');
  if (sheet1) {
    await withRetry(() => sheets.spreadsheets.batchUpdate({ spreadsheetId: ssId, requestBody: { requests: [{ deleteSheet: { sheetId: sheet1.properties.sheetId } }] } }));
    log('  [SK03] Deleted default "ชีต1" sheet');
  }
}

// ─── Master orchestrator ──────────────────────────────────────────

/**
 * Create the full SK03 spreadsheet for a given month.
 * Sequence mirrors the original three GAS scripts:
 *   1. staff GAS  → creates spreadsheet + staff sheet + patches Supabase row_num
 *   2. intern GAS → adds intern sheet + patches Supabase row_num
 *   3. dep GAS    → adds one dept sheet per department
 */
async function createSK03(drive, sheets, supabasePersons, monthKey, supabase) {
  const [beYearStr, monthStr] = monthKey.split('_');
  const beYear = parseInt(beYearStr);
  const month  = parseInt(monthStr);
  // Sheet names derived directly from monthKey — same source as the merge step
  const sheetBase       = monthKeyToSheetBase(monthKey);  // e.g. "ธ.ค. 68"
  const staffSheetName  = `${sheetBase} - staff`;
  const internSheetName = `${sheetBase} - intern`;
  const S = 8; // first data row in overall sheets

  // ── Sort people (DEPT_COLORS order → name within dept) ───────────
  const staffArray  = [];
  const internArray = [];
  for (const dept of Object.keys(DEPT_COLORS)) {
    const sorted = supabasePersons
      .filter(p => p.department === dept)
      .sort((a, b) => normaliseName(a.fullname).localeCompare(normaliseName(b.fullname), 'th'));
    if (dept === 'INTERN') internArray.push(...sorted);
    else staffArray.push(...sorted);
  }

  // ── Find old copies, create new spreadsheet, then delete old ones ──
  const existing = await driveListAll(drive, {
    q: `'${CONFIG.sk03FolderId}' in parents and name='sk03 - ${monthKey}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`,
    fields: 'nextPageToken, files(id)', pageSize: 10,
  });

  log(`  [SK03] Creating "sk03 - ${monthKey}"…`);
  const ssId = (await withRetry(() => drive.files.create({
    requestBody: { name: `sk03 - ${monthKey}`, mimeType: 'application/vnd.google-apps.spreadsheet', parents: [CONFIG.sk03FolderId] },
    fields: 'id',
  }))).data.id;
  log(`  [SK03] id: ${ssId}`);

  for (const old of existing) {
    await withRetry(() => drive.files.delete({ fileId: old.id }));
    log(`  [SK03] Deleted old file: ${old.id}`);
  }

  // ── Fetch template sheet IDs ──────────────────────────────────────
  const tmplSheets = (await withRetry(() => sheets.spreadsheets.get({ spreadsheetId: CONFIG.sk03TemplateId, fields: 'sheets.properties' }))).data.sheets;
  const findTmpl   = name => {
    const s = tmplSheets.find(sh => sh.properties.title === name);
    if (!s) throw new Error(`Template sheet "${name}" not found`);
    return s.properties.sheetId;
  };
  const overallTmplId = findTmpl('overall_template');
  const depTmplId     = findTmpl('dep_template');

  // Get all existing sheets (to delete after copying the first new sheet)
  const ssSheets   = (await withRetry(() => sheets.spreadsheets.get({ spreadsheetId: ssId, fields: 'sheets.properties' }))).data.sheets;
  const oldSheetIds = ssSheets.map(sh => sh.properties.sheetId);

  // ══════════════════════════════════════════════════════════════════
  //  1. Staff sheet
  // ══════════════════════════════════════════════════════════════════
  log(`  [SK03] Staff sheet (${staffArray.length} persons)…`);
  const staffSheetId = (await withRetry(() => sheets.spreadsheets.sheets.copyTo({
    spreadsheetId: CONFIG.sk03TemplateId, sheetId: overallTmplId,
    requestBody: { destinationSpreadsheetId: ssId },
  }))).data.sheetId;

  // Delete all pre-existing sheets first (avoids name conflict on rename),
  // then rename the newly copied sheet to the target name.
  if (oldSheetIds.length > 0) {
    await withRetry(() => sheets.spreadsheets.batchUpdate({ spreadsheetId: ssId, requestBody: { requests:
      oldSheetIds.map(id => ({ deleteSheet: { sheetId: id } })),
    }}));
  }
  await withRetry(() => sheets.spreadsheets.batchUpdate({ spreadsheetId: ssId, requestBody: { requests: [
    { updateSheetProperties: { properties: { sheetId: staffSheetId, title: staffSheetName }, fields: 'title' } },
  ]}}));

  // A3 header
  await withRetry(() => sheets.spreadsheets.values.update({ spreadsheetId: ssId, range: `'${staffSheetName}'!A3`, valueInputOption: 'RAW', requestBody: { values: [[`ประจำเดือน  ${THAI_MONTHS[month]}  พ.ศ.  ${beYear}`]] } }));

  // Data rows
  const staffRows = buildOverallRows(staffArray, S, false);
  if (staffRows.length > 0) {
    await withRetry(() => sheets.spreadsheets.values.update({ spreadsheetId: ssId, range: `'${staffSheetName}'!A${S}:V${S + staffRows.length - 1}`, valueInputOption: 'USER_ENTERED', requestBody: { values: staffRows } }));
  }
  const staffLastRow = S + staffArray.length - 1;

  await writeOverallMeta(sheets, ssId, staffSheetName, staffLastRow, false, internSheetName);
  await formatOverallSheet(sheets, ssId, staffSheetId, staffArray, staffLastRow, false);

  // Totals row: M, N, O = SUM immediately below last person, above signature.
  // Use values.append so it auto-extends the grid — no grid limit errors.
  const RED = { red: 0.933, green: 0.294, blue: 0.169 }; // #EE4B2B
  const totalsRow = Array.from({ length: 26 }, () => '');  // 26 blank cells (cols A-Z)
  totalsRow[12] = `=SUM(M${S}:M${staffLastRow})`;  // col M (idx 12)
  totalsRow[13] = `=SUM(N${S}:N${staffLastRow})`;  // col N (idx 13)
  totalsRow[14] = `=SUM(O${S}:O${staffLastRow})`;  // col O (idx 14)

  const appendRes = await withRetry(() => sheets.spreadsheets.values.append({
    spreadsheetId: ssId,
    range: `'${staffSheetName}'!A${staffLastRow + 1}`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [totalsRow] },
  }));
  // Find the actual row number appended (1-based)
  const staffTotalRowRange = appendRes.data.updates?.updatedRange ?? '';
  const staffTotalRowMatch = staffTotalRowRange.match(/:?[A-Z](\d+)$/);
  const staffTotalRow = staffTotalRowMatch ? parseInt(staffTotalRowMatch[1]) : staffLastRow + 1;

  // Paint M:O of the totals row red; clear M:O below it (template residual colours)
  await withRetry(() => sheets.spreadsheets.batchUpdate({
    spreadsheetId: ssId,
    requestBody: { requests: [
      {
        repeatCell: {
          range: gridRange(staffSheetId, staffTotalRow - 1, 12, staffTotalRow - 1, 14),
          cell: { userEnteredFormat: { backgroundColor: RED } },
          fields: 'userEnteredFormat.backgroundColor',
        },
      },
      {
        repeatCell: {
          range: gridRange(staffSheetId, staffTotalRow, 12, staffTotalRow + 19, 14),
          cell: { userEnteredFormat: { backgroundColor: { red: 1, green: 1, blue: 1 } } },
          fields: 'userEnteredFormat.backgroundColor',
        },
      },
    ]},
  }));

  await writeOverallSignature(sheets, ssId, staffSheetName, staffSheetId, staffTotalRow);

  // Patch Supabase row_num for staff
  log(`  [SK03] Patching row_num for ${staffArray.length} staff…`);
  for (let i = 0; i < staffArray.length; i++) {
    if (staffArray[i].index == null) continue;
    const { error } = await supabase.from(monthKey).update({ row_num: S + i }).eq(SB.index, staffArray[i].index);
    if (error) log(`  [SK03] ⚠ row_num patch failed for index ${staffArray[i].index}: ${error.message}`, 'warn');
  }

  // ══════════════════════════════════════════════════════════════════
  //  2. Intern sheet
  // ══════════════════════════════════════════════════════════════════
  log(`  [SK03] Intern sheet (${internArray.length} persons)…`);
  const internSheetId = (await withRetry(() => sheets.spreadsheets.sheets.copyTo({
    spreadsheetId: CONFIG.sk03TemplateId, sheetId: overallTmplId,
    requestBody: { destinationSpreadsheetId: ssId },
  }))).data.sheetId;

  await withRetry(() => sheets.spreadsheets.batchUpdate({ spreadsheetId: ssId, requestBody: { requests: [
    { updateSheetProperties: { properties: { sheetId: internSheetId, title: internSheetName }, fields: 'title' } },
  ]}}));

  await withRetry(() => sheets.spreadsheets.values.update({ spreadsheetId: ssId, range: `'${internSheetName}'!A3`, valueInputOption: 'RAW', requestBody: { values: [[`ประจำเดือน  ${THAI_MONTHS[month]}  พ.ศ.  ${beYear}`]] } }));

  const internRows = buildOverallRows(internArray, S, true);
  if (internRows.length > 0) {
    await withRetry(() => sheets.spreadsheets.values.update({ spreadsheetId: ssId, range: `'${internSheetName}'!A${S}:V${S + internRows.length - 1}`, valueInputOption: 'USER_ENTERED', requestBody: { values: internRows } }));
  }
  const internLastRow = S + internArray.length - 1;

  await writeOverallMeta(sheets, ssId, internSheetName, internLastRow, true, internSheetName);
  await formatOverallSheet(sheets, ssId, internSheetId, internArray, internLastRow, true);

  // Intern totals row: N and O only (no M — management fee not applicable for interns)
  const internTotalsRow = Array.from({ length: 26 }, () => '');
  internTotalsRow[13] = `=SUM(N${S}:N${internLastRow})`;  // col N (idx 13)
  internTotalsRow[14] = `=SUM(O${S}:O${internLastRow})`;  // col O (idx 14)

  const internAppendRes = await withRetry(() => sheets.spreadsheets.values.append({
    spreadsheetId: ssId,
    range: `'${internSheetName}'!A${internLastRow + 1}`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [internTotalsRow] },
  }));
  const internTotalRowRange = internAppendRes.data.updates?.updatedRange ?? '';
  const internTotalRowMatch = internTotalRowRange.match(/:?[A-Z](\d+)$/);
  const internTotalRow = internTotalRowMatch ? parseInt(internTotalRowMatch[1]) : internLastRow + 1;

  await withRetry(() => sheets.spreadsheets.batchUpdate({
    spreadsheetId: ssId,
    requestBody: { requests: [
      {
        repeatCell: {
          range: gridRange(internSheetId, internTotalRow - 1, 13, internTotalRow - 1, 14),
          cell: { userEnteredFormat: { backgroundColor: RED } },
          fields: 'userEnteredFormat.backgroundColor',
        },
      },
      {
        repeatCell: {
          range: gridRange(internSheetId, internTotalRow, 13, internTotalRow + 19, 14),
          cell: { userEnteredFormat: { backgroundColor: { red: 1, green: 1, blue: 1 } } },
          fields: 'userEnteredFormat.backgroundColor',
        },
      },
    ]},
  }));

  await writeOverallSignature(sheets, ssId, internSheetName, internSheetId, internTotalRow);

  // Patch Supabase row_num for intern
  log(`  [SK03] Patching row_num for ${internArray.length} interns…`);
  for (let i = 0; i < internArray.length; i++) {
    if (internArray[i].index == null) continue;
    const { error } = await supabase.from(monthKey).update({ row_num: S + i }).eq(SB.index, internArray[i].index);
    if (error) log(`  [SK03] ⚠ row_num patch failed for index ${internArray[i].index}: ${error.message}`, 'warn');
  }

  // ══════════════════════════════════════════════════════════════════
  //  3. Department sheets
  // ══════════════════════════════════════════════════════════════════
  // Build row_num lookup: "prefix firstname  lastname" → row in staff/intern sheet
  const rowNumMap = {};
  const addToRowNumMap = (p, row) => {
    const key = `${p.prefix} ${p.firstname}  ${p.lastname}`;
    if (key in rowNumMap) log(`  [SK03] ⚠ rowNumMap duplicate key: "${key}" — overwriting row ${rowNumMap[key]} with ${row}`, 'warn');
    rowNumMap[key] = row;
  };
  staffArray.forEach( (p, i) => addToRowNumMap(p, S + i));
  internArray.forEach((p, i) => addToRowNumMap(p, S + i));

  await buildDeptSheets(sheets, ssId, depTmplId, supabasePersons, beYear, month, staffSheetName, internSheetName, rowNumMap);

  log(`  [SK03] ✓ Done → https://docs.google.com/spreadsheets/d/${ssId}`);
  return ssId;
}

// ═══════════════════════════════════════════════════════════════════
//  DATAFLOW DIAGRAM
// ═══════════════════════════════════════════════════════════════════
function showDataflow() {
  console.log('');
  console.log('  ┌──────────────────────┐     ┌──────────────────────┐');
  console.log('  │     Google Drive      │     │       Supabase        │');
  console.log('  │  (monthly .xlsx files)│     │  (persons + scores)   │');
  console.log('  └──────────┬───────────┘     └───────────┬───────────┘');
  console.log('             │ Step 2: list files           │ Step 3: fetch persons');
  console.log('             │                             │');
  console.log('             └──────────────┬──────────────┘');
  console.log('                            │');
  console.log('               Step 4: validate lists match');
  console.log('                            │');
  console.log('                            ▼');
  console.log('          Step 4.5: fill missing scores          ←── Google Drive (Excel)');
  console.log('               (score=null → extract from        ──► Supabase (score update)');
  console.log('                physician Excel, save to SB)');
  console.log('                            │');
  console.log('                            ▼');
  console.log('               Step 5: merge → one workbook');
  console.log('                    (1 sheet per person,');
  console.log('                     sorted by department)');
  console.log('                            │');
  console.log('                            ▼');
  console.log('               Step 6: create SK03 spreadsheet');
  console.log('                    (staff / intern / dept sheets)');
  console.log('                            │');
  console.log('             ┌──────────────┴──────────────┐');
  console.log('             ▼                             ▼');
  console.log('  ┌──────────────────────┐     ┌──────────────────────┐');
  console.log('  │     Google Drive      │     │       Supabase        │');
  console.log('  │  (merged workbook)    │     │  (row_num updated)    │');
  console.log('  └──────────────────────┘     └──────────────────────┘');
  console.log('');
}

// ═══════════════════════════════════════════════════════════════════
//  MAIN
// ═══════════════════════════════════════════════════════════════════
async function main() {
  const missing = [
    ['GOOGLE_CLIENT_ID',         CONFIG.google.clientId],
    ['GOOGLE_CLIENT_SECRET',     CONFIG.google.clientSecret],
    ['GOOGLE_REFRESH_TOKEN',     CONFIG.google.refreshToken],
    ['SUPABASE_URL',             CONFIG.supabase.url],
    ['SUPABASE_KEY',             CONFIG.supabase.key],
    ['GOOGLE_ROOT_FOLDER_ID',    CONFIG.rootFolderId],
    ['GOOGLE_MERGE_FOLDER_ID',   CONFIG.outputFolderId],
    ['GOOGLE_SK03_FOLDER_ID',    CONFIG.sk03FolderId],
    ['GOOGLE_SK03_TEMPLATE_ID',  CONFIG.sk03TemplateId],
  ].filter(([, v]) => !v).map(([k]) => k);

  if (missing.length > 0) {
    console.error('\n❌  Missing variables in .env:');
    missing.forEach(k => console.error(`     • ${k}`));
    process.exit(1);
  }

  const drive    = createDriveClient();
  const sheets   = createSheetsClient();
  const supabase = createClient(CONFIG.supabase.url, CONFIG.supabase.key);
  const months   = getTargetMonths();

  // Pre-flight: verify SK03 template has required sheets before processing any month
  log('[Pre-flight] Verifying SK03 template sheets…');
  const tmplCheck = await withRetry(() => sheets.spreadsheets.get({ spreadsheetId: CONFIG.sk03TemplateId, fields: 'sheets.properties' }));
  const tmplTitles = tmplCheck.data.sheets.map(s => s.properties.title);
  for (const required of ['overall_template', 'dep_template']) {
    if (!tmplTitles.includes(required)) {
      console.error(`\n❌  SK03 template is missing sheet: "${required}"`);
      process.exit(1);
    }
  }
  log('[Pre-flight] ✓ Template OK');

  console.log('══════════════════════════════════════════════════════');
  console.log(' P4P Excel Merge + SK03 Pipeline');
  console.log('══════════════════════════════════════════════════════');
  showDataflow();
  console.log('Month(s):', months.map(m => m.key).join(', '));
  console.log('');

  const results = { processed: 0, skipped: 0, failed: [] };

  // Cap full-pipeline runs to the latest N months that pass Step 4 validation
  // (i.e. months whose Drive list matches Supabase — all physicians complete).
  // Months iterate newest-first, so this keeps the most recent complete months
  // and skips older ones to stay within the GitHub Actions 30-minute budget.
  const LATEST_COMPLETE_MONTHS = 2;
  let completedCount = 0;

  for (const monthInfo of months) {
    const { key } = monthInfo;
    console.log(`\n┌─ ${key} ${'─'.repeat(46 - key.length)}`);
    try {
      // Step 2
      log('  [Step 2] Google Drive…');
      const driveData = await getDriveMonthData(drive, monthInfo);
      if (!driveData) { results.skipped++; console.log('└─ skipped\n'); continue; }
      log(`  → ${driveData.names.length} files`);

      // Step 3
      log('  [Step 3] Supabase…');
      const sbData = await getSupabaseMonthData(supabase, key);
      if (!sbData) { results.skipped++; console.log('└─ skipped\n'); continue; }
      log(`  → ${sbData.length} persons`);

      // Step 4
      log('  [Step 4] Comparing lists…');
      if (!listsMatch(driveData.names, sbData)) {
        log('  → ✗ Lists do not match — skipping', 'warn');
        results.skipped++; console.log('└─ skipped\n'); continue;
      }
      log('  → ✓ Match');

      // Stop once we've processed the latest N complete months.
      // Older complete months are skipped — they've typically been processed
      // by a prior run and re-processing them would blow the timeout.
      if (completedCount >= LATEST_COMPLETE_MONTHS) {
        log(`  → Already processed latest ${LATEST_COMPLETE_MONTHS} complete months — stopping`);
        results.skipped++;
        console.log('└─ skipped (limit reached)\n');
        break;
      }

      // Step 4.5
      log('  [Step 4.5] Filling missing scores…');
      await fillMissingScores(drive, driveData.files, sbData, key, supabase);

      // Guard: all persons must have a non-null score before merge / SK03
      const nullAfterFill = sbData.filter(p => p.score === null && p.perf_score === 0);
      if (nullAfterFill.length > 0) {
        log(`  → ✗ ${nullAfterFill.length} person(s) still have no score after Step 4.5 — skipping merge & SK03`, 'warn');
        nullAfterFill.forEach(p => log(`    • ${p.fullname}`, 'warn'));
        results.skipped++; console.log('└─ skipped (null scores)\n'); continue;
      }
      log('  → ✓ All scores present');

      // Step 5
      log('  [Step 5] Merging…');
      const uploaded = await mergeAndUpload(drive, driveData.files, sbData, key);
      if (!uploaded) { results.skipped++; console.log('└─ skipped\n'); continue; }

      // Step 6
      log('  [Step 6] Creating SK03 spreadsheet…');
      await createSK03(drive, sheets, sbData, key, supabase);

      results.processed++;
      completedCount++;
      console.log('└─ ✓ complete\n');
    } catch (err) {
      console.error(`  [ERROR] ${err.stack ?? err.message}`);
      results.failed.push({ key, error: err.message });
      console.log('└─ failed\n');
    }
  }

  console.log('══════════════════════════════════════════════════════');
  console.log(` ✓ Processed: ${results.processed}`);
  console.log(` ⊘ Skipped  : ${results.skipped}`);
  console.log(` ✗ Failed   : ${results.failed.length}`);
  results.failed.forEach(f => console.log(`     ${f.key}: ${f.error}`));
  console.log('');
}

main().catch(err => { console.error('\nFatal:', err.stack ?? err); process.exit(1); });