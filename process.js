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

const { google }       = require('googleapis');
const { createClient } = require('@supabase/supabase-js');
const ExcelJS          = require('exceljs');
const JSZip            = require('jszip');
const { Readable }     = require('stream');

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
// Key: firstname + " " + lastname (no prefix). Provides col I (staff) and col E + I (dept).
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
];

// O(1) lookup: "firstname lastname" → { remark, amount }
const MGMT_LOOKUP = Object.fromEntries(
  MANAGEMENT_DATA.map(d => [normaliseName(d.name), { remark: d.remark, amount: d.amount }])
);

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
  level:     'level',           // rank in GAS
  emp_type:  'employee_type',   // type in GAS
  std:       'standard_score',
  boss:      'boss_score',
  perf:      'performance_score', // fallback col J
  score:     'score',               // col J override — used when not null
  mgmt:      'management_fee',
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
    const res = await drive.files.list({ ...params, pageToken });
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
  const res = await drive.files.get({ fileId, alt: 'media' }, { responseType: 'arraybuffer' });
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
      await drive.files.delete({ fileId: d.id });
      log(`  [Drive] Deleted duplicate id: ${d.id}`, 'warn');
    }
    log(`  [Drive] Overwriting "${fileName}" (id: ${first.id})`);
    const res = await drive.files.update({
      fileId: first.id,
      requestBody: { name: fileName },
      media: { mimeType: mime, body: Readable.from([buffer]) },
      fields: 'id, name, webViewLink',
    });
    return res.data;
  }
  const res = await drive.files.create({
    requestBody: { name: fileName, parents: [folderId], mimeType: mime },
    media: { mimeType: mime, body: Readable.from([buffer]) },
    fields: 'id, name, webViewLink',
  });
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
    emp_type:   sbVal(r, SB.emp_type, 'ข้าราชการ'),
    std_score:  sbVal(r, SB.std,      2200),
    boss_score: sbVal(r, SB.boss,     0),
    perf_score: sbVal(r, SB.perf,     0),
    score:      sbVal(r, SB.score,    null), // overrides perf_score in col J when not null
    mgmt_fee:   sbVal(r, SB.mgmt,     0),
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
    normDrive.forEach(n => { if (!normSB.includes(n))   log(`    ✗ Drive but not SB: "${n}"`, 'warn'); });
    normSB.forEach(n   => { if (!normDrive.includes(n)) log(`    ✗ SB but not Drive: "${n}"`, 'warn'); });
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
//  STEP 5 — Merge Excel files
// ═══════════════════════════════════════════════════════════════════
function isSheetEmpty(ws) {
  let count = 0;
  ws.eachRow(row => { row.eachCell(cell => { if (cell.value !== null && cell.value !== undefined && cell.value !== '') count++; }); });
  return count <= 3;
}

async function stripFormulas(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const paths = Object.keys(zip.files).filter(p => /^xl\/worksheets\/sheet\d+\.xml$/.test(p));
  for (const path of paths) {
    let xml = await zip.files[path].async('string');
    xml = xml.replace(/<f\b[^>]*>[\s\S]*?<\/f>/g, '');
    xml = xml.replace(/<f\b[^/]*\/>/g, '');
    xml = xml.replace(/\s+si="\d+"/g, '');
    xml = xml.replace(/\s+t="shared"/g, '');
    zip.file(path, xml);
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function loadWorkbook(buffer) {
  const cleaned = await stripFormulas(buffer);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(cleaned);
  return wb;
}

function resolveCellValue(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if ('richText' in v) return v.richText.map(r => r.text).join('');
    if ('text'     in v) return v.text;
    if ('result'   in v) return v.result;
    if ('error'    in v) return null;
  }
  return v;
}

function copyWorksheet(srcWS, dstWS) {
  srcWS.columns.forEach((srcCol, idx) => {
    const dstCol = dstWS.getColumn(idx + 1);
    if (srcCol.width)        dstCol.width       = srcCol.width;
    if (srcCol.hidden)       dstCol.hidden       = srcCol.hidden;
    if (srcCol.outlineLevel) dstCol.outlineLevel = srcCol.outlineLevel;
  });
  srcWS.eachRow({ includeEmpty: true }, (srcRow, rowNumber) => {
    const dstRow = dstWS.getRow(rowNumber);
    if (srcRow.height)       dstRow.height      = srcRow.height;
    if (srcRow.hidden)       dstRow.hidden       = srcRow.hidden;
    if (srcRow.outlineLevel) dstRow.outlineLevel = srcRow.outlineLevel;
    srcRow.eachCell({ includeEmpty: true }, (srcCell, colNumber) => {
      const dstCell = dstWS.getCell(rowNumber, colNumber);
      dstCell.value = resolveCellValue(srcCell.value);
      if (srcCell.style && Object.keys(srcCell.style).length > 0)
        dstCell.style = JSON.parse(JSON.stringify(srcCell.style));
    });
    dstRow.commit();
  });
  if (srcWS.model?.merges) {
    for (const m of srcWS.model.merges) { try { dstWS.mergeCells(m); } catch (_) {} }
  }
}

function sortByDeptThenName(a, b, deptMap) {
  const dA   = deptMap[a.normName] ?? '';
  const dB   = deptMap[b.normName] ?? '';
  const idxA = DEPT_ORDER[dA] ?? DEPT_ORDER_MAX;
  const idxB = DEPT_ORDER[dB] ?? DEPT_ORDER_MAX;
  if (idxA !== idxB) return idxA - idxB;
  return a.normName.localeCompare(b.normName, 'th');
}

/** Merge all Excel files → one workbook with one sheet per person */
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

  const mergedWB       = new ExcelJS.Workbook();
  const usedSheetNames = new Set();
  let   totalRows = 0, sheetCount = 0;

  for (const file of filesWithName) {
    log(`\n  [Merge] ↓ ${file.origName}`);
    const buffer = await downloadFile(drive, file.id);
    const wb     = await loadWorkbook(buffer);

    let targetWS = null;
    for (const ws of wb.worksheets) {
      if (!isSheetEmpty(ws)) { targetWS = ws; log(`      → using sheet: "${ws.name}"`); break; }
      log(`      → sheet "${ws.name}" is empty, trying next`);
    }
    if (!targetWS) { log(`  ⚠ All sheets empty — skipped`, 'warn'); continue; }

    let sheetName = toSheetName(file.origName);
    if (usedSheetNames.has(sheetName)) {
      let c = 2;
      while (usedSheetNames.has(`${sheetName}_${c}`)) c++;
      sheetName = `${sheetName}_${c}`.substring(0, 31);
    }
    usedSheetNames.add(sheetName);

    const dept    = deptMap[file.normName] ?? '';
    const tabArgb = hexToArgb(DEPT_COLORS[dept]);
    const outWS   = mergedWB.addWorksheet(sheetName, {
      properties: tabArgb ? { tabColor: { argb: tabArgb } } : {},
    });
    sheetCount++;
    copyWorksheet(targetWS, outWS);
    totalRows += Math.max(0, outWS.rowCount - 1);
    log(`      → ${outWS.rowCount} rows copied to "${sheetName}" [${dept || '?'}]${tabArgb ? ' 🎨' : ''}`);
  }

  if (sheetCount === 0) { log('  [Merge] ⚠ No data — nothing to upload', 'warn'); return null; }

  const buf  = await mergedWB.xlsx.writeBuffer();
  const name = `merged_${monthKey}.xlsx`;
  log(`\n  [Merge] Uploading "${name}" — ${sheetCount} sheets, ${totalRows} data rows…`);
  const uploaded = await uploadFileToDrive(drive, CONFIG.outputFolderId, name, buf);
  log(`  [Merge] ✓ id: ${uploaded.id}\n           🔗 ${uploaded.webViewLink}`);
  return uploaded;
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
      p.level     || '',                                     // F: rank
      p.emp_type  || '',                                     // G: type
      p.std_score  ?? 2200,                                  // H
      (getMgmt(p)?.amount ?? p.boss_score ?? 0),             // I: mgmt amount or boss score
      p.score ?? p.perf_score ?? 0,                          // J: score (SB 'score' overrides 'performance_score')
      `=I${r}+J${r}`,                                       // K
      `=K${r}-H${r}`,                                       // L
      p.mgmt_fee   ?? 0,                                     // M
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
  const S      = 8;
  const intern = `'${internSheetName}'`;

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
          ['เฉพาะ staff', `=Y9-${intern}!Y16`],
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
          ['P4P staff & intern (ไม่รวมค่าบริหาร)', `=Y17+${intern}!Y16`],
          ['P4P staff & intern (รวมค่าบริหาร)',     '=Y19+Y18'],
        ]},
        { range: `'${sheetName}'!X22:Y22`, values: [['คงคืน', '']] },
        { range: `'${sheetName}'!X23:Y23`, values: [['=Y9-Y19', '']] },
      ];

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: ssId,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: [
        { range: `'${sheetName}'!S5`, values: [[`=AVERAGE(L${S}:L${lastRow})`]] },
        { range: `'${sheetName}'!U5`, values: [[`=STDEV.P(L${S}:L${lastRow})`]] },
        ...panelData,
      ],
    },
  });
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

  // J column (idx 9) background per dept per row
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

  await sheets.spreadsheets.batchUpdate({ spreadsheetId: ssId, requestBody: { requests: fmtReqs } });
}

/**
 * Append 6 blank rows then signature rows at lastRow+4 and lastRow+5.
 * Mirrors GAS: both staff and intern scripts.
 */
async function writeOverallSignature(sheets, ssId, sheetName, sheetId, lastRow) {
  const blankRows = Array.from({ length: 6 }, () => Array.from({ length: 26 }, () => ' '));
  await sheets.spreadsheets.values.append({
    spreadsheetId: ssId,
    range: `'${sheetName}'!A${lastRow + 1}`,
    valueInputOption: 'RAW',
    requestBody: { values: blankRows },
  });

  const sig1 = lastRow + 4;
  const sig2 = lastRow + 5;

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: ssId,
    requestBody: {
      valueInputOption: 'RAW',
      data: [
        { range: `'${sheetName}'!A${sig1}`, values: [['ลงชื่อ................................................................................................................หัวหน้างาน/กลุ่มงาน/งาน']] },
        { range: `'${sheetName}'!A${sig2}`, values: [['(.....................................................................................................................................)']] },
      ],
    },
  });

  await sheets.spreadsheets.batchUpdate({
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
      ],
    },
  });
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
    log(`  [SK03] Dept sheet: "${dept}" (${persons.length})`);

    // Copy dep_template
    const copyRes = await sheets.spreadsheets.sheets.copyTo({
      spreadsheetId: CONFIG.sk03TemplateId, sheetId: depTmplSheetId,
      requestBody: { destinationSpreadsheetId: ssId },
    });
    const sheetId = copyRes.data.sheetId;

    // Rename + tab colour
    const rgb = DEPT_COLORS[dept] ? hexToRgb(DEPT_COLORS[dept]) : null;
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: ssId,
      requestBody: { requests: [{
        updateSheetProperties: {
          properties: { sheetId, title: dept, ...(rgb ? { tabColorStyle: { rgbColor: rgb } } : {}) },
          fields: rgb ? 'title,tabColorStyle' : 'title',
        },
      }]},
    });

    // Header: A2 = month/year, A3 = dept name (mirrors GAS)
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: ssId,
      requestBody: {
        valueInputOption: 'RAW',
        data: [
          { range: `'${dept}'!A2`, values: [[`ประจำเดือน  ${THAI_MONTHS[month]}  พ.ศ.  ${beYear}`]] },
          { range: `'${dept}'!A3`, values: [[`หน่วยงาน  ${dept} กลุ่มภารกิจ ด้านตติยภูมิ`]] },
        ],
      },
    });

    // Data rows: [count, "prefix firstname  lastname", "position+rank", type, 0, 0, 0, " ", " "]
    // Note: double space before lastname — this is the key for rowNumMap lookup
    const rows = persons.map((p, idx) => {
      const mgmt = getMgmt(p);
      return [
        idx + 1,
        `${p.prefix} ${p.firstname}  ${p.lastname}`,  // double space before lastname
        `${p.position}${p.level}`,                     // concatenated, no separator
        p.emp_type || '',
        mgmt?.amount ?? 0,   // E: management amount (or 0)
        0, 0, ' ',
        mgmt?.remark ?? ' ', // I: remark
      ];
    });
    await sheets.spreadsheets.values.append({
      spreadsheetId: ssId,
      range: `'${dept}'!A${D}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: rows },
    });

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
        // E: use mgmt amount (already written as value) or fall back to staff!M formula
        const mgmt = getMgmt(p);
        if (!mgmt) {
          formulaData.push({ range: `'${dept}'!E${r}`, values: [[`='${staffSheetName}'!M${rowNum}`]] });
        }
        formulaData.push({ range: `'${dept}'!F${r}`, values: [[`='${staffSheetName}'!N${rowNum}`]] });
        formulaData.push({ range: `'${dept}'!G${r}`, values: [[`=E${r}+F${r}`]] });
      }
    }
    if (formulaData.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: ssId,
        requestBody: { valueInputOption: 'USER_ENTERED', data: formulaData },
      });
    }

    // Borders + number format (mirrors GAS dep: A6:I with inner, E6:G number format)
    const R_S = D - 1; // 0-based data start
    const R_E = lastRow - 1;
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: ssId,
      requestBody: { requests: [
        bordersReq(sheetId, R_S, 0, R_E, 8, true),
        { repeatCell: {
            range: gridRange(sheetId, R_S, 4, R_E, 6),
            cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '###,##0.00' } } },
            fields: 'userEnteredFormat.numberFormat',
        }},
      ]},
    });

    // 5 blank rows then signatures at lastRow+3 (two sections) and lastRow+5 (one)
    const blankRows = Array.from({ length: 5 }, () => Array.from({ length: 10 }, () => ' '));
    await sheets.spreadsheets.values.append({
      spreadsheetId: ssId,
      range: `'${dept}'!A${lastRow + 1}`,
      valueInputOption: 'RAW',
      requestBody: { values: blankRows },
    });

    const sig1 = lastRow + 3;
    const sig2 = lastRow + 5;
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: ssId,
      requestBody: {
        valueInputOption: 'RAW',
        data: [
          { range: `'${dept}'!A${sig1}`, values: [['................................................................................................................หัวหน้างาน/กลุ่มงาน/ฝ่าย']] },
          { range: `'${dept}'!E${sig1}`, values: [['................................................................................................................หัวหน้ากลุ่มภารกิจ']] },
          { range: `'${dept}'!A${sig2}`, values: [['................................................................................................................ประธานคณะกรรมการตรวจสอบค่าคะแนน']] },
        ],
      },
    });
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: ssId,
      requestBody: { requests: [
        { mergeCells: { range: gridRange(sheetId, sig1-1, 0, sig1-1, 3), mergeType: 'MERGE_ALL' } }, // A:D
        { mergeCells: { range: gridRange(sheetId, sig1-1, 4, sig1-1, 8), mergeType: 'MERGE_ALL' } }, // E:I
        { mergeCells: { range: gridRange(sheetId, sig2-1, 0, sig2-1, 8), mergeType: 'MERGE_ALL' } }, // A:I
        { repeatCell: { range: gridRange(sheetId, sig1-1, 0, sig1-1, 8), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } },
        { repeatCell: { range: gridRange(sheetId, sig2-1, 0, sig2-1, 8), cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', textFormat: { bold: true } } }, fields: 'userEnteredFormat(horizontalAlignment,textFormat)' } },
      ]},
    });
  }

  // Delete "ชีต1" if still present (created with the spreadsheet, normally deleted in staff step)
  const meta   = await sheets.spreadsheets.get({ spreadsheetId: ssId, fields: 'sheets.properties' });
  const sheet1 = meta.data.sheets.find(s => s.properties.title === 'ชีต1');
  if (sheet1) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: ssId, requestBody: { requests: [{ deleteSheet: { sheetId: sheet1.properties.sheetId } }] } });
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

  // ── Delete old spreadsheet + create new (mirrors staff GAS) ──────
  const existing = await driveListAll(drive, {
    q: `'${CONFIG.sk03FolderId}' in parents and name='sk03 - ${monthKey}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`,
    fields: 'nextPageToken, files(id)', pageSize: 10,
  });
  for (const f of existing) { await drive.files.delete({ fileId: f.id }); log(`  [SK03] Deleted old: ${f.id}`); }

  log(`  [SK03] Creating "sk03 - ${monthKey}"…`);
  const ssId = (await drive.files.create({
    requestBody: { name: `sk03 - ${monthKey}`, mimeType: 'application/vnd.google-apps.spreadsheet', parents: [CONFIG.sk03FolderId] },
    fields: 'id',
  })).data.id;
  log(`  [SK03] id: ${ssId}`);

  // ── Fetch template sheet IDs ──────────────────────────────────────
  const tmplSheets = (await sheets.spreadsheets.get({ spreadsheetId: CONFIG.sk03TemplateId, fields: 'sheets.properties' })).data.sheets;
  const findTmpl   = name => {
    const s = tmplSheets.find(sh => sh.properties.title === name);
    if (!s) throw new Error(`Template sheet "${name}" not found`);
    return s.properties.sheetId;
  };
  const overallTmplId = findTmpl('overall_template');
  const depTmplId     = findTmpl('dep_template');

  // Get "ชีต1" sheetId of new spreadsheet (to delete after copying staff)
  const ssSheets      = (await sheets.spreadsheets.get({ spreadsheetId: ssId, fields: 'sheets.properties' })).data.sheets;
  const defaultSheetId = ssSheets[0]?.properties.sheetId;

  // ══════════════════════════════════════════════════════════════════
  //  1. Staff sheet
  // ══════════════════════════════════════════════════════════════════
  log(`  [SK03] Staff sheet (${staffArray.length} persons)…`);
  const staffSheetId = (await sheets.spreadsheets.sheets.copyTo({
    spreadsheetId: CONFIG.sk03TemplateId, sheetId: overallTmplId,
    requestBody: { destinationSpreadsheetId: ssId },
  })).data.sheetId;

  // Rename + delete default sheet
  await sheets.spreadsheets.batchUpdate({ spreadsheetId: ssId, requestBody: { requests: [
    { updateSheetProperties: { properties: { sheetId: staffSheetId, title: staffSheetName }, fields: 'title' } },
    ...(defaultSheetId != null ? [{ deleteSheet: { sheetId: defaultSheetId } }] : []),
  ]}});

  // A3 header
  await sheets.spreadsheets.values.update({ spreadsheetId: ssId, range: `'${staffSheetName}'!A3`, valueInputOption: 'RAW', requestBody: { values: [[`ประจำเดือน  ${THAI_MONTHS[month]}  พ.ศ.  ${beYear}`]] } });

  // Data rows
  const staffRows = buildOverallRows(staffArray, S, false);
  if (staffRows.length > 0) {
    await sheets.spreadsheets.values.update({ spreadsheetId: ssId, range: `'${staffSheetName}'!A${S}:V${S + staffRows.length - 1}`, valueInputOption: 'USER_ENTERED', requestBody: { values: staffRows } });
  }
  const staffLastRow = S + staffArray.length - 1;

  await writeOverallMeta(sheets, ssId, staffSheetName, staffLastRow, false, internSheetName);
  await formatOverallSheet(sheets, ssId, staffSheetId, staffArray, staffLastRow, false);

  // Totals row: N and O = SUM of all persons, red background
  const staffTotalRow = staffLastRow + 1;
  const RED = { red: 0.933, green: 0.294, blue: 0.169 }; // #EE4B2B
  await sheets.spreadsheets.values.update({
    spreadsheetId: ssId,
    range: `'${staffSheetName}'!N${staffTotalRow}:O${staffTotalRow}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[`=SUM(N${S}:N${staffLastRow})`, `=SUM(O${S}:O${staffLastRow})`]] },
  });
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: ssId,
    requestBody: { requests: [{
      repeatCell: {
        range: gridRange(staffSheetId, staffTotalRow - 1, 13, staffTotalRow - 1, 14), // N:O (0-based cols 13-14)
        cell: { userEnteredFormat: { backgroundColor: RED } },
        fields: 'userEnteredFormat.backgroundColor',
      },
    }]},
  });

  await writeOverallSignature(sheets, ssId, staffSheetName, staffSheetId, staffTotalRow);

  // Patch Supabase row_num for staff
  log(`  [SK03] Patching row_num for ${staffArray.length} staff…`);
  for (let i = 0; i < staffArray.length; i++) {
    if (staffArray[i].index == null) continue;
    await supabase.from(monthKey).update({ row_num: S + i }).eq(SB.index, staffArray[i].index);
  }

  // ══════════════════════════════════════════════════════════════════
  //  2. Intern sheet
  // ══════════════════════════════════════════════════════════════════
  log(`  [SK03] Intern sheet (${internArray.length} persons)…`);
  const internSheetId = (await sheets.spreadsheets.sheets.copyTo({
    spreadsheetId: CONFIG.sk03TemplateId, sheetId: overallTmplId,
    requestBody: { destinationSpreadsheetId: ssId },
  })).data.sheetId;

  await sheets.spreadsheets.batchUpdate({ spreadsheetId: ssId, requestBody: { requests: [
    { updateSheetProperties: { properties: { sheetId: internSheetId, title: internSheetName }, fields: 'title' } },
  ]}});

  await sheets.spreadsheets.values.update({ spreadsheetId: ssId, range: `'${internSheetName}'!A3`, valueInputOption: 'RAW', requestBody: { values: [[`ประจำเดือน  ${THAI_MONTHS[month]}  พ.ศ.  ${beYear}`]] } });

  const internRows = buildOverallRows(internArray, S, true);
  if (internRows.length > 0) {
    await sheets.spreadsheets.values.update({ spreadsheetId: ssId, range: `'${internSheetName}'!A${S}:V${S + internRows.length - 1}`, valueInputOption: 'USER_ENTERED', requestBody: { values: internRows } });
  }
  const internLastRow = S + internArray.length - 1;

  await writeOverallMeta(sheets, ssId, internSheetName, internLastRow, true, internSheetName);
  await formatOverallSheet(sheets, ssId, internSheetId, internArray, internLastRow, true);
  await writeOverallSignature(sheets, ssId, internSheetName, internSheetId, internLastRow);

  // Patch Supabase row_num for intern
  log(`  [SK03] Patching row_num for ${internArray.length} interns…`);
  for (let i = 0; i < internArray.length; i++) {
    if (internArray[i].index == null) continue;
    await supabase.from(monthKey).update({ row_num: S + i }).eq(SB.index, internArray[i].index);
  }

  // ══════════════════════════════════════════════════════════════════
  //  3. Department sheets
  // ══════════════════════════════════════════════════════════════════
  // Build row_num lookup: "prefix firstname  lastname" → row in staff/intern sheet
  const rowNumMap = {};
  staffArray.forEach( (p, i) => { rowNumMap[`${p.prefix} ${p.firstname}  ${p.lastname}`] = S + i; });
  internArray.forEach((p, i) => { rowNumMap[`${p.prefix} ${p.firstname}  ${p.lastname}`] = S + i; });

  await buildDeptSheets(sheets, ssId, depTmplId, supabasePersons, beYear, month, staffSheetName, internSheetName, rowNumMap);

  log(`  [SK03] ✓ Done → https://docs.google.com/spreadsheets/d/${ssId}`);
  return ssId;
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

  console.log('══════════════════════════════════════════════════════');
  console.log(' P4P Excel Merge + SK03 Pipeline');
  console.log('══════════════════════════════════════════════════════');
  console.log('Month(s):', months.map(m => m.key).join(', '));
  console.log('');

  const results = { processed: 0, skipped: 0, failed: [] };

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

      // Step 5
      log('  [Step 5] Merging…');
      const uploaded = await mergeAndUpload(drive, driveData.files, sbData, key);
      if (!uploaded) { results.skipped++; console.log('└─ skipped\n'); continue; }

      // Step 6
      log('  [Step 6] Creating SK03 spreadsheet…');
      await createSK03(drive, sheets, sbData, key, supabase);

      results.processed++;
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