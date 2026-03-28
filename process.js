#!/usr/bin/env node
/**
 * P4P Excel Merge Pipeline
 * Step 1  — 6-month window in BE format
 * Step 2  — Google Drive: find month folder, list Excel files
 * Step 3  — Supabase: fetch person list + metadata
 * Step 4  — Compare Drive vs Supabase lists
 * Step 5  — Merge individual Excel files → one workbook per person sheet
 * Step 6  — Build SK03 summary workbook → upload to SK03 folder
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
  rootFolderId:   process.env.GOOGLE_ROOT_FOLDER_ID,    // source year/month folders
  outputFolderId: process.env.GOOGLE_MERGE_FOLDER_ID,  // merged file destination
  sk03FolderId:   process.env.GOOGLE_SK03_FOLDER_ID,    // SK03 destination
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
const DEPT_COLORS = {
  'กุมารเวชกรรม':                          '#C8D3B8',
  'จักษุวิทยา':                            '#EEE7D3',
  'จิตเวชและยาเสพติด':                     '#D9B4E2',
  'นิติเวช':                               '#B1D1E3',
  'ผู้ป่วยนอก':                             '#EAD7C3',
  'พยาธิวิทยากายวิภาค':                    '#BEAFE1',
  'รังสีวิทยา':                             '#B0E0E6',
  'วิสัญญีวิทยา':                           '#FFFAE6',
  'ศัลยกรรม':                              '#FFE5EC',
  'ศัลยกรรมออร์โธปิดิกส์':                  '#CFDFEF',
  'สูติ-นรีเวชกรรม':                        '#F9ECD9',
  'อาชีวเวชกรรม':                           '#FFDDE4',
  'อายุรกรรม':                              '#A3DDBD',
  'เทคนิคการแพทย์และพยาธิวิทยาคลินิก':     '#F7D5C2',
  'เวชกรรมฟื้นฟู':                          '#F3BBDC',
  'เวชศาสตร์ฉุกเฉิน':                       '#9ACFC4',
  'โสต ศอ นาสิก':                           '#FADA5E',
  'INTERN':                               '#ECC1D1',
};
const DEPT_ORDER     = Object.fromEntries(Object.keys(DEPT_COLORS).map((k, i) => [k, i]));
const DEPT_ORDER_MAX = Object.keys(DEPT_COLORS).length;

// ─── Supabase column name mapping (update if DB differs) ─────────
const SB = {
  prefix:    'prefix',            // พญ. / นพ.
  position:  'position',          // นายแพทย์
  level:     'level',             // ชำนาญการ / เชี่ยวชาญ
  emp_type:  'employee_type',     // ข้าราชการ
  std:       'standard_score',    // col H  แต้มประกันมาตรฐาน (default 2200)
  boss:      'boss_score',        // col I  คะแนนประกันหัวหน้างาน
  perf:      'performance_score', // col J  คะแนนปฏิบัติงาน
  mgmt:      'management_fee',    // col M  ค่าตอบแทนบริหาร
};

// ═══════════════════════════════════════════════════════════════════
//  STEP 1 — 6-month window (current month + 5 prior)
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
function stripExt(filename) { return filename.replace(/\.(xlsx|xls)$/i, '').trim(); }
function normaliseName(name) { return (name ?? '').trim().replace(/\s+/g, ' '); }
function toSheetName(name)   { return name.replace(/[\\/:?*[\]]/g, '').substring(0, 31).trim(); }
function hexToArgb(hex) {
  if (!hex) return undefined;
  const c = hex.replace('#', '').toUpperCase();
  return c.length === 6 ? `FF${c}` : undefined;
}
function log(msg, level = 'info') {
  (level === 'warn' ? console.error : console.log)((level === 'warn' ? '⚠  ' : '') + msg);
}
/** Safely read a value from a Supabase row by column name; returns defaultVal if absent */
function sbVal(row, colName, defaultVal = null) {
  const v = row[colName];
  return (v !== null && v !== undefined && v !== '') ? v : defaultVal;
}

// ═══════════════════════════════════════════════════════════════════
//  Google Drive helpers
// ═══════════════════════════════════════════════════════════════════
function createDriveClient() {
  const auth = new google.auth.OAuth2(CONFIG.google.clientId, CONFIG.google.clientSecret);
  auth.setCredentials({ refresh_token: CONFIG.google.refreshToken });
  return google.drive({ version: 'v3', auth });
}

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

/** Upload buffer; overwrite if a file with the same name already exists in folderId */
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
//  STEP 2 — Google Drive: find month folder + list Excel files
// ═══════════════════════════════════════════════════════════════════
async function getDriveMonthData(drive, { beYear, month }) {
  const yearFolders = await listFolders(drive, CONFIG.rootFolderId);
  const yearFolder  = yearFolders.find(f => f.name === String(beYear));
  if (!yearFolder) { log(`  [Drive] Year folder "${beYear}" not found → skipping`, 'warn'); return null; }

  const thai       = THAI_MONTHS[month];
  const candidates = [`${month} - ${thai}`, `${String(month).padStart(2,'0')} - ${thai}`];
  const monthFolders = await listFolders(drive, yearFolder.id);
  const monthFolder  = monthFolders.find(f => candidates.includes(f.name));
  if (!monthFolder) { log(`  [Drive] Month folder not found (tried: ${candidates.join(' | ')}) → skipping`, 'warn'); return null; }

  const files = await listExcelFiles(drive, monthFolder.id);
  if (files.length === 0) { log(`  [Drive] No Excel files in "${monthFolder.name}" → skipping`, 'warn'); return null; }

  const names = files.map(f => normaliseName(stripExt(f.name)));
  return { names, files, folderId: monthFolder.id };
}

// ═══════════════════════════════════════════════════════════════════
//  STEP 3 — Supabase: fetch all person data
// ═══════════════════════════════════════════════════════════════════
async function getSupabaseMonthData(supabase, tableKey) {
  const { data, error } = await supabase.from(tableKey).select('*');
  if (error) {
    if (error.code === '42P01' || error.message?.includes('does not exist')) {
      log(`  [Supabase] Table "${tableKey}" not found → skipping`, 'warn'); return null;
    }
    throw new Error(`Supabase error (${tableKey}): ${error.message}`);
  }
  if (!data || data.length === 0) { log(`  [Supabase] Table "${tableKey}" is empty → skipping`, 'warn'); return null; }

  return data.map(r => ({
    fullname:      normaliseName(`${r.firstname ?? ''} ${r.lastname ?? ''}`),
    firstname:     r.firstname ?? '',
    lastname:      r.lastname  ?? '',
    department:    (r.department ?? '').trim(),
    prefix:        sbVal(r, SB.prefix,   ''),
    position:      sbVal(r, SB.position, 'นายแพทย์'),
    level:         sbVal(r, SB.level,    ''),
    emp_type:      sbVal(r, SB.emp_type, 'ข้าราชการ'),
    std_score:     sbVal(r, SB.std,      2200),
    boss_score:    sbVal(r, SB.boss,     0),
    perf_score:    sbVal(r, SB.perf,     0),
    mgmt_fee:      sbVal(r, SB.mgmt,     0),
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
//  STEP 5 — Merge individual Excel files
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
    if (srcRow.height)       dstRow.height       = srcRow.height;
    if (srcRow.hidden)       dstRow.hidden        = srcRow.hidden;
    if (srcRow.outlineLevel) dstRow.outlineLevel  = srcRow.outlineLevel;
    srcRow.eachCell({ includeEmpty: true }, (srcCell, colNumber) => {
      const dstCell = dstWS.getCell(rowNumber, colNumber);
      dstCell.value = resolveCellValue(srcCell.value);
      if (srcCell.style && Object.keys(srcCell.style).length > 0)
        dstCell.style = JSON.parse(JSON.stringify(srcCell.style));
    });
    dstRow.commit();
  });
  if (srcWS.model?.merges) {
    for (const m of srcWS.model.merges) {
      try { dstWS.mergeCells(m); } catch (_) {}
    }
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

/** Merge all Excel files → workbook with one sheet per person. Returns { uploaded, mergedWB }. */
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

  if (sheetCount === 0) { log('  [Merge] ⚠ No data — nothing to upload', 'warn'); return { uploaded: null, mergedWB }; }

  const buf  = await mergedWB.xlsx.writeBuffer();
  const name = `merged_${monthKey}.xlsx`;
  log(`\n  [Merge] Uploading "${name}" — ${sheetCount} sheets, ${totalRows} data rows…`);
  const uploaded = await uploadFileToDrive(drive, CONFIG.outputFolderId, name, buf);
  log(`  [Merge] ✓ id: ${uploaded.id}\n           🔗 ${uploaded.webViewLink}`);
  return { uploaded, mergedWB };
}

// ═══════════════════════════════════════════════════════════════════
//  STEP 6 — Build SK03 summary workbook
// ═══════════════════════════════════════════════════════════════════

/**
 * Write a single SK03-format sheet.
 * @param {object} opts
 *   isIntern      – true for the intern summary sheet
 *   internSheetName – name of the intern sheet (for staff cross-references)
 *   rateCell      – 'Y14' for staff, 'Y13' for intern (per-P4P rate cell)
 */
function buildSK03Sheet(ws, persons, beYear, month, opts = {}) {
  const { isIntern = false, internSheetName = '' } = opts;
  const monthFull = THAI_MONTHS[month];
  const S = 8;                         // first data row
  const n = persons.length;
  const L = S + n - 1;                 // last data row
  const T = L + 1;                     // totals row
  const rateCell = isIntern ? 'Y13' : 'Y14';

  // ── Column widths ──────────────────────────────────────────────
  const widths = { A:5.29, B:4.14, C:15.29, D:18.43, E:11.71, F:23.14, G:18.71,
    H:9.29, I:13, J:10.43, K:13, L:11.14, M:13.29, N:13.71, O:13,
    P:9.14, Q:8.71, R:13, S:13, T:13, U:13, V:13, W:7.86, X:33.71, Y:21 };
  Object.entries(widths).forEach(([col, w]) => { ws.getColumn(col).width = w; });

  // ── Row heights ────────────────────────────────────────────────
  [1,2,3,4].forEach(r => { ws.getRow(r).height = 21; });
  ws.getRow(5).height = 18; ws.getRow(6).height = 60.75; ws.getRow(7).height = 30;
  for (let r = S; r <= T + 3; r++) ws.getRow(r).height = 30;

  // ── Helper: set bold centered header cell ──────────────────────
  const hdr = (coord, value) => {
    const c = ws.getCell(coord);
    c.value = value;
    c.font  = { bold: true };
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  };

  // ── Rows 1-3 ──────────────────────────────────────────────────
  ws.mergeCells('A1:P1');
  hdr('A1', 'แบบสรุปภาระงานและประเมินผล (แต้มคะแนน)');
  ws.mergeCells('A2:P2');
  ws.getCell('A2').value = 'หน่วยงาน...........องค์กรแพทย์ ................                  กลุ่มภารกิจ.....................................................';
  ws.getCell('A2').alignment = { horizontal: 'left' };
  ws.mergeCells('A3:P3');
  ws.getCell('A3').value = `ประจำเดือน  ${monthFull}  พ.ศ.  ${beYear}`;
  ws.getCell('A3').alignment = { horizontal: 'center' };

  // ── Row 5: main headers (with merges spanning rows 5-7) ───────
  ws.mergeCells('A5:A7');  hdr('A5', 'ที่');
  ws.mergeCells('B5:D7');  hdr('B5', 'ชื่อ-สกุล');
  ws.mergeCells('E5:E7');  hdr('E5', 'ตำแหน่ง ');
  ws.mergeCells('F5:F7');  hdr('F5', 'ระดับ');
  ws.mergeCells('G5:G7');  hdr('G5', 'ประเภท\n(ข้าราชการ/พนักงานราชการ/ลูกจ้างประจำ/ลูกจ้างชั่วคราว)');
  ws.mergeCells('H5:H6');  hdr('H5', 'แต้มประกันมาตรฐาน\n(แต้ม)');
  ws.mergeCells('I5:I6');  hdr('I5', 'คะแนนประกันหัวหน้างาน\n(แต้ม)');
  ws.mergeCells('J5:J6');  hdr('J5', 'คะแนน \nปฏิบัติงาน\n(แต้ม)');
  ws.mergeCells('K5:K6');  hdr('K5', 'รวมคะแนน\n(แต้มที่ทำได้ทั้งหมด)\n(แต้ม)');
  ws.mergeCells('L5:O5');  hdr('L5', 'ประเมินผล');
  ws.mergeCells('P5:P7');  hdr('P5', 'หมายเหตุ');
  ws.mergeCells('V5:V7');  hdr('V5', 'MAX');

  // stats headers (not merged across rows)
  hdr('R5', 'AVG'); hdr('T5', 'STD');
  ws.getCell('S5').value = `=AVERAGE(L${S}:L${L})`;
  ws.getCell('U5').value = `=_xlfn.STDEV.P(L${S}:L${L})`;

  // ── Row 6: sub-headers ────────────────────────────────────────
  hdr('L6', 'คะแนน\nที่เกินมาตรฐาน (แต้ม)');
  hdr('M6', 'ค่าตอบแทนบริหาร');
  hdr('N6', 'ค่าตอบแทนตามสัดส่วนวิชาชีพ');
  hdr('O6', 'รวมจำนวน\nเงินที่จ่าย');
  hdr('Q6', 'AVG-2SD'); hdr('R6', 'AVG-1SD'); hdr('S6', 'AVG '); hdr('T6', 'AVG+1SD'); hdr('U6', 'AVG+2SD');

  // ── Row 7: column labels + SD statistics ──────────────────────
  ['H7:A','I7:B','J7:C','K7:D = B+C','L7: E = D-A','M7:(บาท)','N7:(บาท)','O7:(บาท)'].forEach(s => {
    const [coord, ...rest] = s.split(':');
    ws.getCell(coord).value = rest.join(':');
    ws.getCell(coord).alignment = { horizontal: 'center', vertical: 'middle' };
  });
  ws.getCell('Q7').value = '=S5-2*U5';
  ws.getCell('R7').value = '=S5-U5';
  ws.getCell('S7').value = '=S5';
  ws.getCell('T7').value = '=S5+U5';
  ws.getCell('U7').value = '=S5+2*U5';

  // ── Data rows ─────────────────────────────────────────────────
  persons.forEach((p, idx) => {
    const r = S + idx;
    ws.getCell(`A${r}`).value = idx + 1;
    ws.getCell(`B${r}`).value = p.prefix    || '';
    ws.getCell(`C${r}`).value = p.firstname || '';
    ws.getCell(`D${r}`).value = p.lastname  || '';
    ws.getCell(`E${r}`).value = p.position  || 'นายแพทย์';
    ws.getCell(`F${r}`).value = isIntern ? '' : (p.level || '');
    ws.getCell(`G${r}`).value = p.emp_type  || '';
    ws.getCell(`H${r}`).value = p.std_score  ?? 2200;
    ws.getCell(`I${r}`).value = p.boss_score ?? 0;
    ws.getCell(`J${r}`).value = p.perf_score ?? 0;
    ws.getCell(`K${r}`).value = `=I${r}+J${r}`;
    ws.getCell(`L${r}`).value = `=K${r}-H${r}`;
    ws.getCell(`M${r}`).value = p.mgmt_fee  ?? 0;
    ws.getCell(`N${r}`).value = `=ROUNDDOWN(V${r}*${rateCell},0)`;
    ws.getCell(`O${r}`).value = `=M${r}+N${r}`;
    ws.getCell(`P${r}`).value = ' ';
    ws.getCell(`Q${r}`).value = `=IF(L${r}>0,IF(L${r}>Q7,0.9,0),0)`;
    ws.getCell(`R${r}`).value = `=IF(L${r}>R7,0.95,0)`;
    ws.getCell(`S${r}`).value = `=IF(L${r}>=S7,1,0)`;
    ws.getCell(`T${r}`).value = `=IF(L${r}>T7,1.05,0)`;
    ws.getCell(`U${r}`).value = `=IF(L${r}>U7,1.1,0)`;
    ws.getCell(`V${r}`).value = `=MAX(Q${r}:U${r})`;
    ws.getRow(r).commit();
  });

  // ── Totals row ────────────────────────────────────────────────
  ws.getCell(`N${T}`).value = `=SUM(N${S}:N${L})`;
  ws.getCell(`O${T}`).value = `=SUM(O${S}:O${L})`;
  ws.getRow(T).commit();

  // ── Side panel ────────────────────────────────────────────────
  if (isIntern) {
    // Intern side panel — Y9=budget, Y12=sum points, Y13=rate, Y16=total payout
    ws.mergeCells('X8:Y8');
    ws.getCell('X8').value = 'งบจัดสรร';
    ws.getCell('X9').value = 'เฉพาะ intern';
    ws.getCell('Y9').value = 0;               // ← fill in intern budget
    ws.mergeCells('X11:Y11');
    ws.getCell('X11').value = 'ค่าที่คำนวณได้';
    ws.getCell('X12').value = 'sum point';
    ws.getCell('Y12').value = `=SUM(V${S}:V${L})`;
    ws.getCell('X13').value = 'per P4P';
    ws.getCell('Y13').value = '=Y9/Y12';
    ws.mergeCells('X15:Y15');
    ws.getCell('X15').value = 'ผลรวม';
    ws.getCell('X16').value = 'P4P intern';
    ws.getCell('Y16').value = `=SUM(O${S}:O${L})`;  // ← referenced by staff sheet
  } else {
    // Staff side panel
    ws.mergeCells('X8:Y8');
    ws.getCell('X8').value = 'งบจัดสรร';
    ws.getCell('X9').value = 'งบทั้งหมด';
    ws.getCell('Y9').value = 0;               // ← fill in total budget
    ws.getCell('X10').value = 'เฉพาะ staff';
    ws.getCell('Y10').value = `=Y9-'${internSheetName}'!Y16`;
    ws.mergeCells('X12:Y12');
    ws.getCell('X12').value = 'ค่าที่คำนวณได้';
    ws.getCell('X13').value = 'sum point';
    ws.getCell('Y13').value = `=SUM(V${S}:V${L})`;
    ws.getCell('X14').value = 'per P4P';
    ws.getCell('Y14').value = '=Y10/Y13';
    ws.mergeCells('X16:Y16');
    ws.getCell('X16').value = 'ผลรวม';
    ws.getCell('X17').value = 'ค่าตอบแทนตามสัดส่วนวิชาชีพ';
    ws.getCell('Y17').value = `=SUM(N${S}:N${L})`;
    ws.getCell('X18').value = 'ค่าตอบแทนบริหาร';
    ws.getCell('Y18').value = `=SUM(M${S}:M${L})`;
    ws.getCell('X19').value = 'P4P staff & intern (ไม่รวมค่าบริหาร)';
    ws.getCell('Y19').value = `=Y17+'${internSheetName}'!Y16`;
    ws.getCell('X20').value = 'P4P staff & intern (รวมค่าบริหาร)';
    ws.getCell('Y20').value = '=Y19+Y18';
    ws.mergeCells('X22:Y22');
    ws.getCell('X22').value = 'คงคืน';
    ws.mergeCells('X23:Y23');
    ws.getCell('X23').value = '=Y9-Y19';
  }

  // ── Signature rows ────────────────────────────────────────────
  const sig1 = T + 3, sig2 = T + 4;
  ws.mergeCells(`A${sig1}:P${sig1}`);
  ws.getCell(`A${sig1}`).value = 'ลงชื่อ................................................................................................................หัวหน้างาน/กลุ่มงาน/งาน';
  ws.getRow(sig1).commit();
  ws.mergeCells(`A${sig2}:P${sig2}`);
  ws.getCell(`A${sig2}`).value = '(....................................................................................................................................)';
  ws.getRow(sig2).commit();
}

/** Create SK03 workbook and upload to sk03FolderId */
async function createAndUploadSK03(drive, supabasePersons, monthInfo) {
  const { key, beYear, month } = monthInfo;
  const abbr        = THAI_MONTH_ABBR[month];
  const yr2         = String(beYear).slice(-2);
  const staffName   = `${abbr} ${yr2} - staff`;
  const internName  = `${abbr} ${yr2} - intern`;

  const deptMap = {};
  supabasePersons.forEach(p => { deptMap[normaliseName(p.fullname)] = p.department; });

  // Sort everyone same as merged file
  const sorted = [...supabasePersons].sort((a, b) => {
    const dA = a.department, dB = b.department;
    const idxA = DEPT_ORDER[dA] ?? DEPT_ORDER_MAX;
    const idxB = DEPT_ORDER[dB] ?? DEPT_ORDER_MAX;
    if (idxA !== idxB) return idxA - idxB;
    return normaliseName(a.fullname).localeCompare(normaliseName(b.fullname), 'th');
  });

  const staff  = sorted.filter(p => p.department !== 'INTERN');
  const intern = sorted.filter(p => p.department === 'INTERN');

  const sk03WB = new ExcelJS.Workbook();

  // Sheet 1: staff summary
  const staffWS = sk03WB.addWorksheet(staffName);
  buildSK03Sheet(staffWS, staff, beYear, month, { isIntern: false, internSheetName: internName });

  // Sheet 2: intern summary
  const internWS = sk03WB.addWorksheet(internName);
  buildSK03Sheet(internWS, intern, beYear, month, { isIntern: true });

  // Sheets 3+: one per department (DEPT_COLORS order, INTERN last)
  for (const dept of Object.keys(DEPT_COLORS)) {
    const deptPersons = sorted.filter(p => p.department === dept);
    if (deptPersons.length === 0) continue;
    const argb  = hexToArgb(DEPT_COLORS[dept]);
    const deptWS = sk03WB.addWorksheet(dept.substring(0, 31), {
      properties: argb ? { tabColor: { argb } } : {},
    });
    // Dept sheets are standalone (no intern cross-ref, use their own rate Y14)
    buildSK03Sheet(deptWS, deptPersons, beYear, month,
      { isIntern: dept === 'INTERN', internSheetName: internName });
  }

  const buf      = await sk03WB.xlsx.writeBuffer();
  const fileName = `sk03_${key}.xlsx`;
  log(`\n  [SK03] Uploading "${fileName}"…`);
  const uploaded = await uploadFileToDrive(drive, CONFIG.sk03FolderId, fileName, buf);
  log(`  [SK03] ✓ id: ${uploaded.id}\n           🔗 ${uploaded.webViewLink}`);
  return uploaded;
}

// ═══════════════════════════════════════════════════════════════════
//  MAIN
// ═══════════════════════════════════════════════════════════════════
async function main() {
  const missing = [
    ['GOOGLE_CLIENT_ID',      CONFIG.google.clientId],
    ['GOOGLE_CLIENT_SECRET',  CONFIG.google.clientSecret],
    ['GOOGLE_REFRESH_TOKEN',  CONFIG.google.refreshToken],
    ['SUPABASE_URL',          CONFIG.supabase.url],
    ['SUPABASE_KEY',          CONFIG.supabase.key],
    ['GOOGLE_ROOT_FOLDER_ID',    CONFIG.rootFolderId],
    ['GOOGLE_MERGE_FOLDER_ID',  CONFIG.outputFolderId],
    ['GOOGLE_SK03_FOLDER_ID',    CONFIG.sk03FolderId],
  ].filter(([, v]) => !v).map(([k]) => k);

  if (missing.length > 0) {
    console.error('\n❌  Missing variables in .env:');
    missing.forEach(k => console.error(`     • ${k}`));
    process.exit(1);
  }

  const drive    = createDriveClient();
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
      const { uploaded: mergeUploaded } = await mergeAndUpload(drive, driveData.files, sbData, key);
      if (!mergeUploaded) { results.skipped++; console.log('└─ skipped\n'); continue; }

      // Step 6
      log('  [Step 6] Building SK03…');
      await createAndUploadSK03(drive, sbData, monthInfo);

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