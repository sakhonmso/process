#!/usr/bin/env node
/**
 * P4P Excel Merge Pipeline
 * 1. Target month in BE format XXXX_XX
 * 2. Find Excel files in Google Drive month folder
 * 3. Fetch person list from Supabase table
 * 4. Compare lists — proceed only if they match exactly
 * 5. Merge: one sheet per person (named after them), sorted by
 *    department asc (intern last) then name asc. Upload to Drive.
 */

require('dotenv').config();

const { google }       = require('googleapis');
const { createClient } = require('@supabase/supabase-js');
const ExcelJS          = require('exceljs');
const JSZip            = require('jszip');
const { Readable }     = require('stream');

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
  rootFolderId:   process.env.GOOGLE_ROOT_FOLDER_ID,    // source: year/month folders
  outputFolderId: process.env.GOOGLE_OUTPUT_FOLDER_ID,  // destination: merged files
};

const THAI_MONTHS = {
  1:'มกราคม', 2:'กุมภาพันธ์', 3:'มีนาคม', 4:'เมษายน',
  5:'พฤษภาคม', 6:'มิถุนายน', 7:'กรกฎาคม', 8:'สิงหาคม',
  9:'กันยายน', 10:'ตุลาคม', 11:'พฤศจิกายน', 12:'ธันวาคม',
};

// Tab colours per department — keys match Supabase "department" values exactly
const DEPT_COLORS = {
  'INTERN':                               '#ECC1D1',
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
};

/**
 * Convert #RRGGBB to the ARGB string ExcelJS tabColor expects (e.g. "FFAABBCC").
 * Returns undefined if the hex is invalid so ExcelJS uses the default tab colour.
 */
function hexToArgb(hex) {
  if (!hex) return undefined;
  const clean = hex.replace('#', '').toUpperCase();
  return clean.length === 6 ? `FF${clean}` : undefined;
}

// ── STEP 1 ────────────────────────────────────────────────────────
/**
 * Return the 6-month window ending with the current month.
 * Index 0 = current month, index 5 = 5 months ago.
 * Years are in Buddhist Era (CE + 543).
 */
function getTargetMonths() {
  const now    = new Date();
  const result = [];
  for (let i = 0; i < 6; i++) {
    const d      = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const beYear = d.getFullYear() + 543;
    const month  = d.getMonth() + 1; // 1–12
    const key    = `${beYear}_${String(month).padStart(2, '0')}`;
    result.push({ key, beYear, month });
  }
  return result;
}

// ── Utility ───────────────────────────────────────────────────────
function stripExt(filename) {
  return filename.replace(/\.(xlsx|xls)$/i, '').trim();
}

/** Trim + collapse internal whitespace for reliable name comparison */
function normaliseName(name) {
  return (name ?? '').trim().replace(/\s+/g, ' ');
}

/**
 * Sanitise to a valid Excel sheet name (max 31 chars).
 * FIX #2: added backslash to illegal-char set.
 */
function toSheetName(name) {
  return name.replace(/[\\/:?*[\]]/g, '').substring(0, 31).trim();
}

/**
 * FIX #9: warnings → stderr, info → stdout.
 */
function log(msg, level = 'info') {
  (level === 'warn' ? console.error : console.log)((level === 'warn' ? '⚠  ' : '') + msg);
}

// ── Google Drive helpers ──────────────────────────────────────────
function createDriveClient() {
  const auth = new google.auth.OAuth2(CONFIG.google.clientId, CONFIG.google.clientSecret);
  auth.setCredentials({ refresh_token: CONFIG.google.refreshToken });
  return google.drive({ version: 'v3', auth });
}

/**
 * FIX #6: paginate through ALL Drive results — no silent truncation.
 */
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
    fields: 'nextPageToken, files(id, name)',
    pageSize: 100,
  });
}

async function listExcelFiles(drive, folderId) {
  const mimes = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
  ].map(m => `mimeType='${m}'`).join(' or ');
  return driveListAll(drive, {
    q: `'${folderId}' in parents and (${mimes}) and trashed=false`,
    fields: 'nextPageToken, files(id, name)',
    pageSize: 100,
  });
}

async function downloadFile(drive, fileId) {
  const res = await drive.files.get({ fileId, alt: 'media' }, { responseType: 'arraybuffer' });
  return Buffer.from(res.data);
}

/**
 * Upload buffer as xlsx to folderId.
 * If a file with the same name already exists in that folder, overwrite it
 * (Drive update) instead of creating a duplicate.
 * FIX #1: Readable.from([buffer]) emits the whole buffer as one chunk.
 */
async function uploadFileToDrive(drive, folderId, fileName, buffer) {
  const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  // Check for an existing file with the same name in the destination folder
  const existing = await driveListAll(drive, {
    q: `'${folderId}' in parents and name='${fileName.replace(/'/g, "\\'")}' and trashed=false`,
    fields: 'nextPageToken, files(id, name)',
    pageSize: 10,
  });

  if (existing.length > 0) {
    // Overwrite: update the first match (delete extras if somehow duplicated)
    const [first, ...dupes] = existing;
    for (const dupe of dupes) {
      await drive.files.delete({ fileId: dupe.id });
      log(`  [Drive] Deleted duplicate file id: ${dupe.id}`, 'warn');
    }
    log(`  [Drive] Overwriting existing file: "${fileName}" (id: ${first.id})`);
    const res = await drive.files.update({
      fileId:      first.id,
      requestBody: { name: fileName },
      media:       { mimeType: mime, body: Readable.from([buffer]) },
      fields:      'id, name, webViewLink',
    });
    return res.data;
  }

  // No existing file — create new
  const res = await drive.files.create({
    requestBody: { name: fileName, parents: [folderId], mimeType: mime },
    media:       { mimeType: mime, body: Readable.from([buffer]) },
    fields:      'id, name, webViewLink',
  });
  return res.data;
}

// ── STEP 2 ────────────────────────────────────────────────────────
async function getDriveMonthData(drive, { beYear, month }) {
  const yearFolders = await listFolders(drive, CONFIG.rootFolderId);
  const yearFolder  = yearFolders.find(f => f.name === String(beYear));
  if (!yearFolder) {
    log(`  [Drive] Year folder "${beYear}" not found → skipping`, 'warn');
    return null;
  }

  // FIX #5: try both zero-padded and plain month numbers
  const thai      = THAI_MONTHS[month];
  const candidates = [`${month} - ${thai}`, `${String(month).padStart(2,'0')} - ${thai}`];
  const monthFolders = await listFolders(drive, yearFolder.id);
  const monthFolder  = monthFolders.find(f => candidates.includes(f.name));
  if (!monthFolder) {
    log(`  [Drive] Month folder not found (tried: ${candidates.join(' | ')}) → skipping`, 'warn');
    return null;
  }

  const files = await listExcelFiles(drive, monthFolder.id);
  if (files.length === 0) {
    log(`  [Drive] No Excel files in "${monthFolder.name}" → skipping`, 'warn');
    return null;
  }

  // FIX #10: normalise names here so they match Supabase names reliably
  const names = files.map(f => normaliseName(stripExt(f.name)));
  return { names, files, folderId: monthFolder.id };
}

// ── STEP 3 ────────────────────────────────────────────────────────
async function getSupabaseMonthData(supabase, tableKey) {
  const { data, error } = await supabase.from(tableKey).select('firstname, lastname, department');
  if (error) {
    if (error.code === '42P01' || error.message?.includes('does not exist')) {
      log(`  [Supabase] Table "${tableKey}" not found → skipping`, 'warn');
      return null;
    }
    throw new Error(`Supabase error (${tableKey}): ${error.message}`);
  }
  if (!data || data.length === 0) {
    log(`  [Supabase] Table "${tableKey}" is empty → skipping`, 'warn');
    return null;
  }
  return data.map(r => ({
    // FIX #3: trim each field before joining; guard nulls
    fullname:   normaliseName(`${r.firstname ?? ''} ${r.lastname ?? ''}`),
    department: (r.department ?? '').trim(),
  }));
}

// ── STEP 4 ────────────────────────────────────────────────────────
function listsMatch(driveNames, supabasePersons) {
  // FIX #4: normalise both sides before any comparison
  const normDrive = driveNames.map(normaliseName);
  const normSB    = supabasePersons.map(p => normaliseName(p.fullname));

  if (normDrive.length !== normSB.length) {
    log(`  [Compare] Length mismatch — Drive: ${normDrive.length}, Supabase: ${normSB.length}`, 'warn');
    normDrive.forEach(n => { if (!normSB.includes(n))   log(`    ✗ In Drive but not Supabase: "${n}"`, 'warn'); });
    normSB.forEach(n   => { if (!normDrive.includes(n)) log(`    ✗ In Supabase but not Drive: "${n}"`, 'warn'); });
    return false;
  }

  const sorted1 = [...normDrive].sort();
  const sorted2 = [...normSB].sort();
  let match = true;
  for (let i = 0; i < sorted1.length; i++) {
    if (sorted1[i] !== sorted2[i]) {
      log(`  [Compare] Value mismatch: "${sorted1[i]}" ≠ "${sorted2[i]}"`, 'warn');
      match = false; // don't break — report all mismatches
    }
  }
  return match;
}

// ── STEP 5 ────────────────────────────────────────────────────────
function isSheetEmpty(ws) {
  let count = 0;
  ws.eachRow(row => {
    row.eachCell(cell => {
      if (cell.value !== null && cell.value !== undefined && cell.value !== '') count++;
    });
  });
  return count <= 3;
}

/**
 * Pre-strip all formulas from xlsx XML before ExcelJS parses it.
 * FIX #8: use [\s\S]*? (lazy dotall) so formulas containing '>'
 *         (e.g. comparison: A1>0) are also matched and removed.
 */
async function stripFormulas(buffer) {
  const zip        = await JSZip.loadAsync(buffer);
  const sheetPaths = Object.keys(zip.files).filter(
    p => /^xl\/worksheets\/sheet\d+\.xml$/.test(p),
  );
  for (const path of sheetPaths) {
    let xml = await zip.files[path].async('string');
    xml = xml.replace(/<f\b[^>]*>[\s\S]*?<\/f>/g, ''); // <f ...>content</f>
    xml = xml.replace(/<f\b[^/]*\/>/g, '');              // <f ... />
    xml = xml.replace(/\s+si="\d+"/g, '');               // shared formula index attr
    xml = xml.replace(/\s+t="shared"/g, '');             // shared formula type attr
    zip.file(path, xml);
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function loadWorkbook(buffer) {
  const cleaned = await stripFormulas(buffer);
  const wb      = new ExcelJS.Workbook();
  await wb.xlsx.load(cleaned);
  return wb;
}

/**
 * Resolve an ExcelJS cell value to a plain primitive.
 * Formulas are replaced with their cached result (preserved by stripFormulas).
 */
function resolveCellValue(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if ('richText' in v) return v;              // keep rich-text object — ExcelJS writes it correctly
    if ('text'     in v) return v.text;         // hyperlink → plain text
    if ('result'   in v) return v.result;       // cached formula result
    if ('error'    in v) return null;           // formula error → blank
  }
  return v;
}

/**
 * Copy a worksheet faithfully from src to dst, preserving:
 *   - Cell values (resolved) and rich-text
 *   - Cell styles (font, fill, border, alignment, numFmt)
 *   - Row heights and hidden state
 *   - Column widths, hidden state, and outline level
 *   - Merged cell regions
 */
function copyWorksheet(srcWS, dstWS) {
  // ── Column widths / properties ──────────────────────────────────
  srcWS.columns.forEach((srcCol, idx) => {
    const dstCol = dstWS.getColumn(idx + 1);
    if (srcCol.width)        dstCol.width        = srcCol.width;
    if (srcCol.hidden)       dstCol.hidden        = srcCol.hidden;
    if (srcCol.outlineLevel) dstCol.outlineLevel  = srcCol.outlineLevel;
  });

  // ── Rows + cells ────────────────────────────────────────────────
  srcWS.eachRow({ includeEmpty: true }, (srcRow, rowNumber) => {
    const dstRow = dstWS.getRow(rowNumber);

    // Row-level properties
    if (srcRow.height)  dstRow.height  = srcRow.height;
    if (srcRow.hidden)  dstRow.hidden  = srcRow.hidden;
    if (srcRow.outlineLevel) dstRow.outlineLevel = srcRow.outlineLevel;

    srcRow.eachCell({ includeEmpty: true }, (srcCell, colNumber) => {
      const dstCell = dstWS.getCell(rowNumber, colNumber);

      // Value
      dstCell.value = resolveCellValue(srcCell.value);

      // Style — deep-copy so objects are not shared between workbooks
      if (srcCell.style && Object.keys(srcCell.style).length > 0) {
        dstCell.style = JSON.parse(JSON.stringify(srcCell.style));
      }
    });

    dstRow.commit();
  });

  // ── Merged cells ────────────────────────────────────────────────
  if (srcWS.model && srcWS.model.merges) {
    for (const mergeRange of srcWS.model.merges) {
      try { dstWS.mergeCells(mergeRange); } catch (_) { /* skip overlapping */ }
    }
  }
}

function sortByDeptThenName(a, b, deptMap) {
  // FIX #10: look up by normName (normalised) so deptMap always hits
  const dA = (deptMap[a.normName] ?? '').toLowerCase();
  const dB = (deptMap[b.normName] ?? '').toLowerCase();
  if (dA === 'intern' && dB !== 'intern') return  1;
  if (dA !== 'intern' && dB === 'intern') return -1;
  if (dA !== dB) return dA.localeCompare(dB, 'th');
  return a.normName.localeCompare(b.normName, 'th');
}

async function mergeAndUpload(drive, driveFiles, supabasePersons, monthKey) {
  // FIX #10: key the deptMap by normalised fullname
  const deptMap = {};
  supabasePersons.forEach(p => { deptMap[normaliseName(p.fullname)] = p.department; });

  const filesWithName = driveFiles.map(f => ({
    ...f,
    origName: stripExt(f.name),
    normName: normaliseName(stripExt(f.name)),
  }));
  filesWithName.sort((a, b) => sortByDeptThenName(a, b, deptMap));

  log('\n  [Merge] Sorted order:');
  filesWithName.forEach((f, i) => {
    log(`    ${String(i+1).padStart(2)}. ${f.origName.padEnd(32)} [${deptMap[f.normName] ?? '?'}]`);
  });

  const mergedWB       = new ExcelJS.Workbook();
  const usedSheetNames = new Set();
  let   totalRows      = 0;
  let   sheetCount     = 0;

  for (const file of filesWithName) {
    log(`\n  [Merge] ↓ ${file.origName}`);
    const buffer = await downloadFile(drive, file.id);
    const wb     = await loadWorkbook(buffer);

    let targetWS = null;
    for (const ws of wb.worksheets) {
      if (!isSheetEmpty(ws)) { targetWS = ws; log(`      → using sheet: "${ws.name}"`); break; }
      log(`      → sheet "${ws.name}" is empty, trying next`);
    }
    if (!targetWS) {
      log(`  [Merge] ⚠ All sheets empty in "${file.origName}" — skipped`, 'warn');
      continue;
    }

    // FIX #7: deduplicate sheet names with a counter suffix
    let sheetName = toSheetName(file.origName);
    if (usedSheetNames.has(sheetName)) {
      let c = 2;
      while (usedSheetNames.has(`${sheetName}_${c}`)) c++;
      sheetName = `${sheetName}_${c}`.substring(0, 31);
    }
    usedSheetNames.add(sheetName);

    const dept     = deptMap[file.normName] ?? '';
    const tabArgb  = hexToArgb(DEPT_COLORS[dept]);
    const outWS    = mergedWB.addWorksheet(sheetName, {
      properties: tabArgb ? { tabColor: { argb: tabArgb } } : {},
    });
    sheetCount++;

    copyWorksheet(targetWS, outWS);

    const dataRows = outWS.rowCount;
    totalRows += Math.max(0, dataRows - 1);
    log(`      → ${dataRows} rows copied to sheet "${sheetName}" [${dept || '?'}]${tabArgb ? ' 🎨' : ''}`);
  }

  if (sheetCount === 0) {
    log('  [Merge] ⚠ No data accumulated — nothing to upload', 'warn');
    return null;
  }

  const outputBuffer = await mergedWB.xlsx.writeBuffer();
  const outputName   = `merged_${monthKey}.xlsx`;

  log(`\n  [Merge] Uploading "${outputName}" — ${sheetCount} sheets, ${totalRows} data rows…`);
  const uploaded = await uploadFileToDrive(drive, CONFIG.outputFolderId, outputName, outputBuffer);
  log(`  [Merge] ✓ Uploaded — id: ${uploaded.id}`);
  log(`           🔗 ${uploaded.webViewLink}`);
  return uploaded;
}

// ── MAIN ──────────────────────────────────────────────────────────
async function main() {
  const missing = [
    ['GOOGLE_CLIENT_ID',      CONFIG.google.clientId],
    ['GOOGLE_CLIENT_SECRET',  CONFIG.google.clientSecret],
    ['GOOGLE_REFRESH_TOKEN',  CONFIG.google.refreshToken],
    ['SUPABASE_URL',          CONFIG.supabase.url],
    ['SUPABASE_KEY',          CONFIG.supabase.key],
    ['GOOGLE_ROOT_FOLDER_ID',    CONFIG.rootFolderId],
    ['GOOGLE_OUTPUT_FOLDER_ID',  CONFIG.outputFolderId],
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
  console.log(' P4P Excel Merge Pipeline');
  console.log('══════════════════════════════════════════════════════');
  console.log('Month(s):', months.map(m => m.key).join(', '));
  console.log('');

  const results = { processed: 0, skipped: 0, failed: [] };

  for (const monthInfo of months) {
    const { key } = monthInfo;
    console.log(`\n┌─ ${key} ${'─'.repeat(46 - key.length)}`);
    try {
      log('  [Step 2] Google Drive — finding month folder…');
      const driveData = await getDriveMonthData(drive, monthInfo);
      if (!driveData) { results.skipped++; console.log('└─ skipped\n'); continue; }
      log(`  → ${driveData.names.length} file(s): ${driveData.names.join(', ')}`);

      log('  [Step 3] Supabase — querying table…');
      const sbData = await getSupabaseMonthData(supabase, key);
      if (!sbData) { results.skipped++; console.log('└─ skipped\n'); continue; }
      log(`  → ${sbData.length} person(s): ${sbData.map(p => p.fullname).join(', ')}`);

      log('  [Step 4] Comparing lists…');
      if (!listsMatch(driveData.names, sbData)) {
        log('  → ✗ Lists do not match — skipping merge', 'warn');
        results.skipped++; console.log('└─ skipped\n'); continue;
      }
      log('  → ✓ Lists match');

      log('  [Step 5] Merging and uploading…');
      const uploaded = await mergeAndUpload(drive, driveData.files, sbData, key);
      if (uploaded) results.processed++;
      console.log('└─ ✓ complete\n');
    } catch (err) {
      console.error(`  [ERROR] ${err.stack ?? err.message}`);
      results.failed.push({ key, error: err.message });
      console.log('└─ failed\n');
    }
  }

  console.log('══════════════════════════════════════════════════════');
  console.log(` ✓ Merged  : ${results.processed}`);
  console.log(` ⊘ Skipped : ${results.skipped}`);
  console.log(` ✗ Failed  : ${results.failed.length}`);
  results.failed.forEach(f => console.log(`     ${f.key}: ${f.error}`));
  console.log('');
}

main().catch(err => { console.error('\nFatal:', err.stack ?? err); process.exit(1); });