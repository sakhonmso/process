'use strict';
/**
 * report.js — Missing submission tracker
 *
 * For each of the last 6 months, compares Supabase person list against
 * Google Drive Excel files (Supabase as reference). Reports physicians
 * who exist in Supabase but have no corresponding Excel file in Drive.
 * Complete months (no missing files) are omitted from the report.
 *
 * Output: missing_submissions.xlsx → uploaded (overwritten) to REPORT_FOLDER_ID daily.
 *
 * Steps:
 *   1. Get 6-month window
 *   2. For each month: fetch SB persons + Drive file names
 *   3. Diff: SB persons not found in Drive
 *   4. Skip month if diff is empty (complete)
 *   5. Build Excel (Month | Name | Department), months descending
 *   6. Upload / overwrite in Drive
 */

const { google }       = require('googleapis');
const { createClient } = require('@supabase/supabase-js');
const ExcelJS          = require('exceljs');
const { Readable }     = require('stream');

// ═══════════════════════════════════════════════════════════════════
//  Config
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
  reportFolderId: '1vbEX7-RRCPij2UkkHwnc9oSbQ2HRfXkq',
  reportFileName: 'missing_submissions.xlsx',
};

// ═══════════════════════════════════════════════════════════════════
//  Thai month names
// ═══════════════════════════════════════════════════════════════════
const THAI_MONTHS = {
  1: 'มกราคม',   2: 'กุมภาพันธ์', 3: 'มีนาคม',
  4: 'เมษายน',   5: 'พฤษภาคม',   6: 'มิถุนายน',
  7: 'กรกฎาคม',  8: 'สิงหาคม',   9: 'กันยายน',
  10: 'ตุลาคม', 11: 'พฤศจิกายน', 12: 'ธันวาคม',
};

// ═══════════════════════════════════════════════════════════════════
//  Utilities
// ═══════════════════════════════════════════════════════════════════
function normaliseName(name) { return (name ?? '').trim().replace(/\s+/g, ' '); }
function stripExt(filename)  { return filename.replace(/\.(xlsx|xls)$/i, '').trim(); }
function log(msg, level = 'info') {
  (level === 'warn' ? console.error : console.log)((level === 'warn' ? '⚠  ' : '') + msg);
}
const sleep = ms => new Promise(r => setTimeout(r, ms));

async function withRetry(fn, maxRetries = 7) {
  let delay = 3000;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return await fn();
    } catch (err) {
      const status   = err?.response?.status ?? err?.status ?? 0;
      const msg      = err?.message ?? '';
      const isQuota  = status === 429 || msg.includes('Quota exceeded') || msg.includes('RESOURCE_EXHAUSTED');
      const isServer = status >= 500 && status < 600;
      if ((isQuota || isServer) && attempt < maxRetries) {
        const wait = delay + Math.random() * 1000;
        log(`  [Retry] ${isQuota ? 'Quota' : 'Server'} — waiting ${(wait / 1000).toFixed(1)}s (attempt ${attempt + 1}/${maxRetries})`, 'warn');
        await sleep(wait);
        delay = Math.min(delay * 2, 60000);
      } else {
        throw err;
      }
    }
  }
}

// ═══════════════════════════════════════════════════════════════════
//  Google API clients
// ═══════════════════════════════════════════════════════════════════
function createAuth() {
  const auth = new google.auth.OAuth2(CONFIG.google.clientId, CONFIG.google.clientSecret);
  auth.setCredentials({ refresh_token: CONFIG.google.refreshToken });
  return auth;
}
function createDriveClient() { return google.drive({ version: 'v3', auth: createAuth() }); }

// ═══════════════════════════════════════════════════════════════════
//  Drive helpers
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

// ═══════════════════════════════════════════════════════════════════
//  Step 1 — 6-month window (newest first)
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
//  Step 2a — Get normalised Drive file name set for a month
// ═══════════════════════════════════════════════════════════════════
async function getDriveNameSet(drive, { beYear, month }) {
  const yearFolders = await listFolders(drive, CONFIG.rootFolderId);
  const yearFolder  = yearFolders.find(f => f.name === String(beYear));
  if (!yearFolder) return null;

  const thai       = THAI_MONTHS[month];
  const candidates = [`${month} - ${thai}`, `${String(month).padStart(2, '0')} - ${thai}`];
  const monthFolders = await listFolders(drive, yearFolder.id);
  const monthFolder  = monthFolders.find(f => candidates.includes(f.name));
  if (!monthFolder) return null;

  const files = await listExcelFiles(drive, monthFolder.id);
  return new Set(files.map(f => normaliseName(stripExt(f.name))));
}

// ═══════════════════════════════════════════════════════════════════
//  Step 2b — Get Supabase persons for a month
// ═══════════════════════════════════════════════════════════════════
async function getSupabasePersons(supabase, tableKey) {
  const { data, error } = await supabase
    .from(tableKey)
    .select('firstname, lastname, department');

  if (error) {
    if (error.code === '42P01' || error.message?.includes('does not exist')) return null;
    throw new Error(`Supabase error (${tableKey}): ${error.message}`);
  }
  if (!data || data.length === 0) return null;

  return data.map(r => ({
    fullname:   normaliseName(`${r.firstname ?? ''} ${r.lastname ?? ''}`),
    department: (r.department ?? '').trim(),
  }));
}

// ═══════════════════════════════════════════════════════════════════
//  Step 5 — Build Excel workbook
// ═══════════════════════════════════════════════════════════════════
async function buildExcel(reportRows) {
  const wb = new ExcelJS.Workbook();
  wb.creator  = 'Missing Submission Tracker';
  wb.created  = new Date();
  wb.modified = new Date();

  const ws = wb.addWorksheet('Missing Submissions');

  ws.columns = [
    { header: 'เดือน',       key: 'month',      width: 24 },
    { header: 'ชื่อ-นามสกุล', key: 'fullname',   width: 36 },
    { header: 'กลุ่มงาน',    key: 'department', width: 32 },
  ];

  // Header row styling
  const headerRow = ws.getRow(1);
  headerRow.height = 22;
  headerRow.eachCell(cell => {
    cell.font      = { bold: true, size: 11 };
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border    = {
      top:    { style: 'medium' },
      left:   { style: 'medium' },
      bottom: { style: 'medium' },
      right:  { style: 'medium' },
    };
  });

  if (reportRows.length === 0) {
    const row = ws.addRow(['ไม่มีข้อมูลที่ขาดส่ง', '', '']);
    ws.mergeCells(`A${row.number}:C${row.number}`);
    row.getCell(1).alignment = { horizontal: 'center' };
    row.getCell(1).font      = { italic: true, color: { argb: 'FF888888' } };
  } else {
    let prevMonth = null;
    reportRows.forEach((r, idx) => {
      // Only show month label on first row of each group
      const monthCell = r.monthLabel === prevMonth ? '' : r.monthLabel;
      prevMonth = r.monthLabel;

      const row = ws.addRow({ month: monthCell, fullname: r.fullname, department: r.department });
      row.height = 18;
      row.eachCell({ includeEmpty: true }, cell => {
        cell.border = {
          top:    { style: 'thin' },
          left:   { style: 'thin' },
          bottom: { style: 'thin' },
          right:  { style: 'thin' },
        };
        cell.alignment = { vertical: 'middle' };
      });

      // Shade alternating month groups for readability
      const groupIdx = reportRows.filter((x, i) => i <= idx && x.monthLabel === r.monthLabel).length;
      const isOddGroup = (reportRows.findIndex(x => x.monthLabel === r.monthLabel) % 2 === 0);
      if (isOddGroup) {
        row.getCell('month').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
      }
    });
  }

  // Freeze header row
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

// ═══════════════════════════════════════════════════════════════════
//  Step 6 — Upload / overwrite in Drive
// ═══════════════════════════════════════════════════════════════════
async function uploadReport(drive, buffer) {
  const mime     = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const fileName = CONFIG.reportFileName;
  const folderId = CONFIG.reportFolderId;

  const existing = await driveListAll(drive, {
    q: `'${folderId}' in parents and name='${fileName}' and trashed=false`,
    fields: 'nextPageToken, files(id, name)',
    pageSize: 10,
  });

  if (existing.length > 0) {
    const [first, ...dupes] = existing;
    for (const d of dupes) {
      await withRetry(() => drive.files.delete({ fileId: d.id }));
      log(`  [Drive] Deleted duplicate: ${d.id}`, 'warn');
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

  log(`  [Drive] Creating "${fileName}"`);
  const res = await withRetry(() => drive.files.create({
    requestBody: { name: fileName, parents: [folderId], mimeType: mime },
    media: { mimeType: mime, body: Readable.from([buffer]) },
    fields: 'id, name, webViewLink',
  }));
  return res.data;
}

// ═══════════════════════════════════════════════════════════════════
//  Main
// ═══════════════════════════════════════════════════════════════════
async function main() {
  const missingEnv = [
    ['GOOGLE_CLIENT_ID',       CONFIG.google.clientId],
    ['GOOGLE_CLIENT_SECRET',   CONFIG.google.clientSecret],
    ['GOOGLE_REFRESH_TOKEN',   CONFIG.google.refreshToken],
    ['SUPABASE_URL',           CONFIG.supabase.url],
    ['SUPABASE_KEY',           CONFIG.supabase.key],
    ['GOOGLE_ROOT_FOLDER_ID',  CONFIG.rootFolderId],
  ].filter(([, v]) => !v).map(([k]) => k);

  if (missingEnv.length > 0) {
    console.error('\n❌  Missing env vars:', missingEnv.join(', '));
    process.exit(1);
  }

  const drive    = createDriveClient();
  const supabase = createClient(CONFIG.supabase.url, CONFIG.supabase.key);
  const months   = getTargetMonths(); // newest first

  console.log('══════════════════════════════════════════════════════');
  console.log(' Missing Submission Tracker');
  console.log('══════════════════════════════════════════════════════');
  console.log('Months:', months.map(m => m.key).join(', '));
  console.log('');

  const reportRows = []; // will be month-descending (months iterated newest first)

  for (const monthInfo of months) {
    const { key, beYear, month } = monthInfo;
    const monthLabel = `${THAI_MONTHS[month]} ${beYear}`;
    log(`\n── ${key} (${monthLabel})`);

    // Fetch Supabase persons
    const sbPersons = await getSupabasePersons(supabase, key);
    if (!sbPersons) {
      log(`  [SB] Table not found or empty — skipping`);
      continue;
    }
    log(`  [SB] ${sbPersons.length} persons`);

    // Fetch Drive file names
    const driveNames = await getDriveNameSet(drive, monthInfo);
    if (driveNames === null) {
      // Folder not found — treat all SB persons as missing
      log(`  [Drive] Folder not found — all ${sbPersons.length} persons have no file`);
      sbPersons.forEach(p => reportRows.push({ monthLabel, fullname: p.fullname, department: p.department }));
      continue;
    }
    log(`  [Drive] ${driveNames.size} files`);

    // Diff: in Supabase but not in Drive
    const missing = sbPersons.filter(p => !driveNames.has(normaliseName(p.fullname)));

    if (missing.length === 0) {
      log(`  ✓ Complete — omitting from report`);
      continue;
    }

    log(`  ✗ ${missing.length} physician(s) have no Drive file:`);
    missing.forEach(p => {
      log(`    • ${p.fullname} [${p.department}]`);
      reportRows.push({ monthLabel, fullname: p.fullname, department: p.department });
    });
  }

  console.log('\n══════════════════════════════════════════════════════');
  console.log(` Total missing: ${reportRows.length} physician-month entries`);
  console.log('');

  // Build and upload Excel
  log('[Excel] Building workbook…');
  const buffer   = await buildExcel(reportRows);
  log('[Drive] Uploading report…');
  const uploaded = await uploadReport(drive, buffer);
  log(`\n✓ Report saved: ${uploaded.webViewLink}`);
}

main().catch(err => {
  console.error('\n❌ Fatal:', err.stack ?? err.message);
  process.exit(1);
});
