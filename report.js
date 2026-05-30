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
  reportFileName: 'รายชื่อแพทย์ค้างส่ง P4P.xlsx',
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

const THAI_MONTH_ABBR = {
  1: 'ม.ค.',  2: 'ก.พ.',  3: 'มี.ค.',
  4: 'เม.ย.', 5: 'พ.ค.',  6: 'มิ.ย.',
  7: 'ก.ค.',  8: 'ส.ค.',  9: 'ก.ย.',
  10: 'ต.ค.', 11: 'พ.ย.', 12: 'ธ.ค.',
};

/** Format current Bangkok time as "25 เม.ย. 69, 14.14" */
function formatRunTime() {
  // Bangkok = UTC+7
  const now = new Date(Date.now() + 7 * 60 * 60 * 1000);
  const day    = now.getUTCDate();
  const month  = now.getUTCMonth() + 1;
  const year   = String(now.getUTCFullYear() + 543).slice(-2);
  const hour   = String(now.getUTCHours()).padStart(2, '0');
  const minute = String(now.getUTCMinutes()).padStart(2, '0');
  return `${day} ${THAI_MONTH_ABBR[month]} ${year}, ${hour}.${minute}`;
}

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
//  Step 1 — 6-month window, excluding the current month (newest first)
// ═══════════════════════════════════════════════════════════════════
function getTargetMonths() {
  const now = new Date();
  // i starts at 1 to skip the current month entirely
  return Array.from({ length: 6 }, (_, i) => {
    const d      = new Date(now.getFullYear(), now.getMonth() - (i + 1), 1);
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
//  Department sort helper — Thai ascending, INTERN last
// ═══════════════════════════════════════════════════════════════════
function sortDepartments(depts) {
  const nonIntern = [...depts].filter(d => d !== 'INTERN').sort((a, b) => a.localeCompare(b, 'th'));
  return depts.includes('INTERN') ? [...nonIntern, 'INTERN'] : nonIntern;
}

// ═══════════════════════════════════════════════════════════════════
//  Shared cell style helpers
// ═══════════════════════════════════════════════════════════════════
const BORDER_THIN   = { style: 'thin' };
const BORDER_MEDIUM = { style: 'medium' };
const FULL_THIN     = { top: BORDER_THIN,   left: BORDER_THIN,   bottom: BORDER_THIN,   right: BORDER_THIN };
const FULL_MEDIUM   = { top: BORDER_MEDIUM, left: BORDER_MEDIUM, bottom: BORDER_MEDIUM, right: BORDER_MEDIUM };

// ═══════════════════════════════════════════════════════════════════
//  Step 5 — Build Excel workbook
//    Sheet 1  : ภาพรวม  (summary counts by dept × month)
//    Sheet 2+ : one sheet per incomplete month
//               - table: ชื่อ-นามสกุล | กลุ่มงาน
//               - footnote merged row immediately below table
// ═══════════════════════════════════════════════════════════════════
async function buildExcel(monthGroups, runTime, allDepts) {
  // monthGroups: [{ monthLabel, rows: [{ fullname, department }] }]
  // already in descending month order; rows sorted by Thai name asc
  // allDepts: Set of every department seen in Supabase across all months

  const wb = new ExcelJS.Workbook();
  wb.creator  = 'Missing Submission Tracker';
  wb.created  = new Date();
  wb.modified = new Date();

  // ── Sheet 1: ภาพรวม ───────────────────────────────────────────────
  {
    const ws = wb.addWorksheet('ภาพรวม');

    // Departments with at least one missing physician (incomplete)
    const incompleteDeptSet = new Set();
    monthGroups.forEach(g => g.rows.forEach(r => incompleteDeptSet.add(r.department)));
    const incompleteDepts = sortDepartments([...incompleteDeptSet]);

    // Departments where all physicians submitted (complete) — from allDepts minus incomplete
    const completeDepts = sortDepartments(
      [...(allDepts ?? [])].filter(d => !incompleteDeptSet.has(d))
    );

    // Column layout: กลุ่มงาน | month1 | month2 | ... | รวม
    ws.columns = [
      { header: 'กลุ่มงาน', key: 'dept', width: 34 },
      ...monthGroups.map(g => ({ header: g.monthLabel, key: g.monthLabel, width: 20 })),
      { header: 'รวม', key: 'total', width: 10 },
    ];

    // Style header row
    const hdr = ws.getRow(1);
    hdr.height = 22;
    hdr.eachCell({ includeEmpty: true }, cell => {
      cell.font      = { bold: true, size: 11 };
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border    = FULL_MEDIUM;
    });

    // Incomplete dept rows
    for (const dept of incompleteDepts) {
      const counts  = monthGroups.map(g => g.rows.filter(r => r.department === dept).length);
      const total   = counts.reduce((s, c) => s + c, 0);
      const rowData = { dept, total };
      monthGroups.forEach((g, i) => { rowData[g.monthLabel] = counts[i]; });
      const row = ws.addRow(rowData);
      row.height = 18;
      row.eachCell({ includeEmpty: true }, cell => {
        cell.border    = FULL_THIN;
        cell.alignment = { vertical: 'middle', horizontal: cell.col === 1 ? 'left' : 'center' };
      });
    }

    // รวม row (totals)
    const colTotals = monthGroups.map(g => g.rows.length);
    const grandTotal = colTotals.reduce((s, c) => s + c, 0);
    const totalRowData = { dept: 'รวม', total: grandTotal };
    monthGroups.forEach((g, i) => { totalRowData[g.monthLabel] = colTotals[i]; });
    const totalRow = ws.addRow(totalRowData);
    totalRow.height = 18;
    totalRow.eachCell({ includeEmpty: true }, cell => {
      cell.font      = { bold: true };
      cell.border    = FULL_MEDIUM;
      cell.alignment = { vertical: 'middle', horizontal: cell.col === 1 ? 'left' : 'center' };
    });

    // Footnote: completed depts merged across all columns, row immediately below รวม
    if (completeDepts.length > 0) {
      const totalColCount = 1 + monthGroups.length + 1; // กลุ่มงาน + months + รวม
      const lastColLetter = ws.getColumn(totalColCount).letter;
      const footnoteRowNum = totalRow.number + 1;
      ws.mergeCells(`A${footnoteRowNum}:${lastColLetter}${footnoteRowNum}`);
      const fnCell = ws.getCell(`A${footnoteRowNum}`);
      fnCell.value     = `กลุ่มงานที่ครบถ้วน : ${completeDepts.join('  |  ')}`;
      fnCell.border    = FULL_THIN;
      fnCell.alignment = { horizontal: 'left', vertical: 'middle' };
      fnCell.font      = { italic: true, size: 10, color: { argb: 'FF000000' } };
      fnCell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
      ws.getRow(footnoteRowNum).height = 18;
    }

    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  // ── Sheets 2+: one per incomplete month ───────────────────────────
  for (const group of monthGroups) {
    const ws = wb.addWorksheet(group.monthLabel);

    ws.columns = [
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
      cell.border    = FULL_MEDIUM;
    });

    if (group.rows.length === 0) {
      const row = ws.addRow(['ไม่มีข้อมูลที่ขาดส่ง', '']);
      ws.mergeCells(`A${row.number}:B${row.number}`);
      row.getCell(1).alignment = { horizontal: 'center' };
      row.getCell(1).font      = { italic: true, color: { argb: 'FF888888' } };
    } else {
      group.rows.forEach(r => {
        const row = ws.addRow({ fullname: r.fullname, department: r.department });
        row.height = 18;
        row.eachCell({ includeEmpty: true }, cell => {
          cell.border    = FULL_THIN;
          cell.alignment = { vertical: 'middle' };
        });
      });
    }

    // Footnote: merged A:B, row immediately below last data row, black bg / white text
    const lastDataRowNum = 1 + Math.max(group.rows.length, 1); // header + data (min 1)
    const footnoteRowNum = lastDataRowNum + 1;
    ws.mergeCells(`A${footnoteRowNum}:B${footnoteRowNum}`);
    const footnoteCell = ws.getCell(`A${footnoteRowNum}`);
    footnoteCell.value     = `ตรวจสอบเมื่อ : ${runTime}`;
    footnoteCell.border    = FULL_THIN;
    footnoteCell.alignment = { horizontal: 'center', vertical: 'middle' };
    footnoteCell.font      = { italic: true, size: 10, color: { argb: 'FFFFFFFF' } };
    footnoteCell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } };
    ws.getRow(footnoteRowNum).height = 18;

    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

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

  // monthGroups: one entry per incomplete month, in descending order
  const monthGroups = [];
  const allDeptSet  = new Set(); // all departments seen across all SB months
  let totalMissing  = 0;

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

    // Accumulate all departments seen in Supabase
    sbPersons.forEach(p => allDeptSet.add(p.department));

    // Fetch Drive file names
    const driveNames = await getDriveNameSet(drive, monthInfo);
    if (driveNames === null) {
      // Folder not found — all SB persons are missing
      log(`  [Drive] Folder not found — all ${sbPersons.length} persons have no file`);
      const rows = [...sbPersons].sort((a, b) => a.fullname.localeCompare(b.fullname, 'th'));
      monthGroups.push({ monthLabel, rows });
      totalMissing += rows.length;
      rows.forEach(p => log(`    • ${p.fullname} [${p.department}]`));
      continue;
    }
    log(`  [Drive] ${driveNames.size} files`);

    // Diff: in Supabase but not in Drive
    const missing = sbPersons.filter(p => !driveNames.has(normaliseName(p.fullname)));

    if (missing.length === 0) {
      log(`  ✓ Complete — omitting from report`);
      continue;
    }

    // Sort missing by name ascending (Thai locale)
    const rows = missing.sort((a, b) => a.fullname.localeCompare(b.fullname, 'th'));
    log(`  ✗ ${rows.length} physician(s) have no Drive file:`);
    rows.forEach(p => log(`    • ${p.fullname} [${p.department}]`));
    monthGroups.push({ monthLabel, rows });
    totalMissing += rows.length;
  }

  console.log('\n══════════════════════════════════════════════════════');
  console.log(` Incomplete months : ${monthGroups.length}`);
  console.log(` Total missing     : ${totalMissing} physician-month entries`);
  console.log('');

  // Build and upload Excel
  const runTime = formatRunTime();
  log(`[Excel] Building workbook (ตรวจสอบเมื่อ: ${runTime})…`);
  const buffer   = await buildExcel(monthGroups, runTime, allDeptSet);
  log('[Drive] Uploading report…');
  const uploaded = await uploadReport(drive, buffer);
  log(`\n✓ Report saved: ${uploaded.webViewLink}`);
}

main().catch(err => {
  console.error('\n❌ Fatal:', err.stack ?? err.message);
  process.exit(1);
});
