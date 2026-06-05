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
const fs               = require('fs');
const os               = require('os');
const path             = require('path');

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

    // Column layout: กลุ่มงาน | month1 | month2 | ...
    ws.columns = [
      { header: 'กลุ่มงาน', key: 'dept', width: 34 },
      ...monthGroups.map(g => ({ header: g.monthLabel, key: g.monthLabel, width: 20 })),
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
      const rowData = { dept };
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
    const totalRowData = { dept: 'รวม' };
    monthGroups.forEach((g, i) => { totalRowData[g.monthLabel] = colTotals[i]; });
    const totalRow = ws.addRow(totalRowData);
    totalRow.height = 18;
    totalRow.eachCell({ includeEmpty: true }, cell => {
      cell.font      = { bold: true };
      cell.border    = FULL_MEDIUM;
      cell.alignment = { vertical: 'middle', horizontal: cell.col === 1 ? 'left' : 'center' };
    });


    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  // ── Sheets 2+: one per incomplete month ───────────────────────────
  for (const group of monthGroups) {
    const ws = wb.addWorksheet(group.monthLabel);

    ws.columns = [
      { header: 'ชื่อ-นามสกุล', key: 'fullname',   width: 36 },
      { header: 'กลุ่มงาน',    key: 'department', width: 32 },
    ];

    // Row 1: Header styling
    const headerRow = ws.getRow(1);
    headerRow.height = 22;
    headerRow.eachCell(cell => {
      cell.font      = { bold: true, size: 11 };
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border    = FULL_MEDIUM;
    });

    // Row 2: Footnote — merged A:B, black bg / white text, sticky with header
    const footnoteRow = ws.addRow(['', '']);
    ws.mergeCells(`A${footnoteRow.number}:B${footnoteRow.number}`);
    const footnoteCell = ws.getCell(`A${footnoteRow.number}`);
    footnoteCell.value     = `ตรวจสอบเมื่อ : ${runTime}`;
    footnoteCell.border    = FULL_THIN;
    footnoteCell.alignment = { horizontal: 'center', vertical: 'middle' };
    footnoteCell.font      = { italic: true, size: 10, color: { argb: 'FF000000' } };
    footnoteCell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB8CCE4' } };
    footnoteRow.height     = 18;

    // Rows 3+: Data
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

    // Freeze rows 1 (header) and 2 (footnote)
    ws.views = [{ state: 'frozen', ySplit: 2 }];
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
//  PNG — HTML template
// ═══════════════════════════════════════════════════════════════════
// One PNG per month — receives a single group
function buildHtml(group, runTime) {
  const rows = group.rows.length === 0
    ? `<tr><td class="no-data" colspan="3">ไม่มีข้อมูล</td></tr>`
    : group.rows.map((r, i) => `
        <tr class="${i % 2 === 1 ? 'alt' : ''}">
          <td class="num">${i + 1}</td>
          <td class="name">${r.fullname}</td>
          <td class="dept">${r.department}</td>
        </tr>`).join('');

  const emptyMsg = '';

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Prompt:ital,wght@0,400;0,600;0,700;1,400&display=swap');
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'Prompt', 'Noto Sans Thai', sans-serif;
    background: #f0f4f9;
    color: #1e2d3d;
    width: 960px;
    font-size: 28px;
    line-height: 1.5;
  }

  /* ── Header bar ── */
  .page-header {
    background: #1e3a5f;
    color: #fff;
    padding: 28px 40px 22px;
  }
  .page-header h1 { font-size: 38px; font-weight: 700; margin-bottom: 4px; }
  .page-header .month-tag {
    font-size: 26px;
    color: #a8c4e0;
    font-weight: 600;
  }

  /* ── Content area ── */
  .content { padding: 28px 40px 12px; }

  /* ── Month bar ── */
  .month-bar {
    background: #4472C4;
    color: #fff;
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 12px 20px;
    border-radius: 6px 6px 0 0;
  }
  .month-label { font-size: 28px; font-weight: 700; }
  .month-count {
    font-size: 23px;
    background: rgba(255,255,255,0.2);
    padding: 3px 14px;
    border-radius: 14px;
    font-weight: 600;
  }

  /* ── Table ── */
  table {
    width: 100%;
    border-collapse: collapse;
    background: #fff;
    border-radius: 0 0 6px 6px;
    overflow: hidden;
    box-shadow: 0 2px 6px rgba(0,0,0,0.08);
    table-layout: auto;
  }
  thead tr { background: #D9E1F2; }
  th {
    padding: 12px 18px;
    text-align: left;
    font-weight: 700;
    font-size: 25px;
    color: #2c3e50;
    border-bottom: 3px solid #b4c6e0;
    white-space: nowrap;
  }
  td {
    padding: 11px 18px;
    border-bottom: 1px solid #e8edf5;
    vertical-align: middle;
    font-size: 27px;
  }
  tr.alt { background: #EBF1FB; }
  tr:last-child td { border-bottom: none; }
  .num    { color: #aaa; font-size: 21px; text-align: center; width: 50px; white-space: nowrap; }
  .th-num { width: 50px; text-align: center; }
  .name   { }
  .dept   { color: #4472C4; font-weight: 600; white-space: nowrap; }
  .no-data {
    text-align: center; color: #bbb; font-style: italic;
    padding: 22px; font-size: 25px;
  }

  /* ── Footer ── */
  .footer {
    background: #1e3a5f;
    color: #a8c4e0;
    font-size: 21px;
    font-style: italic;
    padding: 14px 40px;
    text-align: right;
  }
</style>
</head>
<body>
  <div class="page-header">
    <h1>รายชื่อแพทย์ค้างส่ง P4P</h1>
    <div class="month-tag">${group.monthLabel}</div>
  </div>

  <div class="content">
    <div class="month-bar">
      <span class="month-label">แพทย์ที่ยังไม่ส่งไฟล์</span>
      <span class="month-count">${group.rows.length} คน</span>
    </div>
    <table>
      <thead>
        <tr>
          <th class="th-num">#</th>
          <th>ชื่อ-นามสกุล</th>
          <th>กลุ่มงาน</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </div>

  <div class="footer">ตรวจสอบเมื่อ : ${runTime}</div>
</body>
</html>`;
}

// ═══════════════════════════════════════════════════════════════════
//  PNG — render via Puppeteer
// ═══════════════════════════════════════════════════════════════════
async function renderPng(html) {
  const puppeteer = require('puppeteer');
  const browser   = await puppeteer.launch({
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
  });
  try {
    const page = await browser.newPage();
    await page.setViewport({ width: 960, height: 800, deviceScaleFactor: 2 }); // 1920px wide = Full HD
    await page.setContent(html, { waitUntil: 'networkidle0', timeout: 30000 });
    // screenshot() returns Uint8Array in Puppeteer v22+ — convert explicitly
    const raw = await page.screenshot({ fullPage: true, type: 'png' });
    return Buffer.isBuffer(raw) ? raw : Buffer.from(raw);
  } finally {
    await browser.close();
  }
}

// ═══════════════════════════════════════════════════════════════════
//  PNG — upload / overwrite in Drive
// ═══════════════════════════════════════════════════════════════════
async function uploadPng(drive, buffer, fileName) {
  const mime     = 'image/png';
  const folderId = CONFIG.reportFolderId;

  // Write to temp file — fs.createReadStream is more reliable than
  // Readable.from() for large binary uploads via googleapis
  const tmpPath = path.join(os.tmpdir(), `report_png_${Date.now()}.png`);
  fs.writeFileSync(tmpPath, buffer);

  try {
    const existing = await driveListAll(drive, {
      q: `'${folderId}' in parents and name='${fileName}' and trashed=false`,
      fields: 'nextPageToken, files(id, name)',
      pageSize: 10,
    });

    if (existing.length > 0) {
      const [first, ...dupes] = existing;
      for (const d of dupes) await withRetry(() => drive.files.delete({ fileId: d.id }));
      log(`  [Drive] Overwriting "${fileName}" (id: ${first.id})`);
      const res = await withRetry(() => drive.files.update({
        fileId: first.id,
        requestBody: { name: fileName },
        media: { mimeType: mime, body: fs.createReadStream(tmpPath) },
        fields: 'id, name, webViewLink',
      }));
      return res.data;
    }

    log(`  [Drive] Creating "${fileName}"`);
    const res = await withRetry(() => drive.files.create({
      requestBody: { name: fileName, parents: [folderId], mimeType: mime },
      media: { mimeType: mime, body: fs.createReadStream(tmpPath) },
      fields: 'id, name, webViewLink',
    }));
    return res.data;
  } finally {
    fs.unlinkSync(tmpPath);
  }
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
    const monthLabel    = `${THAI_MONTHS[month]} ${beYear}`;
    const abbMonthLabel = `${THAI_MONTH_ABBR[month]} ${String(beYear).slice(-2)}`;
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
      monthGroups.push({ monthLabel, abbMonthLabel, key, rows });
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
    monthGroups.push({ monthLabel, abbMonthLabel, key, rows });
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
  log(`\n✓ Excel saved : ${uploaded.webViewLink}`);

  // Remove all existing PNG files in report folder before uploading new ones
  log('\n[PNG] Clearing existing PNG files in report folder…');
  const existingPngs = await driveListAll(drive, {
    q: `'${CONFIG.reportFolderId}' in parents and mimeType='image/png' and trashed=false`,
    fields: 'nextPageToken, files(id, name)',
    pageSize: 100,
  });
  for (const f of existingPngs) {
    await withRetry(() => drive.files.delete({ fileId: f.id }));
    log(`  ✗ Deleted: ${f.name}`);
  }
  if (existingPngs.length === 0) log('  (none found)');

  // Build and upload one PNG per incomplete month
  log('\n[PNG] Rendering per-month images via Puppeteer…');
  for (const group of monthGroups) {
    const pngFileName = `ค้างส่ง ${group.key}.png`;
    log(`  ↳ ${pngFileName}`);
    const html      = buildHtml(group, runTime);
    const pngBuffer = await renderPng(html);
    const up        = await uploadPng(drive, pngBuffer, pngFileName);
    log(`    ✓ ${up.webViewLink}`);
  }
}

main().catch(err => {
  console.error('\n❌ Fatal:', err.stack ?? err.message);
  process.exit(1);
});
