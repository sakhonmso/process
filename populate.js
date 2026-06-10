#!/usr/bin/env node
/**
 * Populate year + month subfolders inside the Google Drive root folder.
 * Mirrors the GAS populateFolders() function.
 * Run once on 1 January each year via GitHub Actions.
 */

require('dotenv').config();

const { google } = require('googleapis');

const CONFIG = {
  google: {
    clientId:     process.env.GOOGLE_CLIENT_ID,
    clientSecret: process.env.GOOGLE_CLIENT_SECRET,
    refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
  },
  rootFolderId: process.env.GOOGLE_ROOT_FOLDER_ID,
};

const THAI_MONTHS = [
  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน',
  'พฤษภาคม', 'มิถุนายน',  'กรกฎาคม', 'สิงหาคม',
  'กันยายน', 'ตุลาคม',    'พฤศจิกายน', 'ธันวาคม',
];

function log(msg) { console.log(msg); }

async function main() {
  if (!CONFIG.rootFolderId) throw new Error('GOOGLE_ROOT_FOLDER_ID is not set');

  const auth = new google.auth.OAuth2(CONFIG.google.clientId, CONFIG.google.clientSecret);
  auth.setCredentials({ refresh_token: CONFIG.google.refreshToken });
  const drive = google.drive({ version: 'v3', auth });

  const beYear = (new Date().getFullYear() + 543).toString();
  log(`[populate] Target year (BE): ${beYear}`);

  // Check if year folder already exists
  const existing = await drive.files.list({
    q: `'${CONFIG.rootFolderId}' in parents and name = '${beYear}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
    fields: 'files(id, name)',
    pageSize: 1,
  });

  if (existing.data.files.length > 0) {
    log(`[populate] Folder '${beYear}' already exists (${existing.data.files[0].id}). Nothing to do.`);
    return;
  }

  // Create year folder
  const yearRes = await drive.files.create({
    requestBody: {
      name: beYear,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [CONFIG.rootFolderId],
    },
    fields: 'id',
  });
  const yearFolderId = yearRes.data.id;
  log(`[populate] Created year folder '${beYear}' → ${yearFolderId}`);

  // Create 12 month subfolders
  for (let i = 0; i < 12; i++) {
    const monthName = `${i + 1} - ${THAI_MONTHS[i]}`;
    const res = await drive.files.create({
      requestBody: {
        name: monthName,
        mimeType: 'application/vnd.google-apps.folder',
        parents: [yearFolderId],
      },
      fields: 'id',
    });
    log(`[populate]   Created '${monthName}' → ${res.data.id}`);
  }

  log('[populate] Done.');
}

main().catch(err => {
  console.error('[populate] ERROR:', err.message || err);
  process.exit(1);
});
