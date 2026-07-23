/**
 * MSETC GRID VIEWER — Backend (Google Apps Script + Google Sheets as DB)
 * ------------------------------------------------------------------
 * SETUP:
 * 1. Go to script.google.com > New project. Paste this whole file in as Code.gs.
 *    (No need to create a Sheet first — this script makes its own.)
 * 2. In the function dropdown (top toolbar), select "setupSheets" and click Run.
 *    - Grant permissions when asked (first run only).
 *    - This creates a new Google Sheet called "MSETC Grid Viewer Data" with
 *      3 tabs (Substations, Lines, ICTs), seeded with data from your 400kV
 *      overview image. Check View > Logs (or Executions) for the Sheet's URL.
 * 3. Deploy > New deployment > type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the deployed Web App URL and send it back — it gets hardcoded into index.html.
 *
 * DATA MODEL:
 *   Substations: id, name, voltage, region, x, y, notes
 *   Lines:       id, name, from, to, voltage, flowDir, status, notes
 *                 flowDir: "from->to" or "to->from" (direction of NORMAL/usual power flow)
 *                 status:  "in-service" or "tripped"
 *   ICTs:        id, substationId, hvSide, lvSide, name, notes
 *                 e.g. a 765/400 ICT at WRDHAPG: hvSide=765, lvSide=400
 */

// ---------------------------------------------------------------------
// SPREADSHEET CONNECTION
// This script creates its own Google Sheet automatically the first time
// it runs (no need to make a Sheet yourself or paste in an ID). The
// created Sheet's ID is remembered in Script Properties from then on.
// ---------------------------------------------------------------------
const SHEET_NAME = 'MSETC Grid Viewer Data';

function getSS() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID');

  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (err) {
      // stored ID no longer valid (sheet deleted) — fall through and recreate
    }
  }

  // No spreadsheet yet (or it's gone) — create one and remember it.
  const ss = SpreadsheetApp.create(SHEET_NAME);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  Logger.log('Created new spreadsheet: %s', ss.getUrl());
  return ss;
}

// Utility: run this any time from the editor to print the Sheet's URL
// to Logs (View > Logs) without touching any data.
function getSpreadsheetUrl() {
  Logger.log(getSS().getUrl());
}

const SHEETS = { SUB: 'Substations', LINE: 'Lines', ICT: 'ICTs' };

const SCHEMAS = {
  Substations: ['id', 'name', 'voltage', 'region', 'x', 'y', 'notes'],
  Lines: ['id', 'name', 'from', 'to', 'voltage', 'flowDir', 'status', 'notes'],
  ICTs: ['id', 'substationId', 'hvSide', 'lvSide', 'name', 'notes']
};

// ---------------------------------------------------------------------
// HTTP ENTRY POINTS
// ---------------------------------------------------------------------

function doGet(e) {
  try {
    const data = {
      substations: readSheet(SHEETS.SUB),
      lines: readSheet(SHEETS.LINE),
      icts: readSheet(SHEETS.ICT)
    };
    return jsonOut({ ok: true, data: data });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// NOTE: index.html posts with contentType 'text/plain' to avoid CORS
// preflight (Apps Script web apps can't handle OPTIONS). We parse the
// raw body as JSON manually here.
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const payload = body.payload || {};
    let result;

    switch (action) {
      case 'addSubstation':    result = addRow(SHEETS.SUB, payload); break;
      case 'updateSubstation': result = updateRow(SHEETS.SUB, payload); break;
      case 'deleteSubstation': result = deleteRow(SHEETS.SUB, payload.id); break;

      case 'addLine':          result = addRow(SHEETS.LINE, payload); break;
      case 'updateLine':       result = updateRow(SHEETS.LINE, payload); break;
      case 'deleteLine':       result = deleteRow(SHEETS.LINE, payload.id); break;
      case 'tripLine':         result = updateRow(SHEETS.LINE, { id: payload.id, status: 'tripped' }); break;
      case 'restoreLine':      result = updateRow(SHEETS.LINE, { id: payload.id, status: 'in-service' }); break;

      case 'addICT':           result = addRow(SHEETS.ICT, payload); break;
      case 'updateICT':        result = updateRow(SHEETS.ICT, payload); break;
      case 'deleteICT':        result = deleteRow(SHEETS.ICT, payload.id); break;

      default: throw new Error('Unknown action: ' + action);
    }

    const data = {
      substations: readSheet(SHEETS.SUB),
      lines: readSheet(SHEETS.LINE),
      icts: readSheet(SHEETS.ICT)
    };
    return jsonOut({ ok: true, result: result, data: data });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------
// SHEET CRUD HELPERS
// ---------------------------------------------------------------------

function readSheet(name) {
  const sh = getSS().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1)
    .filter(r => r[0] !== '')
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
}

function addRow(sheetName, payload) {
  const sh = getSS().getSheetByName(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (!payload.id) payload.id = Utilities.getUuid().slice(0, 8);
  const row = headers.map(h => payload[h] !== undefined ? payload[h] : '');
  sh.appendRow(row);
  return payload;
}

function updateRow(sheetName, payload) {
  const sh = getSS().getSheetByName(sheetName);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(payload.id)) {
      headers.forEach((h, c) => {
        if (payload[h] !== undefined) sh.getRange(i + 1, c + 1).setValue(payload[h]);
      });
      return payload;
    }
  }
  throw new Error('Row not found: ' + payload.id);
}

function deleteRow(sheetName, id) {
  const sh = getSS().getSheetByName(sheetName);
  const values = sh.getDataRange().getValues();
  const idCol = values[0].indexOf('id');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(id)) {
      sh.deleteRow(i + 1);
      return { id: id };
    }
  }
  throw new Error('Row not found: ' + id);
}

// ---------------------------------------------------------------------
// ONE-TIME SETUP + SEED DATA (400kV network, from SLDC overview)
// Run this once manually from the Apps Script editor.
// ---------------------------------------------------------------------

function setupSheets() {
  const ss = getSS();

  Object.keys(SCHEMAS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
    sh = ss.insertSheet(name);
    sh.appendRow(SCHEMAS[name]);
    sh.setFrozenRows(1);
  });

  // remove the blank default tab ("Sheet1") that comes with a new spreadsheet
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  const subSheet = ss.getSheetByName(SHEETS.SUB);
  const lineSheet = ss.getSheetByName(SHEETS.LINE);
  const ictSheet = ss.getSheetByName(SHEETS.ICT);

  // region-based grid coords (cleaner reorganized layout, not geo-exact)
  const substations = [
    // Konkan / coastal
    ['DABHOL', 'Dabhol', 400, 'Konkan', 60, 620],
    ['BOISRPG', 'Boisar PG', 400, 'Konkan', 60, 120],
    ['NGOTNE', 'NGOTNE', 400, 'Konkan', 60, 480],
    ['TARPUR', 'Tarapur', 400, 'Konkan', 160, 200],
    ['PADGE2', 'Padge-2', 400, 'Konkan', 160, 280],
    ['KUDUS', 'Kudus', 400, 'Konkan', 160, 360],
    ['KALWA2', 'Kalwa-2', 400, 'Konkan', 260, 400],
    ['KARGAR', 'Kargar', 400, 'Konkan', 260, 320],
    ['TALEGAONPG', 'Talegaon PG', 400, 'W. Maha', 260, 480],
    ['PUNEGIS', 'Pune GIS', 400, 'W. Maha', 360, 480],
    // Western Maharashtra
    ['KOYNAN', 'Koynan', 400, 'W. Maha', 260, 620],
    ['KLPHR3', 'Kolhapur-3', 400, 'W. Maha', 260, 700],
    ['JSW', 'JSW', 400, 'W. Maha', 360, 700],
    ['ALKUD', 'Alkud', 400, 'W. Maha', 360, 760],
    ['RAICHR', 'Raichur Interface', 400, 'W. Maha', 460, 760],
    ['KARAD2', 'Karad-2', 400, 'W. Maha', 460, 620],
    ['SOLPRPG', 'Solapur PG', 400, 'W. Maha', 560, 620],
    ['SOLPR3', 'Solapur-3', 400, 'W. Maha', 560, 700],
    ['CHAKAN', 'Chakan', 400, 'W. Maha', 360, 400],
    ['LNKND2', 'Lonikand-2', 400, 'W. Maha', 460, 400],
    ['DHULE', 'Dhule', 400, 'N. Maha', 360, 200],
    ['DHULEBD', 'Dhule BD', 400, 'N. Maha', 460, 200],
    ['KHNDWA', 'Khandwa', 400, 'N. Maha', 460, 120],
    ['IBNASIK', 'IB-Nasik', 400, 'N. Maha', 360, 120],
    // Vidarbha
    ['AKOLA', 'Akola', 400, 'Vidarbha', 760, 200],
    ['AKOLAAD', 'Akola-AD', 400, 'Vidarbha', 660, 200],
    ['BHSWL2', 'Bhusawal-2', 400, 'Vidarbha', 660, 120],
    ['BHSWL3', 'Bhusawal-3', 400, 'Vidarbha', 760, 120],
    ['KORADI3AD', 'Koradi-3 AD', 400, 'Vidarbha', 860, 280],
    ['KORADI2', 'Koradi-2', 400, 'Vidarbha', 960, 320],
    ['WRDHAPG', 'Wardha PG', 400, 'Vidarbha', 960, 400],
    ['WARORA', 'Warora', 400, 'Vidarbha', 960, 480],
    ['CHDPR', 'Chandrapur', 400, 'Vidarbha', 1060, 480],
    ['CHDPR2', 'Chandrapur-2', 400, 'Vidarbha', 1060, 560],
    ['CHDPRSW', 'Chandrapur SW', 400, 'Vidarbha', 1060, 640],
    ['NANDED', 'Nanded', 400, 'Marathwada', 760, 640],
    ['PARLY2', 'Parli-2', 400, 'Marathwada', 660, 640],
    ['PARLYPG', 'Parli PG', 400, 'Marathwada', 660, 700],
    ['NEWPARLYPG', 'New Parli PG', 400, 'Marathwada', 760, 700],
    ['KORADY', 'Koradi Y', 400, 'Vidarbha', 1060, 200],
    ['KHPKD2', 'Khaparkheda-2', 400, 'Vidarbha', 1060, 280],
    ['BAITUL', 'Baitul', 400, 'Vidarbha', 960, 120],
    ['MOUDA', 'Mouda', 400, 'Vidarbha', 1160, 320],
    ['RAIPUR', 'Raipur Interface', 400, 'Vidarbha', 1160, 400],
    ['BDRVT', 'Bhadravati', 400, 'Vidarbha', 1160, 480]
  ].map(([id, name, voltage, region, x, y]) => ({ id, name, voltage, region, x, y, notes: '' }));

  const lines = [
    ['BOISRPG-TARPUR', 'Boisar PG - Tarapur', 'BOISRPG', 'TARPUR', 400, 'from->to', 'in-service'],
    ['TARPUR-PADGE2', 'Tarapur - Padge-2', 'TARPUR', 'PADGE2', 400, 'from->to', 'in-service'],
    ['PADGE2-KUDUS', 'Padge-2 - Kudus', 'PADGE2', 'KUDUS', 400, 'from->to', 'in-service'],
    ['KUDUS-KARGAR', 'Kudus - Kargar', 'KUDUS', 'KARGAR', 400, 'from->to', 'in-service'],
    ['KARGAR-KALWA2', 'Kargar - Kalwa-2', 'KARGAR', 'KALWA2', 400, 'from->to', 'in-service'],
    ['KALWA2-NGOTNE', 'Kalwa-2 - NGOTNE', 'KALWA2', 'NGOTNE', 400, 'to->from', 'in-service'],
    ['NGOTNE-DABHOL', 'NGOTNE - Dabhol', 'NGOTNE', 'DABHOL', 400, 'from->to', 'in-service'],
    ['KARGAR-TALEGAONPG', 'Kargar - Talegaon PG', 'KARGAR', 'TALEGAONPG', 400, 'from->to', 'in-service'],
    ['TALEGAONPG-PUNEGIS', 'Talegaon PG - Pune GIS', 'TALEGAONPG', 'PUNEGIS', 400, 'from->to', 'in-service'],
    ['TALEGAONPG-CHAKAN', 'Talegaon PG - Chakan', 'TALEGAONPG', 'CHAKAN', 400, 'to->from', 'in-service'],
    ['CHAKAN-LNKND2', 'Chakan - Lonikand-2', 'CHAKAN', 'LNKND2', 400, 'from->to', 'in-service'],
    ['LNKND2-KARAD2', 'Lonikand-2 - Karad-2', 'LNKND2', 'KARAD2', 400, 'from->to', 'in-service'],
    ['KARAD2-KOYNAN', 'Karad-2 - Koynan', 'KARAD2', 'KOYNAN', 400, 'to->from', 'in-service'],
    ['KOYNAN-KLPHR3', 'Koynan - Kolhapur-3', 'KOYNAN', 'KLPHR3', 400, 'from->to', 'in-service'],
    ['KLPHR3-JSW', 'Kolhapur-3 - JSW', 'KLPHR3', 'JSW', 400, 'from->to', 'in-service'],
    ['JSW-ALKUD', 'JSW - Alkud', 'JSW', 'ALKUD', 400, 'from->to', 'in-service'],
    ['ALKUD-RAICHR', 'Alkud - Raichur Interface', 'ALKUD', 'RAICHR', 400, 'from->to', 'in-service'],
    ['KARAD2-SOLPRPG', 'Karad-2 - Solapur PG', 'KARAD2', 'SOLPRPG', 400, 'from->to', 'in-service'],
    ['SOLPRPG-SOLPR3', 'Solapur PG - Solapur-3', 'SOLPRPG', 'SOLPR3', 400, 'from->to', 'in-service'],
    ['SOLPRPG-RAICHR', 'Solapur PG - Raichur Interface', 'SOLPRPG', 'RAICHR', 400, 'to->from', 'in-service'],
    ['DHULE-IBNASIK', 'Dhule - IB-Nasik', 'DHULE', 'IBNASIK', 400, 'from->to', 'in-service'],
    ['DHULE-DHULEBD', 'Dhule - Dhule BD', 'DHULE', 'DHULEBD', 400, 'from->to', 'in-service'],
    ['DHULEBD-KHNDWA', 'Dhule BD - Khandwa', 'DHULEBD', 'KHNDWA', 400, 'to->from', 'in-service'],
    ['IBNASIK-KUDUS', 'IB-Nasik - Kudus', 'IBNASIK', 'KUDUS', 400, 'from->to', 'in-service'],
    ['DHULE-AKOLAAD', 'Dhule - Akola-AD', 'DHULE', 'AKOLAAD', 400, 'from->to', 'in-service'],
    ['AKOLAAD-AKOLA', 'Akola-AD - Akola', 'AKOLAAD', 'AKOLA', 400, 'from->to', 'in-service'],
    ['AKOLAAD-BHSWL2', 'Akola-AD - Bhusawal-2', 'AKOLAAD', 'BHSWL2', 400, 'to->from', 'in-service'],
    ['BHSWL2-BHSWL3', 'Bhusawal-2 - Bhusawal-3', 'BHSWL2', 'BHSWL3', 400, 'from->to', 'in-service'],
    ['BHSWL3-AKOLA', 'Bhusawal-3 - Akola', 'BHSWL3', 'AKOLA', 400, 'from->to', 'in-service'],
    ['AKOLA-KORADI3AD', 'Akola - Koradi-3 AD', 'AKOLA', 'KORADI3AD', 400, 'from->to', 'in-service'],
    ['KORADI3AD-KORADI2', 'Koradi-3 AD - Koradi-2', 'KORADI3AD', 'KORADI2', 400, 'from->to', 'in-service'],
    ['KORADI2-WRDHAPG', 'Koradi-2 - Wardha PG', 'KORADI2', 'WRDHAPG', 400, 'from->to', 'in-service'],
    ['WRDHAPG-WARORA', 'Wardha PG - Warora', 'WRDHAPG', 'WARORA', 400, 'from->to', 'in-service'],
    ['WARORA-CHDPR', 'Warora - Chandrapur', 'WARORA', 'CHDPR', 400, 'from->to', 'in-service'],
    ['CHDPR-CHDPR2', 'Chandrapur - Chandrapur-2', 'CHDPR', 'CHDPR2', 400, 'from->to', 'in-service'],
    ['CHDPR2-CHDPRSW', 'Chandrapur-2 - Chandrapur SW', 'CHDPR2', 'CHDPRSW', 400, 'from->to', 'in-service'],
    ['CHDPRSW-NANDED', 'Chandrapur SW - Nanded', 'CHDPRSW', 'NANDED', 400, 'to->from', 'in-service'],
    ['NANDED-PARLY2', 'Nanded - Parli-2', 'NANDED', 'PARLY2', 400, 'to->from', 'in-service'],
    ['PARLY2-PARLYPG', 'Parli-2 - Parli PG', 'PARLY2', 'PARLYPG', 400, 'from->to', 'in-service'],
    ['PARLYPG-NEWPARLYPG', 'Parli PG - New Parli PG', 'PARLYPG', 'NEWPARLYPG', 400, 'from->to', 'in-service'],
    ['NEWPARLYPG-NANDED', 'New Parli PG - Nanded', 'NEWPARLYPG', 'NANDED', 400, 'from->to', 'in-service'],
    ['KORADI2-KORADY', 'Koradi-2 - Koradi Y', 'KORADI2', 'KORADY', 400, 'from->to', 'in-service'],
    ['KORADY-KHPKD2', 'Koradi Y - Khaparkheda-2', 'KORADY', 'KHPKD2', 400, 'from->to', 'in-service'],
    ['BHSWL3-BAITUL', 'Bhusawal-3 - Baitul', 'BHSWL3', 'BAITUL', 400, 'from->to', 'in-service'],
    ['BAITUL-WRDHAPG', 'Baitul - Wardha PG', 'BAITUL', 'WRDHAPG', 400, 'to->from', 'in-service'],
    ['MOUDA-KORADI3AD', 'Mouda - Koradi-3 AD', 'MOUDA', 'KORADI3AD', 400, 'from->to', 'in-service'],
    ['MOUDA-RAIPUR', 'Mouda - Raipur Interface', 'MOUDA', 'RAIPUR', 400, 'to->from', 'in-service'],
    ['RAIPUR-CHDPR2', 'Raipur Interface - Chandrapur-2', 'RAIPUR', 'CHDPR2', 400, 'from->to', 'in-service'],
    ['CHDPR2-BDRVT', 'Chandrapur-2 - Bhadravati', 'CHDPR2', 'BDRVT', 400, 'from->to', 'in-service']
  ].map(([id, name, from, to, voltage, flowDir, status]) => ({ id, name, from, to, voltage, flowDir, status, notes: '' }));

  const icts = [
    ['ICT-WRDHAPG-1', 'WRDHAPG', 765, 400, 'Wardha ICT-1', ''],
    ['ICT-KORADI-1', 'KORADI3AD', 765, 400, 'Koradi ICT-1', ''],
    ['ICT-SOLPRPG-1', 'SOLPRPG', 765, 400, 'Solapur ICT-1', ''],
    ['ICT-TALEGAONPG-1', 'TALEGAONPG', 765, 400, 'Talegaon ICT-1', '']
  ].map(([id, substationId, hvSide, lvSide, name, notes]) => ({ id, substationId, hvSide, lvSide, name, notes }));

  writeRows(subSheet, SCHEMAS.Substations, substations);
  writeRows(lineSheet, SCHEMAS.Lines, lines);
  writeRows(ictSheet, SCHEMAS.ICTs, icts);

  SpreadsheetApp.flush();
  Logger.log('Setup complete: %s substations, %s lines, %s ICTs',
    substations.length, lines.length, icts.length);
  Logger.log('Your data Sheet: %s', ss.getUrl());
}

function writeRows(sheet, headers, rows) {
  const data = rows.map(r => headers.map(h => r[h] !== undefined ? r[h] : ''));
  if (data.length) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
}
