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

const SHEETS = { SUB: 'Substations', LINE: 'Lines', ICT: 'ICTs', TOWER: 'Towers', SET: 'Settings' };

const SCHEMAS = {
  Substations: ['id', 'fullName', 'voltage', 'owner', 'region', 'x', 'y', 'busWidth', 'lat', 'lon', 'color', 'sameAs', 'hvdc', 'source', 'verify', 'notes'],
  Lines: ['id', 'name', 'from', 'to', 'voltage', 'circuit', 'flowDir', 'status', 'color', 'weight', 'verify', 'notes'],
  ICTs: ['id', 'substationId', 'hvSide', 'lvSide', 'name', 'rating', 'verify', 'notes'],
  Towers: ['id', 'lineId', 'seq', 'lat', 'lon', 'name', 'notes'],
  Settings: ['key', 'value']
};

// ---------------------------------------------------------------------
// HTTP ENTRY POINTS
// ---------------------------------------------------------------------

function doGet(e) {
  try {
    return jsonOut({ ok: true, data: readAll() });
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

      case 'addTower':         result = addRow(SHEETS.TOWER, payload); break;
      case 'updateTower':      result = updateRow(SHEETS.TOWER, payload); break;
      case 'deleteTower':      result = deleteRow(SHEETS.TOWER, payload.id); break;
      case 'replaceTowers':    result = replaceTowers(payload); break;
      case 'setSetting':       result = setSetting(payload.key, payload.value); break;
      case 'addICT':           result = addRow(SHEETS.ICT, payload); break;
      case 'updateICT':        result = updateRow(SHEETS.ICT, payload); break;
      case 'deleteICT':        result = deleteRow(SHEETS.ICT, payload.id); break;

      default: throw new Error('Unknown action: ' + action);
    }

    return jsonOut({ ok: true, result: result, data: readAll() });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function readAll() {
  return {
    substations: readSheet(SHEETS.SUB),
    lines: readSheet(SHEETS.LINE),
    icts: readSheet(SHEETS.ICT),
    towers: readSheet(SHEETS.TOWER),
    settings: readSheet(SHEETS.SET)
  };
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

// Replace the full tower chain of one line in a single call
function replaceTowers(payload) {
  const sh = getSS().getSheetByName(SHEETS.TOWER);
  const values = sh.getDataRange().getValues();
  const idCol = values[0].indexOf('lineId');
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][idCol]) === String(payload.lineId)) sh.deleteRow(i + 1);
  }
  (payload.towers || []).forEach((t, n) => {
    addRow(SHEETS.TOWER, {
      id: payload.lineId + '-T' + (n + 1),
      lineId: payload.lineId,
      seq: n + 1,
      lat: t.lat, lon: t.lon,
      name: t.name || ('Tower ' + (n + 1)),
      notes: ''
    });
  });
  return { lineId: payload.lineId, count: (payload.towers || []).length };
}

function setSetting(key, value) {
  const sh = getSS().getSheetByName(SHEETS.SET);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(key)) {
      sh.getRange(i + 1, 2).setValue(value);
      return { key: key, value: value };
    }
  }
  sh.appendRow([key, value]);
  return { key: key, value: value };
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

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  const subSheet = ss.getSheetByName(SHEETS.SUB);
  const lineSheet = ss.getSheetByName(SHEETS.LINE);
  const ictSheet = ss.getSheetByName(SHEETS.ICT);
  // ------------------------------------------------------------------
  // SUBSTATIONS
  // sameAs : if set, this row is an OFF-PAGE CONNECTOR — the same physical
  //          station drawn again elsewhere on the sheet because a line
  //          couldn't be routed. Contingency logic merges it with the
  //          canonical row, so it counts as ONE node.
  // hvdc   : 'yes' if an HVDC converter/terminal sits at this station
  // source : generation at the bus — thermal | hydro | gas | solar | wind | ''
  // verify : 'yes' = still unconfirmed, please correct in the Sheet
  // ------------------------------------------------------------------
  const substations = [
    // id, fullName, voltage, owner, region, x, y, busWidth, lat, lon, sameAs, hvdc, source, verify, notes
    // lat/lon are APPROXIMATE town-level estimates — correct them on the Map page.
    ['PADGHAPG',   '765/400 kV Padgha PG',              765, 'PG',        'Konkan',     140, 1040, 300, 19.42, 73.13, '', 'yes', '',        'no',  'Western terminal of the Chandrapur–Padghe HVDC bipole'],
    ['BOISRPG',    '400 kV Boisar PG',                  400, 'PG',        'Konkan',      140, 140, 250, 19.8, 72.76, '', 'no',  '',        'no',  ''],
    ['TARAPUR',    '400 kV Tarapur',                    400, 'MSETCL',    'Konkan',     140, 440, 250, 19.83, 72.66, '', 'no',  'nuclear', 'no',  'TAPS nuclear generation'],
    ['PADGE2',     '400 kV Padghe-2',                   400, 'MSETCL',    'Konkan',     140, 740, 250, 19.41, 73.12, '', 'no',  '',        'no',  ''],
    ['KUDUS',      '400 kV Kudus',                      400, 'MSETCL',    'Konkan',     780, 740, 250, 19.55, 73.15, '', 'no',  '',        'no',  ''],
    ['KARGAR',     '400 kV Kargar',                     400, 'MSETCL',    'Konkan',     780, 1040, 250, 19.03, 73.06, '', 'no',  '',        'no',  ''],
    ['VIKHROLI',   '400 kV Vikhroli',                   400, 'MSETCL',    'Konkan',     780, 1340, 250, 19.11, 72.93, '', 'no',  '',        'yes', 'label partly unreadable'],
    ['KALWA2',     '400 kV Kalwa-2',                    400, 'MSETCL',    'Konkan',      140, 1340, 250, 19.2, 73.0, '', 'no',  '',        'no',  'Kalwa frequency reference bus'],
    ['NAGOTHANE',  '400 kV Nagothane',                  400, 'MSETCL',    'Konkan',      140, 1640, 250, 18.53, 73.13, '', 'no',  '',        'no',  ''],
    ['DABHOL',     '400 kV Dabhol',                     400, 'IPP',       'Konkan',      140, 1940, 250, 17.58, 73.18, '', 'no',  'gas',     'no',  'RGPPL Dabhol CCGT'],
    ['NVMUMPG',    '400 kV Navi Mumbai PG',             400, 'PG',        'Konkan',     780, 1640, 250, 19.03, 73.02, '', 'no',  '',        'no',  ''],

    ['DHULE',      '400 kV Dhule',                      400, 'MSETCL',    'N. Maha',    1420, 1040, 250, 20.9, 74.77, '', 'no',  '',        'no',  ''],
    ['DHULEBD',    '400 kV Dhule BD',                   400, 'MSETCL',    'N. Maha',    1420, 740, 250, 20.91, 74.8, '', 'no',  '',        'no',  ''],
    ['IBNASIK',    '400 kV Indiabulls Nashik',                   400, 'MSETCL',    'N. Maha',    2060, 440, 250, 19.85, 74.0, '', 'no',  'thermal', 'yes', 'Indiabulls (RattanIndia) Nashik, Sinnar'],
    ['KHNDWA',     '400 kV Khandwa',                    400, 'INTERSTATE','N. Maha',    2060, 140, 250, 21.82, 76.35, '', 'no',  '',        'no',  'MP interface'],
    ['SDSRV',      '400 kV Sardar Sarovar',             400, 'INTERSTATE','N. Maha',    1420, 440, 250, 21.83, 73.75, '', 'no',  'hydro',   'no',  'Sardar Sarovar hydro, Gujarat'],
    ['VDODRAPG',   '400 kV Vadodara PG',                400, 'PG',        'N. Maha',    1420, 140, 250, 22.31, 73.18, '', 'no',  '',        'no',  'Gujarat interface'],
    ['EKTUNI',     '400 kV Ektuni',                     400, 'MSETCL',    'N. Maha',    2060, 1040, 250, 19.75, 75.5, '', 'no',  '',        'no',  ''],
    ['BABHLESHWAR','400 kV Babhaleshwar',               400, 'MSETCL',    'N. Maha',    2060, 740, 330, 19.55, 74.6, '', 'no',  '',        'yes', 'large multi-bay bus — confirm connections'],

    ['TALEGAONPG', '765/400 kV Talegaon PG',            765, 'PG',        'W. Maha',    1420, 1640, 290, 18.73, 73.68, '', 'no',  '',        'no',  ''],
    ['CHAKAN',     '400 kV Chakan',                     400, 'MSETCL',    'W. Maha',    1420, 1940, 250, 18.76, 73.86, '', 'no',  '',        'no',  ''],
    ['PUNEGIS',    '400 kV Pune GIS',                   400, 'MSETCL',    'W. Maha',    1420, 2240, 250, 18.52, 73.86, '', 'no',  '',        'no',  ''],
    ['LONIKAND',   '400 kV Lonikand',                   400, 'MSETCL',    'W. Maha',    1420, 2540, 250, 18.6, 73.98, '', 'no',  '',        'no',  ''],
    ['LNKND2',     '400 kV Lonikand-2',                 400, 'MSETCL',    'W. Maha',    2060, 2540, 250, 18.61, 73.99, '', 'no',  '',        'no',  ''],
    ['JEJURY',     '400 kV Jejuri',                     400, 'MSETCL',    'W. Maha',    1420, 2840, 250, 18.28, 74.16, '', 'no',  '',        'no',  ''],
    ['KOYNA4',     '400 kV Koyna-4',                    400, 'MSETCL',    'W. Maha',    1420, 3140, 250, 17.4, 73.75, '', 'no',  'hydro',   'no',  'Koyna HEP Stage IV'],
    ['KOYNAAN',    '400 kV Koyna AN',                   400, 'MSETCL',    'W. Maha',    1420, 3440, 250, 17.39, 73.76, '', 'no',  'hydro',   'no',  'separate from Koyna-4'],
    ['KARAD2',     '400 kV Karad-2',                    400, 'MSETCL',    'W. Maha',    2060, 3140, 250, 17.29, 74.18, '', 'no',  '',        'no',  ''],
    ['KARJAT',     '400 kV Karjat',                     400, 'MSETCL',    'W. Maha',    2060, 2840, 250, 18.92, 75.01, '', 'no',  '',        'no',  ''],
    ['KLPHR3',     '400 kV Kolhapur-3',                 400, 'MSETCL',    'W. Maha',    1420, 3740, 250, 16.7, 74.24, '', 'no',  '',        'no',  ''],
    ['JSW',        '400 kV JSW',                        400, 'IPP',       'W. Maha',    1420, 4040, 250, 17.28, 73.23, '', 'no',  'thermal', 'no',  'JSW Energy'],
    ['ALKUD',      '400 kV Alkud',                      400, 'MSETCL',    'W. Maha',    2060, 3740, 250, 17.15, 74.6, '', 'no',  '',        'no',  ''],
    ['KUPWADPG',   '400 kV Kupwad PG',                  400, 'PG',        'W. Maha',    2060, 4040, 250, 16.87, 74.62, '', 'no',  '',        'no',  ''],
    ['MAPUSA',     '400 kV Mapusa',                     400, 'INTERSTATE','W. Maha',    2700, 4340, 250, 15.59, 73.81, '', 'no',  '',        'no',  'Goa interface'],
    ['NANDRA',     '400 kV Nandra',                     400, 'MSETCL',    'W. Maha',    2060, 4340, 250, 16.95, 74.5, '', 'no',  '',        'yes', ''],
    ['RAICHR',     '400 kV Raichur',                    400, 'INTERSTATE','W. Maha',    2700, 4040, 250, 16.2, 77.36, '', 'no',  'thermal', 'no',  'Karnataka interface'],
    ['SOLPRPG',    '765/400 kV Solapur PG',             765, 'PG',        'W. Maha',    2700, 3440, 290, 17.66, 75.9, '', 'no',  '',        'no',  ''],
    ['SOLPR3',     '400 kV Solapur-3',                  400, 'MSETCL',    'W. Maha',    2700, 3140, 250, 17.67, 75.91, '', 'no',  'thermal', 'no',  'NTPC Solapur nearby'],
    ['KALAMBPG',   '400 kV Kalamb PG',                  400, 'PG',        'W. Maha',    2700, 3740, 250, 18.1, 76.0, '', 'no',  '',        'yes', ''],

    ['AURNGBDPG',  '765/400 kV Aurangabad PG',          765, 'PG',        'Marathwada', 3340, 2840, 290, 19.88, 75.34, '', 'no',  '',        'no',  'appears at several places on the SLDC sheet as an off-page connector'],
    ['PARLY2',     '400 kV Parli-2',                    400, 'MSETCL',    'Marathwada', 3340, 3440, 250, 18.85, 76.53, '', 'no',  'thermal', 'no',  'Parli TPS'],
    ['PARLYPG',    '765/400 kV Parli PG',               765, 'PG',        'Marathwada', 3340, 3740, 290, 18.86, 76.54, '', 'no',  '',        'no',  ''],
    ['NEWPARLYPG', '400 kV New Parli PG',               400, 'PG',        'Marathwada', 3340, 4040, 250, 18.87, 76.55, '', 'no',  '',        'no',  ''],
    ['NANDED',     '400 kV Nanded',                     400, 'MSETCL',    'Marathwada', 3980, 3440, 250, 19.15, 77.32, '', 'no',  '',        'no',  ''],
    ['DHRWL',      '400 kV Dharwal',                    400, 'MSETCL',    'Marathwada', 3980, 3140, 250, 19.2, 76.8, '', 'no',  '',        'yes', 'spelling unconfirmed'],

    ['BHSWL2',     '400 kV Bhusawal-2',                 400, 'MSETCL',    'Vidarbha',   3340, 140, 270, 21.04, 75.79, '', 'no',  'thermal', 'no',  'Bhusawal TPS'],
    ['BHSWL3',     '400 kV Bhusawal-3',                 400, 'MSETCL',    'Vidarbha',   3340, 440, 250, 21.05, 75.8, '', 'no',  'thermal', 'no',  ''],
    ['AKOLAAD',    '400 kV Akola-AD',                   400, 'MSETCL',    'Vidarbha',   3340, 740, 250, 20.7, 77.0, '', 'no',  '',        'no',  ''],
    ['AKOLA',      '400 kV Akola',                      400, 'MSETCL',    'Vidarbha',   3980, 740, 250, 20.71, 77.01, '', 'no',  '',        'no',  ''],
    ['RTNIND',     '400 kV RattanIndia Amravati',       400, 'IPP',       'Vidarbha',   3340, 1040, 250, 20.93, 77.76, '', 'no',  'thermal', 'yes', 'AMT read as Amravati — please confirm'],
    ['ARGBD3',     '400 kV Aurangabad-3',               400, 'MSETCL',    'Vidarbha',   3340, 3140, 250, 19.89, 75.35, '', 'no',  '',        'no',  ''],
    ['BAITUL',     '400 kV Baitul',                     400, 'INTERSTATE','Vidarbha',   3980, 140, 250, 21.9, 77.9, '', 'no',  '',        'no',  'MP interface'],
    ['KORADY',     '400 kV Koradi Y',                   400, 'MSETCL',    'Vidarbha',   4620, 140, 250, 21.24, 79.09, '', 'no',  'thermal', 'yes', 'relation to Koradi-2/-3 to confirm'],
    ['KHPKD2',     '400 kV Khaparkheda-2',              400, 'MSETCL',    'Vidarbha',   4620, 440, 250, 21.28, 79.06, '', 'no',  'thermal', 'no',  'Khaparkheda TPS'],
    ['SATPURA',    '400 kV Satpura',                    400, 'INTERSTATE','Vidarbha',   5260, 140, 250, 22.1, 78.17, '', 'no',  'thermal', 'no',  'MP interface'],
    ['BHILY',      '400 kV Bhilai',                     400, 'INTERSTATE','Vidarbha',  5260, 440, 250, 21.19, 81.38, '', 'no',  '',        'yes', 'read as BHILY — unconfirmed'],
    ['MOUDA',      '400 kV Mouda',                      400, 'IPP',       'Vidarbha',   5260, 740, 250, 21.29, 79.35, '', 'no',  'thermal', 'no',  'NTPC Mouda'],
    ['KORADI3AD',  '400 kV Koradi-3 AD',                400, 'MSETCL',    'Vidarbha',   3980, 1040, 280, 21.245, 79.095, '', 'no',  'thermal', 'no',  'Koradi TPS'],
    ['KORADI2',    '400 kV Koradi-2',                   400, 'MSETCL',    'Vidarbha',   4620, 1040, 250, 21.25, 79.1, '', 'no',  'thermal', 'no',  ''],
    ['APML',       '400 kV APML Tiroda',                400, 'IPP',       'Vidarbha',   3980, 1340, 250, 21.41, 79.94, '', 'no',  'thermal', 'yes', 'Adani Power Maharashtra Ltd, Tiroda'],
    ['WRDHAPG',    '765/400 kV Wardha PG',              765, 'PG',        'Vidarbha',   4620, 1340, 300, 20.75, 78.6, '', 'no',  '',        'no',  ''],
    ['WARORAPG',   '400 kV Warora PG',                  400, 'PG',        'Vidarbha',   4620, 1640, 250, 20.24, 79.01, '', 'no',  '',        'no',  ''],
    ['WARORA',     '400 kV Warora',                     400, 'MSETCL',    'Vidarbha',   4620, 1940, 250, 20.23, 79.0, '', 'no',  '',        'no',  ''],
    ['NZBOL',      '400 kV Nizamabad',                  400, 'INTERSTATE','Vidarbha',   5260, 1640, 250, 18.67, 78.09, '', 'no',  '',        'yes', 'read as NZBOL — unconfirmed'],
    ['DURG',       '400 kV Durg',                       400, 'INTERSTATE','Vidarbha',   5260, 1940, 250, 21.19, 81.28, '', 'no',  '',        'no',  'Chhattisgarh interface'],
    ['SEONI',      '400 kV Seoni',                      400, 'INTERSTATE','Vidarbha',   5260, 1340, 250, 22.09, 79.55, '', 'no',  '',        'no',  'MP interface'],
    ['RAIPUR',     '400 kV Raipur',                     400, 'INTERSTATE','Vidarbha',  5900, 1040, 250, 21.25, 81.63, '', 'no',  '',        'no',  'Chhattisgarh interface'],
    ['CHDPR',      '400 kV Chandrapur',                 400, 'MSETCL',    'Vidarbha',  5260, 2240, 280, 19.97, 79.3, '', 'yes', 'thermal', 'no',  'Chandrapur TPS + eastern terminal of the Chandrapur–Padghe HVDC bipole'],
    ['CHDPR2',     '400 kV Chandrapur-2',               400, 'MSETCL',    'Vidarbha',  5260, 2540, 250, 19.975, 79.305, '', 'no',  'thermal', 'no',  ''],
    ['CHDPRSW',    '400 kV Chandrapur SW',              400, 'MSETCL',    'Vidarbha',   4620, 2240, 250, 19.98, 79.31, '', 'no',  '',        'no',  'switchyard'],
    ['BDRVT',      '400 kV Bhadravati',                 400, 'MSETCL',    'Vidarbha',  5900, 2240, 250, 20.1, 79.12, '', 'no',  '',        'no',  ''],

    // ---- OFF-PAGE CONNECTORS (same physical station, drawn again) ----
    ['AURNGBDPG_B','765/400 kV Aurangabad PG',          765, 'PG',        'Vidarbha',  5900, 1340, 250, 19.88, 75.34, 'AURNGBDPG', 'no', '', 'no', 'off-page connector for Aurangabad PG'],
    ['AURNGBDPG_C','765/400 kV Aurangabad PG',          765, 'PG',        'N. Maha',    2060, 1340, 250, 19.88, 75.34, 'AURNGBDPG', 'no', '', 'no', 'off-page connector for Aurangabad PG'],
    ['PADGHAPG_B', '765/400 kV Padgha PG',              765, 'PG',        'W. Maha',    780, 1940, 250, 19.42, 73.13, 'PADGHAPG',  'no', '', 'no', 'off-page connector for Padgha PG']
  ].map(([id, fullName, voltage, owner, region, x, y, busWidth, lat, lon, sameAs, hvdc, source, verify, notes]) =>
    ({ id, fullName, voltage, owner, region, x, y, busWidth, lat, lon, color: '', sameAs, hvdc, source, verify, notes }));

  // ------------------------------------------------------------------
  // LINES  — one row per CIRCUIT.
  // Where two circuits run between the same pair, they appear as
  // separate rows with circuit = "Ckt 1" / "Ckt 2".
  // ------------------------------------------------------------------
  const rawLines = [
    // from, to, circuit, flowDir, verify
    ['BOISRPG','TARAPUR','Ckt 1','from->to','no'],
    ['BOISRPG','TARAPUR','Ckt 2','from->to','yes'],
    ['TARAPUR','PADGE2','Ckt 1','from->to','no'],
    ['TARAPUR','PADGE2','Ckt 2','from->to','yes'],
    ['PADGE2','PADGHAPG','Ckt 1','from->to','no'],
    ['PADGHAPG','KUDUS','Ckt 1','from->to','no'],
    ['PADGHAPG','KUDUS','Ckt 2','from->to','yes'],
    ['KUDUS','KARGAR','Ckt 1','from->to','no'],
    ['KUDUS','KARGAR','Ckt 2','from->to','yes'],
    ['KARGAR','VIKHROLI','Ckt 1','from->to','yes'],
    ['VIKHROLI','KALWA2','Ckt 1','from->to','yes'],
    ['KARGAR','KALWA2','Ckt 1','from->to','no'],
    ['KALWA2','NAGOTHANE','Ckt 1','to->from','no'],
    ['NAGOTHANE','DABHOL','Ckt 1','to->from','no'],
    ['NAGOTHANE','DABHOL','Ckt 2','to->from','yes'],
    ['KARGAR','NVMUMPG','Ckt 1','from->to','yes'],
    ['PADGHAPG','TALEGAONPG','Ckt 1','from->to','no'],
    ['TALEGAONPG','PUNEGIS','Ckt 1','from->to','no'],
    ['TALEGAONPG','PUNEGIS','Ckt 2','from->to','yes'],
    ['TALEGAONPG','CHAKAN','Ckt 1','to->from','no'],
    ['TALEGAONPG','LONIKAND','Ckt 1','from->to','no'],
    ['LONIKAND','LNKND2','Ckt 1','from->to','no'],
    ['CHAKAN','LNKND2','Ckt 1','from->to','no'],
    ['PUNEGIS','JEJURY','Ckt 1','from->to','no'],
    ['JEJURY','KARAD2','Ckt 1','from->to','no'],
    ['LNKND2','KARJAT','Ckt 1','from->to','yes'],
    ['KARJAT','PARLY2','Ckt 1','from->to','yes'],
    ['KARAD2','KOYNA4','Ckt 1','to->from','no'],
    ['KOYNA4','KOYNAAN','Ckt 1','from->to','yes'],
    ['KOYNAAN','KLPHR3','Ckt 1','from->to','no'],
    ['KLPHR3','JSW','Ckt 1','to->from','no'],
    ['KLPHR3','ALKUD','Ckt 1','from->to','no'],
    ['ALKUD','KUPWADPG','Ckt 1','from->to','yes'],
    ['KUPWADPG','MAPUSA','Ckt 1','from->to','yes'],
    ['KUPWADPG','NANDRA','Ckt 1','from->to','yes'],
    ['ALKUD','RAICHR','Ckt 1','from->to','no'],
    ['KARAD2','SOLPRPG','Ckt 1','from->to','no'],
    ['SOLPRPG','SOLPR3','Ckt 1','from->to','no'],
    ['SOLPRPG','RAICHR','Ckt 1','to->from','no'],
    ['SOLPRPG','KALAMBPG','Ckt 1','from->to','yes'],
    ['SOLPR3','PARLY2','Ckt 1','from->to','yes'],
    ['DHULE','IBNASIK','Ckt 1','from->to','no'],
    ['DHULE','DHULEBD','Ckt 1','from->to','no'],
    ['DHULEBD','KHNDWA','Ckt 1','to->from','no'],
    ['DHULE','SDSRV','Ckt 1','from->to','yes'],
    ['SDSRV','VDODRAPG','Ckt 1','to->from','yes'],
    ['IBNASIK','BABHLESHWAR','Ckt 1','from->to','yes'],
    ['BABHLESHWAR','PADGHAPG','Ckt 1','to->from','yes'],
    ['BABHLESHWAR','EKTUNI','Ckt 1','from->to','yes'],
    ['BABHLESHWAR','AURNGBDPG','Ckt 1','from->to','yes'],
    ['AURNGBDPG','TALEGAONPG','Ckt 1','from->to','yes'],
    ['EKTUNI','BHSWL3','Ckt 1','from->to','no'],
    ['BHSWL2','BHSWL3','Ckt 1','from->to','no'],
    ['BHSWL2','AKOLAAD','Ckt 1','to->from','no'],
    ['BHSWL3','ARGBD3','Ckt 1','from->to','yes'],
    ['AKOLAAD','AKOLA','Ckt 1','from->to','no'],
    ['AKOLAAD','RTNIND','Ckt 1','from->to','yes'],
    ['AKOLA','BAITUL','Ckt 1','from->to','no'],
    ['BAITUL','KORADY','Ckt 1','from->to','yes'],
    ['KORADY','SATPURA','Ckt 1','to->from','no'],
    ['KORADY','BHILY','Ckt 1','to->from','yes'],
    ['KORADY','KHPKD2','Ckt 1','from->to','no'],
    ['KHPKD2','MOUDA','Ckt 1','to->from','no'],
    ['AKOLA','KORADI3AD','Ckt 1','from->to','no'],
    ['AKOLA','KORADI3AD','Ckt 2','from->to','yes'],
    ['MOUDA','KORADI3AD','Ckt 1','from->to','no'],
    ['KORADI3AD','KORADI2','Ckt 1','from->to','no'],
    ['KORADI3AD','KORADI2','Ckt 2','from->to','yes'],
    ['KORADI3AD','APML','Ckt 1','to->from','yes'],
    ['APML','TALEGAONPG','Ckt 1','from->to','yes'],
    ['KORADI2','WRDHAPG','Ckt 1','from->to','no'],
    ['WRDHAPG','AURNGBDPG_B','Ckt 1','from->to','yes'],
    ['WRDHAPG','WARORAPG','Ckt 1','from->to','no'],
    ['WARORAPG','WARORA','Ckt 1','from->to','no'],
    ['WARORAPG','NZBOL','Ckt 1','from->to','yes'],
    ['WARORAPG','DURG','Ckt 1','to->from','no'],
    ['WARORA','SEONI','Ckt 1','from->to','yes'],
    ['WARORA','CHDPR','Ckt 1','from->to','no'],
    ['CHDPR','CHDPR2','Ckt 1','from->to','no'],
    ['CHDPR','BDRVT','Ckt 1','from->to','no'],
    ['BDRVT','RAIPUR','Ckt 1','to->from','no'],
    ['CHDPR2','CHDPRSW','Ckt 1','from->to','no'],
    ['CHDPRSW','NANDED','Ckt 1','to->from','no'],
    ['CHDPRSW','DHRWL','Ckt 1','from->to','yes'],
    ['DHRWL','PARLY2','Ckt 1','from->to','yes'],
    ['NANDED','PARLY2','Ckt 1','to->from','no'],
    ['PARLY2','PARLYPG','Ckt 1','from->to','no'],
    ['PARLYPG','NEWPARLYPG','Ckt 1','from->to','no'],
    ['NEWPARLYPG','NANDED','Ckt 1','from->to','no'],
    ['PARLYPG','KALAMBPG','Ckt 1','to->from','yes'],
    ['CHDPR','PADGHAPG','HVDC Pole 1','from->to','no'],
    ['CHDPR','PADGHAPG','HVDC Pole 2','from->to','no']
  ];

  const subName = {};
  substations.forEach(s => subName[s.id] = s.fullName);

  const lines = rawLines.map(([from, to, circuit, flowDir, verify]) => ({
    id: from + '__' + to + '__' + circuit.replace(/\s+/g, ''),
    name: stripKV(subName[from]) + ' - ' + stripKV(subName[to]) + ' ' + circuit,
    from: from,
    to: to,
    voltage: circuit.indexOf('HVDC') !== -1 ? 500 : 400,
    circuit: circuit,
    flowDir: flowDir,
    status: 'in-service',
    color: '',
    weight: '',
    verify: verify,
    notes: ''
  }));

  // ------------------------------------------------------------------
  // ICTs — every station that steps between voltage classes.
  // 765/400 stations get 765/400 ICTs; the rest get 400/220 ICTs
  // feeding the (separate) 220 kV network.
  // Bay counts are ESTIMATES from the diagram — please correct.
  // ------------------------------------------------------------------
  const ictSpec = [
    // substationId, hvSide, lvSide, count
    ['PADGHAPG',   765, 400, 3],
    ['TALEGAONPG', 765, 400, 3],
    ['SOLPRPG',    765, 400, 3],
    ['WRDHAPG',    765, 400, 4],
    ['PARLYPG',    765, 400, 2],
    ['AURNGBDPG',  765, 400, 2],

    ['BOISRPG',    400, 220, 2],
    ['TARAPUR',    400, 220, 2],
    ['PADGE2',     400, 220, 3],
    ['KUDUS',      400, 220, 3],
    ['KARGAR',     400, 220, 2],
    ['VIKHROLI',   400, 220, 2],
    ['KALWA2',     400, 220, 3],
    ['NAGOTHANE',     400, 220, 2],
    ['NVMUMPG',    400, 220, 2],
    ['DHULE',      400, 220, 2],
    ['IBNASIK',    400, 220, 2],
    ['EKTUNI',     400, 220, 2],
    ['BABHLESHWAR',400, 220, 4],
    ['CHAKAN',     400, 220, 2],
    ['PUNEGIS',    400, 220, 3],
    ['LONIKAND',   400, 220, 2],
    ['LNKND2',     400, 220, 2],
    ['JEJURY',     400, 220, 2],
    ['KARAD2',     400, 220, 3],
    ['KARJAT',     400, 220, 2],
    ['KLPHR3',     400, 220, 2],
    ['ALKUD',      400, 220, 2],
    ['KUPWADPG',    400, 220, 2],
    ['SOLPR3',     400, 220, 2],
    ['KALAMBPG',   400, 220, 2],
    ['BHSWL2',     400, 220, 3],
    ['BHSWL3',     400, 220, 2],
    ['AKOLAAD',    400, 220, 2],
    ['AKOLA',      400, 220, 3],
    ['ARGBD3',     400, 220, 2],
    ['KORADY',     400, 220, 2],
    ['KHPKD2',     400, 220, 3],
    ['KORADI3AD',  400, 220, 3],
    ['KORADI2',    400, 220, 3],
    ['WARORAPG',   400, 220, 2],
    ['WARORA',     400, 220, 2],
    ['CHDPR',      400, 220, 3],
    ['CHDPR2',     400, 220, 3],
    ['CHDPRSW',    400, 220, 3],
    ['BDRVT',      400, 220, 2],
    ['NANDED',     400, 220, 3],
    ['PARLY2',     400, 220, 3],
    ['NEWPARLYPG', 400, 220, 2],
    ['DHRWL',      400, 220, 2],
    ['KOYNA4',     400, 220, 2],
    ['KOYNAAN',     400, 220, 2]
  ];

  const icts = [];
  ictSpec.forEach(([substationId, hvSide, lvSide, count]) => {
    for (let n = 1; n <= count; n++) {
      icts.push({
        id: 'ICT-' + substationId + '-' + hvSide + '-' + n,
        substationId: substationId,
        hvSide: hvSide,
        lvSide: lvSide,
        name: hvSide + '/' + lvSide + ' kV ICT-' + n,
        rating: hvSide === 765 ? '1500 MVA' : '315 MVA',
        verify: 'yes',
        notes: 'Bay count and rating are estimates — please confirm'
      });
    }
  });

  const settings = [
    ['color.sub.765', '#e14fd6'], ['color.sub.400', '#ff5555'],
    ['color.sub.220', '#f5c542'], ['color.sub.132', '#7ee8c9'],
    ['color.line.765', '#e14fd6'], ['color.line.400', '#ff5555'],
    ['color.line.220', '#f5c542'], ['color.line.132', '#7ee8c9'],
    ['color.line.hvdc', '#4cc4ff'],
    ['color.tripped', '#46525f'], ['color.isolated', '#ff3b3b']
  ].map(([key, value]) => ({ key: key, value: value }));

  writeRows(getSS().getSheetByName(SHEETS.SET), SCHEMAS.Settings, settings);
  writeRows(subSheet, SCHEMAS.Substations, substations);
  writeRows(lineSheet, SCHEMAS.Lines, lines);
  writeRows(ictSheet, SCHEMAS.ICTs, icts);

  SpreadsheetApp.flush();
  Logger.log('Setup complete: %s substations, %s line circuits, %s ICTs, %s settings',
    substations.length, lines.length, icts.length, settings.length);
  Logger.log('Your data Sheet: %s', ss.getUrl());
}

// "400 kV Chandrapur" -> "Chandrapur"  (used to build line names)
function stripKV(fullName) {
  if (!fullName) return '';
  return String(fullName).replace(/^[\d\/]+\s*kV\s*/i, '');
}

function writeRows(sheet, headers, rows) {
  const data = rows.map(r => headers.map(h => r[h] !== undefined ? r[h] : ''));
  if (data.length) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
}
