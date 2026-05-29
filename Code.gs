/**
 * Yacout Tours — Google Sheets ↔ GitHub sync v3
 *
 * Setup:
 * 1. Create a new Google Apps Script project (Extensions > Apps Script).
 * 2. Paste this file as Code.gs.
 * 3. Add the following script properties (File > Project settings > Script properties):
 *    - GITHUB_TOKEN : your GitHub personal access token (classic, repo scope)
 *    - GITHUB_OWNER : repository owner (username or org)
 *    - GITHUB_REPO  : repository name
 *    - GITHUB_BRANCH: branch name (default: master)
 *    - GITHUB_PATH  : path inside repo (default: public/data/cars.json)
 *    - PWA_SECRET   : un mot de passe secret pour sécuriser la PWA (ex: remons2024)
 * 4. Run setupSheet() once to initialise the sheet.
 * 5. Click "Installer le trigger automatique" in the menu — done!
 * 6. Deploy > New deployment > Web App
 *    - Execute as : Me
 *    - Who has access : Anyone
 *    Copier l'URL et la coller dans la PWA.
 */

// ── Sheet configuration ──────────────────────────────────────────────────────
const SHEET_NAME = 'Cars';
const COLUMNS = [
  'Car Name',
  'Category',
  'Price (EUR)',
  'Duration',
  'Seats',
  'Transmission',
  'Doors',
  'Fuel',
  'Image URL',
];

const RESERVATION_SHEET_NAME = 'Reservations';
const RESERVATION_COLUMNS = [
  'Created At',
  'Status',
  'Customer Name',
  'Email',
  'Phone',
  'Car Name',
  'Car Category',
  'Season',
  'Duration (Days)',
  'Daily Rate (EUR)',
  'Rental Total (EUR)',
  'Transport (EUR)',
  'Total (EUR)',
  'Location',
  'Start Date',
  'End Date',
  'Calendar Event ID',
  'Last Email Status',
  'Notes',
];
const RESERVATION_STATUSES = ['En attente', 'Confirmee', 'Refusee', 'Annulee', 'Terminee'];

// ── Menu ──────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Yacout Cars')
    .addItem('Initialiser la feuille Cars', 'setupSheet')
    .addSeparator()
    .addItem('Installer le trigger auto (Cars)', 'installTrigger')
    .addItem('Desinstaller le trigger (Cars)', 'uninstallTrigger')
    .addSeparator()
    .addItem('Synchroniser vers GitHub', 'manualSync')
    .addItem('Inspecter les donnees', 'testExport')
    .addSeparator()
    .addItem('Verifier connexion GitHub', 'checkGitHubConnection')
    .addItem('Resume du catalogue', 'showSummary')
    .addSeparator()
    .addItem('Vider le catalogue', 'clearCars')
    .addItem('A propos', 'showAbout')
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu('Yacout Reservations')
    .addItem('Initialiser la feuille Reservations', 'setupReservationsSheet')
    .addItem('Installer le trigger reservations', 'installReservationTrigger')
    .addItem('Desinstaller le trigger reservations', 'uninstallReservationTrigger')
    .addToUi();
}

// ── Initialiser la feuille Cars ──────────────────────────────────────────────
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  sheet.clear();
  sheet.appendRow(COLUMNS);

  const headerRange = sheet.getRange(1, 1, 1, COLUMNS.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1E293B');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  sheet.autoResizeColumns(1, COLUMNS.length);

  Logger.log('Sheet "' + SHEET_NAME + '" is ready.');
  SpreadsheetApp.getUi().alert('Feuille "' + SHEET_NAME + '" initialisee avec succes !');
}

// ── Trigger onChange (Cars) ──────────────────────────────────────────────────
function installTrigger() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onSheetChange') {
      ui.alert('Trigger deja installe', 'Le trigger automatique est deja actif.', ui.ButtonSet.OK);
      return;
    }
  }

  ScriptApp.newTrigger('onSheetChange')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange()
    .create();

  ui.alert('Trigger installe avec succes !');
}

function uninstallTrigger() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onSheetChange') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }
  ui.alert(removed > 0 ? removed + ' trigger(s) supprime(s).' : 'Aucun trigger trouve.');
}

function onSheetChange(e) {
  if (e.changeType === 'OTHER') return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('Sheet "' + SHEET_NAME + '" not found.'); return; }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) { Logger.log('No data rows.'); return; }

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(cell => cell === '')) continue;

    const car = {
      id: i,
      name:        String(row[0] || '').trim(),
      category:    String(row[1] || 'CAT A').trim(),
      price:       parseFloat(String(row[2]).replace(/[^0-9.]/g, '')) || 0,
      duration:    String(row[3] || 'jour').trim(),
      seats:       parseInt(String(row[4]), 10) || 4,
      transmission:String(row[5] || 'Manuelle').trim(),
      doors:       parseInt(String(row[6]), 10) || 4,
      fuel:        String(row[7] || 'Essence').trim(),
      image:       String(row[8] || '/images/car-placeholder.jpg').trim(),
    };

    if (!car.name) continue;
    rows.push(car);
  }

  if (rows.length === 0) { Logger.log('No valid cars.'); return; }

  Logger.log('[DEBUG] parsed cars: ' + JSON.stringify(rows));
  const json = JSON.stringify(rows, null, 2);
  pushToGitHub(json);
}

// ── Push JSON vers GitHub ─────────────────────────────────────────────────────
function pushToGitHub(jsonContent) {
  const token  = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const owner  = PropertiesService.getScriptProperties().getProperty('GITHUB_OWNER');
  const repo   = PropertiesService.getScriptProperties().getProperty('GITHUB_REPO');
  const branch = PropertiesService.getScriptProperties().getProperty('GITHUB_BRANCH') || 'master';
  const path   = PropertiesService.getScriptProperties().getProperty('GITHUB_PATH') || 'public/data/cars.json';

  if (!token || !owner || !repo) {
    Logger.log('Missing GITHUB_TOKEN, GITHUB_OWNER, or GITHUB_REPO.');
    return;
  }

  const api = 'https://api.github.com';
  let fileSha = null;
  try {
    const fileRes = UrlFetchApp.fetch(
      api + '/repos/' + owner + '/' + repo + '/contents/' + path + '?ref=' + branch,
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (fileRes.getResponseCode() === 200) {
      fileSha = JSON.parse(fileRes.getContentText()).sha;
    }
  } catch (e) { /* file may not exist yet */ }

  if (!jsonContent || jsonContent.trim() === '') { Logger.log('jsonContent empty.'); return; }

  const body = {
    message: 'Auto-sync cars from Google Sheets',
    branch: branch,
    content: Utilities.base64Encode(jsonContent),
  };
  if (fileSha) body.sha = fileSha;

  const commitRes = UrlFetchApp.fetch(
    api + '/repos/' + owner + '/' + repo + '/contents/' + path,
    {
      method: 'PUT',
      headers: { Authorization: 'Bearer ' + token },
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    }
  );

  const code = commitRes.getResponseCode();
  if (code === 200 || code === 201) {
    Logger.log('Successfully pushed cars.json to GitHub.');
  } else {
    Logger.log('GitHub API error: ' + commitRes.getContentText());
  }
}

// ── Sync manuelle ─────────────────────────────────────────────────────────────
function manualSync() {
  const ui = SpreadsheetApp.getUi();
  try {
    onSheetChange({ changeType: 'EDIT' });
    ui.alert('Synchronisation reussie !', '', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Erreur de synchronisation', e.message, ui.ButtonSet.OK);
  }
}

function testExport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert('Feuille introuvable.'); return; }

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(r => r[0] && r[0] !== '');
  const msg = rows.length > 0
    ? rows.map(r => String(r[0] || '(sans nom)') + ' | ' + r[2] + ' EUR | ' + r[1]).join('\n')
    : 'Aucune ligne valide.';
  SpreadsheetApp.getUi().alert('Donnees (' + rows.length + ' voiture(s))', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── GitHub Connection Check ───────────────────────────────────────────────────
function checkGitHubConnection() {
  const ui = SpreadsheetApp.getUi();
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const owner = PropertiesService.getScriptProperties().getProperty('GITHUB_OWNER');
  const repo  = PropertiesService.getScriptProperties().getProperty('GITHUB_REPO');
  if (!token || !owner || !repo) {
    ui.alert('Configuration manquante', 'Verifie GITHUB_TOKEN, GITHUB_OWNER, GITHUB_REPO.', ui.ButtonSet.OK);
    return;
  }
  try {
    const res = UrlFetchApp.fetch('https://api.github.com/repos/' + owner + '/' + repo, {
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code === 200) {
      const d = JSON.parse(res.getContentText());
      ui.alert('Connexion GitHub OK', 'Depot : ' + d.full_name + '\nBranche : ' + d.default_branch, ui.ButtonSet.OK);
    } else if (code === 401) { ui.alert('Token invalide.', '', ui.ButtonSet.OK); }
    else if (code === 404) { ui.alert('Depot introuvable.', '', ui.ButtonSet.OK); }
    else { ui.alert('Erreur ' + code, res.getContentText(), ui.ButtonSet.OK); }
  } catch (e) { ui.alert('Erreur reseau', e.message, ui.ButtonSet.OK); }
}

function showSummary() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { ui.alert('Feuille introuvable.'); return; }

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1).filter(r => r[0] && r[0] !== '');
  const prices = rows.map(r => parseFloat(String(r[2]).replace(/[^0-9.]/g, '')) || 0);
  const total = rows.length;
  const avg = prices.length ? (prices.reduce((a, b) => a + b, 0) / prices.length).toFixed(0) : 0;
  const min = prices.length ? Math.min(...prices) : 0;
  const max = prices.length ? Math.max(...prices) : 0;

  const categories = {};
  rows.forEach(r => { const c = String(r[1] || 'N/A').trim(); categories[c] = (categories[c] || 0) + 1; });
  const catStr = Object.entries(categories).map(([k, v]) => k + ': ' + v).join(', ');

  ui.alert(
    'Resume du catalogue',
    'Total voitures : ' + total + '\n\n'
    + 'Prix EUR : moyen ' + avg + ' | min ' + min + ' | max ' + max + '\n\n'
    + 'Categories : ' + (catStr || 'N/A'),
    ui.ButtonSet.OK
  );
}

function clearCars() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Vider le catalogue ?', '', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { ui.alert('Feuille introuvable.'); return; }
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  ui.alert('Catalogue vide.', '', ui.ButtonSet.OK);
}

function showAbout() {
  SpreadsheetApp.getUi().alert(
    'Yacout Tours — Sync Tool v3',
    'Sync automatique Cars + Reservations\n'
    + 'Google Sheets ↔ GitHub\n'
    + 'Categories : CAT A / B / C / D\n'
    + 'Prix en EUR | Saison Normale / Haute\n'
    + 'PWA Web App API activee.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ════════════════════════════════════════════════════════════════════════════════
// ██████  WEB APP API — PWA Controller  ████████████████████████████████████████
// ════════════════════════════════════════════════════════════════════════════════

function _getCarsFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(cell => cell === '')) continue;
    const name = String(row[0] || '').trim();
    if (!name) continue;
    rows.push({
      id: i,
      name,
      category:     String(row[1] || 'CAT A').trim(),
      price:        parseFloat(String(row[2]).replace(/[^0-9.]/g, '')) || 0,
      duration:     String(row[3] || 'jour').trim(),
      seats:        parseInt(String(row[4]), 10) || 4,
      transmission: String(row[5] || 'Manuelle').trim(),
      doors:        parseInt(String(row[6]), 10) || 4,
      fuel:         String(row[7] || 'Essence').trim(),
      image:        String(row[8] || '/images/car-placeholder.jpg').trim(),
    });
  }
  return rows;
}

function _getReservationsList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESERVATION_SHEET_NAME);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(cell => cell === '')) continue;
    rows.push({
      rowNumber: i + 1,
      name:      String(row[2] || '').trim(),
      email:     String(row[3] || '').trim(),
      phone:     String(row[4] || '').trim(),
      carName:   String(row[5] || '').trim(),
      carPrice:  row[9] || 0,
      carDuration: row[8] || 'jour',
      startDate: row[14] || '',
      endDate:   row[15] || '',
      status:    String(row[1] || 'En attente').trim(),
    });
  }
  return rows;
}

function _updateReservationRowStatus(rowNumber, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESERVATION_SHEET_NAME);
  if (!sheet) throw new Error('Sheet Reservations introuvable.');
  sheet.getRange(rowNumber, 2).setValue(newStatus);
  const reservation = _getReservationFromRow(sheet, rowNumber);
  _sendReservationStatusEmail(reservation);
  sheet.getRange(rowNumber, 18).setValue(newStatus);
  _updateReservationCalendarEvent(reservation);
}

function _writeCarsToSheet(cars) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" introuvable.');

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  if (cars.length === 0) return;

  const rows = cars.map(c => [
    c.name         || '',
    c.category     || 'CAT A',
    c.price        || 0,
    c.duration     || 'jour',
    c.seats        || 4,
    c.transmission || 'Manuelle',
    c.doors        || 4,
    c.fuel         || 'Essence',
    c.image        || '/images/car-placeholder.jpg',
  ]);

  sheet.getRange(2, 1, rows.length, COLUMNS.length).setValues(rows);
}

function _jsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function _auth(paramSecret, bodySecret) {
  const stored = PropertiesService.getScriptProperties().getProperty('PWA_SECRET');
  if (!stored) return true;
  return (paramSecret === stored) || (bodySecret === stored);
}

// ── Handler GET ───────────────────────────────────────────────────────────────
function doGet(e) {
  const paramSecret = e.parameter.pwa_secret || '';

  // Mutations via payload base64
  if (e.parameter.payload) {
    let body = {};
    try {
      const decoded = decodeURIComponent(escape(
        Utilities.newBlob(Utilities.base64Decode(e.parameter.payload)).getDataAsString()
      ));
      body = JSON.parse(decoded);
    } catch (err) {
      return _jsonResponse({ ok: false, error: 'Payload invalide : ' + err.message });
    }

    const bodySecret = body.pwa_secret || '';
    if (!_auth(paramSecret, bodySecret)) {
      return _jsonResponse({ ok: false, error: 'Non autorise — secret invalide.' });
    }

    const action = body.action || '';

    try {
      if (action === 'createReservation') {
        return _jsonResponse(createReservation(body.reservation));
      }

      if (action === 'add') {
        const car = body.car;
        if (!car || !car.name) return _jsonResponse({ ok: false, error: 'car.name manquant.' });
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
        if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });
        sheet.appendRow([
          car.name         || '',
          car.category     || 'CAT A',
          car.price        || 0,
          car.duration     || 'jour',
          car.seats        || 4,
          car.transmission || 'Manuelle',
          car.doors        || 4,
          car.fuel         || 'Essence',
          car.image        || '/images/car-placeholder.jpg',
        ]);
        onSheetChange({ changeType: 'EDIT' });
        return _jsonResponse({ ok: true, message: 'Voiture ajoutee et synchronisee.' });
      }

      if (action === 'update') {
        const car = body.car;
        if (!car || !car.id) return _jsonResponse({ ok: false, error: 'car.id manquant.' });
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
        if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });
        const rowNum = parseInt(car.id, 10) + 1;
        sheet.getRange(rowNum, 1, 1, COLUMNS.length).setValues([[
          car.name         || '',
          car.category     || 'CAT A',
          car.price        || 0,
          car.duration     || 'jour',
          car.seats        || 4,
          car.transmission || 'Manuelle',
          car.doors        || 4,
          car.fuel         || 'Essence',
          car.image        || '/images/car-placeholder.jpg',
        ]]);
        onSheetChange({ changeType: 'EDIT' });
        return _jsonResponse({ ok: true, message: 'Voiture modifiee et synchronisee.' });
      }

      if (action === 'delete') {
        if (!body.id) return _jsonResponse({ ok: false, error: 'id manquant.' });
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
        if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });
        sheet.deleteRow(parseInt(body.id, 10) + 1);
        onSheetChange({ changeType: 'EDIT' });
        return _jsonResponse({ ok: true, message: 'Voiture supprimee et synchronisee.' });
      }

      if (action === 'clear') {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
        if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
        onSheetChange({ changeType: 'EDIT' });
        return _jsonResponse({ ok: true, message: 'Catalogue vide et synchronise.' });
      }

      if (action === 'reorder') {
        if (!Array.isArray(body.cars)) return _jsonResponse({ ok: false, error: 'cars (tableau) manquant.' });
        _writeCarsToSheet(body.cars);
        onSheetChange({ changeType: 'EDIT' });
        return _jsonResponse({ ok: true, message: 'Voitures reordonnees et synchronisees.' });
      }

      if (action === 'listReservations') {
        return _jsonResponse({ ok: true, reservations: _getReservationsList() });
      }

      if (action === 'updateReservationStatus') {
        if (!body.rowNumber || !body.status) return _jsonResponse({ ok: false, error: 'rowNumber et status requis.' });
        _updateReservationRowStatus(body.rowNumber, body.status);
        return _jsonResponse({ ok: true, message: 'Statut mis a jour.' });
      }

      return _jsonResponse({ ok: false, error: 'Action inconnue : ' + action });

    } catch (err) {
      Logger.log('[payload error] ' + err.message);
      return _jsonResponse({ ok: false, error: err.message });
    }
  }

  // GET classiques (lecture)
  if (!_auth(paramSecret, '')) {
    return _jsonResponse({ ok: false, error: 'Non autorise.' });
  }

  const action = e.parameter.action || 'list';

  try {
    if (action === 'list') {
      return _jsonResponse({ ok: true, cars: _getCarsFromSheet() });
    }

    if (action === 'summary') {
      const cars = _getCarsFromSheet();
      const prices = cars.map(c => c.price).filter(p => p > 0);
      return _jsonResponse({
        ok: true,
        total: cars.length,
        avgPrice: prices.length ? Math.round(prices.reduce((a, b) => a + b, 0) / prices.length) : 0,
        minPrice: prices.length ? Math.min(...prices) : 0,
        maxPrice: prices.length ? Math.max(...prices) : 0,
      });
    }

    if (action === 'sync') {
      onSheetChange({ changeType: 'EDIT' });
      return _jsonResponse({ ok: true, message: 'Synchronise vers GitHub.' });
    }

    if (action === 'listReservations') {
      return _jsonResponse({ ok: true, reservations: _getReservationsList() });
    }

    return _jsonResponse({ ok: false, error: 'Action inconnue : ' + action });

  } catch (err) {
    Logger.log('[doGet error] ' + err.message);
    return _jsonResponse({ ok: false, error: err.message });
  }
}

// ── Handler POST ──────────────────────────────────────────────────────────────
function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return _jsonResponse({ ok: false, error: 'JSON invalide.' });
  }

  const bodySecret  = body.pwa_secret || '';
  const paramSecret = (e.parameter && e.parameter.pwa_secret) || '';
  if (!_auth(paramSecret, bodySecret)) {
    return _jsonResponse({ ok: false, error: 'Non autorise.' });
  }

  const action = body.action || '';

  try {
    if (action === 'createReservation') {
      return _jsonResponse(createReservation(body.reservation));
    }

    if (action === 'add') {
      const car = body.car;
      if (!car || !car.name) return _jsonResponse({ ok: false, error: 'car.name manquant.' });
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });
      sheet.appendRow([
        car.name         || '',
        car.category     || 'CAT A',
        car.price        || 0,
        car.duration     || 'jour',
        car.seats        || 4,
        car.transmission || 'Manuelle',
        car.doors        || 4,
        car.fuel         || 'Essence',
        car.image        || '/images/car-placeholder.jpg',
      ]);
      onSheetChange({ changeType: 'EDIT' });
      return _jsonResponse({ ok: true, message: 'Voiture ajoutee et synchronisee.' });
    }

    if (action === 'update') {
      const car = body.car;
      if (!car || !car.id) return _jsonResponse({ ok: false, error: 'car.id manquant.' });
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });
      const rowNum = parseInt(car.id, 10) + 1;
      sheet.getRange(rowNum, 1, 1, COLUMNS.length).setValues([[
        car.name         || '',
        car.category     || 'CAT A',
        car.price        || 0,
        car.duration     || 'jour',
        car.seats        || 4,
        car.transmission || 'Manuelle',
        car.doors        || 4,
        car.fuel         || 'Essence',
        car.image        || '/images/car-placeholder.jpg',
      ]]);
      onSheetChange({ changeType: 'EDIT' });
      return _jsonResponse({ ok: true, message: 'Voiture modifiee et synchronisee.' });
    }

    if (action === 'delete') {
      if (!body.id) return _jsonResponse({ ok: false, error: 'id manquant.' });
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });
      sheet.deleteRow(parseInt(body.id, 10) + 1);
      onSheetChange({ changeType: 'EDIT' });
      return _jsonResponse({ ok: true, message: 'Voiture supprimee et synchronisee.' });
    }

    if (action === 'clear') {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      onSheetChange({ changeType: 'EDIT' });
      return _jsonResponse({ ok: true, message: 'Catalogue vide et synchronise.' });
    }

    if (action === 'reorder') {
      if (!Array.isArray(body.cars)) return _jsonResponse({ ok: false, error: 'cars (tableau) manquant.' });
      _writeCarsToSheet(body.cars);
      onSheetChange({ changeType: 'EDIT' });
      return _jsonResponse({ ok: true, message: 'Voitures reordonnees et synchronisees.' });
    }

    if (action === 'listReservations') {
      return _jsonResponse({ ok: true, reservations: _getReservationsList() });
    }

    if (action === 'updateReservationStatus') {
      if (!body.rowNumber || !body.status) return _jsonResponse({ ok: false, error: 'rowNumber et status requis.' });
      _updateReservationRowStatus(body.rowNumber, body.status);
      return _jsonResponse({ ok: true, message: 'Statut mis a jour.' });
    }

    return _jsonResponse({ ok: false, error: 'Action inconnue : ' + action });

  } catch (err) {
    Logger.log('[doPost error] ' + err.message);
    return _jsonResponse({ ok: false, error: err.message });
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// ██████  Reservations module  █████████████████████████████████████████████████
// ════════════════════════════════════════════════════════════════════════════════

function setupReservationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(RESERVATION_SHEET_NAME);

  sheet.getRange(1, 1, 1, RESERVATION_COLUMNS.length).setValues([RESERVATION_COLUMNS]);
  sheet.setFrozenRows(1);

  const headerRange = sheet.getRange(1, 1, 1, RESERVATION_COLUMNS.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1E293B');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  const statusRange = sheet.getRange(2, 2, Math.max(sheet.getMaxRows() - 1, 1), 1);
  statusRange.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(RESERVATION_STATUSES, true).setAllowInvalid(false).build()
  );

  sheet.autoResizeColumns(1, RESERVATION_COLUMNS.length);
  Logger.log('Reservations sheet ready.');
  try {
    SpreadsheetApp.getUi().alert('Feuille Reservations configuree avec succes.');
  } catch (err) {
    Logger.log('[setupReservationsSheet ui] ' + err.message);
  }
}

function installReservationTrigger() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onReservationEdit') {
      ui.alert('Trigger reservations deja installe.');
      return;
    }
  }
  ScriptApp.newTrigger('onReservationEdit').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  ui.alert('Trigger reservations installe.');
}

function uninstallReservationTrigger() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onReservationEdit') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }
  ui.alert(removed + ' trigger(s) reservations supprime(s).');
}

function onReservationEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== RESERVATION_SHEET_NAME) return;
  if (e.range.getRow() <= 1 || e.range.getColumn() !== 2) return;

  const rowNumber = e.range.getRow();
  const status = String(sheet.getRange(rowNumber, 2).getValue() || '').trim();
  const lastEmailStatus = String(sheet.getRange(rowNumber, 18).getValue() || '').trim();
  if (!status || status === lastEmailStatus) return;

  const reservation = _getReservationFromRow(sheet, rowNumber);
  if (!reservation.email) { Logger.log('Row ' + rowNumber + ' has no email.'); return; }

  _sendReservationStatusEmail(reservation);
  sheet.getRange(rowNumber, 18).setValue(status);
  _updateReservationCalendarEvent(reservation);
}

function createReservation(reservation) {
  if (!reservation) throw new Error('Reservation manquante.');
  if (!reservation.name) throw new Error('Nom client manquant.');
  if (!reservation.email) throw new Error('Email client manquant.');
  if (!reservation.phone) throw new Error('Telephone client manquant.');
  if (!reservation.carName) throw new Error('Voiture manquante.');
  if (!reservation.startDate || !reservation.endDate) throw new Error('Dates manquantes.');

  setupReservationsSheet();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESERVATION_SHEET_NAME);
  const status = 'En attente';
  const row = [
    new Date(),
    status,
    String(reservation.name || '').trim(),
    String(reservation.email || '').trim(),
    String(reservation.phone || '').trim(),
    String(reservation.carName || '').trim(),
    String(reservation.carCategory || '').trim(),
    String(reservation.season || '').trim(),
    reservation.durationDays || reservation.carDuration || '',
    reservation.dailyRateEUR || reservation.carPrice || '',
    reservation.rentalTotalEUR || '',
    reservation.transportEUR || 0,
    reservation.totalEUR || '',
    String(reservation.location || '').trim(),
    reservation.startDate || '',
    reservation.endDate || '',
    '',
    '',
    '',
  ];

  sheet.appendRow(row);
  const rowNumber = sheet.getLastRow();
  const saved = _getReservationFromRow(sheet, rowNumber);
  const eventId = _createReservationCalendarEvent(saved);
  if (eventId) {
    sheet.getRange(rowNumber, 17).setValue(eventId);
    saved.calendarEventId = eventId;
  }

  try {
    _sendReservationStatusEmail(saved);
    sheet.getRange(rowNumber, 18).setValue(status);
  } catch (err) {
    Logger.log('[reservation initial email] ' + err.message);
  }

  return { ok: true, row: rowNumber, status: status };
}

function _getReservationFromRow(sheet, rowNumber) {
  const values = sheet.getRange(rowNumber, 1, 1, RESERVATION_COLUMNS.length).getValues()[0];
  return {
    rowNumber: rowNumber,
    createdAt: values[0],
    status: String(values[1] || '').trim(),
    name: String(values[2] || '').trim(),
    email: String(values[3] || '').trim(),
    phone: String(values[4] || '').trim(),
    carName: String(values[5] || '').trim(),
    carCategory: String(values[6] || '').trim(),
    season: String(values[7] || '').trim(),
    durationDays: values[8],
    dailyRateEUR: values[9],
    rentalTotalEUR: values[10],
    transportEUR: values[11],
    totalEUR: values[12],
    location: String(values[13] || '').trim(),
    startDate: values[14],
    endDate: values[15],
    calendarEventId: String(values[16] || '').trim(),
    notes: String(values[18] || '').trim(),
  };
}

function _parseReservationDate(value) {
  if (value instanceof Date) return value;
  if (!value) return null;
  const text = String(value).trim();
  const parts = text.split('-');
  if (parts.length === 3) {
    return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  }
  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function _getCalendarEndDate(value) {
  const end = _parseReservationDate(value);
  if (!end) return null;
  end.setDate(end.getDate() + 1);
  return end;
}

function _createReservationCalendarEvent(reservation) {
  const start = _parseReservationDate(reservation.startDate);
  const end = _getCalendarEndDate(reservation.endDate);
  if (!start || !end) return '';

  const title = 'Reservation - ' + reservation.carName + ' - ' + reservation.name;
  const description = [
    'Client: ' + reservation.name,
    'Email: ' + reservation.email,
    'Telephone: ' + reservation.phone,
    'Voiture: ' + reservation.carName + ' (' + (reservation.carCategory || '') + ')',
    'Saison: ' + (reservation.season || '-'),
    'Duree: ' + (reservation.durationDays || '-') + ' jours',
    'Total: ' + (reservation.totalEUR || '-') + ' EUR',
    'Lieu: ' + (reservation.location || '-'),
    'Statut: ' + reservation.status,
  ].join('\n');

  const event = CalendarApp.getDefaultCalendar().createAllDayEvent(title, start, end, {
    description: description,
  });
  return event.getId();
}

function _updateReservationCalendarEvent(reservation) {
  if (!reservation.calendarEventId) return;
  try {
    const event = CalendarApp.getDefaultCalendar().getEventById(reservation.calendarEventId);
    if (!event) return;
    event.setTitle('Reservation ' + reservation.status + ' - ' + reservation.carName + ' - ' + reservation.name);
    event.setDescription([
      'Client: ' + reservation.name,
      'Email: ' + reservation.email,
      'Telephone: ' + reservation.phone,
      'Voiture: ' + reservation.carName + ' (' + (reservation.carCategory || '') + ')',
      'Saison: ' + (reservation.season || '-'),
      'Duree: ' + (reservation.durationDays || '-') + ' jours',
      'Total: ' + (reservation.totalEUR || '-') + ' EUR',
      'Lieu: ' + (reservation.location || '-'),
      'Statut: ' + reservation.status,
      'Notes: ' + reservation.notes,
    ].join('\n'));
  } catch (err) {
    Logger.log('[calendar update] ' + err.message);
  }
}

function _getStatusEmailContent(reservation) {
  const status = reservation.status || 'En attente';
  const statusMessages = {
    'En attente': 'Votre demande de reservation a bien ete recue et elle est en cours de traitement.',
    'Confirmee':  'Votre reservation est confirmee. Notre equipe vous attendra aux dates prevues.',
    'Refusee':    'Nous sommes desoles, votre reservation ne peut pas etre confirmee pour ces dates.',
    'Annulee':    'Votre reservation a ete annulee. Vous pouvez nous contacter pour une nouvelle demande.',
    'Terminee':   'Votre reservation est marquee comme terminee. Merci pour votre confiance.',
  };
  const message = statusMessages[status] || 'Le statut de votre reservation a ete mis a jour.';
  const subject = 'Mise a jour de votre reservation - ' + reservation.carName;

  const details =
    '<p style="margin:0 0 10px;"><strong>Statut:</strong> ' + _escapeHtml(status) + '</p>'
    + '<p style="margin:0 0 10px;"><strong>Voiture:</strong> ' + _escapeHtml(reservation.carName)
    + (reservation.carCategory ? ' (' + _escapeHtml(reservation.carCategory) + ')' : '') + '</p>'
    + '<p style="margin:0 0 10px;"><strong>Saison:</strong> ' + _escapeHtml(reservation.season || '-') + '</p>'
    + '<p style="margin:0 0 10px;"><strong>Duree:</strong> ' + (reservation.durationDays || '-') + ' jours</p>'
    + '<p style="margin:0 0 10px;"><strong>Lieu:</strong> ' + _escapeHtml(reservation.location || '-') + '</p>'
    + '<p style="margin:0 0 10px;"><strong>Total:</strong> ' + (reservation.totalEUR || '-') + ' EUR</p>'
    + '<p style="margin:0 0 10px;"><strong>Date depart:</strong> ' + _formatReservationDate(reservation.startDate) + '</p>'
    + '<p style="margin:0;"><strong>Date retour:</strong> ' + _formatReservationDate(reservation.endDate) + '</p>';

  const htmlBody =
    '<div style="margin:0;padding:24px;background:#f6f7f9;font-family:Arial,sans-serif;color:#1f2937;">'
    + '<div style="max-width:640px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e5e7eb;">'
    + '<div style="background:#1E293B;color:#ffffff;padding:24px 28px;">'
    + '<h1 style="margin:0;font-size:22px;line-height:1.3;">Yacout Tours</h1>'
    + '<p style="margin:8px 0 0;color:#cbd5e1;">Mise a jour de reservation</p>'
    + '</div>'
    + '<div style="padding:28px;">'
    + '<p style="font-size:16px;margin:0 0 16px;">Bonjour ' + _escapeHtml(reservation.name) + ',</p>'
    + '<p style="font-size:15px;line-height:1.7;margin:0 0 22px;">' + _escapeHtml(message) + '</p>'
    + '<div style="border:1px solid #e5e7eb;border-radius:10px;padding:18px;background:#fafafa;">'
    + details
    + '</div>'
    + '<p style="font-size:14px;color:#64748b;line-height:1.6;margin:22px 0 0;">Pour toute question, repondez a cet email ou contactez notre equipe.</p>'
    + '</div>'
    + '</div>'
    + '</div>';

  return { subject: subject, htmlBody: htmlBody };
}

function _sendReservationStatusEmail(reservation) {
  const content = _getStatusEmailContent(reservation);
  MailApp.sendEmail({
    to: reservation.email,
    subject: content.subject,
    htmlBody: content.htmlBody,
  });
}

function _formatReservationDate(value) {
  const date = _parseReservationDate(value);
  if (!date) return _escapeHtml(String(value || ''));
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function _escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
