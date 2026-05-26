/**
 * Remons Car Rental — Google Sheets ↔ GitHub sync
 *
 * Setup:
 * 1. Create a new Google Apps Script project (Extensions > Apps Script).
 * 2. Paste this file as Code.gs.
 * 3. Add the following script properties (File > Project settings > Script properties):
 *    - GITHUB_TOKEN : your GitHub personal access token (classic, repo scope)
 *    - GITHUB_OWNER : repository owner (username or org)
 *    - GITHUB_REPO  : repository name
 *    - GITHUB_BRANCH: branch name (default: master)
 *    - GITHUB_PATH  : path inside repo (default: data/cars.json)
 *    - PWA_SECRET   : un mot de passe secret pour sécuriser la PWA (ex: remons2024)
 * 4. Run setupSheet() once to initialise the sheet.
 * 5. Click "⚡ Installer le trigger automatique" in the menu — done!
 * 6. Deploy > New deployment > Web App
 *    - Execute as : Me
 *    - Who has access : Anyone
 *    Copier l'URL et la coller dans la PWA.
 */

// ── Sheet configuration ──────────────────────────────────────────────────────
const SHEET_NAME = 'Cars';
const COLUMNS = [
  'Car Name',
  'Price (MAD)',
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
  'Price (MAD)',
  'Duration',
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
    .createMenu('🚗 Remons Car Rental')
    .addItem('📋 Initialiser la feuille', 'setupSheet')
    .addSeparator()
    .addItem('⚡ Installer le trigger automatique', 'installTrigger')
    .addItem('🗑️ Désinstaller le trigger', 'uninstallTrigger')
    .addSeparator()
    .addItem('🔄 Synchroniser vers GitHub', 'manualSync')
    .addItem('🔍 Inspecter les données brutes', 'testExport')
    .addSeparator()
    .addItem('✅ Vérifier la connexion GitHub', 'checkGitHubConnection')
    .addItem('📊 Résumé du catalogue', 'showSummary')
    .addSeparator()
    .addItem('🗑️ Vider le catalogue (garder l\'en-tête)', 'clearCars')
    .addItem('ℹ️ À propos', 'showAbout')
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu('Reservations')
    .addItem('Setup reservations', 'setupReservationsSheet')
    .addItem('Installer le trigger reservations', 'installReservationTrigger')
    .addItem('Desinstaller le trigger reservations', 'uninstallReservationTrigger')
    .addToUi();
}

// ── Initialiser la feuille ────────────────────────────────────────────────────
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  sheet.clear();
  sheet.appendRow(COLUMNS);

  const headerRange = sheet.getRange(1, 1, 1, COLUMNS.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1E293B');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  sheet.autoResizeColumns(1, COLUMNS.length);

  Logger.log('✅ Sheet "' + SHEET_NAME + '" is ready.');
  SpreadsheetApp.getUi().alert('✅ Feuille "' + SHEET_NAME + '" initialisée avec succès !');
}

// ── Installer le trigger onChange ─────────────────────────────────────────────
function installTrigger() {
  const ui = SpreadsheetApp.getUi();

  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onSheetChange') {
      ui.alert(
        '⚡ Trigger déjà installé',
        'Le trigger automatique est déjà actif.\nChaque modification du sheet déclenchera une sync vers GitHub.',
        ui.ButtonSet.OK
      );
      return;
    }
  }

  ScriptApp.newTrigger('onSheetChange')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange()
    .create();

  ui.alert(
    '✅ Trigger installé avec succès !',
    'Désormais, chaque modification du sheet "Cars" sera automatiquement synchronisée vers GitHub.\n\nAucune action supplémentaire requise.',
    ui.ButtonSet.OK
  );
}

// ── Désinstaller le trigger ───────────────────────────────────────────────────
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

  if (removed > 0) {
    ui.alert('✅ Trigger désinstallé', removed + ' trigger(s) supprimé(s). La sync automatique est désactivée.', ui.ButtonSet.OK);
  } else {
    ui.alert('ℹ️ Aucun trigger trouvé', 'Aucun trigger automatique n\'était installé.', ui.ButtonSet.OK);
  }
}

// ── Trigger: se déclenche automatiquement à chaque modification ───────────────
function onSheetChange(e) {
  if (e.changeType === 'OTHER') return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log('Sheet "' + SHEET_NAME + '" not found. Run setupSheet() first.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('No data rows to export.');
    return;
  }

  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(cell => cell === '')) continue;

    const car = {
      id: i,
      name: String(row[0] || '').trim(),
      price: parseFloat(String(row[1]).replace(/[^0-9.]/g, '')) || 0,
      duration: String(row[2] || 'Per Day').trim(),
      seats: parseInt(String(row[3]), 10) || 4,
      transmission: String(row[4] || 'Automatic').trim(),
      doors: parseInt(String(row[5]), 10) || 4,
      fuel: String(row[6] || 'Petrol').trim(),
      image: String(row[7] || '/images/car-placeholder.jpg').trim(),
    };

    if (!car.name) continue;

    rows.push(car);
  }

  if (rows.length === 0) {
    Logger.log('No valid cars found.');
    return;
  }

  Logger.log('[DEBUG] raw data rows: ' + JSON.stringify(data));
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
    Logger.log('❌ Missing GITHUB_TOKEN, GITHUB_OWNER, or GITHUB_REPO in script properties.');
    return;
  }

  const api = 'https://api.github.com';

  // 1. Récupérer le SHA du fichier existant (si présent)
  let fileSha = null;
  try {
    const fileRes = UrlFetchApp.fetch(
      api + '/repos/' + owner + '/' + repo + '/contents/' + path + '?ref=' + branch,
      {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true,
      }
    );
    Logger.log('[GitHub GET Response] ' + fileRes.getContentText());
    if (fileRes.getResponseCode() === 200) {
      const fileData = JSON.parse(fileRes.getContentText());
      fileSha = fileData.sha;
    }
  } catch (e) {
    // Le fichier n'existe pas encore — c'est normal
  }

  // 2. Valider le contenu
  if (!jsonContent || jsonContent.trim() === '') {
    Logger.log('❌ jsonContent vide — rien à pusher.');
    return;
  }

  // 3. Créer ou mettre à jour le fichier
  const contentBase64 = Utilities.base64Encode(jsonContent);
  const body = {
    message: '🔄 Auto-sync cars from Google Sheets',
    branch: branch,
    content: contentBase64,
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

  Logger.log('[GitHub PUT Response] ' + commitRes.getContentText());

  const code = commitRes.getResponseCode();
  if (code === 200 || code === 201) {
    Logger.log('✅ Successfully pushed cars.json to GitHub.');
  } else {
    Logger.log('❌ GitHub API error: ' + commitRes.getContentText());
  }
}

// ── Sync manuelle (bouton menu) ───────────────────────────────────────────────
function manualSync() {
  const ui = SpreadsheetApp.getUi();
  try {
    onSheetChange({ changeType: 'EDIT' });
    ui.alert('✅ Synchronisation réussie !', 'Les données ont été envoyées vers GitHub.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Erreur de synchronisation', e.message, ui.ButtonSet.OK);
  }
}

// ── Inspecter les données brutes ──────────────────────────────────────────────
function testExport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log('❌ Sheet "' + SHEET_NAME + '" not found.');
    SpreadsheetApp.getUi().alert('❌ Feuille "' + SHEET_NAME + '" introuvable. Lance setupSheet() d\'abord.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  Logger.log('[testExport] raw data: ' + JSON.stringify(data));

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(cell => cell === '')) continue;
    rows.push({
      id: i,
      name: String(row[0] || '').trim(),
      price: row[1],
      transmission: String(row[4] || '').trim(),
      fuel: String(row[6] || '').trim(),
    });
  }

  Logger.log('[testExport] parsed rows: ' + JSON.stringify(rows));

  const msg = rows.length > 0
    ? rows.map(r => `#${r.id} — ${r.name || '(sans nom)'} | ${r.price} MAD | ${r.transmission} | ${r.fuel}`).join('\n')
    : '⚠️ Aucune ligne valide trouvée.';

  SpreadsheetApp.getUi().alert('🔍 Données brutes (' + rows.length + ' voiture(s))', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── Vérifier la connexion GitHub ──────────────────────────────────────────────
function checkGitHubConnection() {
  const ui = SpreadsheetApp.getUi();
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  const owner = PropertiesService.getScriptProperties().getProperty('GITHUB_OWNER');
  const repo  = PropertiesService.getScriptProperties().getProperty('GITHUB_REPO');

  if (!token || !owner || !repo) {
    ui.alert(
      '❌ Configuration manquante',
      'Vérifie que GITHUB_TOKEN, GITHUB_OWNER et GITHUB_REPO sont définis dans les propriétés du script.',
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    const res = UrlFetchApp.fetch(
      'https://api.github.com/repos/' + owner + '/' + repo,
      {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true,
      }
    );
    const code = res.getResponseCode();
    if (code === 200) {
      const repoData = JSON.parse(res.getContentText());
      ui.alert(
        '✅ Connexion GitHub OK',
        'Dépôt : ' + repoData.full_name + '\n' +
        'Visibilité : ' + (repoData.private ? 'Privé' : 'Public') + '\n' +
        'Branche par défaut : ' + repoData.default_branch,
        ui.ButtonSet.OK
      );
    } else if (code === 401) {
      ui.alert('❌ Token invalide', 'GITHUB_TOKEN est incorrect ou expiré.', ui.ButtonSet.OK);
    } else if (code === 404) {
      ui.alert('❌ Dépôt introuvable', 'Vérifie GITHUB_OWNER et GITHUB_REPO.', ui.ButtonSet.OK);
    } else {
      ui.alert('❌ Erreur ' + code, res.getContentText(), ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('❌ Erreur réseau', e.message, ui.ButtonSet.OK);
  }
}

// ── Résumé du catalogue ───────────────────────────────────────────────────────
function showSummary() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    ui.alert('❌ Feuille "' + SHEET_NAME + '" introuvable.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert('📊 Catalogue vide', 'Aucune voiture enregistrée.', ui.ButtonSet.OK);
    return;
  }

  const rows = data.slice(1).filter(r => r[0] && r[0] !== '');
  const prices = rows.map(r => parseFloat(String(r[1]).replace(/[^0-9.]/g, '')) || 0);
  const totalCars = rows.length;
  const avgPrice = prices.length ? (prices.reduce((a, b) => a + b, 0) / prices.length).toFixed(0) : 0;
  const minPrice = prices.length ? Math.min(...prices) : 0;
  const maxPrice = prices.length ? Math.max(...prices) : 0;

  const fuels = {};
  const transmissions = {};
  rows.forEach(r => {
    const f = String(r[6] || 'N/A').trim();
    const t = String(r[4] || 'N/A').trim();
    fuels[f] = (fuels[f] || 0) + 1;
    transmissions[t] = (transmissions[t] || 0) + 1;
  });

  const fuelStr = Object.entries(fuels).map(([k, v]) => k + ': ' + v).join(', ');
  const transStr = Object.entries(transmissions).map(([k, v]) => k + ': ' + v).join(', ');

  ui.alert(
    '📊 Résumé du catalogue',
    '🚗 Total voitures : ' + totalCars + '\n\n' +
    '💰 Prix moyen : ' + avgPrice + ' MAD\n' +
    '   Min : ' + minPrice + ' MAD | Max : ' + maxPrice + ' MAD\n\n' +
    '⛽ Carburants : ' + (fuelStr || 'N/A') + '\n' +
    '⚙️ Transmissions : ' + (transStr || 'N/A'),
    ui.ButtonSet.OK
  );
}

// ── Vider le catalogue ────────────────────────────────────────────────────────
function clearCars() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Confirmer la suppression',
    'Vider toutes les voitures du catalogue ? L\'en-tête sera conservé.',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    ui.alert('❌ Feuille "' + SHEET_NAME + '" introuvable.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  ui.alert('✅ Catalogue vidé', 'Toutes les voitures ont été supprimées. L\'en-tête est conservé.', ui.ButtonSet.OK);
}

// ── À propos ──────────────────────────────────────────────────────────────────
function showAbout() {
  SpreadsheetApp.getUi().alert(
    'ℹ️ Remons Car Rental — Sync Tool',
    'Version : 2.0\n' +
    'Auteur  : Remons\n\n' +
    'Ce script synchronise automatiquement le catalogue de voitures\n' +
    'de Google Sheets vers GitHub (data/cars.json) à chaque modification.\n\n' +
    'Fonctionnalités du menu :\n' +
    '• Initialiser la feuille\n' +
    '• Installer / désinstaller le trigger automatique\n' +
    '• Sync manuel vers GitHub\n' +
    '• Inspecter les données brutes\n' +
    '• Vérifier la connexion GitHub\n' +
    '• Résumé du catalogue\n' +
    '• Vider le catalogue\n\n' +
    '🌐 PWA Web App API activée\n' +
    'Deploy > Web App pour obtenir l\'URL de connexion.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ════════════════════════════════════════════════════════════════════════════════
// ██████████████████████  WEB APP API — PWA Controller  ██████████████████████
// ════════════════════════════════════════════════════════════════════════════════
//
// DÉPLOIEMENT :
//   Apps Script > Deploy > New deployment > Web App
//   Execute as  : Me
//   Who has access : Anyone
//
// PROPRIÉTÉ À AJOUTER dans Script Properties :
//   PWA_SECRET = votre_mot_de_passe_secret
//
// La PWA enverra ce secret dans chaque requête via le paramètre "pwa_secret".
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Lit et retourne le tableau de voitures depuis le sheet.
 */
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
      price:        parseFloat(String(row[1]).replace(/[^0-9.]/g, '')) || 0,
      duration:     String(row[2] || 'Per Day').trim(),
      seats:        parseInt(String(row[3]), 10) || 4,
      transmission: String(row[4] || 'Automatic').trim(),
      doors:        parseInt(String(row[5]), 10) || 4,
      fuel:         String(row[6] || 'Petrol').trim(),
      image:        String(row[7] || '/images/car-placeholder.jpg').trim(),
    });
  }
  return rows;
}

/**
 * Réécrit toutes les lignes du sheet à partir d'un tableau de voitures.
 */
function _writeCarsToSheet(cars) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" introuvable.');

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  if (cars.length === 0) return;

  const rows = cars.map(c => [
    c.name        || '',
    c.price       || 0,
    c.duration    || 'Per Day',
    c.seats       || 4,
    c.transmission|| 'Automatic',
    c.doors       || 4,
    c.fuel        || 'Petrol',
    c.image       || '/images/car-placeholder.jpg',
  ]);

  sheet.getRange(2, 1, rows.length, COLUMNS.length).setValues(rows);
}

/**
 * Construit une réponse JSON pour le client.
 */
function _jsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * Vérifie le secret PWA envoyé par le client.
 * Le secret peut venir :
 *  - en paramètre GET  : ?pwa_secret=xxx
 *  - dans le body POST : { "pwa_secret": "xxx", ... }
 */
function _auth(paramSecret, bodySecret) {
  const stored = PropertiesService.getScriptProperties().getProperty('PWA_SECRET');
  if (!stored) return true; // Pas de secret configuré = accès libre (non recommandé)
  return (paramSecret === stored) || (bodySecret === stored);
}

// ── Handler GET ───────────────────────────────────────────────────────────────
// Actions disponibles via GET :
//   ?action=list          → retourne toutes les voitures
//   ?action=summary       → retourne les statistiques du catalogue
//   ?action=sync          → déclenche la sync manuelle vers GitHub
//   ?payload=<base64>     → mutations PWA (add, update, delete, clear, reorder)
//                           encodées en base64 pour éviter les problèmes CORS POST
// ─────────────────────────────────────────────────────────────────────────────
function doGet(e) {
  const paramSecret = e.parameter.pwa_secret || '';

  // ── Mutations via payload base64 (remplace doPost pour la PWA) ──
  if (e.parameter.payload) {
    let body = {};
    try {
      const decoded = decodeURIComponent(escape(Utilities.newBlob(Utilities.base64Decode(e.parameter.payload)).getDataAsString()));
      body = JSON.parse(decoded);
    } catch (err) {
      return _jsonResponse({ ok: false, error: 'Payload base64 invalide : ' + err.message });
    }

    const bodySecret = body.pwa_secret || '';
    if (!_auth(paramSecret, bodySecret)) {
      return _jsonResponse({ ok: false, error: 'Non autorisé — secret invalide.' });
    }

    const action = body.action || '';

    try {
      // ── Ajouter une voiture ──────────────────────────────────────
      if (action === 'createReservation') {
        const result = createReservation(body.reservation);
        return _jsonResponse(result);
      }

      if (action === 'updateReservationStatus') {
        const result = updateReservationStatus(body.rowNumber, body.status);
        return _jsonResponse(result);
      }

      if (action === 'checkAvailability') {
        const result = checkReservationAvailability(body.reservation || body);
        return _jsonResponse(result);
      }

      if (action === 'setupReservations') {
        setupReservationsSheet();
        return _jsonResponse({ ok: true, message: 'Feuille Reservations configuree.' });
      }

      if (action === 'add') {
        const car = body.car;
        if (!car || !car.name) return _jsonResponse({ ok: false, error: 'Champ car.name manquant.' });
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
        if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });
        sheet.appendRow([
          car.name         || '',
          car.price        || 0,
          car.duration     || 'Per Day',
          car.seats        || 4,
          car.transmission || 'Automatic',
          car.doors        || 4,
          car.fuel         || 'Petrol',
          car.image        || '/images/car-placeholder.jpg',
        ]);
        onSheetChange({ changeType: 'EDIT' });
        return _jsonResponse({ ok: true, message: 'Voiture ajoutée et synchronisée.' });
      }

      // ── Modifier une voiture ─────────────────────────────────────
      if (action === 'update') {
        const car = body.car;
        if (!car || !car.id) return _jsonResponse({ ok: false, error: 'Champ car.id manquant.' });
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
        if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });
        const rowNum = parseInt(car.id, 10) + 1;
        sheet.getRange(rowNum, 1, 1, COLUMNS.length).setValues([[
          car.name         || '',
          car.price        || 0,
          car.duration     || 'Per Day',
          car.seats        || 4,
          car.transmission || 'Automatic',
          car.doors        || 4,
          car.fuel         || 'Petrol',
          car.image        || '/images/car-placeholder.jpg',
        ]]);
        onSheetChange({ changeType: 'EDIT' });
        return _jsonResponse({ ok: true, message: 'Voiture modifiée et synchronisée.' });
      }

      // ── Supprimer une voiture ────────────────────────────────────
      if (action === 'delete') {
        const id = body.id;
        if (!id) return _jsonResponse({ ok: false, error: 'Champ id manquant.' });
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
        if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });
        sheet.deleteRow(parseInt(id, 10) + 1);
        onSheetChange({ changeType: 'EDIT' });
        return _jsonResponse({ ok: true, message: 'Voiture supprimée et synchronisée.' });
      }

      // ── Vider tout le catalogue ──────────────────────────────────
      if (action === 'clear') {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
        if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
        onSheetChange({ changeType: 'EDIT' });
        return _jsonResponse({ ok: true, message: 'Catalogue vidé et synchronisé.' });
      }

      // ── Réordonner les voitures ──────────────────────────────────
      if (action === 'reorder') {
        const cars = body.cars;
        if (!Array.isArray(cars)) return _jsonResponse({ ok: false, error: 'Champ cars (tableau) manquant.' });
        _writeCarsToSheet(cars);
        onSheetChange({ changeType: 'EDIT' });
        return _jsonResponse({ ok: true, message: 'Voitures réordonnées et synchronisées.' });
      }

      return _jsonResponse({ ok: false, error: 'Action inconnue : ' + action });

    } catch (err) {
      Logger.log('[payload action error] ' + err.message);
      return _jsonResponse({ ok: false, error: err.message });
    }
  }

  // ── Requêtes GET classiques (lecture) ────────────────────────────
  if (!_auth(paramSecret, '')) {
    return _jsonResponse({ ok: false, error: 'Non autorisé — secret invalide.' });
  }

  const action = e.parameter.action || 'list';

  try {

    // ── Lister les voitures ──────────────────────────────────────
    if (action === 'list') {
      const cars = _getCarsFromSheet();
      return _jsonResponse({ ok: true, cars: cars });
    }

    if (action === 'listReservations') {
      const reservations = listReservations();
      return _jsonResponse({ ok: true, reservations: reservations });
    }

    if (action === 'checkAvailability') {
      const result = checkReservationAvailability(e.parameter);
      return _jsonResponse(result);
    }

    // ── Résumé / statistiques ────────────────────────────────────
    if (action === 'summary') {
      const cars   = _getCarsFromSheet();
      const prices = cars.map(c => c.price).filter(p => p > 0);
      const total  = cars.length;
      const avg    = prices.length ? Math.round(prices.reduce((a, b) => a + b, 0) / prices.length) : 0;
      const min    = prices.length ? Math.min(...prices) : 0;
      const max    = prices.length ? Math.max(...prices) : 0;

      const fuels = {};
      const transmissions = {};
      cars.forEach(c => {
        fuels[c.fuel]                = (fuels[c.fuel]                || 0) + 1;
        transmissions[c.transmission]= (transmissions[c.transmission]|| 0) + 1;
      });

      return _jsonResponse({
        ok: true,
        total: total,
        avgPrice: avg,
        minPrice: min,
        maxPrice: max,
        fuels: fuels,
        transmissions: transmissions,
      });
    }

    // ── Sync manuelle vers GitHub ────────────────────────────────
    if (action === 'setupReservations') {
      setupReservationsSheet();
      return _jsonResponse({ ok: true, message: 'Feuille Reservations configuree.' });
    }

    if (action === 'sync') {
      onSheetChange({ changeType: 'EDIT' });
      return _jsonResponse({ ok: true, message: 'Synchronisé vers GitHub avec succès.' });
    }

    return _jsonResponse({ ok: false, error: 'Action inconnue : ' + action });

  } catch (err) {
    Logger.log('[doGet error] ' + err.message);
    return _jsonResponse({ ok: false, error: err.message });
  }
}

// ── Handler POST ──────────────────────────────────────────────────────────────
// Actions disponibles via POST (body JSON) :
//   { action: "add",     pwa_secret: "...", car: { name, price, ... } }
//   { action: "update",  pwa_secret: "...", car: { id, name, price, ... } }
//   { action: "delete",  pwa_secret: "...", id: 3 }
//   { action: "clear",   pwa_secret: "..." }
//   { action: "reorder", pwa_secret: "...", cars: [ {...}, {...} ] }
// ─────────────────────────────────────────────────────────────────────────────
function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return _jsonResponse({ ok: false, error: 'Corps JSON invalide.' });
  }

  const bodySecret  = body.pwa_secret || '';
  const paramSecret = (e.parameter && e.parameter.pwa_secret) || '';
  if (!_auth(paramSecret, bodySecret)) {
    return _jsonResponse({ ok: false, error: 'Non autorisé — secret invalide.' });
  }

  const action = body.action || '';

  try {

    // ── Ajouter une voiture ──────────────────────────────────────
    if (action === 'createReservation') {
      const result = createReservation(body.reservation);
      return _jsonResponse(result);
    }

    if (action === 'updateReservationStatus') {
      const result = updateReservationStatus(body.rowNumber, body.status);
      return _jsonResponse(result);
    }

    if (action === 'checkAvailability') {
      const result = checkReservationAvailability(body.reservation || body);
      return _jsonResponse(result);
    }

    if (action === 'setupReservations') {
      setupReservationsSheet();
      return _jsonResponse({ ok: true, message: 'Feuille Reservations configuree.' });
    }

    if (action === 'add') {
      const car = body.car;
      if (!car || !car.name) return _jsonResponse({ ok: false, error: 'Champ car.name manquant.' });

      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });

      sheet.appendRow([
        car.name         || '',
        car.price        || 0,
        car.duration     || 'Per Day',
        car.seats        || 4,
        car.transmission || 'Automatic',
        car.doors        || 4,
        car.fuel         || 'Petrol',
        car.image        || '/images/car-placeholder.jpg',
      ]);

      onSheetChange({ changeType: 'EDIT' });
      return _jsonResponse({ ok: true, message: 'Voiture ajoutée et synchronisée.' });
    }

    // ── Modifier une voiture (par id = numéro de ligne - 1) ──────
    if (action === 'update') {
      const car = body.car;
      if (!car || !car.id) return _jsonResponse({ ok: false, error: 'Champ car.id manquant.' });

      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });

      const rowNum = parseInt(car.id, 10) + 1; // id=1 → ligne 2 (après l'en-tête)
      sheet.getRange(rowNum, 1, 1, COLUMNS.length).setValues([[
        car.name         || '',
        car.price        || 0,
        car.duration     || 'Per Day',
        car.seats        || 4,
        car.transmission || 'Automatic',
        car.doors        || 4,
        car.fuel         || 'Petrol',
        car.image        || '/images/car-placeholder.jpg',
      ]]);

      onSheetChange({ changeType: 'EDIT' });
      return _jsonResponse({ ok: true, message: 'Voiture modifiée et synchronisée.' });
    }

    // ── Supprimer une voiture (par id) ───────────────────────────
    if (action === 'delete') {
      const id = body.id;
      if (!id) return _jsonResponse({ ok: false, error: 'Champ id manquant.' });

      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });

      const rowNum = parseInt(id, 10) + 1;
      sheet.deleteRow(rowNum);

      onSheetChange({ changeType: 'EDIT' });
      return _jsonResponse({ ok: true, message: 'Voiture supprimée et synchronisée.' });
    }

    // ── Vider tout le catalogue ──────────────────────────────────
    if (action === 'clear') {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
      if (!sheet) return _jsonResponse({ ok: false, error: 'Sheet introuvable.' });

      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

      onSheetChange({ changeType: 'EDIT' });
      return _jsonResponse({ ok: true, message: 'Catalogue vidé et synchronisé.' });
    }

    // ── Réordonner les voitures (envoyer le tableau complet réordonné) ──
    if (action === 'reorder') {
      const cars = body.cars;
      if (!Array.isArray(cars)) return _jsonResponse({ ok: false, error: 'Champ cars (tableau) manquant.' });

      _writeCarsToSheet(cars);
      onSheetChange({ changeType: 'EDIT' });
      return _jsonResponse({ ok: true, message: 'Voitures réordonnées et synchronisées.' });
    }

    return _jsonResponse({ ok: false, error: 'Action inconnue : ' + action });

  } catch (err) {
    Logger.log('[doPost error] ' + err.message);
    return _jsonResponse({ ok: false, error: err.message });
  }
}

// Reservations module.
function setupReservationsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RESERVATION_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RESERVATION_SHEET_NAME);
  }

  sheet.getRange(1, 1, 1, RESERVATION_COLUMNS.length).setValues([RESERVATION_COLUMNS]);
  sheet.setFrozenRows(1);

  const headerRange = sheet.getRange(1, 1, 1, RESERVATION_COLUMNS.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1E293B');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setHorizontalAlignment('center');

  const statusRange = sheet.getRange(2, 2, Math.max(sheet.getMaxRows() - 1, 1), 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(RESERVATION_STATUSES, true)
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(rule);

  sheet.autoResizeColumns(1, RESERVATION_COLUMNS.length);
  Logger.log('Reservations sheet is ready.');
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

  ScriptApp.newTrigger('onReservationEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  ui.alert('Trigger reservations installe. Les changements de statut enverront un email client.');
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
  const lastEmailStatus = String(sheet.getRange(rowNumber, 12).getValue() || '').trim();
  if (!status || status === lastEmailStatus) return;

  const reservation = _getReservationFromRow(sheet, rowNumber);
  if (!reservation.email) {
    Logger.log('Reservation row ' + rowNumber + ' has no email.');
    return;
  }

  _sendReservationStatusEmail(reservation);
  sheet.getRange(rowNumber, 12).setValue(status);
  _updateReservationCalendarEvent(reservation);
}

function listReservations() {
  setupReservationsSheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESERVATION_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const reservations = [];
  for (let rowNumber = 2; rowNumber <= lastRow; rowNumber++) {
    const reservation = _getReservationFromRow(sheet, rowNumber);
    if (!reservation.name && !reservation.email && !reservation.carName) continue;
    reservations.push(_serializeReservation(reservation));
  }

  reservations.sort(function(a, b) {
    return String(b.createdAt || '').localeCompare(String(a.createdAt || ''));
  });
  Logger.log('[reservations list] ' + reservations.length + ' reservation(s).');
  return reservations;
}

function checkReservationAvailability(reservation) {
  _validateReservationDates(reservation);
  const carName = String(reservation.carName || '').trim();
  const start = _parseReservationDate(reservation.startDate);
  const end = _parseReservationDate(reservation.endDate);
  const ignoreRowNumber = parseInt(reservation.ignoreRowNumber || reservation.rowNumber || 0, 10);
  const conflicts = _findReservationConflicts(carName, start, end, ignoreRowNumber);
  return {
    ok: true,
    available: conflicts.length === 0,
    conflicts: conflicts.map(_serializeReservation),
  };
}

function updateReservationStatus(rowNumber, status) {
  setupReservationsSheet();
  rowNumber = parseInt(rowNumber, 10);
  status = String(status || '').trim();

  if (!rowNumber || rowNumber < 2) throw new Error('Numero de ligne reservation invalide.');
  if (RESERVATION_STATUSES.indexOf(status) === -1) throw new Error('Statut reservation invalide: ' + status);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESERVATION_SHEET_NAME);
  if (rowNumber > sheet.getLastRow()) throw new Error('Reservation introuvable.');

  const previous = _getReservationFromRow(sheet, rowNumber);
  sheet.getRange(rowNumber, 2).setValue(status);
  const reservation = _getReservationFromRow(sheet, rowNumber);

  if (reservation.email && previous.status !== status) {
    try {
      _sendReservationStatusEmail(reservation);
      sheet.getRange(rowNumber, 12).setValue(status);
    } catch (err) {
      Logger.log('[reservation status email] row ' + rowNumber + ': ' + err.message);
    }
  }

  _updateReservationCalendarEvent(reservation);
  Logger.log('[reservation status] row ' + rowNumber + ' -> ' + status);
  return { ok: true, reservation: _serializeReservation(reservation) };
}

function createReservation(reservation) {
  if (!reservation) throw new Error('Reservation manquante.');
  if (!reservation.name) throw new Error('Nom client manquant.');
  if (!reservation.email) throw new Error('Email client manquant.');
  if (!reservation.phone) throw new Error('Telephone client manquant.');
  if (!reservation.carName) throw new Error('Voiture manquante.');
  _validateReservationDates(reservation);

  setupReservationsSheet();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESERVATION_SHEET_NAME);
  const availability = checkReservationAvailability(reservation);
  if (!availability.available) {
    Logger.log('[reservation conflict] ' + JSON.stringify(availability.conflicts));
    return {
      ok: false,
      error: 'Vehicule indisponible pour ces dates.',
      conflicts: availability.conflicts,
    };
  }

  const status = 'En attente';
  const row = [
    new Date(),
    status,
    String(reservation.name || '').trim(),
    String(reservation.email || '').trim(),
    String(reservation.phone || '').trim(),
    String(reservation.carName || '').trim(),
    reservation.carPrice || '',
    reservation.carDuration || '',
    reservation.startDate || '',
    reservation.endDate || '',
    '',
    '',
    String(reservation.notes || '').trim(),
  ];

  sheet.appendRow(row);
  const rowNumber = sheet.getLastRow();
  const saved = _getReservationFromRow(sheet, rowNumber);
  const eventId = _createReservationCalendarEvent(saved);
  if (eventId) {
    sheet.getRange(rowNumber, 11).setValue(eventId);
    saved.calendarEventId = eventId;
  }

  try {
    _sendReservationStatusEmail(saved);
    sheet.getRange(rowNumber, 12).setValue(status);
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
    carPrice: values[6],
    carDuration: String(values[7] || '').trim(),
    startDate: values[8],
    endDate: values[9],
    calendarEventId: String(values[10] || '').trim(),
    notes: String(values[12] || '').trim(),
  };
}

function _serializeReservation(reservation) {
  return {
    rowNumber: reservation.rowNumber,
    createdAt: _dateToIsoDateTime(reservation.createdAt),
    status: reservation.status || 'En attente',
    name: reservation.name,
    email: reservation.email,
    phone: reservation.phone,
    carName: reservation.carName,
    carPrice: reservation.carPrice,
    carDuration: reservation.carDuration,
    startDate: _dateToIsoDate(reservation.startDate),
    endDate: _dateToIsoDate(reservation.endDate),
    calendarEventId: reservation.calendarEventId,
    notes: reservation.notes,
  };
}

function _dateToIsoDate(value) {
  const date = _parseReservationDate(value);
  if (!date) return String(value || '');
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _dateToIsoDateTime(value) {
  const date = _parseReservationDate(value);
  if (!date) return String(value || '');
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function _validateReservationDates(reservation) {
  if (!reservation) throw new Error('Reservation manquante.');
  if (!reservation.startDate || !reservation.endDate) throw new Error('Dates de reservation manquantes.');

  const start = _parseReservationDate(reservation.startDate);
  const end = _parseReservationDate(reservation.endDate);
  if (!start || !end) throw new Error('Format de date invalide.');
  if (end.getTime() < start.getTime()) throw new Error('La date fin doit etre apres la date debut.');
}

function _findReservationConflicts(carName, start, end, ignoreRowNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESERVATION_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2 || !carName || !start || !end) return [];

  const conflicts = [];
  for (let rowNumber = 2; rowNumber <= sheet.getLastRow(); rowNumber++) {
    if (rowNumber === ignoreRowNumber) continue;
    const current = _getReservationFromRow(sheet, rowNumber);
    if (current.carName !== carName) continue;
    if (['Refusee', 'Annulee', 'Terminee'].indexOf(current.status) !== -1) continue;

    const currentStart = _parseReservationDate(current.startDate);
    const currentEnd = _parseReservationDate(current.endDate);
    if (!currentStart || !currentEnd) continue;

    if (start.getTime() <= currentEnd.getTime() && end.getTime() >= currentStart.getTime()) {
      conflicts.push(current);
    }
  }
  return conflicts;
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
    'Voiture: ' + reservation.carName,
    'Prix: ' + reservation.carPrice + ' MAD / ' + reservation.carDuration,
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
      'Voiture: ' + reservation.carName,
      'Prix: ' + reservation.carPrice + ' MAD / ' + reservation.carDuration,
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
    'Confirmee': 'Votre reservation est confirmee. Notre equipe vous attendra aux dates prevues.',
    'Refusee': 'Nous sommes desoles, votre reservation ne peut pas etre confirmee pour ces dates.',
    'Annulee': 'Votre reservation a ete annulee. Vous pouvez nous contacter pour une nouvelle demande.',
    'Terminee': 'Votre reservation est marquee comme terminee. Merci pour votre confiance.',
  };
  const message = statusMessages[status] || 'Le statut de votre reservation a ete mis a jour.';
  const subject = 'Mise a jour de votre reservation - ' + reservation.carName;
  const htmlBody =
    '<div style="margin:0;padding:24px;background:#f6f7f9;font-family:Arial,sans-serif;color:#1f2937;">' +
      '<div style="max-width:640px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e5e7eb;">' +
        '<div style="background:#1E293B;color:#ffffff;padding:24px 28px;">' +
          '<h1 style="margin:0;font-size:22px;line-height:1.3;">Remons Car Rental</h1>' +
          '<p style="margin:8px 0 0;color:#cbd5e1;">Mise a jour de reservation</p>' +
        '</div>' +
        '<div style="padding:28px;">' +
          '<p style="font-size:16px;margin:0 0 16px;">Bonjour ' + _escapeHtml(reservation.name) + ',</p>' +
          '<p style="font-size:15px;line-height:1.7;margin:0 0 22px;">' + _escapeHtml(message) + '</p>' +
          '<div style="border:1px solid #e5e7eb;border-radius:10px;padding:18px;background:#fafafa;">' +
            '<p style="margin:0 0 10px;"><strong>Statut:</strong> ' + _escapeHtml(status) + '</p>' +
            '<p style="margin:0 0 10px;"><strong>Voiture:</strong> ' + _escapeHtml(reservation.carName) + '</p>' +
            '<p style="margin:0 0 10px;"><strong>Date de depart:</strong> ' + _formatReservationDate(reservation.startDate) + '</p>' +
            '<p style="margin:0;"><strong>Date de retour:</strong> ' + _formatReservationDate(reservation.endDate) + '</p>' +
          '</div>' +
          '<p style="font-size:14px;color:#64748b;line-height:1.6;margin:22px 0 0;">Pour toute question, repondez simplement a cet email ou contactez notre equipe.</p>' +
        '</div>' +
      '</div>' +
    '</div>';

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
