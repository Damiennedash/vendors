/**
 * ╔══════════════════════════════════════════════╗
 * ║  Dépôt : fanmilk-scoreboard                 ║
 * ║  Fichier : backend_apps_script.gs            ║
 * ╚══════════════════════════════════════════════╝
 *
 * TABLEAU DE SCORE – FAN MILK / DANONE
 * Backend Google Apps Script → Google Sheets
 * ─────────────────────────────────────────────
 * INSTRUCTIONS DE DÉPLOIEMENT :
 *
 * 1. Va sur https://script.google.com → Nouveau projet
 * 2. Colle ce code entier (remplace tout)
 * 3. Clique sur "Déployer" → "Nouveau déploiement"
 * 4. Type : Application Web
 *    - Description : Tableau de Score Fan Milk
 *    - Exécuter en tant que : Moi
 *    - Qui peut accéder : Tout le monde
 * 5. Clique "Déployer" → Copie l'URL générée
 * 6. Colle cette URL dans le fichier HTML (variable APPS_SCRIPT_URL)
 */

// ID de ton Google Sheet (crée-en un nouveau vide, copie l'ID depuis l'URL)
// URL du sheet : https://docs.google.com/spreadsheets/d/CECI_EST_L_ID/edit
var SHEET_ID = '1w4xgyq97TmkgsmftM7yFmGL21_deg5azNgEnc-5FKEY';
var SHEET_NAME = 'Réponses';

// ── Colonnes du tableau ──────────────────────────────────────────────────────
var COLONNES = [
  'Horodatage',
  // Période
  'Année', 'Mois', 'Semaine', 'Jour', 'Zone', 'Sous-zone', 'AM', 'Nom du dépôt',
  // Ventes
  'Sell-In Plan (Fcfa)', 'Sell-In MTD (Fcfa)', 'Sell-In %',
  'Total Revendeurs', 'Revendeurs Actifs', 'Actifs/Jour',
  // Recrutement
  'Recrutement FY', 'Recrutement YTD', 'Recrutement Reste',
  'Score Semaine Antérieure',
  // Hotspot Écoles
  'Écoles Dispo', 'Écoles Actif', 'Écoles Assignés', 'Écoles Inactif',
  // Hotspot Marchés
  'Marchés Dispo', 'Marchés Actif', 'Marchés Assignés', 'Marchés Inactif',
  // Hotspot Carrefours
  'Carrefours Dispo', 'Carrefours Actif', 'Carrefours Assignés', 'Carrefours Inactif',
  // Gestion Revendeurs - Données
  'Ventes Target', 'Ventes Actuel', 'QMVP Plan', 'QMVP Actuel',
  // Scores Gestion Revendeurs
  'Score Taux Actifs Jour', 'Score QMVP', 'Score Recrutement',
  // Équipements
  'Equip Total', 'Equip Actif', 'Equip Plan', 'Equip Actuel',
  'Score Equip Propre', 'Score Zéro Inactif', 'Score Equip Sécurisé',
  // Chaîne du froid
  'Congé Total', 'Congé Actif', 'Congé Défectueux',
  'Score Congé Propre', 'Score Congé Accès', 'Score Retours',
  // Hygiène
  'Lavage Mains', 'Trousse Secours',
  'Score Propreté PDV', 'Score Sol Propreté', 'Score Poubelle',
  // Résultat
  'Score Global', 'Mention',
  // Plan d'action
  'Actions Correctives',
  'Observations'
];

// ── Requête GET : test de vie, ou récupération des données (?action=data) ───
function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : null;
  if (action === 'data') {
    return _obtenirDonnees();
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'API Fan Milk active ✅' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Lecture brute du sheet pour le dashboard ─────────────────────────────────
// Important : on lit l'en-tête RÉELLE de la ligne 1 du sheet (pas la constante
// COLONNES ci-dessus), pour rester valide même si les noms de colonnes du
// sheet diffèrent légèrement de ceux du script.
function _obtenirDonnees() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return _reponse({ success: true, header: [], rows: [] });

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return _reponse({ success: true, header: [], rows: [] });

    var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var header = values[0];
    var rows = values.slice(1);

    return _reponse({ success: true, header: header, rows: rows });
  } catch (err) {
    return _reponse({ success: false, error: err.toString() });
  }
}


// ── Requête POST : réception des données du formulaire ───────────────────────
var NOTIF_EMAIL = 'djatadamienne5@gmail.com';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    _ecrireLigne(data);
    _envoyerNotification(data);
    return _reponse({ success: true, message: 'Données enregistrées ✅' });
  } catch (err) {
    return _reponse({ success: false, error: err.toString() });
  }
}

// ── Notification par email à chaque nouvelle saisie ──────────────────────────
// Si l'envoi échoue (quota Gmail dépassé, etc.), ça n'empêche jamais
// l'enregistrement de la donnée dans le Sheet.
function _envoyerNotification(data) {
  try {
    var sujet = '📊 Nouvelle saisie sur le Dashboard — ' + (data.am || '?') + ' / ' + (data.nom_depot || '?');
    var corps =
      'Une nouvelle saisie vient d\'être enregistrée sur le Tableau de Score Fan Milk.\n' +
      'Le dashboard va s\'actualiser automatiquement avec cette donnée.\n\n' +
      'AM : ' + (data.am || '—') + '\n' +
      'Zone : ' + (data.territoire || '—') + '\n' +
      'Sous-zone : ' + (data.sous_zone || '—') + '\n' +
      'Dépôt : ' + (data.nom_depot || '—') + '\n' +
      'Date : ' + new Date().toLocaleString('fr-FR') + '\n' +
      'Score Global : ' + (data.score_global || '—') + ' / 60\n' +
      'Mention : ' + (data.mention || '—') + '\n' +
      'Sell-In MTD : ' + (data.sellin_mtd || '—') + ' Fcfa\n\n' +
      'Voir le dashboard : https://vendors-black.vercel.app/dashboard.html';
    MailApp.sendEmail(NOTIF_EMAIL, sujet, corps);
  } catch (err) {
    Logger.log('Erreur envoi notification email : ' + err.toString());
  }
}

// ── Écriture dans le sheet ───────────────────────────────────────────────────
function _ecrireLigne(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  // Créer la feuille si elle n'existe pas encore
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Vérifier l'en-tête à CHAQUE envoi : si la ligne 1 est vide ou incorrecte,
  // on la (re)crée automatiquement. Corrige le cas où l'en-tête a été
  // supprimée par erreur, sans avoir besoin d'intervenir manuellement.
  var lastRow = sheet.getLastRow();
  var premiereCase = lastRow > 0 ? sheet.getRange(1, 1).getValue() : '';
  if (premiereCase !== COLONNES[0]) {
    if (lastRow > 0) {
      sheet.insertRowBefore(1); // pousse les données existantes vers le bas sans les perdre
    }
    sheet.getRange(1, 1, 1, COLONNES.length).setValues([COLONNES]);
    var headerRange = sheet.getRange(1, 1, 1, COLONNES.length);
    headerRange.setBackground('#0A1E5E');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(10);
    sheet.setFrozenRows(1);
  }

  // Construire la ligne dans l'ordre des colonnes
  var ligne = [
    new Date().toLocaleString('fr-FR'),
    // Période
    data.annee || '', data.mois || '', data.semaine || '', data.jour || '',
    data.territoire || '', data.sous_zone || '', data.am || '', data.nom_depot || '',
    // Ventes
    data.sellin_plan || '', data.sellin_mtd || '', data.sellin_pct || '',
    data.rv_total || '', data.rv_actifs || '', data.rv_actifs_jour || '',
    // Recrutement
    data.recr_fy || '', data.recr_ytd || '', data.recr_reste || '',
    data.score_semaine_precedente || '',
    // Hotspot Écoles
    data.hs_ecole_dispo || '', data.hs_ecole_actif || '',
    data.hs_ecole_assign || '', data.hs_ecole_inactif || '',
    // Hotspot Marchés
    data.hs_marche_dispo || '', data.hs_marche_actif || '',
    data.hs_marche_assign || '', data.hs_marche_inactif || '',
    // Hotspot Carrefours
    data.hs_carrefour_dispo || '', data.hs_carrefour_actif || '',
    data.hs_carrefour_assign || '', data.hs_carrefour_inactif || '',
    // Gestion Revendeurs
    data.ventes_target || '', data.ventes_actuel || '',
    data.qmvp_plan || '', data.qmvp_actuel || '',
    data.score_rv_actifs_jour || '', data.score_qmvp || '',
    data.score_recrutement || '',
    // Équipements
    data.equip_total || '', data.equip_actif || '',
    data.equip_plan || '', data.equip_actuel || '',
    data.score_equip_propre || '', data.score_actif_zero || '',
    data.score_equip_securise || '',
    // Chaîne du froid
    data.conge_total || '', data.conge_actif || '', data.conge_defect || '',
    data.score_conge_propre || '', data.score_conge_acces || '',
    data.score_retours || '',
    // Hygiène
    data.hygiene_lavage || '', data.hygiene_trousse || '',
    data.score_hygiene_proprete || '', data.score_sol_proprete || '',
    data.score_poubelle || '',
    // Résultat
    data.score_global || '', data.mention || '',
    // Plan d'action
    data.plan_action || '', data.observations || ''
  ];

  sheet.appendRow(ligne);

  // Mise en forme alternée des lignes
  var lastRow = sheet.getLastRow();
  if (lastRow % 2 === 0) {
    sheet.getRange(lastRow, 1, 1, COLONNES.length).setBackground('#F4F7FC');
  }

  // Colorer Score Global + Mention selon Bon (vert) / Moyen (jaune) / Faible (rouge)
  _colorerScoreEtMention(sheet, lastRow, data.mention);
}

// ── Coloration conditionnelle Score Global / Mention ─────────────────────────
function _colorerScoreEtMention(sheet, ligneNum, mention) {
  var couleurs = {
    'Bon':    '#D9EAD3', // vert clair
    'Moyen':  '#FFF2CC', // jaune clair
    'Faible': '#F4CCCC'  // rouge clair
  };
  var bg = couleurs[mention];
  if (!bg) return;

  var colScore = COLONNES.indexOf('Score Global') + 1;
  var colMention = COLONNES.indexOf('Mention') + 1;
  if (colScore > 0) sheet.getRange(ligneNum, colScore).setBackground(bg);
  if (colMention > 0) sheet.getRange(ligneNum, colMention).setBackground(bg);
}

// ── Helper réponse JSON avec CORS ────────────────────────────────────────────
function _reponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
