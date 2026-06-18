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
var SHEET_ID = 'REMPLACE_PAR_TON_SHEET_ID';
var SHEET_NAME = 'Réponses';

// ── Colonnes du tableau ──────────────────────────────────────────────────────
var COLONNES = [
  'Horodatage',
  // Période
  'Année', 'Mois', 'Semaine', 'Jour', 'Zone', 'AM', 'Nom du dépôt',
  // Ventes
  'Sell-In Plan (Fcfa)', 'Sell-In MTD (Fcfa)', 'Sell-In %',
  'Total Revendeurs', 'Revendeurs Actifs', 'Actifs/Jour',
  // Recrutement
  'Recrutement FY', 'Recrutement YTD', 'Recrutement Reste',
  'Semaines Consécutives',
  // Hotspot Écoles
  'Écoles Dispo', 'Écoles Actif', 'Écoles Assignés', 'Écoles Inactif',
  // Hotspot Marchés
  'Marchés Dispo', 'Marchés Actif', 'Marchés Assignés', 'Marchés Inactif',
  // Hotspot Carrefours
  'Carrefours Dispo', 'Carrefours Actif', 'Carrefours Assignés', 'Carrefours Inactif',
  // Gestion Revendeurs - Données
  'Ventes Target', 'Ventes Actuel', 'VMPS Plan', 'VMPS Actuel',
  // Scores Gestion Revendeurs
  'Score Taux Actifs Jour', 'Score VMPS', 'Score Recrutement',
  // Équipements
  'Equip Total', 'Equip Actif', 'Equip Plan', 'Equip Actuel',
  'Score Equip Propre', 'Score Zéro Inactif', 'Score Equip Sécurisé',
  // Chaîne du froid
  'Congé Total', 'Congé Actif', 'Congé Défectueux',
  'Score Congé Propre', 'Score Congé Accès', 'Score Retours',
  // Hygiène
  'Lavage Mains', 'Trousse Secours',
  'Score Propreté PDV', 'Score Drains', 'Score Poubelle',
  // Must Do's
  "MustDo Fieldpro", 'MustDo Maintenance', 'MustDo Navigation', 'MustDo Drains',
  // Résultat
  'Score Global', 'Mention',
  // Plan d'action
  'Actions Correctives',
  'Observations'
];

// ── Requête GET : test de vie ────────────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'API Fan Milk active ✅' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Requête POST : réception des données du formulaire ───────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    _ecrireLigne(data);
    return _reponse({ success: true, message: 'Données enregistrées ✅' });
  } catch (err) {
    return _reponse({ success: false, error: err.toString() });
  }
}

// ── Écriture dans le sheet ───────────────────────────────────────────────────
function _ecrireLigne(data) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  // Créer la feuille + en-têtes si elle n'existe pas encore
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLONNES);
    // Style en-têtes
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
    data.territoire || '', data.am || '', data.nom_depot || '',
    // Ventes
    data.sellin_plan || '', data.sellin_mtd || '', data.sellin_pct || '',
    data.rv_total || '', data.rv_actifs || '', data.rv_actifs_jour || '',
    // Recrutement
    data.recr_fy || '', data.recr_ytd || '', data.recr_reste || '',
    data.semaines_consecutives || '',
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
    data.vmps_plan || '', data.vmps_actuel || '',
    data.score_rv_actifs_jour || '', data.score_vmps || '',
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
    data.score_hygiene_proprete || '', data.score_drains || '',
    data.score_poubelle || '',
    // Must Do's
    data.mustdo_fieldpro || '', data.mustdo_maintenance || '',
    data.mustdo_navigation || '', data.mustdo_drains || '',
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
}

// ── Helper réponse JSON avec CORS ────────────────────────────────────────────
function _reponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
