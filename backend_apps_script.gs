/**
 * ╔══════════════════════════════════════════════╗
 * ║  Dépôt : fanmilk-scoreboard                 ║
 * ║  Fichier : backend_apps_script.gs            ║
 * ╚══════════════════════════════════════════════╝
 *
 * TABLEAU DE SCORE – FAN MILK / DANONE
 * Backend Google Apps Script → Google Sheets + Dashboard
 *
 * INSTRUCTIONS :
 * 1. Colle ce code dans Apps Script (script.google.com)
 * 2. Remplace SHEET_ID par l'ID de ton Google Sheet
 * 3. Déployer → Nouveau déploiement → Application Web
 *    - Exécuter en tant que : Moi
 *    - Qui peut accéder : Tout le monde
 * 4. Copie l'URL et colle-la dans le HTML (APPS_SCRIPT_URL)
 */

var SHEET_ID   = 'REMPLACE_PAR_TON_SHEET_ID';
var SHEET_DATA = 'Réponses';
var SHEET_DASH = 'Dashboard';
var SHEET_REF  = 'Ref_Depots';

// ── Colonnes de la feuille Réponses ─────────────────────────────────────────
var COLONNES = [
  'Horodatage', 'Annee', 'Mois', 'NumMois', 'Semaine', 'Jour', 'Date_Saisie',
  'AM', 'Zone', 'Nom_Depot',
  'SellinObjectif', 'SellinMTD', 'SellinPct',
  'Total_Revendeurs', 'Revendeurs_Actifs', 'Actifs_Jour',
  'Recr_FY', 'Recr_YTD', 'Recr_Reste', 'Semaines_Consecutives',
  'HS_Ecole_Dispo', 'HS_Ecole_Actif', 'HS_Ecole_Couverture', 'HS_Ecole_Inactif',
  'HS_Marche_Dispo', 'HS_Marche_Actif', 'HS_Marche_Couverture', 'HS_Marche_Inactif',
  'HS_Carrefour_Dispo', 'HS_Carrefour_Actif', 'HS_Carrefour_Couverture', 'HS_Carrefour_Inactif',
  'HS_Coins_Dispo', 'HS_Coins_Actif', 'HS_Coins_Couverture', 'HS_Coins_Inactif',
  'Ventes_Target', 'Ventes_Actuel', 'VMPS_Plan', 'VMPS_Actuel',
  'Score_Taux_Actifs', 'Score_VMPS', 'Score_Recrutement',
  'Equip_Total', 'Equip_Actif', 'Equip_Plan', 'Equip_Actuel',
  'Score_Equip_Propre', 'Score_Zero_Inactif', 'Score_Equip_Securise',
  'Conge_Total', 'Conge_Actif', 'Conge_Defectueux',
  'Score_Conge_Propre', 'Score_Conge_Acces', 'Score_Retours',
  'Hygiene_Lavage', 'Hygiene_Trousse',
  'Score_Proprete_PDV', 'Score_Drains', 'Score_Poubelle',
  'MustDo_Fieldpro', 'MustDo_Maintenance', 'MustDo_Navigation', 'MustDo_Drains',
  'Score_Global', 'Mention',
  'Plan_Action', 'Observations'
];

var MOIS_FR = ['','Janvier','Février','Mars','Avril','Mai','Juin',
               'Juillet','Août','Septembre','Octobre','Novembre','Décembre'];

// ── Mapping AM → Zone et AM → Dépôts ────────────────────────────────────────
var AM_ZONE = {
  'SYLVAIN':    'Agoé',
  'CLAUDE':     'Lome Centre',
  'GANDE LOME': 'BAGUIDA',
  'ALI':        'ANEHO-KPALIME',
  'GANDE INT':  'GRAND NORD'
};

// ── GET : test de vie ────────────────────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({status:'ok', message:'API Fan Milk Scoreboard ✅'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST : réception des données ─────────────────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.openById(SHEET_ID);

    // Initialiser la feuille Réponses
    var sheet = _initSheet(ss, SHEET_DATA, COLONNES);

    // Calculs auto côté serveur
    var now     = new Date();
    var numMois = parseInt(data.num_mois) || now.getMonth() + 1;
    var pct     = (data.sellin_plan && data.sellin_mtd)
                  ? ((parseFloat(data.sellin_mtd) / parseFloat(data.sellin_plan)) * 100).toFixed(1) + '%'
                  : data.sellin_pct || '';

    // Score global : additionner tous les scores si non fourni
    var scoreFields = ['score_rv_actifs_jour','score_vmps','score_recrutement',
      'score_equip_propre','score_actif_zero','score_equip_securise',
      'score_conge_propre','score_conge_acces','score_retours',
      'score_hygiene_proprete','score_drains','score_poubelle'];
    var scoreCalc = 0;
    scoreFields.forEach(function(f){ scoreCalc += parseInt(data[f]) || 0; });
    var scoreGlobal = parseInt(data.score_global) || scoreCalc;
    var mention = scoreGlobal < 30 ? 'Faible' : scoreGlobal < 45 ? 'Moyen' : 'Bon';

    // Couverture hotspot = actif/dispo * 100
    function couv(actif, dispo) {
      var a = parseInt(actif)||0, d = parseInt(dispo)||0;
      return d > 0 ? ((a/d)*100).toFixed(1)+'%' : '0%';
    }

    var ligne = [
      now.toLocaleString('fr-FR'),
      data.annee || now.getFullYear(),
      data.mois  || MOIS_FR[numMois],
      numMois,
      data.semaine || '',
      data.jour || now.getDate(),
      data.date_saisie || now.toLocaleDateString('fr-FR'),
      data.am || '',
      data.territoire || AM_ZONE[data.am] || '',
      data.nom_depot || '',
      data.sellin_plan || '',
      data.sellin_mtd  || '',
      pct,
      data.rv_total || '', data.rv_actifs || '', data.rv_actifs_jour || '',
      data.recr_fy || '', data.recr_ytd || '', data.recr_reste || '',
      data.semaines_consecutives || '',
      // Hotspot Écoles
      data.hs_ecole_dispo||'', data.hs_ecole_actif||'',
      couv(data.hs_ecole_actif, data.hs_ecole_dispo), data.hs_ecole_inactif||'',
      // Hotspot Marchés
      data.hs_marche_dispo||'', data.hs_marche_actif||'',
      couv(data.hs_marche_actif, data.hs_marche_dispo), data.hs_marche_inactif||'',
      // Hotspot Carrefours
      data.hs_carrefour_dispo||'', data.hs_carrefour_actif||'',
      couv(data.hs_carrefour_actif, data.hs_carrefour_dispo), data.hs_carrefour_inactif||'',
      // Hotspot Coins affluents
      data.hs_coins_dispo||'', data.hs_coins_actif||'',
      couv(data.hs_coins_actif, data.hs_coins_dispo), data.hs_coins_inactif||'',
      // Perf revendeurs
      data.ventes_target||'', data.ventes_actuel||'',
      data.vmps_plan||'', data.vmps_actuel||'',
      data.score_rv_actifs_jour||'', data.score_vmps||'', data.score_recrutement||'',
      // Équipements
      data.equip_total||'', data.equip_actif||'', data.equip_plan||'', data.equip_actuel||'',
      data.score_equip_propre||'', data.score_actif_zero||'', data.score_equip_securise||'',
      // Chaîne du froid
      data.conge_total||'', data.conge_actif||'', data.conge_defect||'',
      data.score_conge_propre||'', data.score_conge_acces||'', data.score_retours||'',
      // Hygiène
      data.hygiene_lavage||'', data.hygiene_trousse||'',
      data.score_hygiene_proprete||'', data.score_drains||'', data.score_poubelle||'',
      // Must Do's
      data.mustdo_fieldpro||'', data.mustdo_maintenance||'',
      data.mustdo_navigation||'', data.mustdo_drains||'',
      // Score
      scoreGlobal, mention,
      data.plan_action||'', data.observations||''
    ];

    sheet.appendRow(ligne);
    _styleLastRow(sheet, scoreGlobal);

    // Créer/mettre à jour la feuille Ref_Depots
    _initRefDepots(ss);

    // Créer le Dashboard si pas encore fait
    _initDashboard(ss);

    return _rep({success:true, score:scoreGlobal, mention:mention});
  } catch(err) {
    return _rep({success:false, error:err.toString()});
  }
}

// ── Init feuille données ──────────────────────────────────────────────────────
function _initSheet(ss, name, cols) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(cols);
    var h = sheet.getRange(1, 1, 1, cols.length);
    h.setBackground('#0A1E5E').setFontColor('#FFFFFF')
     .setFontWeight('bold').setFontSize(9).setWrap(true);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);
  }
  return sheet;
}

// ── Style dernière ligne selon score ─────────────────────────────────────────
function _styleLastRow(sheet, score) {
  var lr = sheet.getLastRow();
  if (lr < 2) return;
  // Couleur alternée
  if (lr % 2 === 0) sheet.getRange(lr,1,1,COLONNES.length).setBackground('#F4F7FC');
  // Colorier cellule Score Global (col BH = index 63)
  var scoreCol = COLONNES.indexOf('Score_Global') + 1;
  var cell = sheet.getRange(lr, scoreCol);
  if (score < 30)      cell.setBackground('#FFCCCC').setFontColor('#C0392B').setFontWeight('bold');
  else if (score < 45) cell.setBackground('#FFF3CD').setFontColor('#856404').setFontWeight('bold');
  else                 cell.setBackground('#D4EDDA').setFontColor('#155724').setFontWeight('bold');
}

// ── Feuille Ref_Depots ────────────────────────────────────────────────────────
function _initRefDepots(ss) {
  var sheet = ss.getSheetByName(SHEET_REF);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_REF);
    var data = [
      ['CODE','DEPOT','AM','ZONE'],
      ['C0046','Ets Bon Choix','SYLVAIN','Agoé'],
      ['C00199','PERSEVERANCE','SYLVAIN','Agoé'],
      ['C00255','YELADOU','SYLVAIN','Agoé'],
      ['C00299','EGK','SYLVAIN','Agoé'],
      ['C00204','PROSPER','SYLVAIN','Agoé'],
      ['C00059','DIEU POURVOIT','SYLVAIN','Agoé'],
      ['C02065','GRANDE KLASS','SYLVAIN','Agoé'],
      ['C02357','Ets DREAM CENTER','SYLVAIN','Agoé'],
      ['C02124','MORNING SUN','SYLVAIN','Agoé'],
      ['C00214','Ets ROSE FAITH (Ex D C A)','SYLVAIN','Agoé'],
      ['C02000','Ets ENOCK ET FRERE','SYLVAIN','Agoé'],
      ['C02302','LET US PRAISE THE LORD','SYLVAIN','Agoé'],
      ['C02406','ETS GRACIAS NONO','SYLVAIN','Agoé'],
      ['C00237','Ets SUPER DEPOT','SYLVAIN','Agoé'],
      ['C02358','ETS CLECLES 19','SYLVAIN','Agoé'],
      ['C01923','GRANDEUR DE DIEU','SYLVAIN','Agoé'],
      ['C01946','Ets BABA AZAAD','SYLVAIN','Agoé'],
      ['C00158','Ets Mamaline Esso','SYLVAIN','Agoé'],
      ['C00203','POINT D APPUI','SYLVAIN','Agoé'],
      ['C02367','ETS LOMAND','SYLVAIN','Agoé'],
      ['C02373','ST MARTIN','SYLVAIN','Agoé'],
      ['C00005','Ets ABIASS','SYLVAIN','Agoé'],
      ['C00149','Ets Les Saints Archanges','SYLVAIN','Agoé'],
      ['C01924','Ets PROVIDENCE DE DIEU','SYLVAIN','Agoé'],
      ['C01667','Ets CHEZ ROBERT','SYLVAIN','Agoé'],
      ['C00144','Ets LDG','SYLVAIN','Agoé'],
      ['C02403','PLAN DU SEIGNEUR B','SYLVAIN','Agoé'],
      ['C00051','Ets les CANTIQUES ex prov zossimé','SYLVAIN','Agoé'],
      ['C00055','Ets COMPASSION','SYLVAIN','Agoé'],
      ['C02293','PLAN DU SEIGNEUR','SYLVAIN','Agoé'],
      ['C02377','JESUS-CHRIST M A RESTAURE','SYLVAIN','Agoé'],
      ['C02384','DIHEN','SYLVAIN','Agoé'],
      ['C00008','ETS ADJOBI-NEVA','SYLVAIN','Agoé'],
      ['C02407','ETS IREBE AMOUR','SYLVAIN','Agoé'],
      ['C00130','Ets KAYAK','CLAUDE','Lome Centre'],
      ['C02376','ETS NBUKE RAMCO','CLAUDE','Lome Centre'],
      ['C00178','Ets NBUKE','CLAUDE','Lome Centre'],
      ['C00244','SK(BENZ CITY)','CLAUDE','Lome Centre'],
      ['C00091','Ets FUNOKO','CLAUDE','Lome Centre'],
      ['C00157','Ets La Majuscule','CLAUDE','Lome Centre'],
      ['C00241','Ets le Temps de DIEU','CLAUDE','Lome Centre'],
      ['C00096','Ets GISPA','CLAUDE','Lome Centre'],
      ['C00151','Ets LETY','CLAUDE','Lome Centre'],
      ['C00153','Ets Logoti','CLAUDE','Lome Centre'],
      ['C00233','Ets Sovi','CLAUDE','Lome Centre'],
      ['C02121','Ets Audit vision','CLAUDE','Lome Centre'],
      ['C00208','Ets La Régénération','CLAUDE','Lome Centre'],
      ['C00074','Ets ESMA','CLAUDE','Lome Centre'],
      ['C00240','ETS TEMIDE','CLAUDE','Lome Centre'],
      ['C02392','ETS SODOBEL','CLAUDE','Lome Centre'],
      ['C02408','ETS RAHINA SHOP','CLAUDE','Lome Centre'],
      ['C00126','Jsk La grace B','GANDE LOME','BAGUIDA'],
      ['C00242','Ets Toto Délali','GANDE LOME','BAGUIDA'],
      ['C02393','LUMIERE DE L UNIVERS','GANDE LOME','BAGUIDA'],
      ['C00019','Ets APRODI','GANDE LOME','BAGUIDA'],
      ['C00226','Ets SEVERINE ET FILS','GANDE LOME','BAGUIDA'],
      ['C00095','Ets GERM DOSSEH','GANDE LOME','BAGUIDA'],
      ['C00021','Ets Asilewo & Fils','GANDE LOME','BAGUIDA'],
      ['C02292','ETS GLORIA DEI','GANDE LOME','BAGUIDA'],
      ['C00294','ETS Demas AGBODRAFO','GANDE LOME','BAGUIDA'],
      ['C01913','ETS GOD PLAN','GANDE LOME','BAGUIDA'],
      ['C01996','ETS YEHONAM','GANDE LOME','BAGUIDA'],
      ['C00026','Ets AYEGNANOU ET FILS','GANDE LOME','BAGUIDA'],
      ['C02391','NADONIELA','GANDE LOME','BAGUIDA'],
      ['C00133','Ets KDEB-Sepopo ex St Ant','GANDE LOME','BAGUIDA'],
      ['C01995','RACH MACHALLAH','GANDE LOME','BAGUIDA'],
      ['C02391','ETS NADONIELLA','GANDE LOME','BAGUIDA'],
      ['C01735','JACQUES VILLE','GANDE LOME','BAGUIDA'],
      ['C02344','PATISSERIE MANNE DORE','GANDE LOME','BAGUIDA'],
      ['C00318','Ets GAMESSOU','ALI','ANEHO-KPALIME'],
      ['C01943','ETS VALLEE-CREDO','ALI','ANEHO-KPALIME'],
      ['C00276','Ets AYATOU','ALI','ANEHO-KPALIME'],
      ['C02332','IL EST TEMPS','ALI','ANEHO-KPALIME'],
      ['C00377','Ets YELLO 2000','ALI','ANEHO-KPALIME'],
      ['C02371','ETS SI CE N EST DIEU','ALI','ANEHO-KPALIME'],
      ['C02167','Ets EDZESKO','ALI','ANEHO-KPALIME'],
      ['C01876','Ets CAMARADI','ALI','ANEHO-KPALIME'],
      ['C02390','ETS MAWUGNON EGNON','ALI','ANEHO-KPALIME'],
      ['C02389','ETS DODANE','ALI','ANEHO-KPALIME'],
      ['C01982','Ets DAKOWA','ALI','ANEHO-KPALIME'],
      ['C00090','Ets FREU DICH','ALI','ANEHO-KPALIME'],
      ['C02366','ETS NASSARAWA','ALI','ANEHO-KPALIME'],
      ['C02360','ETS LE RESPECT DE WAHALA','ALI','ANEHO-KPALIME'],
      ['C01942','Ets MARCEL A DIEU SEUL LA GLOIRE','ALI','ANEHO-KPALIME'],
      ['C00330','Ets LA VERITE','ALI','ANEHO-KPALIME'],
      ['C00277','Ets AZONKO','ALI','ANEHO-KPALIME'],
      ['C00332','Ets LE MOMENT','ALI','ANEHO-KPALIME'],
      ['C00345','Ets Olive Barth','ALI','ANEHO-KPALIME'],
      ['C00328','Ets La FIESTA','ALI','ANEHO-KPALIME'],
      ['C00342','Ets MIVA','ALI','ANEHO-KPALIME'],
      ['C00350','Ets Pat-Akpe','ALI','ANEHO-KPALIME'],
      ['C01990','Ets ROCIS GROUPE','ALI','ANEHO-KPALIME'],
      ['C02378','Ets CND-BTP','ALI','ANEHO-KPALIME'],
      ['C00323','ETS JHK ET FRERES','ALI','ANEHO-KPALIME'],
      ['C00358','SEIDOU ET FILS A','GANDE INT','GRAND NORD'],
      ['C00358','SEIDOU ET FILS B','GANDE INT','GRAND NORD'],
      ['C00263','AFANA','GANDE INT','GRAND NORD'],
      ['C00374','VIE ET SANTE','GANDE INT','GRAND NORD'],
      ['C00373','VICTORIANUS','GANDE INT','GRAND NORD'],
      ['C00336','MANJOEL','GANDE INT','GRAND NORD'],
      ['C00337','MANJOEL B','GANDE INT','GRAND NORD'],
      ['C02286','AUX ELEGANTS ALI','GANDE INT','GRAND NORD'],
      ['C02369','TCHASSAMA','GANDE INT','GRAND NORD'],
      ['C00329','GRACE DIVINE KANTE 2','GANDE INT','GRAND NORD'],
      ['C00348','OUGADJA KOUKA','GANDE INT','GRAND NORD'],
      ['C02209','SAINTE GLADYS','GANDE INT','GRAND NORD'],
      ['C00329','GRACE DIVINE CINKASSE','GANDE INT','GRAND NORD'],
      ['C00329','GRACE DIVINE MANGO','GANDE INT','GRAND NORD'],
      ['C00266','AU SECOURS NIAMTOUGOU','GANDE INT','GRAND NORD'],
      ['C00322','LE JARDIN BASSAR','GANDE INT','GRAND NORD'],
      ['C00329','GRACE DIVINE BARKOISSI','GANDE INT','GRAND NORD'],
      ['C02345','DRISS DELICES','GANDE INT','GRAND NORD'],
      ['C02401','JOJO LE MEILLEUR','GANDE INT','GRAND NORD'],
      ['C02400','AZIA CAHIER','GANDE INT','GRAND NORD'],
      ['C02402','KHR SERVICE','GANDE INT','GRAND NORD'],
      ['C02398','JOJO LE MEILLEUR KOUKA','GANDE INT','GRAND NORD'],
      ['C02410','BI-IZNILLAH','GANDE INT','GRAND NORD'],
      ['C02411','BI-IZNILLAH MANGO','GANDE INT','GRAND NORD'],
      ['C02412','BI-IZNILLAH BARKOISI','GANDE INT','GRAND NORD'],
      ['C02413','LIBRE COMMUNICATION','GANDE INT','GRAND NORD']
    ];
    data.forEach(function(row){ sheet.appendRow(row); });
    var h = sheet.getRange(1,1,1,4);
    h.setBackground('#0A1E5E').setFontColor('#FFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Dashboard avec KPIs et formules ─────────────────────────────────────────
function _initDashboard(ss) {
  var dash = ss.getSheetByName(SHEET_DASH);
  if (dash) return; // déjà créé

  dash = ss.insertSheet(SHEET_DASH, 0); // en premier

  var NAVY='#0A1E5E', BLUE='#1A3A8F', ACCENT='#00AEEF',
      WHITE='#FFFFFF', OFFWHITE='#F4F7FC', GRAY='#D0D8E8',
      GREEN='#D4EDDA', YELLOW='#FFF3CD', RED='#FFCCCC',
      DARKTEXT='#0D1B3E';

  // Helpers
  function bg(r,c,nr,nc,col){ dash.getRange(r,c,nr,nc).setBackground(col); }
  function fc(r,c,nr,nc,col){ dash.getRange(r,c,nr,nc).setFontColor(col); }
  function fw(r,c,nr,nc){ dash.getRange(r,c,nr,nc).setFontWeight('bold'); }
  function fs(r,c,nr,nc,s){ dash.getRange(r,c,nr,nc).setFontSize(s); }
  function al(r,c,nr,nc,h){ dash.getRange(r,c,nr,nc).setHorizontalAlignment(h); }
  function va(r,c,nr,nc,v){ dash.getRange(r,c,nr,nc).setVerticalAlignment(v); }
  function wr(r,c,nr,nc){ dash.getRange(r,c,nr,nc).setWrap(true); }
  function merge(r,c,nr,nc){ dash.getRange(r,c,nr,nc).merge(); }
  function val(r,c,v){ dash.getRange(r,c).setValue(v); }
  function fmt(r,c,nr,nc,f){ dash.getRange(r,c,nr,nc).setNumberFormat(f); }
  function border(r,c,nr,nc){
    dash.getRange(r,c,nr,nc).setBorder(true,true,true,true,false,false,'#D0D8E8',SpreadsheetApp.BorderStyle.SOLID);
  }

  // Largeurs colonnes
  var colWidths = [8,130,90,90,90,90,90,90,90,90,8];
  colWidths.forEach(function(w,i){ dash.setColumnWidth(i+1, w); });

  // ── TITRE ──────────────────────────────────────────────────────────────────
  merge(1,1,1,11);
  val(1,1,'📊 DASHBOARD – TABLEAU DE SCORE FAN MILK');
  bg(1,1,1,11,NAVY); fc(1,1,1,11,WHITE); fw(1,1,1,11); fs(1,1,1,11,14);
  al(1,1,1,11,'center'); va(1,1,1,11,'middle'); dash.setRowHeight(1,42);

  // ── LIGNE FILTRES INFO ──────────────────────────────────────────────────────
  merge(2,1,1,11);
  val(2,1,'⚙️  Utilisez les Segments Google Sheets (Données > Ajouter un segment) sur les colonnes Zone, Mois, Dépôt de la feuille Réponses pour filtrer ce dashboard');
  bg(2,1,1,11,'#E8F4FD'); fc(2,1,1,11,'#1A3A8F'); fs(2,1,1,11,9); al(2,1,1,11,'center');
  dash.setRowHeight(2,28);

  // ── SECTION KPIs GLOBAUX ───────────────────────────────────────────────────
  merge(4,2,1,9);
  val(4,2,'🎯  KPIs GLOBAUX');
  bg(4,2,1,9,BLUE); fc(4,2,1,9,WHITE); fw(4,2,1,9); fs(4,2,1,9,11);
  al(4,2,1,9,'center'); dash.setRowHeight(4,32);

  // KPI Cards - ligne 5 (labels) + ligne 6 (valeurs)
  var kpiLabels = ['Score Moyen','Nb Soumissions','Score MAX','% Score ≥ 45 (Bon)','Revendeurs Total Moy','Actifs/Jour Moy'];
  var kpiCols   = [2,3,4,5,6,7];
  var kpiFormulas = [
    "=IFERROR(AVERAGE(Réponses!BO2:BO),\"—\")",
    "=IFERROR(COUNTA(Réponses!A2:A),\"—\")",
    "=IFERROR(MAX(Réponses!BO2:BO),\"—\")",
    "=IFERROR(COUNTIF(Réponses!BO2:BO,\">=\"&45)/COUNTA(Réponses!BO2:BO),\"—\")",
    "=IFERROR(AVERAGE(Réponses!N2:N),\"—\")",
    "=IFERROR(AVERAGE(Réponses!P2:P),\"—\")"
  ];
  var kpiFmts = ['0.0','0','0','0.0%','0.0','0.0'];

  kpiLabels.forEach(function(lbl,i){
    var col = kpiCols[i];
    val(5, col, lbl);
    bg(5,col,1,1,OFFWHITE); fc(5,col,1,1,'#5A6A88'); fs(5,col,1,1,8); fw(5,col,1,1);
    al(5,col,1,1,'center');

    dash.getRange(6,col).setFormula(kpiFormulas[i]);
    fmt(6,col,1,1, kpiFmts[i]);
    bg(6,col,1,1,WHITE); fc(6,col,1,1,NAVY); fs(6,col,1,1,18); fw(6,col,1,1);
    al(6,col,1,1,'center'); va(6,col,1,1,'middle');
    border(5,col,2,1);
  });
  dash.setRowHeight(5,22); dash.setRowHeight(6,44);

  // ── SECTION : SCORES PAR ZONE ──────────────────────────────────────────────
  var R = 8;
  merge(R,2,1,9); val(R,2,'📈  SCORES PAR ZONE');
  bg(R,2,1,9,BLUE); fc(R,2,1,9,WHITE); fw(R,2,1,9); fs(R,2,1,9,10);
  al(R,2,1,9,'center'); dash.setRowHeight(R,28);

  var zones = ['Agoé','Lome Centre','BAGUIDA','ANEHO-KPALIME','GRAND NORD'];
  var zoneAMs = ['SYLVAIN','CLAUDE','GANDE LOME','ALI','GANDE INT'];

  // En-têtes
  val(R+1,2,'Zone'); val(R+1,3,'AM'); val(R+1,4,'Score Moy');
  val(R+1,5,'Score Max'); val(R+1,6,'Nb Saisies'); val(R+1,7,'Objectif Moy');
  val(R+1,8,'MTD Moy'); val(R+1,9,'% Réal');
  bg(R+1,2,1,8,NAVY); fc(R+1,2,1,8,WHITE); fw(R+1,2,1,8); fs(R+1,2,1,8,9);

  zones.forEach(function(zone,i){
    var row = R+2+i;
    val(row,2, zone); val(row,3, zoneAMs[i]);
    dash.getRange(row,4).setFormula('=IFERROR(AVERAGEIF(Réponses!I:I,"'+ zone +'",Réponses!BO:BO),"—")');
    dash.getRange(row,5).setFormula('=IFERROR(MAXIFS(Réponses!BO:BO,Réponses!I:I,"'+ zone +'"),"—")');
    dash.getRange(row,6).setFormula('=IFERROR(COUNTIF(Réponses!I:I,"'+ zone +'"),"—")');
    dash.getRange(row,7).setFormula('=IFERROR(AVERAGEIF(Réponses!I:I,"'+ zone +'",Réponses!K:K),"—")');
    dash.getRange(row,8).setFormula('=IFERROR(AVERAGEIF(Réponses!I:I,"'+ zone +'",Réponses!L:L),"—")');
    dash.getRange(row,9).setFormula('=IFERROR(IF(D'+(row)+'="—","—",AVERAGEIF(Réponses!I:I,"'+ zone +'",Réponses!M:M)),"—")');
    bg(row,2,1,8, i%2===0 ? WHITE : OFFWHITE);
    fc(row,4,1,1,NAVY); fw(row,4,1,1); fs(row,4,1,1,11);
    al(row,4,1,6,'center');
    fmt(row,4,1,1,'0.0'); fmt(row,5,1,1,'0'); fmt(row,6,1,1,'0');
    fmt(row,7,1,2,'#,##0'); fmt(row,9,1,1,'0.0%');
    border(row,2,1,8);
    dash.setRowHeight(row,24);
  });

  // ── SECTION : RECRUTEMENT PAR ZONE ────────────────────────────────────────
  R = R + 2 + zones.length + 2;
  merge(R,2,1,9); val(R,2,'👥  RECRUTEMENT REVENDEURS PAR ZONE');
  bg(R,2,1,9,BLUE); fc(R,2,1,9,WHITE); fw(R,2,1,9); fs(R,2,1,9,10);
  al(R,2,1,9,'center'); dash.setRowHeight(R,28);

  val(R+1,2,'Zone'); val(R+1,3,'AM');
  val(R+1,4,'Target Mois Moy'); val(R+1,5,'MTD Recrut Moy'); val(R+1,6,'À Recruter Moy');
  val(R+1,7,'Total RV Moy'); val(R+1,8,'Actifs Moy'); val(R+1,9,'Actifs/Jour Moy');
  bg(R+1,2,1,8,NAVY); fc(R+1,2,1,8,WHITE); fw(R+1,2,1,8); fs(R+1,2,1,8,9);

  zones.forEach(function(zone,i){
    var row = R+2+i;
    val(row,2,zone); val(row,3,zoneAMs[i]);
    dash.getRange(row,4).setFormula('=IFERROR(AVERAGEIF(Réponses!I:I,"'+zone+'",Réponses!Q:Q),"—")');
    dash.getRange(row,5).setFormula('=IFERROR(AVERAGEIF(Réponses!I:I,"'+zone+'",Réponses!R:R),"—")');
    dash.getRange(row,6).setFormula('=IFERROR(AVERAGEIF(Réponses!I:I,"'+zone+'",Réponses!S:S),"—")');
    dash.getRange(row,7).setFormula('=IFERROR(AVERAGEIF(Réponses!I:I,"'+zone+'",Réponses!N:N),"—")');
    dash.getRange(row,8).setFormula('=IFERROR(AVERAGEIF(Réponses!I:I,"'+zone+'",Réponses!O:O),"—")');
    dash.getRange(row,9).setFormula('=IFERROR(AVERAGEIF(Réponses!I:I,"'+zone+'",Réponses!P:P),"—")');
    bg(row,2,1,8, i%2===0 ? WHITE : OFFWHITE);
    al(row,4,1,6,'center'); fmt(row,4,1,6,'0.0');
    border(row,2,1,8); dash.setRowHeight(row,24);
  });

  // ── SECTION : COUVERTURE HOTSPOTS PAR ZONE ────────────────────────────────
  R = R + 2 + zones.length + 2;
  merge(R,2,1,9); val(R,2,'📍  COUVERTURE HOTSPOTS PAR ZONE (%)');
  bg(R,2,1,9,BLUE); fc(R,2,1,9,WHITE); fw(R,2,1,9); fs(R,2,1,9,10);
  al(R,2,1,9,'center'); dash.setRowHeight(R,28);

  val(R+1,2,'Zone'); val(R+1,3,'AM');
  val(R+1,4,'Écoles %'); val(R+1,5,'Marchés %');
  val(R+1,6,'Carrefours %'); val(R+1,7,'Coins Affluents %');
  val(R+1,8,'Equip Total Moy'); val(R+1,9,'Equip Actifs Moy');
  bg(R+1,2,1,8,NAVY); fc(R+1,2,1,8,WHITE); fw(R+1,2,1,8); fs(R+1,2,1,8,9);

  // Colonnes couverture dans Réponses: W=Ecole, AA=Marche, AE=Carrefour, AI=Coins
  var couvCols = ['W','AA','AE','AI'];
  zones.forEach(function(zone,i){
    var row = R+2+i;
    val(row,2,zone); val(row,3,zoneAMs[i]);
    couvCols.forEach(function(col,j){
      dash.getRange(row,4+j).setFormula(
        '=IFERROR(AVERAGEIF(Réponses!I:I,"'+zone+'",Réponses!'+col+':'+col+'),"—")'
      );
      fmt(row,4+j,1,1,'0.0"%"');
    });
    dash.getRange(row,8).setFormula('=IFERROR(AVERAGEIF(Réponses!I:I,"'+zone+'",Réponses!AK:AK),"—")');
    dash.getRange(row,9).setFormula('=IFERROR(AVERAGEIF(Réponses!I:I,"'+zone+'",Réponses!AL:AL),"—")');
    bg(row,2,1,8, i%2===0 ? WHITE : OFFWHITE);
    al(row,4,1,6,'center'); fmt(row,8,1,2,'0.0');
    border(row,2,1,8); dash.setRowHeight(row,24);
  });

  // ── SECTION : ÉQUIPEMENTS PAR ZONE ────────────────────────────────────────
  R = R + 2 + zones.length + 2;
  merge(R,2,1,9); val(R,2,'🧊  ÉQUIPEMENTS (CONGÉLATEURS) PAR ZONE');
  bg(R,2,1,9,BLUE); fc(R,2,1,9,WHITE); fw(R,2,1,9); fs(R,2,1,9,10);
  al(R,2,1,9,'center'); dash.setRowHeight(R,28);

  val(R+1,2,'Zone'); val(R+1,3,'AM');
  val(R+1,4,'Congé Total Moy'); val(R+1,5,'Congé Actifs Moy');
  val(R+1,6,'Congé Inactifs Moy'); val(R+1,7,'Equip Total Moy');
  val(R+1,8,'Equip Actifs Moy'); val(R+1,9,'Equip Inactifs Moy');
  bg(R+1,2,1,8,NAVY); fc(R+1,2,1,8,WHITE); fw(R+1,2,1,8); fs(R+1,2,1,8,9);

  // Colonnes Réponses: AO=CongeTotal, AP=CongeActif, AQ=CongeDefect, AK=EquipTotal, AL=EquipActif
  zones.forEach(function(zone,i){
    var row = R+2+i;
    val(row,2,zone); val(row,3,zoneAMs[i]);
    [['AO',4],['AP',5],['AQ',6],['AK',7],['AL',8]].forEach(function(pair){
      dash.getRange(row,pair[1]).setFormula(
        '=IFERROR(AVERAGEIF(Réponses!I:I,"'+zone+'",Réponses!'+pair[0]+':'+pair[0]+'),"—")'
      );
      fmt(row,pair[1],1,1,'0.0');
    });
    // Equip inactifs calculé
    dash.getRange(row,9).setFormula('=IFERROR(G'+row+'-H'+row+',"—")');
    bg(row,2,1,8, i%2===0 ? WHITE : OFFWHITE);
    al(row,4,1,6,'center'); border(row,2,1,8); dash.setRowHeight(row,24);
  });

  // ── SECTION : TOP 10 DÉPÔTS (SCORE) ──────────────────────────────────────
  R = R + 2 + zones.length + 2;
  merge(R,2,1,9); val(R,2,'🏆  TOP 10 DÉPÔTS PAR SCORE MOYEN');
  bg(R,2,1,9,'#B8860B'); fc(R,2,1,9,WHITE); fw(R,2,1,9); fs(R,2,1,9,10);
  al(R,2,1,9,'center'); dash.setRowHeight(R,28);

  val(R+1,2,'#'); val(R+1,3,'Dépôt'); val(R+1,4,'Zone');
  val(R+1,5,'AM'); val(R+1,6,'Score Moy'); val(R+1,7,'Nb Saisies');
  bg(R+1,2,1,6,NAVY); fc(R+1,2,1,6,WHITE); fw(R+1,2,1,6); fs(R+1,2,1,6,9);

  // Note: Top 10 nécessite un tri dynamique - on utilise LARGE + INDEX/MATCH
  // Colonne helper cachée en col K (11) pour les scores par dépôt
  for (var top=1; top<=10; top++) {
    var tRow = R+1+top;
    val(tRow,2, top);
    // Pour top 10 dynamique, on liste les dépôts dans une zone helper
    dash.getRange(tRow,3).setFormula(
      '=IFERROR(INDEX(Réponses!J:J,MATCH(LARGE(Réponses!BO:BO,'+top+'),Réponses!BO:BO,0)),"—")'
    );
    dash.getRange(tRow,4).setFormula(
      '=IFERROR(INDEX(Réponses!I:I,MATCH(LARGE(Réponses!BO:BO,'+top+'),Réponses!BO:BO,0)),"—")'
    );
    dash.getRange(tRow,5).setFormula(
      '=IFERROR(INDEX(Réponses!H:H,MATCH(LARGE(Réponses!BO:BO,'+top+'),Réponses!BO:BO,0)),"—")'
    );
    dash.getRange(tRow,6).setFormula('=IFERROR(LARGE(Réponses!BO:BO,'+top+'),"—")');
    dash.getRange(tRow,7).setFormula(
      '=IFERROR(COUNTIF(Réponses!J:J,C'+tRow+'),"—")'
    );

    var bgColor = top === 1 ? '#FFD700' : top === 2 ? '#C0C0C0' : top === 3 ? '#CD7F32' : (tRow%2===0 ? WHITE : OFFWHITE);
    bg(tRow,2,1,6, bgColor);
    fw(tRow,6,1,1); fs(tRow,6,1,1,11);
    al(tRow,2,1,6,'center'); al(tRow,3,1,2,'left');
    fmt(tRow,6,1,1,'0');
    border(tRow,2,1,6); dash.setRowHeight(tRow,22);
  }

  // ── PIED DE PAGE ──────────────────────────────────────────────────────────
  var lastR = R + 12;
  merge(lastR,2,1,9);
  val(lastR,2,'fanmilk-scoreboard · Fan Milk / Danone · Dashboard auto-généré · Données : feuille Réponses');
  bg(lastR,2,1,9,NAVY); fc(lastR,2,1,9,GRAY); fs(lastR,2,1,9,8);
  al(lastR,2,1,9,'center'); dash.setRowHeight(lastR,24);

  // Mise en forme générale
  dash.setTabColor(ACCENT.replace('#',''));
  dash.getRange(1,1,lastR,11).setFontFamily('Arial');

  Logger.log('✅ Dashboard créé');
}

// ── Helper réponse JSON ───────────────────────────────────────────────────────
function _rep(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Fonction manuelle pour initialiser le dashboard sur un sheet existant ────
function initialiserDashboard() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  // Supprimer l'ancien dashboard si existant
  var old = ss.getSheetByName(SHEET_DASH);
  if (old) ss.deleteSheet(old);
  _initRefDepots(ss);
  _initDashboard(ss);
  Logger.log('✅ Dashboard initialisé !');
}
