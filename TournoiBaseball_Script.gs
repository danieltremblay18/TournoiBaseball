/**
 * ============================================================================
 *  TOURNOI BASEBALL 13U - BASEBALL QUÉBEC 2026
 *  Gestion complète d'un tournoi : 2 classes (A et B), 3 pools de 4 équipes,
 *  round-robin, bris d'égalité (Art. 42.11 Baseball Québec).
 * ============================================================================
 *
 *  INSTALLATION
 *  ------------
 *  1. Créer une nouvelle Google Sheet vierge (https://sheets.new).
 *  2. Menu  Extensions  >  Apps Script.
 *  3. Supprimer le contenu de Code.gs, puis coller TOUT ce fichier.
 *  4. Enregistrer (icône disquette), puis revenir à la feuille de calcul.
 *  5. Recharger la page de la feuille de calcul (F5).
 *  6. Un menu "🏆 Tournoi Baseball" apparaît dans la barre de menu.
 *  7. Cliquer "Initialiser les feuilles" (autoriser le script au besoin).
 *  8. Coller l'horaire dans la feuille "Configuration" : copier les données de
 *     l'onglet "Horaire globalArbitre" du fichier Excel du tournoi (produit par
 *     le comité du tournoi de l'ABMR) — lignes 2 et suivantes, sans l'en-tête —
 *     et les coller à partir de la cellule A2. Cette étape est la SEULE à refaire
 *     chaque année — le reste se régénère automatiquement.
 *  9. Cliquer "Générer les matchs" : répartit chaque match dans "Résultats A"
 *     ou "Résultats B" selon la colonne "# pool" (ex. "3A" = Pool 3, Classe A).
 * 10. Pour chaque partie jouée : indiquer l'"Équipe Locale" (inconnue à l'avance,
 *     c'est le registraire qui la précise au moment de saisir le score), saisir
 *     les scores, puis "Mettre à jour les classements".
 *
 *  COULEURS
 *  --------
 *  - Cellules à saisie manuelle : jaune clair.
 *  - Cellules calculées          : gris clair.
 *  - 1ers de pool                : vert.
 *  - 2es de pool                 : bleu clair.
 *
 * ============================================================================
 */

// ============================================================================
//  CONSTANTES GLOBALES
// ============================================================================

var SHEET_HELP        = 'Aide';
var SHEET_CONFIG      = 'Configuration';
var SHEET_RESULTS     = { 'A': 'Résultats A', 'B': 'Résultats B' };
var SHEET_STANDINGS   = { 'A': 'Classements A', 'B': 'Classements B' };
var SHEET_INN_DETAIL  = 'Manches_Détail';

var CLASSES = ['A', 'B'];
var POOLS   = [1, 2, 3];
var TEAMS_PER_POOL = 4;
var TOTAL_INNINGS  = 6;   // 13U = 6 manches réglementaires

// Couleurs
var COLOR_INPUT     = '#fff9c4';   // jaune clair  - saisie manuelle
var COLOR_CALC      = '#eeeeee';   // gris clair    - calculé
var COLOR_FIRST     = '#c8e6c9';   // vert          - 1er de pool
var COLOR_SECOND    = '#bbdefb';   // bleu clair    - 2e de pool
var COLOR_HEADER    = '#37474f';   // en-tête foncé
var COLOR_HEADER_TX = '#ffffff';   // texte en-tête
var COLOR_POOL_1    = '#ffe0b2';
var COLOR_POOL_2    = '#d1c4e9';
var COLOR_POOL_3    = '#b2dfdb';
var COLOR_SECTION   = '#cfd8dc';

// Les 6 matchs round-robin d'un pool de 4 équipes (indices 0..3). Le vrai horaire
// (Configuration) ne suit PAS cette matrice : il est collé directement depuis le
// fichier Excel du tournoi (produit par le comité de l'ABMR), peu importe l'ordre
// des matchs. Cette
// matrice ne sert plus qu'à dimensionner les feuilles (3 pools x 6 matchs = 18 par classe).
var GAME_MATRIX = [
  [0, 1],   // E1 vs E2
  [0, 2],   // E1 vs E3
  [0, 3],   // E1 vs E4
  [1, 2],   // E2 vs E3
  [1, 3],   // E2 vs E4
  [2, 3]    // E3 vs E4
];

// ============================================================================
//  MENU & DÉCLENCHEUR D'OUVERTURE
// ============================================================================

/**
 * Crée le menu personnalisé à l'ouverture de la feuille de calcul.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏆 Tournoi Baseball')
    .addItem('Initialiser les feuilles', 'createAllSheets')
    .addItem('Générer les matchs', 'generateGames')
    .addItem('Mettre à jour les classements', 'calculateStandings')
    .addSeparator()
    .addItem('⚡ Activer la mise à jour auto', 'installTriggers')
    .addItem('Effacer les résultats', 'clearResults')
    .addSeparator()
    .addItem('🧪 Simuler résultats de match', 'simulateMatchResults')
    .addToUi();
}

/**
 * Enregistre (ou ré-enregistre) le déclencheur installable qui met à jour les
 * classements en direct. À LANCER UNE FOIS par le propriétaire après chaque
 * collage du code (le déclencheur tourne alors avec l'autorisation du
 * propriétaire — LockService garanti — pour les éditions de tous les postes).
 *
 * Idempotent : supprime d'abord tout déclencheur existant pointant vers
 * handleResultEdit pour ne jamais créer de doublon (qui causerait un double
 * recalcul). Sans danger à cliquer plusieurs fois.
 */
function installTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Supprime les anciens déclencheurs de ce handler.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'handleResultEdit') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // (Re)crée le déclencheur installable onEdit.
  ScriptApp.newTrigger('handleResultEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    'Mise à jour automatique activée',
    'Les classements se mettront maintenant à jour automatiquement, en direct et ' +
    'pour tous les postes, dès qu\'une partie est entrée au complet.\n\n' +
    'À refaire une seule fois après chaque nouveau collage du code.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Handler de l'édition d'une feuille "Résultats". Met à jour le classement EN
 * DIRECT pour TOUS les postes (plusieurs personnes saisissent en parallèle,
 * d'autres regardent les onglets Classements).
 *
 * IMPORTANT — déclencheur INSTALLABLE, pas simple. Cette fonction ne s'appelle
 * volontairement PAS "onEdit" : un déclencheur simple (nom magique "onEdit")
 * tourne en mode d'autorisation restreint où LockService peut échouer, ce qui
 * casserait la protection multi-postes. On l'enregistre plutôt via
 * installTriggers() (menu "Activer la mise à jour auto") comme déclencheur
 * installable : il s'exécute avec l'autorisation du propriétaire (donc
 * LockService fonctionne) tout en se déclenchant pour les éditions de TOUS les
 * collaborateurs. À FAIRE UNE FOIS après chaque collage du code.
 *
 * Conçu pour un usage multi-postes fiable :
 *   - On ne réagit qu'aux éditions des colonnes de saisie H..M (8..13) d'une
 *     feuille "Résultats" (les colonnes A..G copiées et N..R calculées sont
 *     ignorées).
 *   - On ne recalcule QUE lorsque la ligne de la partie est COMPLÈTE
 *     (isRowComplete) : pas de classement qui bouge en pleine saisie, donc pas
 *     de confusion pour ceux qui regardent. Exception : l'effacement d'un score
 *     (qui RETIRE une partie du classement) déclenche aussi un rafraîchissement.
 *   - Tout le recalcul est protégé par un VERROU DE DOCUMENT : deux postes ne
 *     peuvent jamais reconstruire la feuille Classements en même temps (sinon un
 *     responsable verrait un classement à moitié reconstruit). Acquisition
 *     défensive : si LockService est indisponible, on recalcule sans verrou
 *     plutôt que de planter (dégradation gracieuse).
 *   - On n'écrit que dans Classements + les colonnes calculées N..R de la ligne
 *     saisie (écriture ciblée) — jamais toute la feuille Résultats, pour ne pas
 *     gêner la frappe d'un autre poste. Le menu "Mettre à jour les classements"
 *     reste le refresh complet/autoritaire (et le secours pour les corrections
 *     en lot).
 *
 * Les écritures programmées (N..R, Classements) ne redéclenchent pas le handler :
 * aucune boucle possible.
 */
function handleResultEdit(e) {
  if (!e || !e.range) { return; }

  var sheet = e.range.getSheet();
  var name = sheet.getName();

  // Quelle classe correspond à la feuille éditée ? (sinon on ignore)
  var classe = null;
  CLASSES.forEach(function (c) {
    if (SHEET_RESULTS[c] === name) { classe = c; }
  });
  if (!classe) { return; }

  // L'édition doit toucher au moins une colonne de saisie H..M (8..13).
  var startCol = e.range.getColumn();
  var endCol   = startCol + e.range.getNumColumns() - 1;
  if (endCol < 8 || startCol > 13) { return; }

  // Lignes éditées (ignore l'en-tête, ligne 1).
  var startRow = Math.max(e.range.getRow(), 2);
  var endRow   = e.range.getRow() + e.range.getNumRows() - 1;
  if (endRow < 2) { return; }

  // ---- Décide s'il faut recalculer (lecture légère, AVANT de prendre le verrou) ----
  var recalc = false;

  // (a) Au moins une ligne éditée est-elle désormais complète ?
  var hm = sheet.getRange(startRow, 8, endRow - startRow + 1, 6).getValues();
  for (var i = 0; i < hm.length; i++) {
    var r = hm[i];
    if (isRowComplete({ scoreA: r[0], scoreB: r[1], local: r[2],
                        manches: r[3], retraits: r[4], type: r[5] })) {
      recalc = true;
      break;
    }
  }

  // (b) Correction : un score (H ou I) vient d'être EFFACÉ sur une cellule unique.
  //     C'est la seule édition qui RETIRE une partie du classement (getGameResults
  //     ne compte une partie que si ses 2 scores sont présents) ; on rafraîchit
  //     pour la faire disparaître. (Les effacements EN LOT passent par le menu.)
  if (!recalc &&
      e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 &&
      (startCol === 8 || startCol === 9) &&
      (e.value === undefined || e.value === '') &&
      (e.oldValue !== undefined && e.oldValue !== '')) {
    recalc = true;
  }

  if (!recalc) { return; }

  // ---- Recalcul sous verrou de document ----
  // Acquisition défensive : si LockService est indisponible (ex. exécution en
  // mode restreint), lock reste null et on recalcule SANS verrou plutôt que de
  // planter. Si LockService est dispo mais le verrou occupé (waitLock dépasse le
  // délai), un autre poste recalcule déjà — son recalcul relit tout l'état, donc
  // on peut renoncer sans risque.
  var lock = null;
  try {
    lock = LockService.getDocumentLock();
  } catch (err) {
    lock = null;
  }
  if (lock) {
    try {
      lock.waitLock(10000);
    } catch (err) {
      return;   // verrou occupé : le poste qui le détient s'en occupe
    }
  }

  try {
    var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
    var games = getGameResults(classe);

    // Colonnes calculées N..R : seulement les lignes éditées (ciblé).
    var byRow = {};
    games.forEach(function (g) { byRow[g.rowIndex] = g; });
    for (var row = startRow; row <= endRow; row++) {
      writeRowCalc(sheet, row, byRow[row]);
    }

    buildStandingsSheet(ss, classe, games);
    ss.toast('Classement ' + classe + ' mis à jour.', 'Tournoi Baseball', 3);
  } finally {
    if (lock) { lock.releaseLock(); }
  }
}

// ============================================================================
//  CRÉATION DES FEUILLES
// ============================================================================

/**
 * Crée (ou réinitialise) toutes les feuilles nécessaires au tournoi.
 */
function createAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  createHelpSheet(ss);
  createConfigSheet(ss);
  CLASSES.forEach(function (c) { createResultsSheet(ss, c); });
  CLASSES.forEach(function (c) { createStandingsSheet(ss, c); });
  createInningDetailSheet(ss);

  // Supprime la feuille par défaut "Sheet1" / "Feuille1" si elle est vide et inutilisée.
  removeDefaultSheet(ss);

  // Réordonne les feuilles dans un ordre logique.
  reorderSheets(ss);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Feuilles initialisées. Collez l\'horaire dans "Configuration" puis "Générer les matchs".',
    'Tournoi Baseball', 6);
}

/**
 * Feuille Configuration : horaire complet du tournoi (36 matchs), au même
 * format que l'onglet "Horaire globalArbitre" du fichier Excel du tournoi
 * (produit chaque année par le comité du tournoi de l'ABMR). Il suffit de coller
 * les données de cet onglet ici (à partir de A2, sans la ligne d'en-tête) pour
 * que tout le système (Résultats A/B, classements) se régénère via "Générer les
 * matchs".
 */
function createConfigSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_CONFIG);
  sheet.clear();
  clearDataValidations(sheet);

  var headers = [
    '# de match', '# pool', 'Jour', 'Heure', 'Terrain', 'Équipe 1', 'Équipe 2',
    'Entré dans Spordle', 'Arbitre Marbre', 'Arbitre But', 'Marqueur', 'Compte Pitch'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet.getRange(1, 1, 1, headers.length));

  sheet.getRange(1, 1).setNote(
    'Collez ici les données de l\'onglet "Horaire globalArbitre" du fichier Excel ' +
    'du tournoi (produit par le comité du tournoi de l\'ABMR) — les lignes de matchs ' +
    'SEULEMENT, à partir de la ligne 2, la ligne d\'en-tête est déjà écrite ici, dans ' +
    'le même ordre de colonnes. Cliquez ensuite sur "Générer les matchs".');
  sheet.getRange(1, 2).setNote(
    'Pool + Classe combinés (ex. "3A" = Pool 3, Classe A ; "1B" = Pool 1, Classe B). ' +
    'Sert à répartir automatiquement chaque match dans l\'onglet "Résultats A" ou ' +
    '"Résultats B".');

  var widths = [70, 60, 90, 70, 70, 170, 170, 130, 130, 130, 110, 110];
  for (var i = 0; i < widths.length; i++) { sheet.setColumnWidth(i + 1, widths[i]); }
  sheet.setFrozenRows(1);

  // Prévoit large pour le copier-coller (36 matchs habituellement, 24 équipes).
  var nRows = 60;
  sheet.getRange(2, 1, nRows, 7).setBackground(COLOR_INPUT);    // A-G : zone de collage
  sheet.getRange(2, 8, nRows, 5).setBackground(COLOR_CALC);     // H-L : informatif seulement
  sheet.getRange(2, 3, nRows, 1).setNumberFormat('yyyy-mm-dd'); // Jour
  sheet.getRange(2, 4, nRows, 1).setNumberFormat('HH:mm');      // Heure
}

/**
 * Feuille Résultats (une par classe). 18 matchs (6 x 3 pools).
 */
function createResultsSheet(ss, classe) {
  var sheet = getOrCreateSheet(ss, SHEET_RESULTS[classe]);
  sheet.clear();
  clearDataValidations(sheet);

  var headers = [
    'Pool', 'Partie #', 'Jour', 'Heure', 'Terrain', 'Équipe 1', 'Équipe 2',
    'Score Équipe 1', 'Score Équipe 2', 'Équipe Locale', 'Manches complètes',
    'Retraits en fin', 'Type de fin', 'Gagnant', 'MO Éq.1', 'MD Éq.1',
    'MO Éq.2', 'MD Éq.2'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet.getRange(1, 1, 1, headers.length));

  // Info-bulles (survol de l'en-tête) pour les colonnes les plus difficiles à
  // comprendre. Le texte est autonome (pas besoin d'aller voir un autre onglet) ;
  // l'onglet "Aide" reste disponible pour une lecture plus complète au calme.
  var notes = {
    1:  'POOL — Numéro du pool (1, 2 ou 3), extrait automatiquement de la colonne ' +
        '"# pool" de la Configuration (ex. "3A" → Pool 3). Copié automatiquement ; ' +
        'pour corriger, modifiez la Configuration puis "Générer les matchs".',
    2:  'PARTIE # — Numéro de match tel que fourni par l\'horaire du tournoi (colonne ' +
        '"# de match" de la Configuration). Copié automatiquement.',
    3:  'JOUR — Date de la partie, copiée automatiquement depuis la Configuration.',
    4:  'HEURE — Heure de la partie, copiée automatiquement depuis la Configuration.',
    5:  'TERRAIN — Terrain assigné, copié automatiquement depuis la Configuration.',
    6:  'ÉQUIPE 1 — Copiée automatiquement depuis la Configuration. Ne désigne PAS ' +
        'l\'équipe locale (voir colonne J, "Équipe Locale").',
    7:  'ÉQUIPE 2 — Copiée automatiquement depuis la Configuration. Ne désigne PAS ' +
        'l\'équipe locale (voir colonne J, "Équipe Locale").',
    10: 'ÉQUIPE LOCALE — Équipe à domicile, celle qui frappe en dernier (dans le bas ' +
        'de la manche). Connue SEULEMENT une fois la partie jouée — impossible de le ' +
        'savoir à l\'avance, l\'horaire ne précise pas qui reçoit qui. C\'est le ' +
        'registraire qui l\'indique, en même temps que le score, en choisissant dans la ' +
        'liste déroulante (les 2 équipes du match). Cette info sert à calculer les ' +
        'fractions de manches des colonnes O à R.',
    11: 'MANCHES COMPLÈTES — Numéro de la dernière manche jouée (normalement 6, le ' +
        'nombre réglementaire en 13U). Si la partie s\'arrête avant via la règle du ' +
        'marqueur ("Mercy"), indiquer la manche où elle s\'est arrêtée (ex. 5). Si elle ' +
        'se prolonge en supplémentaire, indiquer la dernière manche jouée (ex. 7, 8...).',
    12: 'RETRAITS EN FIN — Nombre de retraits (0, 1 ou 2) déjà comptés dans la dernière ' +
        'demi-manche AU MOMENT où la partie s\'est terminée. Pertinent SEULEMENT si ' +
        'l\'équipe locale a gagné en frappant le point décisif avant d\'avoir complété ' +
        'ses 3 retraits (victoire "walk-off") ; sinon, laisser à 0. Cette valeur donne la ' +
        'fraction de manche utilisée dans les colonnes O à R : 1 retrait = ⅓ de manche, ' +
        '2 retraits = ⅔ de manche.\n' +
        'Ex. : score final 4-3, la locale gagne dans le bas de la 6e avec 1 retrait ' +
        '(K=6, L=1) → la locale obtient 5⅓ manches OFFENSIVES (pas 6, sa dernière manche ' +
        'au bâton a été interrompue) et la visiteuse 5⅓ manches DÉFENSIVES (pas 6, elle ' +
        'n\'a pas eu le temps de compléter ses 3 retraits).',
    13: 'TYPE DE FIN — "Normal" = partie jouée jusqu\'au bout. "Mercy" = arrêtée avant ' +
        'la fin (règle de l\'écart de points). "Forfait" = une équipe ne se présente pas ' +
        'ou abandonne ; le score n\'indique alors que le gagnant, et les manches sont ' +
        'automatiquement comptées 6-6 pour le gagnant et 0-0 pour le perdant. ' +
        '"Supplémentaires" = partie prolongée au-delà des manches réglementaires. IMPORTANT ' +
        '(Note 4, Art. 42.11) : les points marqués/alloués en manches supplémentaires doivent ' +
        'être EXCLUS du ratio de bris d\'égalité. Le système ne le fait pas automatiquement : ' +
        'il signale ces parties par un ⚠ dans les Classements pour que le ratio (RD/RO) soit ' +
        'vérifié/ajusté à la main via la feuille Manches_Détail.',
    14: 'GAGNANT — Calculé automatiquement à partir des scores (colonnes H et I). Ne ' +
        'pas modifier à la main, recalculé par "Mettre à jour les classements".',
    15: 'MO ÉQ.1 (Manches Offensives, Équipe 1) — Manches à la batte jouées par l\'Équipe ' +
        '1, calculé automatiquement. Égal au nombre de manches (col. K), SAUF si l\'Équipe 1 ' +
        'est la locale et gagne par walk-off : sa dernière manche au bâton est alors comptée ' +
        'en fraction de tiers (⅓/⅔) selon la col. L, puisqu\'elle n\'a pas eu besoin de la ' +
        'terminer.',
    16: 'MD ÉQ.1 (Manches Défensives, Équipe 1) — Manches au champ jouées par l\'Équipe 1, ' +
        'calculé automatiquement. Égal au nombre de manches (col. K), SAUF si l\'Équipe 1 est ' +
        'la visiteuse et perd par walk-off : sa dernière manche en défense est alors ' +
        'comptée en fraction de tiers (⅓/⅔) selon la col. L, puisqu\'elle n\'a pas eu le ' +
        'temps de la terminer.',
    17: 'MO ÉQ.2 (Manches Offensives, Équipe 2) — Manches à la batte de l\'Équipe 2, calculé ' +
        'automatiquement. Même logique que la colonne O (MO Éq.1), appliquée à l\'Équipe 2.',
    18: 'MD ÉQ.2 (Manches Défensives, Équipe 2) — Manches au champ de l\'Équipe 2, calculé ' +
        'automatiquement. Même logique que la colonne P (MD Éq.1), appliquée à l\'Équipe 2.'
  };
  Object.keys(notes).forEach(function (col) {
    sheet.getRange(1, Number(col)).setNote(notes[col]);
  });

  // Largeurs de colonnes.
  var widths = [55, 70, 95, 70, 70, 150, 150, 95, 95, 150, 110, 100, 110, 150, 95, 95, 95, 95];
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }

  sheet.setFrozenRows(1);

  // Note : les 18 rangées de matchs sont remplies par generateGames().
  // Ici on prépare seulement le formatage (validations, couleurs) pour 18 lignes.
  var nRows = POOLS.length * GAME_MATRIX.length;  // 18

  // Validation pour "Type de fin" (col M = 13).
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Normal', 'Mercy', 'Forfait', 'Supplémentaires'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 13, nRows, 1).setDataValidation(typeRule);

  // Validation "Manches complètes" (col K = 11) : 1 à 9 (inclut manches supplémentaires).
  var inningRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 9)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 11, nRows, 1).setDataValidation(inningRule);

  // Validation "Retraits en fin" (col L = 12) : 0 à 2.
  var outsRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 2)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 12, nRows, 1).setDataValidation(outsRule);

  // Validation "Équipe Locale" (col J = 10) : la liste (les 2 équipes du match)
  // dépend de chaque ligne — elle est appliquée ligne par ligne dans generateGames(),
  // une fois les deux équipes de chaque match connues. Tant que la partie n'est
  // pas jouée, on ne sait pas qui sera locale : c'est au registraire de la
  // préciser au moment de saisir le score.

  // Couleurs : copiées de Configuration (gris) ; saisie manuelle (jaune) ; calculées (gris).
  sheet.getRange(2, 1, nRows, 7).setBackground(COLOR_CALC);    // A..G : copiées de Configuration
  sheet.getRange(2, 8, nRows, 2).setBackground(COLOR_INPUT);   // H, I : scores
  sheet.getRange(2, 10, nRows, 4).setBackground(COLOR_INPUT);  // J, K, L, M
  sheet.getRange(2, 14, nRows, 5).setBackground(COLOR_CALC);   // N..R : calculées
}

/**
 * Feuille Classements (une par classe). Remplie par calculateStandings().
 */
function createStandingsSheet(ss, classe) {
  var sheet = getOrCreateSheet(ss, SHEET_STANDINGS[classe]);
  sheet.clear();
  sheet.getRange(1, 1).setValue('Cliquez sur "Mettre à jour les classements" pour générer le classement de la Classe ' + classe + '.');
  sheet.getRange(1, 1).setFontStyle('italic');
  for (var c = 1; c <= 12; c++) { sheet.setColumnWidth(c, 110); }
  sheet.setColumnWidth(2, 170);  // colonne Équipe plus large
}

/**
 * Feuille optionnelle Manches_Détail : permet d'entrer les scores par manche
 * pour départager la Priorité 4 (manches en avance) si nécessaire.
 */
function createInningDetailSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_INN_DETAIL);
  sheet.clear();

  var headers = ['Classe', 'Pool', 'Partie #', 'Équipe 1', 'Équipe 2'];
  for (var m = 1; m <= TOTAL_INNINGS; m++) { headers.push('M' + m + ' Loc'); }
  for (var m2 = 1; m2 <= TOTAL_INNINGS; m2++) { headers.push('M' + m2 + ' Vis'); }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet.getRange(1, 1, 1, headers.length));
  sheet.setFrozenRows(1);

  // Cellules de score par manche en jaune (saisie manuelle).
  var nRows = POOLS.length * GAME_MATRIX.length * CLASSES.length;  // 36
  sheet.getRange(2, 6, nRows, TOTAL_INNINGS * 2).setBackground(COLOR_INPUT);

  sheet.getRange(2, 1).setNote(
    'Feuille optionnelle, pour les vérifications MANUELLES de bris d\'égalité. ' +
    'Deux usages :\n' +
    '1) Priorité 4 (manches en avance) : saisir le score cumulatif (ou par manche) de ' +
    'chaque équipe afin de compter les manches complètes où elle menait.\n' +
    '2) Note 4 (Art. 42.11) — parties allées en manches SUPPLÉMENTAIRES : le ratio (RD/RO) ' +
    'affiché dans les Classements inclut à tort les points des supplémentaires. Recalculer ' +
    'ici le ratio en ne comptant QUE les points des manches régulières (les supplémentaires ' +
    'sont exclues).');
}

/**
 * Feuille Aide : explique chaque colonne des feuilles Résultats (à partir de G)
 * et détaille le calcul des fractions de manches (⅓, ⅔), qui est la partie la
 * plus difficile à comprendre du système. Recréée à chaque "Initialiser les
 * feuilles" — ne pas modifier directement, le contenu est régénéré.
 */
function createHelpSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_HELP);
  sheet.clear();
  clearDataValidations(sheet);

  var TOTAL_COLS = 6;
  var row = 1;

  function addTitle(label, bg, fg) {
    sheet.getRange(row, 1, 1, TOTAL_COLS).merge();
    sheet.getRange(row, 1).setValue(label)
      .setFontWeight('bold').setFontSize(12)
      .setBackground(bg).setFontColor(fg || '#000000')
      .setVerticalAlignment('middle');
    row++;
  }

  function addText(content, bold) {
    sheet.getRange(row, 1, 1, TOTAL_COLS).merge();
    var range = sheet.getRange(row, 1).setValue(content).setWrap(true)
      .setVerticalAlignment('top');
    if (bold) { range.setFontWeight('bold'); }
    row++;
  }

  function addColumnDoc(colLabel, description) {
    sheet.getRange(row, 1).setValue(colLabel).setFontWeight('bold')
      .setVerticalAlignment('top').setWrap(true);
    sheet.getRange(row, 2, 1, TOTAL_COLS - 1).merge();
    sheet.getRange(row, 2).setValue(description).setWrap(true).setVerticalAlignment('top');
    row++;
  }

  function addBlank() { row++; }

  addTitle('🏆 AIDE — COMPRENDRE LES FEUILLES RÉSULTATS', COLOR_HEADER, COLOR_HEADER_TX);
  addBlank();

  addTitle('CONFIGURATION : COLLER L\'HORAIRE', COLOR_SECTION);
  addText(
    'La feuille "Configuration" reçoit directement les données de l\'onglet "Horaire ' +
    'globalArbitre" du fichier Excel du tournoi, produit chaque année par le comité du ' +
    'tournoi de l\'ABMR (copier les lignes de matchs à partir de la ligne 2, coller à ' +
    'partir de la cellule A2). La ' +
    'colonne "# pool" (ex. "3A" = Pool 3, Classe A) sert à répartir automatiquement ' +
    'chaque match dans "Résultats A" ou "Résultats B" lorsqu\'on clique sur "Générer ' +
    'les matchs".');

  addBlank();
  addTitle('COLONNES DES FEUILLES "RÉSULTATS A / B" (à partir de J)', COLOR_SECTION);
  addColumnDoc('A à G — Pool, Partie #, Jour, Heure, Terrain, Équipe 1, Équipe 2',
    'Copiées automatiquement depuis la feuille Configuration lors du clic sur "Générer ' +
    'les matchs". Ne pas modifier ici : pour corriger un horaire, modifiez la ' +
    'Configuration puis régénérez. "Équipe 1" / "Équipe 2" ne désignent PAS qui reçoit ' +
    'qui — voir la colonne J, "Équipe Locale".');
  addColumnDoc('J — Équipe Locale',
    'Équipe qui jouait à domicile (qui frappait en dernier, dans le bas de la manche). ' +
    'Champ à remplir SEULEMENT une fois la partie jouée : impossible de le savoir à ' +
    'l\'avance, l\'horaire ne précise pas qui reçoit qui. C\'est le registraire qui ' +
    'l\'indique, en même temps que le score, en choisissant dans la liste déroulante ' +
    '(qui ne propose que les 2 équipes du match).');
  addColumnDoc('K — Manches complètes',
    'Numéro de la dernière manche jouée (1 à 9). Une partie 13U normale dure 6 manches ' +
    'réglementaires. Si la partie se termine plus tôt (règle du marqueur / "Mercy"), ' +
    'indiquer la manche où elle s\'est arrêtée (ex. 5). Si elle se prolonge en manches ' +
    'supplémentaires, indiquer la dernière manche jouée (ex. 7, 8...).');
  addColumnDoc('L — Retraits en fin',
    'Nombre de retraits (0, 1 ou 2) déjà comptés dans la dernière demi-manche AU MOMENT ' +
    'où la partie s\'est terminée — pertinent seulement si l\'équipe locale a gagné en ' +
    'frappant le point gagnant avant d\'avoir complété ses 3 retraits (victoire ' +
    '"walk-off"). Si la partie est allée jusqu\'au bout normalement (3 retraits) ou que ' +
    'la visiteuse a gagné, laisser à 0. C\'est cette valeur qui détermine la fraction de ' +
    'manche calculée plus bas (1 retrait = ⅓ de manche, 2 retraits = ⅔ de manche).');
  addColumnDoc('M — Type de fin',
    '"Normal" = partie jouée selon les règles standards jusqu\'au bout. "Mercy" = partie ' +
    'arrêtée avant la fin en raison de la règle de l\'écart de points. "Forfait" = une ' +
    'équipe ne se présente pas ou abandonne ; le score n\'indique alors que qui gagne, et ' +
    'les manches sont automatiquement comptées 6-6 pour le gagnant et 0-0 pour le perdant. ' +
    '"Supplémentaires" = partie prolongée au-delà des manches réglementaires (voir la ' +
    'section "Manches supplémentaires et Note 4" plus bas).');
  addColumnDoc('N — Gagnant (calculé)',
    'Nom de l\'équipe gagnante, déterminé automatiquement à partir des scores (colonnes ' +
    'H et I). Ne pas modifier à la main — recalculé par "Mettre à jour les classements".');
  addColumnDoc('O / P — MO Éq.1 / MD Éq.1 (calculé)',
    'MO = Manches OFFENSIVES (à la batte) ; MD = Manches DÉFENSIVES (au champ) jouées par ' +
    'l\'Équipe 1, en fractions de tiers si la partie s\'est terminée par un walk-off ou ' +
    'un Mercy en milieu de manche. Voir la section "Fractions de manches" ci-dessous. ' +
    'Calculé automatiquement.');
  addColumnDoc('Q / R — MO Éq.2 / MD Éq.2 (calculé)',
    'Même chose que O / P (MO/MD), mais pour l\'Équipe 2.');

  addBlank();
  addTitle('POURQUOI DES FRACTIONS DE MANCHES (⅓, ⅔) ?', COLOR_SECTION);
  addText(
    'Le classement utilise deux ratios pour départager les égalités (Art. 42.11 ' +
    'Baseball Québec) : le ratio défensif RD = PC / MD (points contre / manches ' +
    'défensives, le plus bas est le meilleur), et le ratio offensif RO = PP / MO (points ' +
    'pour / manches offensives, le plus haut est le meilleur). Ces abréviations (RD, RO, ' +
    'MO, MD, PP, PC) sont celles des colonnes des feuilles Classements. Pour que ces ratios ' +
    'soient justes, il faut compter le nombre RÉEL de manches jouées par chaque équipe — ' +
    'pas seulement le nombre de manches de la partie.');
  addText(
    'Quand la partie se termine par un point gagnant frappé par l\'équipe locale dans le ' +
    'bas d\'une manche (avant ses 3 retraits — une victoire "walk-off"), deux ' +
    'demi-manches restent incomplètes : l\'équipe locale n\'a pas eu besoin de terminer ' +
    'sa manche à la batte, et l\'équipe visiteuse n\'a pas eu besoin de terminer sa ' +
    'manche en défense (elle n\'a pas obtenu ses 3 retraits, le jeu s\'est arrêté dès que ' +
    'le point a compté). Tout le reste de la partie (le haut de cette même manche, et ' +
    'toutes les manches précédentes) s\'est joué normalement et compte comme des manches ' +
    'pleines.');
  addText(
    'La fraction se calcule à partir du nombre de retraits déjà comptés (colonne L) au ' +
    'moment où le point gagnant a été marqué :  0 retrait → la manche ne compte pas pour ' +
    'cette demi-manche  •  1 retrait → ⅓ de manche  •  2 retraits → ⅔ de manche.');
  addBlank();
  addText('EXEMPLE :', true);
  addText(
    'Partie qui se termine après 6 manches : l\'équipe locale gagne 4-3 en frappant le ' +
    'point gagnant dans le bas de la 6e manche, avec 1 retrait déjà compté ' +
    '(colonne K = 6, colonne L = 1) :');
  addText(
    '• Équipe locale : 5⅓ manches OFFENSIVES (5 manches pleines + ⅓ pour la 6e, ' +
    'interrompue) — 6 manches DÉFENSIVES (a joué tout le haut de chaque manche, ' +
    'incluant la 6e, au complet).');
  addText(
    '• Équipe visiteuse : 6 manches OFFENSIVES (a frappé au complet à chaque manche, ' +
    'incluant le haut de la 6e) — 5⅓ manches DÉFENSIVES (n\'a joué qu\'⅓ de la 6e en ' +
    'défense avant que le point gagnant ne mette fin à la partie).');
  addBlank();
  addText(
    'Si la partie va jusqu\'au bout normalement (3 retraits complétés en bas de la ' +
    'dernière manche, ou victoire de la visiteuse), AUCUNE fraction n\'est utilisée : ' +
    'les deux équipes ont joué exactement le même nombre de manches pleines.');
  addText(
    'Cas particulier "Mercy" : si la règle de l\'écart de points est déclenchée alors ' +
    'que l\'équipe locale est à la batte (dans le bas d\'une manche), les mêmes fractions ' +
    's\'appliquent que pour une victoire walk-off. Si elle est déclenchée à un autre ' +
    'moment (manche complète, ou victoire de la visiteuse), il n\'y a pas de fraction : ' +
    'l\'équipe gagnante reçoit simplement 6 manches en défense (le maximum ' +
    'réglementaire), comme si la partie avait été jouée au complet.');

  addBlank();
  addTitle('MANCHES SUPPLÉMENTAIRES ET NOTE 4 (Art. 42.11)', COLOR_SECTION);
  addText(
    'Quand une partie se prolonge en manches supplémentaires, la Note 4 du règlement exige ' +
    'que les points marqués/alloués DANS les manches supplémentaires soient EXCLUS du ratio ' +
    'de bris d\'égalité (RD et RO) : seuls les points des manches régulières comptent.');
  addText(
    'Le système ne fait PAS cet ajustement automatiquement (cas rare, et il faudrait saisir ' +
    'le pointage réglementaire). À la place : indiquez "Supplémentaires" dans la colonne ' +
    '"Type de fin" de la partie. Le classement du pool affichera alors un ⚠ rappelant que ' +
    'les ratios RD/RO de ce pool incluent des points de supplémentaires, et que ces ratios ' +
    'doivent être recalculés à la main (en excluant les supplémentaires) via la feuille ' +
    'Manches_Détail — mais SEULEMENT si un bris d\'égalité se joue réellement sur le ratio.');
  addText(
    'Le gagnant, la fiche victoires-défaites et la fiche tête-à-tête (1er critère de bris ' +
    'd\'égalité) restent toujours corrects : seul le ratio (2e et 3e critères) est concerné.');

  addBlank();
  addTitle('MISE À JOUR AUTOMATIQUE DES CLASSEMENTS', COLOR_SECTION);
  addText(
    'Les classements se mettent à jour TOUT SEULS, en direct et pour tout le monde, dès ' +
    'qu\'une partie est entrée au complet. Plusieurs personnes peuvent saisir des scores ' +
    'en même temps : chaque mise à jour est visible immédiatement par ceux qui regardent ' +
    'les onglets "Classements A / B". Aucune manipulation requise.');
  addText(
    'La mise à jour ne se déclenche que lorsque TOUTES les colonnes de saisie d\'une ' +
    'partie sont remplies (H à M : les 2 scores, Équipe Locale, Manches complètes, ' +
    'Retraits en fin, Type de fin) — sauf pour un "Forfait", où seuls les 2 scores et le ' +
    'Type suffisent. Tant que la ligne n\'est pas complète, le classement ne bouge pas : ' +
    'c\'est voulu, pour qu\'il n\'affiche jamais un résultat à moitié entré.');
  addText(
    'IMPORTANT — pensez à remplir les Manches complètes (colonne K) même pour une partie ' +
    'normale : une journée de pluie peut écourter les parties (ex. 5 manches au lieu de ' +
    '6), et cette valeur change les ratios du classement.');
  addText(
    'Cas particulier : si vous EFFACEZ plusieurs cellules d\'un coup (correction en lot), ' +
    'ou en cas de doute, cliquez sur "🏆 Tournoi Baseball > Mettre à jour les classements" ' +
    'pour forcer un recalcul complet et propre.', true);

  addBlank();
  addTitle('RAPPEL', COLOR_SECTION);
  addText(
    'L\'équipe locale (colonne J) ne peut être connue qu\'une fois la partie jouée — ' +
    'remplissez-la TOUJOURS en même temps que le score, sinon le système ne peut pas ' +
    'départager les fractions correctement (un avertissement est alors inscrit dans le ' +
    'journal d\'exécution, visible via Extensions > Apps Script > Exécutions).');

  sheet.setColumnWidth(1, 220);
  for (var c = 2; c <= TOTAL_COLS; c++) { sheet.setColumnWidth(c, 150); }
}

// ============================================================================
//  GÉNÉRATION DES MATCHS
// ============================================================================

/**
 * Génère les matchs par classe dans les feuilles Résultats à partir de
 * l'horaire collé dans la feuille Configuration. Préserve les scores déjà
 * saisis si possible (appariés par Pool + Partie #).
 */
function generateGames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var schedule = readScheduleRows(ss);   // toutes les parties valides, peu importe la classe

  CLASSES.forEach(function (classe) {
    var sheet = ss.getSheetByName(SHEET_RESULTS[classe]);
    if (!sheet) { return; }

    var matches = schedule.filter(function (m) { return m.classe === classe; });
    // Tri par pool puis par # de match, pour regrouper les parties par pool.
    matches.sort(function (a, b) {
      if (a.pool !== b.pool) { return a.pool - b.pool; }
      return Number(a.partieNum) - Number(b.partieNum);
    });

    var nRows = matches.length;
    if (nRows === 0) {
      var maxRowsEmpty = sheet.getMaxRows();
      if (maxRowsEmpty > 1) { sheet.getRange(2, 1, maxRowsEmpty - 1, 18).clearContent(); }
      return;
    }

    // Conserve les scores / saisies manuelles déjà présents, appariés par clé
    // Pool + Partie # (robuste même si l'horaire est recollé dans un ordre différent).
    var existingByKey = {};
    var prevLast = sheet.getLastRow();
    if (prevLast >= 2) {
      var prevData = sheet.getRange(2, 1, prevLast - 1, 18).getValues();
      prevData.forEach(function (r) {
        var key = r[0] + '|' + r[1];
        existingByKey[key] = r;
      });
    }

    var rows = matches.map(function (m) {
      var key = m.pool + '|' + m.partieNum;
      var prev = existingByKey[key];
      return [
        m.pool, m.partieNum, m.jour, m.heure, m.terrain, m.teamA, m.teamB,
        prev ? prev[7] : '',  // H : Score Équipe 1
        prev ? prev[8] : '',  // I : Score Équipe 2
        prev ? prev[9] : '',  // J : Équipe Locale
        prev ? prev[10] : '', // K : Manches complètes
        prev ? prev[11] : '', // L : Retraits en fin
        prev ? prev[12] : 'Normal', // M : Type de fin
        '',             // N : Gagnant
        '', '', '', ''  // O..R : manches calculées
      ];
    });

    var maxRows = sheet.getMaxRows();
    if (maxRows > 1) { sheet.getRange(2, 1, maxRows - 1, 18).clearContent(); }
    sheet.getRange(2, 1, nRows, 18).setValues(rows);
    sheet.getRange(2, 3, nRows, 1).setNumberFormat('yyyy-mm-dd');  // Jour
    sheet.getRange(2, 4, nRows, 1).setNumberFormat('HH:mm');       // Heure

    // Validation "Équipe Locale" (col J) : liste propre à chaque match (ses 2 équipes).
    for (var i = 0; i < nRows; i++) {
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList([matches[i].teamA, matches[i].teamB], true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2 + i, 10).setDataValidation(rule);
    }
  });

  // Remplit aussi la feuille Manches_Détail (liste des matchs).
  fillInningDetailGames(ss);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Matchs générés à partir de la Configuration (' + schedule.length + ' parties). ' +
    'Indiquez l\'Équipe Locale et les scores au moment de saisir chaque résultat, ' +
    'puis "Mettre à jour les classements".',
    'Tournoi Baseball', 6);
}

/**
 * Remplit la feuille Manches_Détail avec la liste des matchs (sans scores),
 * à partir de l'horaire de la feuille Configuration.
 */
function fillInningDetailGames(ss) {
  var sheet = ss.getSheetByName(SHEET_INN_DETAIL);
  if (!sheet) { return; }

  var schedule = readScheduleRows(ss);
  schedule.sort(function (a, b) {
    if (a.classe !== b.classe) { return a.classe < b.classe ? -1 : 1; }
    if (a.pool !== b.pool) { return a.pool - b.pool; }
    return Number(a.partieNum) - Number(b.partieNum);
  });

  var nCols = 5 + TOTAL_INNINGS * 2;
  var maxRows = sheet.getMaxRows();
  if (maxRows > 1) { sheet.getRange(2, 1, maxRows - 1, nCols).clearContent(); }

  var rows = schedule.map(function (m) {
    var row = [m.classe, m.pool, m.partieNum, m.teamA, m.teamB];
    for (var k = 0; k < TOTAL_INNINGS * 2; k++) { row.push(''); }
    return row;
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, nCols).setValues(rows);
  }
}

// ============================================================================
//  LECTURE DES DONNÉES (Configuration & Résultats)
// ============================================================================

/**
 * Sépare la valeur de la colonne "# pool" (ex. "3A") en numéro de pool et classe.
 * @return {{pool:number, classe:string}|null} null si le format est invalide.
 */
function parsePoolClasse(raw) {
  var s = String(raw).trim().toUpperCase();
  var m = s.match(/^(\d+)\s*([AB])$/);
  if (!m) { return null; }
  return { pool: parseInt(m[1], 10), classe: m[2] };
}

/**
 * Lit et interprète toutes les parties saisies dans la feuille Configuration
 * (format copié-collé depuis l'onglet "Horaire globalArbitre" du fichier Excel
 * du tournoi, produit par le comité de l'ABMR). Ignore les lignes vides ou mal formées.
 * @return {Array<Object>} ex: { partieNum, pool, classe, jour, heure, terrain, teamA, teamB }
 */
function readScheduleRows(ss) {
  var sheet = ss.getSheetByName(SHEET_CONFIG);
  var games = [];
  if (!sheet) { return games; }

  var last = sheet.getLastRow();
  if (last < 2) { return games; }

  var data = sheet.getRange(2, 1, last - 1, 7).getValues();
  data.forEach(function (row) {
    var partieNum  = row[0];
    var poolClasse = row[1];
    var jour       = row[2];
    var heure      = row[3];
    var terrain    = row[4];
    var teamA      = String(row[5]).trim();
    var teamB      = String(row[6]).trim();

    if (String(poolClasse).trim() === '' || teamA === '' || teamB === '') { return; }

    var pc = parsePoolClasse(poolClasse);
    if (!pc) { return; }  // format "# pool" invalide (ni "1A" ni "3B", etc.)

    games.push({
      partieNum: partieNum, pool: pc.pool, classe: pc.classe,
      jour: jour, heure: heure, terrain: terrain, teamA: teamA, teamB: teamB
    });
  });
  return games;
}

/**
 * Retourne les équipes d'une classe regroupées par pool, déduites de l'horaire
 * de la feuille Configuration.
 * @return {Object} ex: { 1: ['Aigles','Lynx','Ours','Loups'], 2:[...], 3:[...] }
 */
function getTeams(classe) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = { 1: [], 2: [], 3: [] };
  readScheduleRows(ss).forEach(function (m) {
    if (m.classe !== classe) { return; }
    var list = result[m.pool];
    if (!list) { return; }
    if (list.indexOf(m.teamA) === -1) { list.push(m.teamA); }
    if (list.indexOf(m.teamB) === -1) { list.push(m.teamB); }
  });
  return result;
}

/**
 * Lit les résultats d'une classe depuis la feuille Résultats.
 * Calcule les manches via calculateInnings() et ignore les parties incomplètes.
 * @return {Array<Object>} liste de parties avec stats calculées.
 */
function getGameResults(classe) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_RESULTS[classe]);
  var games = [];
  if (!sheet) { return games; }

  var last = sheet.getLastRow();
  if (last < 2) { return games; }

  var data = sheet.getRange(2, 1, last - 1, 18).getValues();
  data.forEach(function (row, idx) {
    var pool        = parseInt(row[0], 10);
    var partie      = row[1];
    var teamA       = String(row[5]).trim();
    var teamB       = String(row[6]).trim();
    var scoreA      = row[7];
    var scoreB      = row[8];
    var localSel    = String(row[9]).trim();
    var manches     = parseInt(row[10], 10);
    var retraits    = row[11] === '' ? 0 : parseInt(row[11], 10);
    var type        = String(row[12]).trim() || 'Normal';

    // Ignore les parties sans scores (résultats manquants).
    if (scoreA === '' || scoreB === '' || teamA === '' || teamB === '') {
      return;
    }
    scoreA = Number(scoreA);
    scoreB = Number(scoreB);
    if (isNaN(scoreA) || isNaN(scoreB)) { return; }

    // Manches par défaut = 6 si non saisi (partie complète présumée).
    if (isNaN(manches) || manches < 1) { manches = TOTAL_INNINGS; }
    if (isNaN(retraits) || retraits < 0) { retraits = 0; }

    // L'équipe locale n'est connue qu'une fois la partie jouée (le registraire
    // l'indique en saisissant le score). Si elle n'est pas précisée, impossible
    // de savoir qui a frappé en dernier : on calcule les manches de façon
    // symétrique (aucun avantage walk-off accordé) — sans incidence si la
    // partie s'est jouée au complet (6 manches, 0 retrait), mais à corriger
    // sinon pour la précision des bris d'égalité.
    var local, visiteur, scoreLocal, scoreVisiteur, homeKnown;
    if (localSel === teamA) {
      local = teamA; visiteur = teamB; scoreLocal = scoreA; scoreVisiteur = scoreB; homeKnown = true;
    } else if (localSel === teamB) {
      local = teamB; visiteur = teamA; scoreLocal = scoreB; scoreVisiteur = scoreA; homeKnown = true;
    } else {
      local = teamA; visiteur = teamB; scoreLocal = scoreA; scoreVisiteur = scoreB; homeKnown = false;
      if (type !== 'Forfait' && (manches < TOTAL_INNINGS || retraits > 0)) {
        Logger.log('AVERTISSEMENT : Équipe Locale non indiquée pour Classe ' + classe +
                   ' Pool ' + pool + ' Partie ' + partie + ' (' + teamA + ' vs ' + teamB + ').');
      }
    }

    var inn = calculateInnings(scoreLocal, scoreVisiteur, manches, retraits, type, homeKnown);
    var winner = (type === 'Forfait')
      ? (scoreLocal >= scoreVisiteur ? local : visiteur)   // forfait : le score indique le gagnant
      : (scoreLocal > scoreVisiteur ? local : (scoreVisiteur > scoreLocal ? visiteur : ''));

    games.push({
      pool: pool,
      partie: partie,
      rowIndex: idx + 2,
      local: local,
      visiteur: visiteur,
      scoreLocal: scoreLocal,
      scoreVisiteur: scoreVisiteur,
      manches: manches,
      retraits: retraits,
      type: type,
      winner: winner,
      offLocal: inn.offLocal,
      defLocal: inn.defLocal,
      offVisiteur: inn.offVisiteur,
      defVisiteur: inn.defVisiteur
    });
  });
  return games;
}

/**
 * Détermine si une ligne de Résultats est COMPLÈTE — c.-à-d. si le registraire a
 * fini de saisir la partie et que le recalcul live (handleResultEdit) peut se déclencher
 * sans risquer de bouger les classements en pleine saisie.
 *
 * Règle (colonnes de saisie H..M) :
 *   - Cas général : les 6 colonnes non-vides — Score Éq.1 (H), Score Éq.2 (I),
 *     Équipe Locale (J), Manches complètes (K), Retraits en fin (L),
 *     Type de fin (M). On exige Manches explicitement (K) car une journée de
 *     pluie peut écourter une partie « normale » (ex. 5 manches) — il ne faut pas
 *     se fier au défaut de 6.
 *   - Exception Forfait : si Type de fin = "Forfait", seules H, I et M sont
 *     requises (manches/retraits comptés automatiquement 6-6/0-0, équipe locale
 *     non pertinente).
 *
 * IMPORTANT : la valeur 0 est valide (un score de 0, ou Retraits = 0). On teste
 * donc le VIDE (cell === '' / null), jamais la fausseté (!cell aurait rejeté 0).
 *
 * @param {Object} v  valeurs des colonnes H..M : {scoreA, scoreB, local, manches, retraits, type}
 * @return {boolean}
 */
function isRowComplete(v) {
  function filled(x) { return x !== '' && x !== null && x !== undefined; }

  // Scores toujours requis.
  if (!filled(v.scoreA) || !filled(v.scoreB)) { return false; }

  var type = String(v.type).trim();

  // Type de fin toujours requis (sinon on ne sait pas si c'est un forfait).
  if (!filled(type)) { return false; }

  // Forfait : scores + type suffisent.
  if (type === 'Forfait') { return true; }

  // Cas général : Équipe Locale, Manches et Retraits aussi requis.
  return filled(v.local) && filled(v.manches) && filled(v.retraits);
}

// ============================================================================
//  CALCUL DES MANCHES (fractions)
// ============================================================================

/**
 * Calcule les manches offensives/défensives pour la locale et la visiteuse.
 *
 * Règles :
 *  - Forfait : gagnant 6/6, perdant 0/0.
 *  - Mercy : gagnant reçoit 6 manches défensives ; perdant reçoit les manches
 *            réellement jouées (avec fractions).
 *  - Victoire locale en bas (walk-off) : détectée si Score Local > Score Visiteur
 *    ET (manches < 6 OU retraits > 0). La locale n'a pas frappé en bas de la
 *    dernière manche complète -> fractions.
 *  - Sinon (partie normale complète, ou visiteur gagne) : symétrique = manches/manches.
 *
 * @param {boolean} [homeKnown=true]  si faux (équipe locale non indiquée), on force
 *                  un résultat symétrique : impossible de savoir qui a joué en bas.
 * @return {{offLocal:number, defLocal:number, offVisiteur:number, defVisiteur:number}}
 */
function calculateInnings(scoreLocal, scoreVisiteur, manchesCompletes, retraitsEnFin, typeFin, homeKnown) {
  if (homeKnown === undefined) { homeKnown = true; }
  var N = manchesCompletes;
  var H = retraitsEnFin;
  var frac = H / 3;                 // fraction de manche (0, 1/3, 2/3)
  var partial = (N - 1) + frac;     // ex: N=5, H=1 -> 4.333

  // -------- FORFAIT --------
  if (typeFin === 'Forfait') {
    if (scoreLocal >= scoreVisiteur) {
      // Locale gagne par forfait.
      return { offLocal: 6, defLocal: 6, offVisiteur: 0, defVisiteur: 0 };
    } else {
      // Visiteur gagne par forfait.
      return { offLocal: 0, defLocal: 0, offVisiteur: 6, defVisiteur: 6 };
    }
  }

  // Détermine si la locale a gagné dans le bas (walk-off). Impossible à
  // déterminer si l'équipe locale n'a pas été indiquée (homeKnown = false).
  var localeWinsBottom = homeKnown && (scoreLocal > scoreVisiteur) && (N < TOTAL_INNINGS || H > 0);

  // -------- MERCY --------
  if (typeFin === 'Mercy') {
    if (scoreLocal > scoreVisiteur) {
      // Locale gagne par mercy.
      if (localeWinsBottom) {
        // Mercy déclenché dans le bas : locale n'a pas terminé sa manche offensive.
        // Locale (gagnante) : offensive partielle, défensive = 6.
        // Visiteur (perdant) : offensive = N, défensive = partielle.
        return {
          offLocal: partial, defLocal: TOTAL_INNINGS,
          offVisiteur: N,    defVisiteur: partial
        };
      } else {
        // Mercy déclenché en haut / fin de manche complète, locale gagnante.
        return {
          offLocal: N, defLocal: TOTAL_INNINGS,
          offVisiteur: N, defVisiteur: N
        };
      }
    } else {
      // Visiteur gagne par mercy : visiteur gagnant (6 def), locale perdante.
      return {
        offLocal: N, defLocal: N,
        offVisiteur: N, defVisiteur: TOTAL_INNINGS
      };
    }
  }

  // -------- NORMAL --------
  if (localeWinsBottom) {
    // Victoire locale en bas de la Ne manche (walk-off).
    // Visiteur : N off, partial def.   Locale : partial off, N def.
    return {
      offLocal: partial, defLocal: N,
      offVisiteur: N,    defVisiteur: partial
    };
  }

  // Partie normale : visiteur gagne, ou partie complète symétrique.
  return {
    offLocal: N, defLocal: N,
    offVisiteur: N, defVisiteur: N
  };
}

// ============================================================================
//  CALCUL DES CLASSEMENTS
// ============================================================================

/**
 * Recalcule l'ensemble des classements pour les deux classes et écrit
 * les manches calculées dans les feuilles Résultats.
 */
function calculateStandings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  CLASSES.forEach(function (classe) {
    var games = getGameResults(classe);

    // 1) Écrit les colonnes calculées (J..N) dans la feuille Résultats.
    writeCalculatedResults(ss, classe, games);

    // 2) Construit le classement.
    buildStandingsSheet(ss, classe, games);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('Classements recalculés.', 'Tournoi Baseball', 5);
}

/**
 * Écrit le gagnant et les manches calculées (N..R) dans la feuille Résultats.
 */
function writeCalculatedResults(ss, classe, games) {
  var sheet = ss.getSheetByName(SHEET_RESULTS[classe]);
  if (!sheet) { return; }

  var last = sheet.getLastRow();
  if (last < 2) { return; }

  // Indexe les parties calculées par numéro de ligne.
  var byRow = {};
  games.forEach(function (g) { byRow[g.rowIndex] = g; });

  for (var r = 2; r <= last; r++) {
    writeRowCalc(sheet, r, byRow[r]);
  }
}

/**
 * Écrit (ou efface) les colonnes calculées N..R (14..18) d'UNE ligne de Résultats.
 * Source unique du mapping N..R, partagée par writeCalculatedResults (refresh
 * complet via le menu) et handleResultEdit (écriture ciblée de la ligne saisie).
 *
 * @param {Sheet}  sheet     feuille Résultats A/B
 * @param {number} rowIndex  numéro de ligne (>= 2)
 * @param {Object} game      partie calculée, ou null/undefined pour effacer
 */
function writeRowCalc(sheet, rowIndex, game) {
  if (game) {
    sheet.getRange(rowIndex, 14).setValue(game.winner);                       // N : Gagnant
    sheet.getRange(rowIndex, 15).setValue(formatFraction(game.offLocal));     // O
    sheet.getRange(rowIndex, 16).setValue(formatFraction(game.defLocal));     // P
    sheet.getRange(rowIndex, 17).setValue(formatFraction(game.offVisiteur));  // Q
    sheet.getRange(rowIndex, 18).setValue(formatFraction(game.defVisiteur));  // R
  } else {
    // Partie sans résultat : efface les calculs.
    sheet.getRange(rowIndex, 14, 1, 5).clearContent();
  }
}

/**
 * Calcule les statistiques cumulées d'une équipe sur un ensemble de parties.
 * @param {string} team        nom de l'équipe
 * @param {Array}  games       parties à considérer
 * @param {boolean} excludeForfaitRatios  si vrai, exclut les forfaits des ratios
 * @return {Object} stats
 */
function computeTeamStats(team, games, excludeForfaitRatios) {
  var s = {
    team: team, pj: 0, v: 0, d: 0,
    rs: 0, ra: 0,            // runs scored / allowed
    offInn: 0, defInn: 0,    // manches (avec exclusion forfait si demandé)
    rsRatio: 0, raRatio: 0
  };

  games.forEach(function (g) {
    var isLocal = (g.local === team);
    var isVis   = (g.visiteur === team);
    if (!isLocal && !isVis) { return; }

    s.pj++;
    if (g.winner === team) { s.v++; }
    else if (g.winner !== '' ) { s.d++; }

    var teamScore = isLocal ? g.scoreLocal : g.scoreVisiteur;
    var oppScore  = isLocal ? g.scoreVisiteur : g.scoreLocal;
    s.rs += teamScore;
    s.ra += oppScore;

    // Manches pour les ratios — peut exclure les forfaits.
    var skipRatio = (excludeForfaitRatios && g.type === 'Forfait');
    if (!skipRatio) {
      if (isLocal) {
        s.offInn += g.offLocal;
        s.defInn += g.defLocal;
      } else {
        s.offInn += g.offVisiteur;
        s.defInn += g.defVisiteur;
      }
    }
  });

  // Ratios : RA/DefInn (le plus bas gagne), RS/OffInn (le plus haut gagne).
  s.raRatio = s.defInn > 0 ? (s.ra / s.defInn) : Number.POSITIVE_INFINITY;
  s.rsRatio = s.offInn > 0 ? (s.rs / s.offInn) : 0;
  return s;
}

/**
 * Classement d'un pool (Étape A). Trie les 4 équipes avec bris d'égalité.
 * @return {Array} stats triées (1er en tête) avec champ .rank
 */
function calculatePoolStandings(games, teams) {
  // Stats globales du pool (pour affichage).
  var stats = teams.map(function (t) {
    return computeTeamStats(t, games, false);
  });

  // Tri avec bris d'égalité — Étape A utilise seulement les parties directes
  // entre équipes à égalité (géré dans tiebreaker via useAllGames=false).
  var ordered = orderTeams(teams, games, false);

  // Associe le rang à chaque équipe.
  var rankByTeam = {};
  ordered.forEach(function (t, i) { rankByTeam[t] = i + 1; });

  stats.forEach(function (s) { s.rank = rankByTeam[s.team] || 99; });
  stats.sort(function (a, b) { return a.rank - b.rank; });
  return stats;
}

/**
 * Étape B (meilleur 2e) ou Étape C (classement des 1ers).
 * @param {Array}  teams        noms des équipes à comparer
 * @param {Array}  games        toutes les parties (toutes classes/pools confondues
 *                              pour cette classe) — on filtrera par pool de l'équipe
 * @param {boolean} useAllGames true => Étapes B/C (toutes les parties de pool de
 *                               chaque équipe). false => Étape A (tête-à-tête).
 * @return {Array} équipes ordonnées
 */
function calculateStep(teams, games, useAllGames) {
  return orderTeams(teams, games, useAllGames);
}

/**
 * Ordonne une liste d'équipes en appliquant le bris d'égalité récursif.
 * Sépare d'abord par fiche V-D globale (sur l'ensemble pertinent), puis pour
 * chaque groupe à égalité, applique tiebreaker().
 *
 * @param {Array}  teams
 * @param {Array}  games        parties disponibles
 * @param {boolean} useAllGames true = B/C (toutes parties de pool), false = A (direct)
 * @return {Array} noms ordonnés
 */
function orderTeams(teams, games, useAllGames) {
  if (teams.length <= 1) { return teams.slice(); }

  // Trie l'ensemble via tiebreaker (qui gère récursivement les sous-égalités).
  return tiebreaker(teams.slice(), games, useAllGames);
}

/**
 * Résout récursivement les égalités selon l'Art. 42.11.
 *
 * Priorités :
 *   1. Fiche tête-à-tête (V-D dans les parties entre équipes à égalité)
 *   2. Ratio RA/DefInn (le plus bas gagne)
 *   3. Ratio RS/OffInn (le plus haut gagne)
 *   4. Manches en avance (avertissement : vérification manuelle requise)
 *
 * Résolution partielle : si une priorité sépare le groupe en sous-groupes,
 * chaque sous-groupe encore à égalité est re-traité depuis la Priorité 1,
 * en recalculant sur ses propres parties directes.
 *
 * @param {Array}  tiedTeams    équipes à départager
 * @param {Array}  games        parties disponibles
 * @param {boolean} useAllGames true = B/C, false = A (tête-à-tête direct)
 * @return {Array} équipes ordonnées (meilleure en tête)
 */
function tiebreaker(tiedTeams, games, useAllGames) {
  if (tiedTeams.length === 1) { return tiedTeams.slice(); }

  // Sélectionne les parties pertinentes pour le calcul des ratios.
  // - Étape A (useAllGames=false) : uniquement les parties ENTRE équipes à égalité.
  // - Étapes B/C (useAllGames=true) : toutes les parties de pool de chaque équipe.
  var relevantGames = useAllGames
    ? games.filter(function (g) {
        // On garde toutes les parties impliquant l'une des équipes à égalité.
        return tiedTeams.indexOf(g.local) !== -1 || tiedTeams.indexOf(g.visiteur) !== -1;
      })
    : headToHeadGames(tiedTeams, games);

  // ---------- PRIORITÉ 1 : fiche tête-à-tête ----------
  // Pour A : fiche dans les parties entre équipes à égalité.
  // Pour B/C : ces équipes ne se sont jamais affrontées -> fiche = toutes parties.
  var h2hGames = useAllGames ? relevantGames : headToHeadGames(tiedTeams, games);
  var groups = groupByMetric(tiedTeams, function (team) {
    var st = computeTeamStats(team, h2hGames, useAllGames /* exclure forfait pour B/C */);
    // Différentiel V-D : plus élevé = meilleur. On groupe par valeur identique.
    return st.v - st.d;
  }, true /* descending */);

  if (groups.length > 1) {
    return resolveGroups(groups, games, useAllGames);
  }

  // ---------- PRIORITÉ 2 : ratio RA/DefInn (le plus bas gagne) ----------
  groups = groupByMetric(tiedTeams, function (team) {
    var st = computeTeamStats(team, relevantGames, useAllGames);
    return st.raRatio;
  }, false /* ascending : plus bas = meilleur */);

  if (groups.length > 1) {
    return resolveGroups(groups, games, useAllGames);
  }

  // ---------- PRIORITÉ 3 : ratio RS/OffInn (le plus haut gagne) ----------
  groups = groupByMetric(tiedTeams, function (team) {
    var st = computeTeamStats(team, relevantGames, useAllGames);
    return st.rsRatio;
  }, true /* descending : plus haut = meilleur */);

  if (groups.length > 1) {
    return resolveGroups(groups, games, useAllGames);
  }

  // ---------- PRIORITÉ 4 : manches en avance (vérification manuelle) ----------
  // Données par manche non disponibles automatiquement.
  // On marque les équipes et on conserve l'ordre alphabétique stable.
  Logger.log('AVERTISSEMENT : Priorité 4 atteinte pour : ' + tiedTeams.join(', ') +
             '. Vérification manuelle requise - Voir feuille Manches_Détail.');
  var warned = tiedTeams.slice().sort();
  warned.__needsManualCheck = true;
  return warned;
}

/**
 * Applique les sous-groupes : pour chaque sous-groupe encore à égalité (>1 équipe),
 * relance tiebreaker récursivement sur ses parties directes.
 */
function resolveGroups(groups, games, useAllGames) {
  var ordered = [];
  groups.forEach(function (group) {
    if (group.length === 1) {
      ordered.push(group[0]);
    } else {
      // Sous-égalité : recommence depuis la Priorité 1 sur ce sous-groupe.
      // Pour A on recalcule sur leurs parties directes ; pour B/C on garde toutes.
      var sub = tiebreaker(group, games, useAllGames);
      sub.forEach(function (t) { ordered.push(t); });
    }
  });
  return ordered;
}

/**
 * Groupe des équipes par valeur d'une métrique, et retourne les groupes triés
 * du meilleur au moins bon. Les équipes de même valeur restent dans un sous-groupe.
 *
 * @param {Array}    teams
 * @param {Function} metricFn   team -> nombre
 * @param {boolean}  descending true = plus haut est meilleur
 * @return {Array<Array>} liste de groupes (chaque groupe = équipes à égalité)
 */
function groupByMetric(teams, metricFn, descending) {
  var entries = teams.map(function (t) {
    return { team: t, value: metricFn(t) };
  });

  entries.sort(function (a, b) {
    return descending ? (b.value - a.value) : (a.value - b.value);
  });

  var groups = [];
  var current = [];
  var lastVal = null;
  entries.forEach(function (e, i) {
    if (i === 0 || approxEqual(e.value, lastVal)) {
      current.push(e.team);
    } else {
      groups.push(current);
      current = [e.team];
    }
    lastVal = e.value;
  });
  if (current.length > 0) { groups.push(current); }
  return groups;
}

/**
 * Retourne les parties jouées strictement entre les équipes du groupe.
 */
function headToHeadGames(teams, games) {
  return games.filter(function (g) {
    return teams.indexOf(g.local) !== -1 && teams.indexOf(g.visiteur) !== -1;
  });
}

/** Comparaison de flottants avec tolérance. */
function approxEqual(a, b) {
  if (a === Number.POSITIVE_INFINITY && b === Number.POSITIVE_INFINITY) { return true; }
  return Math.abs(a - b) < 1e-9;
}

// ============================================================================
//  CONSTRUCTION DE LA FEUILLE CLASSEMENTS
// ============================================================================

/**
 * Construit la feuille Classements pour une classe : 3 pools + Étape C + Étape B.
 */
function buildStandingsSheet(ss, classe, games) {
  var sheet = ss.getSheetByName(SHEET_STANDINGS[classe]);
  if (!sheet) { sheet = getOrCreateSheet(ss, SHEET_STANDINGS[classe]); }
  sheet.clear();
  clearDataValidations(sheet);

  var teamsByPool = getTeams(classe);
  var row = 1;

  // Titre principal.
  sheet.getRange(row, 1).setValue('CLASSEMENTS — CLASSE ' + classe);
  sheet.getRange(row, 1, 1, 13).merge();
  sheet.getRange(row, 1).setFontSize(14).setFontWeight('bold')
    .setBackground(COLOR_HEADER).setFontColor(COLOR_HEADER_TX)
    .setHorizontalAlignment('center');
  row += 2;

  // Conserve les vainqueurs/2es de pool pour les Étapes B et C.
  var firsts  = [];   // 1er de chaque pool
  var seconds = [];   // 2e de chaque pool
  var poolStatsByTeam = {};   // team -> stats globales (pour affichage B/C)

  // -------- SECTIONS 1-3 : un classement par pool --------
  POOLS.forEach(function (p) {
    var teams = teamsByPool[p] || [];
    var poolGames = games.filter(function (g) { return g.pool === p; });

    var standings = calculatePoolStandings(poolGames, teams);

    row = writePoolSection(sheet, row, classe, p, standings, poolGames);

    standings.forEach(function (s) {
      poolStatsByTeam[s.team] = s;
      if (s.rank === 1) { firsts.push({ team: s.team, pool: p }); }
      if (s.rank === 2) { seconds.push({ team: s.team, pool: p }); }
    });

    row += 1;  // espace entre sections
  });

  // -------- SECTION 4 : Étape C — classement des 1ers --------
  var firstTeams = firsts.map(function (f) { return f.team; });
  var orderedFirsts = calculateStep(firstTeams, games, true);
  row = writeAdvancementSection(
    sheet, row, classe,
    'SECTION 4 — CLASSEMENT DES 1ers (Positions 1-2-3) — Étape C',
    orderedFirsts, games, poolStatsByTeam, firsts, 1);

  row += 1;

  // -------- SECTION 5 : Étape B — meilleur 2e (position 4) --------
  var secondTeams = seconds.map(function (s) { return s.team; });
  var orderedSeconds = calculateStep(secondTeams, games, true);
  row = writeAdvancementSection(
    sheet, row, classe,
    'SECTION 5 — MEILLEUR 2e (Position 4) — Étape B',
    orderedSeconds, games, poolStatsByTeam, seconds, 4);

  row += 1;

  // -------- RÉCAPITULATIF DEMI-FINALES --------
  writeSemifinalSummary(sheet, row, classe, orderedFirsts, orderedSeconds);

  // Largeurs de colonnes (13 colonnes).
  var widths = [55, 170, 45, 45, 45, 80, 80, 75, 75, 105, 105, 120, 30];
  for (var c = 0; c < widths.length; c++) { sheet.setColumnWidth(c + 1, widths[c]); }
  sheet.setFrozenRows(1);
}

// Info-bulles (survol de l'en-tête) des feuilles Classements — comme dans les
// feuilles Résultats, le texte est autonome pour qu'un responsable comprenne
// chaque abréviation sans aller voir l'onglet Aide. Clés = numéro de colonne.
var POOL_HEADER_NOTES = {
  1:  'RANG — Position de l\'équipe dans son pool, après application des bris d\'égalité ' +
      '(Art. 42.11). 1er et 2e sont surlignés.',
  3:  'PJ — Parties Jouées dans le pool.',
  4:  'V — Victoires.',
  5:  'D — Défaites.',
  6:  'PP — Points POUR : total des points marqués par l\'équipe dans le pool.',
  7:  'PC — Points CONTRE : total des points alloués (encaissés) par l\'équipe dans le pool.',
  8:  'MO — Manches OFFENSIVES jouées (à la batte), en fractions de tiers (⅓/⅔) si la partie ' +
      's\'est terminée hâtivement (walk-off / mercy). Dénominateur du ratio offensif RO.',
  9:  'MD — Manches DÉFENSIVES jouées (au champ), en fractions de tiers (⅓/⅔) si fin hâtive. ' +
      'Dénominateur du ratio défensif RD.',
  10: 'RD — Ratio DÉFENSIF = PC / MD (points contre par manche défensive). Le plus BAS est le ' +
      'meilleur. 2e critère de bris d\'égalité (Art. 42.11), après la fiche tête-à-tête.',
  11: 'RO — Ratio OFFENSIF = PP / MO (points pour par manche offensive). Le plus HAUT est le ' +
      'meilleur. 3e critère de bris d\'égalité (Art. 42.11).',
  12: 'AVANCEMENT — Qualification déduite du rang (ex. 1er de pool, meilleur 2e, etc.).'
};

var ADV_HEADER_NOTES = {
  1:  'POSITION — Rang inter-pool. Étape C : positions 1-2-3 (les 1ers de pool). ' +
      'Étape B : meilleur 2e = position 4.',
  3:  'POOL — Pool d\'origine de l\'équipe.',
  4:  'V — Victoires (sur TOUTES les parties de pool de l\'équipe).',
  5:  'D — Défaites (sur toutes les parties de pool).',
  6:  'RD — Ratio DÉFENSIF = PC / MD (le plus BAS est le meilleur). Calculé sur toutes les ' +
      'parties de pool (forfaits exclus des ratios).',
  7:  'RO — Ratio OFFENSIF = PP / MO (le plus HAUT est le meilleur). Calculé sur toutes les ' +
      'parties de pool (forfaits exclus des ratios).',
  8:  'PP — Points POUR (sur toutes les parties de pool).',
  9:  'PC — Points CONTRE (sur toutes les parties de pool).',
  10: 'NOTE — Avertissement si une vérification manuelle est requise (Priorité 4 « manches ' +
      'en avance » non automatisée — voir la feuille Manches_Détail).'
};

/**
 * Applique des info-bulles (setNote) aux cellules d'en-tête d'une ligne donnée.
 * @param {Sheet}  sheet
 * @param {number} headerRow   numéro de ligne de l'en-tête
 * @param {Object} notesByCol  { numéroColonne: texte }
 */
function applyHeaderNotes(sheet, headerRow, notesByCol) {
  Object.keys(notesByCol).forEach(function (col) {
    sheet.getRange(headerRow, parseInt(col, 10)).setNote(notesByCol[col]);
  });
}

/**
 * Une partie est-elle allée en manches supplémentaires ? (marquée par le
 * registraire via Type de fin = "Supplémentaires"). Sert à signaler, dans les
 * classements, que les ratios incluent des points de supplémentaires à exclure
 * manuellement selon la Note 4 (Art. 42.11).
 */
function gameIsSupp(g) {
  return g && g.type === 'Supplémentaires';
}

/**
 * Écrit une section de classement de pool. Retourne la prochaine ligne libre.
 * @param {Array} poolGames  parties du pool — pour signaler les manches supplémentaires.
 */
function writePoolSection(sheet, startRow, classe, pool, standings, poolGames) {
  var row = startRow;

  // Titre de section.
  sheet.getRange(row, 1).setValue('POOL ' + pool + ' — CLASSE ' + classe);
  sheet.getRange(row, 1, 1, 13).merge();
  sheet.getRange(row, 1).setFontWeight('bold').setBackground(poolColor(pool))
    .setHorizontalAlignment('center');
  row++;

  // En-têtes (13 colonnes).
  var headers = ['Rang', 'Équipe', 'PJ', 'V', 'D', 'PP',
                 'PC', 'MO', 'MD',
                 'RD', 'RO', 'Avancement', ''];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet.getRange(row, 1, 1, headers.length));
  applyHeaderNotes(sheet, row, POOL_HEADER_NOTES);
  row++;

  // Lignes d'équipes.
  standings.forEach(function (s) {
    var advancement = advancementLabel(s.rank);
    var raRatioDisplay = s.defInn > 0 ? s.raRatio.toFixed(3) : '—';
    var rsRatioDisplay = s.offInn > 0 ? s.rsRatio.toFixed(3) : '—';
    var rowData = [
      s.rank, s.team, s.pj, s.v, s.d, s.rs, s.ra,
      formatFraction(s.offInn), formatFraction(s.defInn),
      raRatioDisplay, rsRatioDisplay, advancement, ''
    ];
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);

    // Couleur selon le rang (1er = vert, 2e = bleu clair).
    if (s.rank === 1) {
      sheet.getRange(row, 1, 1, headers.length).setBackground(COLOR_FIRST);
    } else if (s.rank === 2) {
      sheet.getRange(row, 1, 1, headers.length).setBackground(COLOR_SECOND);
    } else {
      sheet.getRange(row, 1, 1, headers.length).setBackground(COLOR_CALC);
    }
    row++;
  });

  // Avertissement Note 4 : si une partie du pool est allée en supplémentaires,
  // les ratios RD/RO ci-dessus incluent des points de supplémentaires à exclure.
  var suppGames = (poolGames || []).filter(gameIsSupp);
  if (suppGames.length > 0) {
    var details = suppGames.map(function (g) {
      return g.local + ' vs ' + g.visiteur +
             (g.partie ? ' (partie #' + g.partie + ')' : '');
    }).join(' ; ');
    sheet.getRange(row, 1, 1, 13).merge();
    sheet.getRange(row, 1)
      .setValue('⚠ Manches supplémentaires : ' + details + '. Les ratios RD/RO ci-dessus ' +
                'incluent les points des supplémentaires — à recalculer à la main en les ' +
                'EXCLUANT (Note 4, Art. 42.11) si un bris d\'égalité par ratio est en jeu. ' +
                'Voir la feuille Manches_Détail.')
      .setWrap(true).setBackground(COLOR_INPUT).setVerticalAlignment('top');
    row++;
  }

  return row;
}

/**
 * Écrit une section d'avancement (Étape B ou C). Retourne la prochaine ligne.
 * @param {number} basePosition  position de départ (1 pour Étape C, 4 pour Étape B)
 */
function writeAdvancementSection(sheet, startRow, classe, title, orderedTeams,
                                 games, poolStatsByTeam, poolInfo, basePosition) {
  var row = startRow;

  // Titre.
  sheet.getRange(row, 1).setValue(title);
  sheet.getRange(row, 1, 1, 13).merge();
  sheet.getRange(row, 1).setFontWeight('bold').setBackground(COLOR_SECTION)
    .setHorizontalAlignment('center');
  row++;

  // En-têtes.
  var headers = ['Position', 'Équipe', 'Pool', 'V', 'D',
                 'RD', 'RO', 'PP',
                 'PC', 'Note', ''];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet.getRange(row, 1, 1, headers.length));
  applyHeaderNotes(sheet, row, ADV_HEADER_NOTES);
  row++;

  // Map team -> pool d'origine.
  var poolOf = {};
  poolInfo.forEach(function (info) { poolOf[info.team] = info.pool; });

  // Détecte si Priorité 4 a été nécessaire.
  var needsManual = orderedTeams.__needsManualCheck === true;

  orderedTeams.forEach(function (team, idx) {
    // Recalcule les stats de cette équipe sur toutes ses parties de pool
    // (forfaits exclus des ratios pour B/C).
    var teamGames = games.filter(function (g) {
      return g.local === team || g.visiteur === team;
    });
    var st = computeTeamStats(team, teamGames, true);

    // Note : vérif. manuelle Priorité 4 et/ou ratio incluant des supplémentaires
    // (Note 4) — les deux peuvent coexister.
    var noteParts = [];
    if (needsManual) { noteParts.push('⚠ Vérif. manuelle (Manches_Détail)'); }
    if (teamGames.some(gameIsSupp)) { noteParts.push('⚠ Supp. — ratio à ajuster (Note 4)'); }
    var note = noteParts.join(' ');
    var rowData = [
      basePosition + idx, team, poolOf[team] || '',
      st.v, st.d,
      isFinite(st.raRatio) ? round3(st.raRatio) : '—',
      round3(st.rsRatio),
      st.rs, st.ra, note, ''
    ];
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);

    // Couleur : positions qualificatives en vert (Étape C) ou bleu (Étape B 1ère place).
    if (basePosition === 1) {
      sheet.getRange(row, 1, 1, headers.length).setBackground(COLOR_FIRST);
    } else {
      // Étape B : seule la 1ère (meilleur 2e) se qualifie -> bleu ; autres gris.
      sheet.getRange(row, 1, 1, headers.length)
        .setBackground(idx === 0 ? COLOR_SECOND : COLOR_CALC);
    }
    row++;
  });

  return row;
}

/**
 * Récapitulatif des demi-finales : croise positions 1-2-3-4.
 */
function writeSemifinalSummary(sheet, startRow, classe, orderedFirsts, orderedSeconds) {
  var row = startRow;

  sheet.getRange(row, 1).setValue('DEMI-FINALES — CLASSE ' + classe);
  sheet.getRange(row, 1, 1, 6).merge();
  sheet.getRange(row, 1).setFontWeight('bold').setBackground(COLOR_HEADER)
    .setFontColor(COLOR_HEADER_TX).setHorizontalAlignment('center');
  row++;

  var p1 = orderedFirsts[0]  || '—';
  var p2 = orderedFirsts[1]  || '—';
  var p3 = orderedFirsts[2]  || '—';
  var p4 = orderedSeconds[0] || '—';   // meilleur 2e

  var lines = [
    ['Position 1', p1],
    ['Position 2', p2],
    ['Position 3', p3],
    ['Position 4 (meilleur 2e)', p4],
    ['', ''],
    ['Demi-finale 1', p1 + '  vs  ' + p4],
    ['Demi-finale 2', p2 + '  vs  ' + p3]
  ];
  lines.forEach(function (ln) {
    sheet.getRange(row, 1).setValue(ln[0]).setFontWeight('bold');
    sheet.getRange(row, 2, 1, 3).merge();
    sheet.getRange(row, 2).setValue(ln[1]);
    row++;
  });
  return row;
}

// ============================================================================
//  EFFACER LES RÉSULTATS
// ============================================================================

/**
 * Efface les scores saisis et les valeurs calculées dans les feuilles Résultats,
 * tout en conservant les matchs générés (équipes, pools).
 */
function clearResults() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Effacer les résultats',
    'Effacer tous les scores et calculs des feuilles Résultats A et B ? ' +
    'Les matchs générés (équipes) seront conservés.',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) { return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  CLASSES.forEach(function (classe) {
    var sheet = ss.getSheetByName(SHEET_RESULTS[classe]);
    if (!sheet) { return; }
    var last = sheet.getLastRow();
    if (last < 2) { return; }
    var n = last - 1;
    // Efface H, I (scores), J (Équipe Locale), K, L (manches/retraits), N..R (calculs).
    sheet.getRange(2, 8, n, 2).clearContent();    // H, I
    sheet.getRange(2, 10, n, 1).clearContent();   // J : Équipe Locale
    sheet.getRange(2, 11, n, 2).clearContent();   // K, L
    sheet.getRange(2, 14, n, 5).clearContent();   // N..R
    // Réinitialise le type de fin à "Normal".
    var normalCol = [];
    for (var i = 0; i < n; i++) { normalCol.push(['Normal']); }
    sheet.getRange(2, 13, n, 1).setValues(normalCol);  // M : Type de fin
  });

  ui.alert('Résultats effacés.');
}

// ============================================================================
//  SIMULATION DE RÉSULTATS
// ============================================================================

/**
 * Simule des résultats de match couvrant tous les cas possibles : Normal,
 * Mercy, Forfait, manches supplémentaires, walk-off. Utilise l'horaire DÉJÀ
 * généré dans les feuilles Résultats (via "Générer les matchs" à partir de
 * la Configuration) — ne touche PAS à la Configuration ni à l'horaire des
 * matchs, seulement aux scores. Les scores sont injectés par position
 * (1re partie du pool, 2e partie du pool, etc.), peu importe les noms
 * d'équipes réels.
 *
 * Scénarios de classement par pool (M1 = 1re partie du pool, M6 = 6e) :
 *  Classe A Pool 1 — Égalité 3 équipes (M1-M2-M4 ... à 2V-1D), une équipe 0V-3D
 *  Classe A Pool 2 — un net 1er (3V-0D), match en manche supplémentaire (7e)
 *  Classe A Pool 3 — un net 1er (3V-0D), un forfait, une victoire walk-off
 *  Classe B Pool 1 — Égalité 3 équipes (2V-1D), une équipe 0V-3D
 *  Classe B Pool 2 — un net 1er, manches supplémentaires (7e walk-off et 8e)
 *  Classe B Pool 3 — deux égalités à 2 équipes, un forfait
 *
 * À réinitialiser (menu "Effacer les résultats") avant le vrai tournoi.
 */
function simulateMatchResults() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Vérifie que l'horaire a déjà été généré (Configuration -> "Générer les matchs").
  var notReady = CLASSES.some(function (classe) {
    var sheet = ss.getSheetByName(SHEET_RESULTS[classe]);
    return !sheet || sheet.getLastRow() < 2 ||
      String(sheet.getRange(2, 6).getValue()).trim() === '';
  });
  if (notReady) {
    ui.alert('🧪 Simuler résultats de match',
      'Aucun match généré dans les feuilles Résultats. Collez d\'abord l\'horaire dans ' +
      '"Configuration", cliquez "Générer les matchs", puis relancez cette simulation.',
      ui.ButtonSet.OK);
    return;
  }

  var resp = ui.alert(
    '🧪 Simuler résultats de match',
    'Ceci va écraser tous les scores actuels (de l\'horaire déjà généré) avec des ' +
    'résultats fictifs couvrant tous les cas (normal, mercy, forfait, supplémentaire), ' +
    'et recalculer les classements. L\'horaire (équipes, dates, terrains) n\'est PAS ' +
    'modifié.\n\nÀ utiliser seulement pour tester le système — PAS pour le vrai ' +
    'tournoi.\n\nContinuer ?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) { return; }

  // ─── Format des lignes : [scoreÉquipe1, scoreÉquipe2, manches, retraits, type] ───
  // (Équipe 1 = équipe locale pour ces données fictives, écrite ci-dessous dans J.)
  // Les 18 lignes sont appliquées par position : Pool1 M1-M6, Pool2 M1-M6, Pool3 M1-M6
  // (correspond à l'ordre des matchs tel qu'écrit par "Générer les matchs").

  var DATA = {
    'A': [
      // ── Pool 1 ── Égalité 3-équipes (T0, T1, T2 tous 2V-1D) ──────────────
      //  Bris d'égalité : tête-à-tête encore à égalité → ratios RA/Inn
      //  T1 ressort 1er (RA 0.500), T2 2e, T0 3e. T3 dernier (0V-3D).
      [ 5,  4, 6, 0, 'Normal'],   // M1 T0 vs T1  → T0 gagne (normal)
      [ 3,  7, 6, 0, 'Normal'],   // M2 T0 vs T2  → T2 gagne
      [ 6,  2, 6, 0, 'Normal'],   // M3 T0 vs T3  → T0 gagne
      [ 8,  1, 6, 0, 'Normal'],   // M4 T1 vs T2  → T1 gagne
      [ 9,  0, 5, 0, 'Mercy' ],   // M5 T1 vs T3  → T1 gagne (Mercy 5e manche)
      [ 5,  3, 6, 0, 'Normal'],   // M6 T2 vs T3  → T2 gagne

      // ── Pool 2 ── T2 1er net ; match en supplémentaire (7e) ──────────────
      //  T2 3V-0D, T0 2V-1D, T1 1V-2D, T3 0V-3D.
      [ 4,  3, 6, 2, 'Normal'],   // M1 T0 vs T1  → T0 gagne (walk-off bas 6e, 2 ret.)
      [ 2,  8, 6, 0, 'Normal'],   // M2 T0 vs T2  → T2 gagne
      [ 7,  1, 6, 0, 'Normal'],   // M3 T0 vs T3  → T0 gagne
      [ 3,  4, 7, 0, 'Supplémentaires'],   // M4 T1 vs T2  → T2 gagne (supplémentaire 7e → ⚠ Note 4)
      [11,  1, 5, 0, 'Mercy' ],   // M5 T1 vs T3  → T1 gagne (Mercy 5e manche)
      [ 6,  2, 6, 0, 'Normal'],   // M6 T2 vs T3  → T2 gagne

      // ── Pool 3 ── T2 1er net ; forfait T1 ; walk-off T0 ─────────────────
      //  T2 3V-0D, T0 2V-1D, T3 1V-2D, T1 0V-3D (forfait).
      [ 7,  3, 6, 0, 'Normal'],   // M1 T0 vs T1  → T0 gagne
      [ 4,  9, 6, 0, 'Normal'],   // M2 T0 vs T2  → T2 gagne
      [ 1,  0, 6, 1, 'Normal'],   // M3 T0 vs T3  → T0 gagne (walk-off bas 6e, 1 ret.)
      [ 5,  8, 6, 0, 'Normal'],   // M4 T1 vs T2  → T2 gagne
      [ 0,  7, 6, 0, 'Forfait'],  // M5 T1 vs T3  → T3 gagne (forfait visiteur)
      [ 6,  1, 6, 0, 'Normal'],   // M6 T2 vs T3  → T2 gagne
    ],
    'B': [
      // ── Pool 1 ── Égalité 3-équipes (T0, T1, T3 tous 2V-1D) ─────────────
      //  Tête-à-tête à égalité → ratios : T0 1er, T3 2e, T1 3e. T2 0V-3D.
      [ 4,  7, 6, 0, 'Normal'],   // M1 T0 vs T1  → T1 gagne
      [16,  0, 5, 0, 'Mercy' ],   // M2 T0 vs T2  → T0 gagne (Mercy, 5e manche)
      [ 6,  4, 6, 1, 'Normal'],   // M3 T0 vs T3  → T0 gagne (walk-off bas 6e, 1 ret.)
      [ 8,  2, 6, 0, 'Normal'],   // M4 T1 vs T2  → T1 gagne
      [ 5,  8, 6, 0, 'Normal'],   // M5 T1 vs T3  → T3 gagne
      [ 3,  5, 6, 0, 'Normal'],   // M6 T2 vs T3  → T3 gagne

      // ── Pool 2 ── T3 1er net ; supplémentaires 7e (walk-off) et 8e ───────
      //  T3 3V-0D, T1 2V-1D, T0 1V-2D, T2 0V-3D.
      [ 6,  8, 6, 0, 'Normal'],   // M1 T0 vs T1  → T1 gagne
      [ 5,  4, 7, 1, 'Supplémentaires'],   // M2 T0 vs T2  → T0 gagne (walk-off supp. 7e → ⚠ Note 4)
      [ 1,  2, 8, 0, 'Supplémentaires'],   // M3 T0 vs T3  → T3 gagne (supplémentaire 8e → ⚠ Note 4)
      [ 9,  0, 4, 0, 'Mercy' ],   // M4 T1 vs T2  → T1 gagne (Mercy 4e manche)
      [ 4,  5, 6, 0, 'Normal'],   // M5 T1 vs T3  → T3 gagne
      [ 3,  8, 6, 0, 'Normal'],   // M6 T2 vs T3  → T3 gagne

      // ── Pool 3 ── Égalité T0/T1 (2V-1D) + égalité T2/T3 (1V-2D) + forfait
      //  T0 1er (bat T1 tête-à-tête), T1 2e, T3 3e (bat T2 forfait), T2 4e.
      [ 8,  3, 6, 0, 'Normal'],   // M1 T0 vs T1  → T0 gagne
      [ 2,  6, 6, 0, 'Normal'],   // M2 T0 vs T2  → T2 gagne
      [ 4,  3, 6, 2, 'Normal'],   // M3 T0 vs T3  → T0 gagne (walk-off bas 6e, 2 ret.)
      [ 7,  3, 6, 0, 'Normal'],   // M4 T1 vs T2  → T1 gagne
      [ 6,  0, 6, 0, 'Normal'],   // M5 T1 vs T3  → T1 gagne
      [ 0,  7, 6, 0, 'Forfait'],  // M6 T2 vs T3  → T3 gagne (forfait visiteur)
    ]
  };

  // Validation étendue pour manches supplémentaires (1-9).
  var inningRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 9).setAllowInvalid(false).build();

  CLASSES.forEach(function (classe) {
    var sheet = ss.getSheetByName(SHEET_RESULTS[classe]);
    if (!sheet) { return; }
    var allRows = DATA[classe];
    var n = Math.min(allRows.length, sheet.getLastRow() - 1);
    var rows = allRows.slice(0, n);
    if (n === 0) { return; }

    // Met à jour la validation manches si nécessaire.
    sheet.getRange(2, 11, n, 1).setDataValidation(inningRule);

    // Écrit les scores (H, I) et paramètres de fin (K, L, M) en deux blocs.
    var scores = rows.map(function (r) { return [r[0], r[1]]; });
    var params = rows.map(function (r) { return [r[2], r[3], r[4]]; });
    sheet.getRange(2, 8, n, 2).setValues(scores);   // H, I
    sheet.getRange(2, 11, n, 3).setValues(params);  // K, L, M

    // Équipe Locale (J) = Équipe 1 pour ces données fictives (correspond au
    // scénario narratif des commentaires ci-dessus, ex. "T0 vs T1 → ... walk-off").
    var teamA = sheet.getRange(2, 6, n, 1).getValues();
    sheet.getRange(2, 10, n, 1).setValues(teamA);
  });

  // Recalcule les classements immédiatement.
  calculateStandings();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Résultats simulés (normal, mercy, forfait, supplémentaire) sur l\'horaire existant. ' +
    'Classements mis à jour.\nUtilisez "Effacer les résultats" avant le vrai tournoi.',
    '🧪 Simulation', 8);
}

// ============================================================================
//  UTILITAIRES
// ============================================================================

/**
 * Convertit un décimal de manches en format lisible (4.333 -> "4 1/3").
 * Gère les tiers (1/3 et 2/3) ; arrondit au tiers le plus proche.
 */
function formatFraction(decimal) {
  if (decimal === '' || decimal === null || decimal === undefined) { return ''; }
  var num = Number(decimal);
  if (isNaN(num)) { return ''; }

  var whole = Math.floor(num + 1e-9);
  var frac = num - whole;

  // Arrondit la fraction au tiers le plus proche.
  var thirds = Math.round(frac * 3);
  if (thirds === 3) { whole += 1; thirds = 0; }

  if (thirds === 0) { return String(whole); }
  if (thirds === 1) { return (whole === 0 ? '' : whole + ' ') + '⅓'; }
  if (thirds === 2) { return (whole === 0 ? '' : whole + ' ') + '⅔'; }
  return String(num);
}

/** Arrondit à 3 décimales pour l'affichage des ratios. */
function round3(x) {
  return Math.round(x * 1000) / 1000;
}

/** Étiquette d'avancement selon le rang dans un pool. */
function advancementLabel(rank) {
  if (rank === 1) { return '1er'; }
  if (rank === 2) { return '2e'; }
  if (rank === 3) { return '3e'; }
  return '4e';
}

/** Couleur associée à un pool. */
function poolColor(pool) {
  if (pool === 1) { return COLOR_POOL_1; }
  if (pool === 2) { return COLOR_POOL_2; }
  return COLOR_POOL_3;
}

/** Récupère une feuille ou la crée si absente. */
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); }
  return sheet;
}

/** Applique le style d'en-tête à une plage. */
function styleHeader(range) {
  range.setBackground(COLOR_HEADER)
       .setFontColor(COLOR_HEADER_TX)
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setWrap(true);
}

/** Supprime toutes les validations de données d'une feuille. */
function clearDataValidations(sheet) {
  var maxR = sheet.getMaxRows();
  var maxC = sheet.getMaxColumns();
  sheet.getRange(1, 1, maxR, maxC).clearDataValidations();
}

/** Supprime la feuille par défaut vide ("Sheet1"/"Feuille1") si présente. */
function removeDefaultSheet(ss) {
  ['Sheet1', 'Feuille1', 'Feuille 1'].forEach(function (name) {
    var s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch (e) { /* ignore */ }
    }
  });
}

/** Réordonne les feuilles dans un ordre logique. */
function reorderSheets(ss) {
  var order = [
    SHEET_HELP,
    SHEET_CONFIG,
    SHEET_RESULTS['A'], SHEET_RESULTS['B'],
    SHEET_STANDINGS['A'], SHEET_STANDINGS['B'],
    SHEET_INN_DETAIL
  ];
  order.forEach(function (name, idx) {
    var s = ss.getSheetByName(name);
    if (s) {
      ss.setActiveSheet(s);
      ss.moveActiveSheet(idx + 1);
    }
  });
  // Replace le focus sur Configuration.
  var cfg = ss.getSheetByName(SHEET_CONFIG);
  if (cfg) { ss.setActiveSheet(cfg); }
}
