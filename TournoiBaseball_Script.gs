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
var SHEET_LEDGER      = 'Grand livre';

// Numéro de version de l'application (le code lui-même). À incrémenter à la main
// lors d'un changement notable ; s'affiche dans le pied de page de l'affichage
// public (doGet) pour savoir quelle version est déployée sur le lien Facebook.
var APP_VERSION = '1.1.1';

var CLASSES = ['A', 'B'];
var POOLS   = [1, 2, 3];
var TEAMS_PER_POOL = 4;
var TOTAL_INNINGS  = 6;   // 13U = 6 manches réglementaires

// Couleurs
var COLOR_INPUT     = '#fff9c4';   // jaune clair  - saisie manuelle
var COLOR_CALC      = '#eeeeee';   // gris clair    - calculé
var COLOR_FIRST     = '#c8e6c9';   // vert          - 1er de pool
var COLOR_SECOND    = '#bbdefb';   // bleu clair    - 2e de pool
var COLOR_WIN       = '#c8e6c9';   // vert          - victoire (Grand livre)
var COLOR_LOSS      = '#ffcdd2';   // rouge clair   - défaite (Grand livre)
var COLOR_LOCAL     = '#e1f5fe';   // bleu très pâle - équipe locale (Grand livre)
var COLOR_VISITOR   = '#ffe0b2';   // orange pâle   - équipe visiteuse (Grand livre)
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
    .addItem('Initialiser les feuilles…', 'createAllSheets')
    .addItem('Initialiser (conserver Configuration)…', 'createAllSheetsKeepConfig')
    .addItem('Générer les matchs…', 'generateGames')
    .addItem('🔄 Mettre à jour les classements', 'calculateStandings')
    .addItem('📒 Grand livre des matchs', 'buildLedgerSheet')
    .addSeparator()
    .addItem('⚡ Activer la mise à jour auto', 'installTriggers')
    .addItem('Effacer les résultats…', 'clearResults')
    .addSeparator()
    .addItem('📦 Exporter Résultats + Classements + Grand livre (ZIP de TSV)', 'exportSheetsToZip')
    .addItem('📱 Lien affichage public (Facebook)', 'showPublicUrl')
    .addSeparator()
    .addItem('🧪 Simuler résultats de match…', 'simulateMatchResults')
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
 * Deux déclencheurs possibles :
 *   - Édition des colonnes de saisie H..O (8..15) d'une feuille "Résultats" (les
 *     colonnes A..G copiées et P..T calculées sont ignorées) — recalcul complet.
 *   - Édition de la colonne « Forcer 2e » (13) d'une feuille "Classements" : forçage
 *     admin du 2e d'un pool (Note 5) — reconstruit le classement de la classe via
 *     recalcStandingsOnly (sans toucher aux Résultats).
 *
 * Conçu pour un usage multi-postes fiable :
 *   - On ne recalcule QUE lorsque la ligne de la partie est COMPLÈTE
 *     (isRowComplete) : pas de classement qui bouge en pleine saisie, donc pas
 *     de confusion pour ceux qui regardent. Exception : l'effacement d'un score
 *     (qui RETIRE une partie du classement) déclenche aussi un rafraîchissement.
 *   - Tout le recalcul est protégé par un VERROU DE DOCUMENT : deux postes ne
 *     peuvent jamais reconstruire la feuille Classements en même temps (sinon un
 *     responsable verrait un classement à moitié reconstruit). Acquisition
 *     défensive : si LockService est indisponible, on recalcule sans verrou
 *     plutôt que de planter (dégradation gracieuse).
 *   - On n'écrit que dans Classements + les colonnes calculées P..T de la ligne
 *     saisie (écriture ciblée) — jamais toute la feuille Résultats, pour ne pas
 *     gêner la frappe d'un autre poste. Le menu "Mettre à jour les classements"
 *     reste le refresh complet/autoritaire (et le secours pour les corrections
 *     en lot).
 *
 * Les écritures programmées (P..T, Classements) ne redéclenchent pas le handler :
 * aucune boucle possible.
 */
function handleResultEdit(e) {
  if (!e || !e.range) { return; }

  var sheet = e.range.getSheet();
  var name = sheet.getName();
  var startCol = e.range.getColumn();
  var endCol   = startCol + e.range.getNumColumns() - 1;

  // --- Édition d'une saisie admin/registraire d'une feuille Classements ? ---
  // Deux colonnes déclenchent un recalcul du classement de la classe :
  //   • col 13 « Forcer 2e » — représentant d'un pool au Meilleur 2e (Note 5) ;
  //   • col 24 « Forcer rang » — ordre manuel de la Priorité 4 (Art. 42.11).
  // buildStandingsSheet relit ces deux forçages (readSecondOverrides / readForcedRanks)
  // AVANT de tout réécrire — aucune écriture dans Résultats.
  var standingsClasse = null;
  CLASSES.forEach(function (c) {
    if (SHEET_STANDINGS[c] === name) { standingsClasse = c; }
  });
  if (standingsClasse) {
    var hitsForce2e   = (startCol <= 13 && endCol >= 13);
    var hitsForceRank = (startCol <= 24 && endCol >= 24);
    if (hitsForce2e || hitsForceRank) { recalcStandingsOnly(e, standingsClasse); }
    return;
  }

  // --- Sinon : édition d'une feuille "Résultats" ? (sinon on ignore) ---
  var classe = null;
  CLASSES.forEach(function (c) {
    if (SHEET_RESULTS[c] === name) { classe = c; }
  });
  if (!classe) { return; }

  // L'édition doit toucher au moins une colonne de saisie H..O (8..15).
  if (endCol < 8 || startCol > 15) { return; }

  // Lignes éditées (ignore l'en-tête, ligne 1).
  var startRow = Math.max(e.range.getRow(), 2);
  var endRow   = e.range.getRow() + e.range.getNumRows() - 1;
  if (endRow < 2) { return; }

  // ---- Décide s'il faut recalculer (lecture légère, AVANT de prendre le verrou) ----
  var recalc = false;

  // (a) Au moins une ligne éditée est-elle désormais complète ?
  //     On lit H..O (8..15) : r[5] = Manches prévues, r[6] = Type de fin,
  //     r[7] = Pointage régl. (suppl.).
  var hm = sheet.getRange(startRow, 8, endRow - startRow + 1, 8).getValues();
  for (var i = 0; i < hm.length; i++) {
    var r = hm[i];
    if (isRowComplete({ scoreA: r[0], scoreB: r[1], local: r[2],
                        manches: r[3], retraits: r[4], type: r[6], suppTie: r[7] })) {
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

  // Retour visuel immédiat au poste : la ligne est complète, le calcul suit. Affiché
  // AVANT de prendre le verrou pour que le poste ne perçoive pas d'attente « morte ».
  (e.source || SpreadsheetApp.getActiveSpreadsheet())
    .toast('Calcul du classement ' + classe + ' en cours…', 'Tournoi Baseball', 3);

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

    // Colonnes calculées P..T : seulement les lignes éditées (ciblé).
    var byRow = {};
    games.forEach(function (g) { byRow[g.rowIndex] = g; });
    for (var row = startRow; row <= endRow; row++) {
      writeRowCalc(sheet, row, byRow[row]);
    }

    buildStandingsSheet(ss, classe, games);
    // NB : le Grand livre (buildLedgerSheet) n'est PAS reconstruit en direct — c'est
    // l'étape la plus lourde (relit les 2 classes + reformate cellule par cellule) et
    // les registraires suivent la feuille Classements pendant la saisie, pas le Grand
    // livre. Il reste à jour via le menu « 📒 Grand livre des matchs » et « Mettre à
    // jour les classements » (calculateStandings). Cf. recalcStandingsOnly (idem).
    ss.toast('Classement ' + classe + ' mis à jour.', 'Tournoi Baseball', 3);
  } finally {
    if (lock) { lock.releaseLock(); }
  }
}

/**
 * Reconstruit UNIQUEMENT la feuille Classements d'une classe (sans toucher aux
 * colonnes calculées P..T des Résultats), sous verrou de document. Déclenché par un
 * forçage saisi dans Classements : « Forcer 2e » (col 13, Note 5) ou « Forcer rang »
 * (col 24, Priorité 4). buildStandingsSheet relit ces forçages (readSecondOverrides /
 * readForcedRanks) AVANT de reconstruire, donc la saisie est prise en compte
 * immédiatement. Acquisition défensive du verrou (cf. handleResultEdit). Les écritures
 * programmées de buildStandingsSheet ne redéclenchent pas le handler.
 */
function recalcStandingsOnly(e, classe) {
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
      return;   // verrou occupé : un autre poste recalcule déjà
    }
  }

  try {
    var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
    var games = getGameResults(classe);
    buildStandingsSheet(ss, classe, games);
    // Grand livre non reconstruit en direct (cf. note dans handleResultEdit).
    ss.toast('Classement ' + classe + ' mis à jour (forçage).', 'Tournoi Baseball', 3);
  } finally {
    if (lock) { lock.releaseLock(); }
  }
}

// ============================================================================
//  CRÉATION DES FEUILLES
// ============================================================================

/**
 * Crée (ou réinitialise) TOUTES les feuilles nécessaires au tournoi, Configuration
 * comprise (l'horaire collé est effacé). Item de menu "Initialiser les feuilles".
 */
function createAllSheets() {
  var ui = SpreadsheetApp.getUi();
  // Double confirmation : c'est l'action la plus destructive (efface l'horaire
  // ET tous les scores, sans annulation possible).
  var resp1 = ui.alert('Initialiser les feuilles',
    'Réinitialiser TOUTES les feuilles ? L\'horaire collé dans Configuration et ' +
    'TOUS les scores seront effacés.',
    ui.ButtonSet.YES_NO);
  if (resp1 !== ui.Button.YES) { return; }
  var resp2 = ui.alert('Initialiser les feuilles',
    '⚠ Action irréversible : ceci EFFACE l\'horaire du tournoi ET tous les scores ' +
    'déjà saisis. Vraiment continuer ?',
    ui.ButtonSet.YES_NO);
  if (resp2 !== ui.Button.YES) { return; }
  rebuildSheets(false);
}

/**
 * Comme "Initialiser les feuilles", mais PRÉSERVE la feuille "Configuration"
 * (l'horaire déjà collé). Pratique pour régénérer Aide / Résultats / Classements
 * après un nouveau collage de code, sans avoir à recoller l'horaire du tournoi.
 * Note : les feuilles Résultats A/B sont quand même reconstruites — les scores
 * déjà saisis sont effacés ; relancez "Générer les matchs" ensuite.
 */
function createAllSheetsKeepConfig() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Initialiser (conserver Configuration)',
    'Reconstruire les feuilles Résultats A/B ? La Configuration (horaire) est ' +
    'conservée, mais TOUS les scores déjà saisis seront effacés. Continuer ?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) { return; }
  rebuildSheets(true);
}

/**
 * Reconstruit les feuilles du tournoi.
 * @param {boolean} keepConfig  true => ne touche pas à "Configuration" si elle
 *                              existe déjà (préserve l'horaire collé) ; false =>
 *                              réinitialise aussi "Configuration".
 */
function rebuildSheets(keepConfig) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  createHelpSheet(ss);
  if (keepConfig && ss.getSheetByName(SHEET_CONFIG)) {
    // Configuration conservée telle quelle (horaire déjà collé préservé).
  } else {
    createConfigSheet(ss);
  }
  CLASSES.forEach(function (c) { createResultsSheet(ss, c); });
  CLASSES.forEach(function (c) { createStandingsSheet(ss, c); });
  createLedgerSheet(ss);

  // Ancien onglet "Manches_Détail" (saisie du pointage par manche) : retiré du système
  // — la Priorité 4 se calcule désormais sur papier puis se saisit dans la colonne
  // "Forcer rang" des Classements. On supprime l'onglet s'il subsiste d'un ancien
  // déploiement (nettoyage idempotent).
  var oldInnDetail = ss.getSheetByName('Manches_Détail');
  if (oldInnDetail) { ss.deleteSheet(oldInnDetail); }

  // Supprime la feuille par défaut "Sheet1" / "Feuille1" si elle est vide et inutilisée.
  removeDefaultSheet(ss);

  // Réordonne les feuilles dans un ordre logique.
  reorderSheets(ss);

  ss.toast(
    keepConfig
      ? 'Feuilles réinitialisées (Configuration conservée). Lancez "Générer les ' +
        'matchs" puis "⚡ Activer la mise à jour auto".'
      : 'Feuilles initialisées. Collez l\'horaire dans "Configuration" puis "Générer les matchs".',
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
    'Retraits en fin', 'Manches prévues', 'Type de fin', 'Pointage régl. (suppl.)',
    'Gagnant', 'MO Éq.1', 'MD Éq.1', 'MO Éq.2', 'MD Éq.2'
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
        'fractions de manches des colonnes Q à T.',
    11: 'MANCHES COMPLÈTES — Numéro de la dernière manche jouée (normalement 6, le ' +
        'nombre réglementaire en 13U). Si la partie s\'arrête avant via la règle du ' +
        'marqueur ("Mercy"), indiquer la manche où elle s\'est arrêtée (ex. 5). Si elle ' +
        'se prolonge en supplémentaire, indiquer la dernière manche jouée (ex. 7, 8...). ' +
        'NE PAS confondre avec "Manches prévues" (col. M).',
    12: 'RETRAITS EN FIN — Nombre de retraits (0, 1 ou 2) déjà comptés dans la dernière ' +
        'demi-manche AU MOMENT où la partie s\'est terminée. Pertinent SEULEMENT si ' +
        'l\'équipe locale a gagné en frappant le point décisif avant d\'avoir complété ' +
        'ses 3 retraits (victoire "walk-off") ; sinon, laisser à 0. Cette valeur donne la ' +
        'fraction de manche utilisée dans les colonnes Q à T : 1 retrait = ⅓ de manche, ' +
        '2 retraits = ⅔ de manche.\n' +
        'Ex. : score final 4-3, la locale gagne dans le bas de la 6e avec 1 retrait ' +
        '(K=6, L=1) → la locale obtient 5⅓ manches OFFENSIVES (pas 6, sa dernière manche ' +
        'au bâton a été interrompue) et la visiteuse 5⅓ manches DÉFENSIVES (pas 6, elle ' +
        'n\'a pas eu le temps de compléter ses 3 retraits).',
    13: 'MANCHES PRÉVUES — Nombre de manches RÉGLEMENTAIRES prévues pour CETTE partie ' +
        '(normalement 6 en 13U ; mettre 5 lors d\'une journée écourtée par la pluie, etc.). ' +
        'Pré-rempli à 6. Sert à créditer le bon nombre de manches DÉFENSIVES à l\'équipe ' +
        'gagnante d\'un Mercy ou d\'un Forfait (la règle BQ donne au gagnant le crédit du ' +
        'nombre réglementaire de manches en défense). À NE PAS confondre avec "Manches ' +
        'complètes" (col. K = dernière manche réellement jouée).',
    14: 'TYPE DE FIN — "Normal" = partie jouée jusqu\'au bout. "Mercy" = arrêtée avant ' +
        'la fin (règle de l\'écart de points). "Forfait" = une équipe ne se présente pas ' +
        'ou abandonne ; indiquez simplement le gagnant en lui donnant un score plus élevé ' +
        '(les chiffres exacts importent peu). Le système attribue alors automatiquement ' +
        'le pointage officiel — 1 point par manche prévue (col. M) au gagnant, 0 au ' +
        'perdant — et crédite le gagnant de "Manches prévues" manches DÉFENSIVES / 0 ' +
        'offensive, et le perdant de 0 manche défensive / "Manches prévues" offensives ' +
        '(Art. 42.11). Rappel (Note 5) : une partie gagnée par forfait n\'est PAS ' +
        'comptabilisée dans les ratios servant au "Meilleur deuxième" (Étapes B et C). ' +
        '"Supplémentaires" = partie prolongée au-delà des manches réglementaires. Dans ce ' +
        'cas, saisissez aussi le "Pointage régl. (suppl.)" (col. O) : le système EXCLUT alors ' +
        'automatiquement les points des manches supplémentaires du ratio de bris d\'égalité ' +
        '(Note 4, Art. 42.11).',
    15: 'POINTAGE RÉGL. (SUPPL.) — À remplir SEULEMENT si Type de fin = "Supplémentaires". ' +
        'C\'est le pointage (égal) atteint à la FIN DES MANCHES RÉGLEMENTAIRES, avant les ' +
        'manches supplémentaires. Ex. : une partie finale 7-5 allée en 7e qui était 5-5 après ' +
        'la 6e → saisir 5. Une partie ne va en supplémentaire que si elle est NULLE au terme ' +
        'du réglementaire (Art. 42.4), donc les deux équipes ont le même total réglementaire : ' +
        'un seul chiffre suffit. Sert à exclure les points des supplémentaires des ratios ' +
        'RD/RO (Note 4). Le gagnant et la fiche V-D, eux, restent basés sur le score FINAL.',
    16: 'GAGNANT — Calculé automatiquement à partir des scores (colonnes H et I). Ne ' +
        'pas modifier à la main, recalculé par "Mettre à jour les classements".',
    17: 'MO ÉQ.1 (Manches Offensives, Équipe 1) — Manches à la batte jouées par l\'Équipe ' +
        '1, calculé automatiquement. Égal au nombre de manches (col. K), SAUF si l\'Équipe 1 ' +
        'est la locale et GAGNE (walk-off, ou elle menait déjà et n\'est pas retournée ' +
        'frapper en bas de la dernière manche) : sa dernière manche au bâton est alors ' +
        'comptée en fraction de tiers (⅓/⅔) selon la col. L, puisqu\'elle ne l\'a pas ' +
        'complétée.',
    18: 'MD ÉQ.1 (Manches Défensives, Équipe 1) — Manches au champ jouées par l\'Équipe 1, ' +
        'calculé automatiquement. Égal au nombre de manches (col. K), SAUF si l\'Équipe 1 est ' +
        'la visiteuse et que l\'équipe locale GAGNE (walk-off, ou la locale menait sans ' +
        'frapper en bas) : la dernière manche en défense de l\'Équipe 1 est alors comptée ' +
        'en fraction de tiers (⅓/⅔) selon la col. L — ou retranchée d\'une manche entière si ' +
        'la locale n\'est pas retournée au bâton — puisqu\'elle n\'a pas eu lieu.',
    19: 'MO ÉQ.2 (Manches Offensives, Équipe 2) — Manches à la batte de l\'Équipe 2, calculé ' +
        'automatiquement. Même logique que la colonne Q (MO Éq.1), appliquée à l\'Équipe 2.',
    20: 'MD ÉQ.2 (Manches Défensives, Équipe 2) — Manches au champ de l\'Équipe 2, calculé ' +
        'automatiquement. Même logique que la colonne R (MD Éq.1), appliquée à l\'Équipe 2.'
  };
  Object.keys(notes).forEach(function (col) {
    sheet.getRange(1, Number(col)).setNote(notes[col]);
  });

  // Largeurs de colonnes.
  var widths = [55, 70, 95, 70, 70, 150, 150, 95, 95, 150, 110, 100, 110, 110, 120, 150, 95, 95, 95, 95];
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }

  sheet.setFrozenRows(1);

  // Note : les 18 rangées de matchs sont remplies par generateGames().
  // Ici on prépare seulement le formatage (validations, couleurs) pour 18 lignes.
  var nRows = POOLS.length * GAME_MATRIX.length;  // 18

  // Validation pour "Type de fin" (col N = 14).
  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Normal', 'Mercy', 'Forfait', 'Supplémentaires'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 14, nRows, 1).setDataValidation(typeRule);

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

  // Validation "Manches prévues" (col M = 13) : 1 à 9 (longueur réglementaire de la partie).
  var regRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 9)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 13, nRows, 1).setDataValidation(regRule);

  // Validation "Pointage régl. (suppl.)" (col O = 15) : entier >= 0. Cellule vide
  // permise (ne sert qu'aux parties "Supplémentaires") — la règle ne se déclenche
  // qu'à la saisie d'une valeur.
  var suppScoreRule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 15, nRows, 1).setDataValidation(suppScoreRule);

  // Validation "Équipe Locale" (col J = 10) : la liste (les 2 équipes du match)
  // dépend de chaque ligne — elle est appliquée ligne par ligne dans generateGames(),
  // une fois les deux équipes de chaque match connues. Tant que la partie n'est
  // pas jouée, on ne sait pas qui sera locale : c'est au registraire de la
  // préciser au moment de saisir le score.

  // Couleurs : copiées de Configuration (gris) ; saisie manuelle (jaune) ; calculées (gris).
  sheet.getRange(2, 1, nRows, 7).setBackground(COLOR_CALC);    // A..G : copiées de Configuration
  sheet.getRange(2, 8, nRows, 2).setBackground(COLOR_INPUT);   // H, I : scores
  sheet.getRange(2, 10, nRows, 6).setBackground(COLOR_INPUT);  // J..O : saisie manuelle (dont Pointage régl.)
  sheet.getRange(2, 16, nRows, 5).setBackground(COLOR_CALC);   // P..T : calculées
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
    'supplémentaires, indiquer la dernière manche jouée (ex. 7, 8...). À NE PAS confondre ' +
    'avec "Manches prévues" (col. M).');
  addColumnDoc('L — Retraits en fin',
    'Nombre de retraits (0, 1 ou 2) déjà comptés dans la dernière demi-manche AU MOMENT ' +
    'où la partie s\'est terminée — pertinent seulement si l\'équipe locale a gagné en ' +
    'frappant le point gagnant avant d\'avoir complété ses 3 retraits (victoire ' +
    '"walk-off"). Si la partie est allée jusqu\'au bout normalement (3 retraits) ou que ' +
    'la visiteuse a gagné, laisser à 0. C\'est cette valeur qui détermine la fraction de ' +
    'manche calculée plus bas (1 retrait = ⅓ de manche, 2 retraits = ⅔ de manche).');
  addColumnDoc('M — Manches prévues',
    'Nombre de manches RÉGLEMENTAIRES prévues pour CETTE partie (pré-rempli à 6 ; mettre 5 ' +
    'lors d\'une journée écourtée par la pluie, etc.). Sert à créditer le bon nombre de ' +
    'manches DÉFENSIVES au gagnant d\'un Mercy ou d\'un Forfait (la règle BQ donne au ' +
    'gagnant le crédit du nombre réglementaire de manches en défense). À NE PAS confondre ' +
    'avec "Manches complètes" (col. K = dernière manche réellement jouée).');
  addColumnDoc('N — Type de fin',
    '"Normal" = partie jouée selon les règles standards jusqu\'au bout. "Mercy" = partie ' +
    'arrêtée avant la fin en raison de la règle de l\'écart de points. "Forfait" = une ' +
    'équipe ne se présente pas ou abandonne ; indiquez simplement le gagnant en lui donnant ' +
    'un score plus élevé (les chiffres exacts importent peu). Le système attribue alors le ' +
    'pointage officiel (1 point par manche prévue, col. M, au gagnant ; 0 au perdant) et ' +
    'crédite le gagnant de "Manches prévues" manches DÉFENSIVES et 0 offensive, le perdant ' +
    'de 0 manche défensive et "Manches prévues" offensives (Art. 42.11). Voir aussi la ' +
    'section "Forfaits (Art. 42.11)" plus bas. ' +
    '"Supplémentaires" = partie prolongée au-delà des manches réglementaires (remplir alors ' +
    'la col. O ; voir la section "Manches supplémentaires et Note 4" plus bas).');
  addColumnDoc('O — Pointage régl. (suppl.)',
    'À remplir SEULEMENT si Type de fin = "Supplémentaires". C\'est le pointage (forcément ' +
    'ÉGAL) atteint à la fin des manches RÉGLEMENTAIRES, avant les supplémentaires. Une partie ' +
    'ne va en supplémentaire que si elle est nulle au terme du réglementaire (Art. 42.4), donc ' +
    'les deux équipes ont le même total réglementaire : un seul chiffre suffit. Ex. : finale ' +
    '7-5 allée en 7e, qui était 5-5 après la 6e → saisir 5. Sert à EXCLURE automatiquement les ' +
    'points des supplémentaires des ratios RD/RO (Note 4 — voir la section plus bas).');
  addColumnDoc('P — Gagnant (calculé)',
    'Nom de l\'équipe gagnante, déterminé automatiquement à partir des scores (colonnes ' +
    'H et I). Ne pas modifier à la main — recalculé par "Mettre à jour les classements".');
  addColumnDoc('Q / R — MO Éq.1 / MD Éq.1 (calculé)',
    'MO = Manches OFFENSIVES (à la batte) ; MD = Manches DÉFENSIVES (au champ) jouées par ' +
    'l\'Équipe 1, en fractions de tiers si la partie s\'est terminée par un walk-off ou ' +
    'un Mercy en milieu de manche. Voir la section "Fractions de manches" ci-dessous. ' +
    'Calculé automatiquement (manches RÉELLES de la partie, supplémentaires incluses).');
  addColumnDoc('S / T — MO Éq.2 / MD Éq.2 (calculé)',
    'Même chose que Q / R (MO/MD), mais pour l\'Équipe 2.');

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
    'Dès que l\'équipe LOCALE gagne, elle n\'a jamais complété sa dernière manche au bâton, ' +
    'et deux demi-manches restent incomplètes. Deux situations mènent à ce résultat : ' +
    '(1) walk-off — la locale marque le point gagnant dans le bas d\'une manche avant ses ' +
    '3 retraits ; (2) la locale menait déjà après le haut de la dernière manche et n\'est ' +
    'donc pas retournée frapper en bas. Dans les deux cas, la locale n\'a pas eu besoin de ' +
    'terminer sa manche à la batte, et la visiteuse n\'a pas joué cette demi-manche en ' +
    'défense. Tout le reste de la partie (le haut de cette même manche et toutes les ' +
    'manches précédentes) s\'est joué normalement et compte comme des manches pleines.');
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
  addText('AUTRE EXEMPLE (la locale menait, sans frapper en bas) :', true);
  addText(
    'Partie de 6 manches : la locale gagne 7-5 ; elle menait déjà après le haut de la 6e et ' +
    'n\'est pas retournée frapper en bas (colonne K = 6, colonne L = 0) :');
  addText(
    '• Équipe locale : 5 manches OFFENSIVES (n\'a pas frappé en bas de la 6e) — 6 manches ' +
    'DÉFENSIVES (a joué tout le haut de chaque manche, incluant la 6e).');
  addText(
    '• Équipe visiteuse : 6 manches OFFENSIVES — 5 manches DÉFENSIVES (n\'a pas joué le bas ' +
    'de la 6e en défense, puisque la locale n\'a pas frappé).');
  addBlank();
  addText(
    'AUCUNE fraction (ni demi-manche) n\'est utilisée UNIQUEMENT quand la VISITEUSE gagne ' +
    '(la locale a alors frappé jusqu\'au bout pour tenter de remonter) : les deux équipes ' +
    'ont joué le même nombre de manches pleines. Une victoire de la locale est, elle, ' +
    'TOUJOURS asymétrique (voir les deux exemples ci-dessus).');
  addText(
    'Cas particulier "Mercy" : si la règle de l\'écart de points est déclenchée alors ' +
    'que l\'équipe locale est à la batte (dans le bas d\'une manche), les mêmes fractions ' +
    's\'appliquent que pour une victoire walk-off. Dans tous les cas, l\'équipe GAGNANTE ' +
    'd\'un Mercy reçoit le crédit du nombre de manches DÉFENSIVES PRÉVUES (colonne M — ' +
    '6 en temps normal, 5 lors d\'une journée écourtée), comme si elle avait lancé toute ' +
    'la partie réglementaire.');

  addBlank();
  addTitle('FORFAITS (Art. 42.11)', COLOR_SECTION);
  addText(
    'Pour une partie gagnée par FORFAIT (une équipe ne se présente pas ou abandonne), ' +
    'choisissez "Forfait" dans la colonne "Type de fin" et indiquez le gagnant en lui ' +
    'donnant un score plus élevé que l\'adversaire (par ex. 1-0). Les chiffres exacts ' +
    'saisis n\'ont pas d\'importance : le système les remplace automatiquement par le ' +
    'pointage officiel.');
  addText(
    'Pointage officiel : 1 point par manche RÉGLEMENTAIRE PRÉVUE (colonne M) est accordé ' +
    'à l\'équipe GAGNANTE, et 0 au perdant. Une partie prévue de 6 manches donne donc 6-0 ' +
    '(5-0 lors d\'une journée écourtée à 5 manches). C\'est ce pointage qui apparaît dans ' +
    'les points pour/contre (PP/PC) des classements.');
  addText(
    'Manches créditées : l\'équipe GAGNANTE reçoit "Manches prévues" manches DÉFENSIVES et ' +
    '0 manche OFFENSIVE ; l\'équipe PERDANTE reçoit 0 manche défensive et "Manches prévues" ' +
    'manches offensives.');
  addText(
    'Note 5 (Art. 42.11) : les parties gagnées par forfait ne sont PAS comptabilisées dans ' +
    'les ratios RD/RO servant à départager les équipes aux fins du « Meilleur deuxième » ' +
    '(Étapes B et C) — le système les exclut alors automatiquement (points ET manches). ' +
    'Elles comptent toutefois normalement dans la fiche victoires-défaites et dans le ' +
    'classement de chaque pool (Étape A).');

  addBlank();
  addTitle('MANCHES SUPPLÉMENTAIRES ET NOTE 4 (Art. 42.11)', COLOR_SECTION);
  addText(
    'Quand une partie se prolonge en manches supplémentaires, la Note 4 du règlement exige ' +
    'que les points marqués/alloués DANS les manches supplémentaires soient EXCLUS des ratios ' +
    'de bris d\'égalité (RD et RO) : seuls les points des manches régulières comptent.');
  addText(
    'Le système applique cet ajustement AUTOMATIQUEMENT, à deux conditions : (1) indiquer ' +
    '"Supplémentaires" dans la colonne "Type de fin" ; (2) saisir le "Pointage régl. (suppl.)" ' +
    '(colonne O) — le pointage nul atteint à la fin des manches réglementaires. Une partie ne ' +
    'va en supplémentaire que si elle est nulle au terme du réglementaire (Art. 42.4), donc un ' +
    'seul chiffre suffit (les deux équipes ont le même total réglementaire).');
  addText(
    'IMPORTANT — l\'exclusion ne touche QUE le bris d\'égalité. Le TABLEAU DE POOL (à gauche) ' +
    'affiche des ratios RD/RO RÉELS : toutes les manches jouées sont comptées, supplémentaires ' +
    'incluses, donc RD = PC ÷ MD y tombe juste. La Note 4 (exclusion des supplémentaires) ne ' +
    's\'applique que dans le TABLEAU DE BRIS D\'ÉGALITÉ (à droite), là où une égalité doit ' +
    'réellement être départagée. Un ℹ sous le pool confirme que la Note 4 y a été appliquée ; ' +
    'si vous oubliez de saisir le Pointage régl. (col. O), un ⚠ apparaît à la place et ' +
    'l\'exclusion n\'est pas faite tant que la colonne reste vide.');
  addText(
    'À noter : dans le tableau de bris d\'égalité, les colonnes PP / PC / MO / MD sont en base ' +
    'RÉGULIÈRE (supplémentaires exclues) — ce sont les chiffres exacts qui produisent ses ratios ' +
    'RD/RO. Pour une partie allée en supplémentaire, ces valeurs (et les RD/RO) diffèrent donc de ' +
    'celles du tableau de pool, qui sont réelles : c\'est normal et voulu (Note 4).');
  addText(
    'Le gagnant, la fiche victoires-défaites et la fiche tête-à-tête (1er critère de bris ' +
    'd\'égalité) restent toujours basés sur le score FINAL : seul le ratio (2e et 3e critères) ' +
    'exclut les supplémentaires.');

  addBlank();
  addTitle('LIRE LES TABLEAUX DE BRIS D\'ÉGALITÉ (à droite des classements)', COLOR_SECTION);
  addText(
    'À DROITE de chaque classement (chaque pool, puis le classement des 1ers — Étape C — et ' +
    'le meilleur 2e — Étape B), un bloc "BRIS D\'ÉGALITÉ" montre COMMENT le système a ' +
    'départagé les équipes à fiche identique. Le classement de gauche ne montre qu\'un ordre ' +
    'final ; ce bloc rend le calcul de l\'Art. 42.11 transparent.');
  addText(
    'N\'y figurent que les GROUPES d\'équipes réellement à égalité (même fiche victoires-' +
    'défaites). Si aucun départage n\'était nécessaire, le bloc affiche simplement "Aucune ' +
    'égalité à départager — rangs établis par la fiche V-D". Les groupes distincts sont ' +
    'séparés par une ligne vide.');
  addText(
    'Chaque ligne montre l\'équipe, sa fiche V-D, ses points pour/contre (PP/PC), ses manches ' +
    'offensives/défensives (MO/MD), ses ratios RD et RO, et surtout le CRITÈRE DÉCISIF — le ' +
    'critère de l\'Art. 42.11 qui l\'a classée devant l\'équipe de la ligne suivante du même ' +
    'groupe. Ces PP/PC/MO/MD sont en base RÉGULIÈRE (supplémentaires exclues, Note 4), si bien ' +
    'que RD = PC ÷ MD tombe juste ici. Comme les équipes d\'un groupe ont la même fiche ' +
    '(Priorité 1), le départage vient de :', true);
  addText(
    '• "RD x < y" — départagées par le ratio DÉFENSIF (Priorité 2 ; le plus bas gagne).  ' +
    '• "RO x > y" — ratios défensifs égaux, départagées par le ratio OFFENSIF (Priorité 3 ; ' +
    'le plus haut gagne).  • "⚠ Manuel (P4)" — ratios défensif ET offensif identiques : le ' +
    'système ne peut pas trancher (Priorité 4 = "manches avec l\'avance au pointage", qui ' +
    'exige la feuille de pointage). Calculez l\'ordre à la main sur les feuilles de pointage ' +
    'papier, puis saisissez-le dans la colonne "Forcer rang" du tableau de bris d\'égalité : ' +
    'le critère passe alors à "🔒 Forcé (P4)".');
  addText(
    'Les RD/RO de ce bloc sont calculés exactement comme par le moteur de bris : sur la ' +
    'portée TÊTE-À-TÊTE entre équipes à égalité pour un pool (Étape A), et sur TOUTES les ' +
    'parties de pool de chaque équipe pour les Étapes B/C. Ils sont en base RÉGULIÈRE : si une ' +
    'partie de la portée est allée en supplémentaires, un ℹ rappelle que les manches ' +
    'supplémentaires sont exclues (Note 4 — voir la section ci-dessus). C\'est pourquoi, en ' +
    'présence d\'une partie supplémentaire, les RD/RO (et les PP/PC/MO/MD) de ce bloc peuvent ' +
    'différer de ceux du tableau de pool à gauche, qui montre les valeurs RÉELLES.');
  addText(
    'Exemple de lecture : trois équipes 2-1 dans un pool. Le bloc montre la 1re sans critère ' +
    '(meneuse), la 2e avec "RD 0.333 < 0.500" (elle a un meilleur ratio défensif), la 3e avec ' +
    '"RO 1.083 > 0.900" (ratios défensifs égaux à la 2e, départagée au ratio offensif). On ' +
    'voit ainsi d\'un coup d\'œil quel critère a fait la différence à chaque rang.');

  addBlank();
  addTitle('QUAND LES POSITIONS DE DEMI-FINALE (1-2-3-4) S\'AFFICHENT', COLOR_SECTION);
  addText(
    'Les positions de demi-finale — 1re, 2e, 3e place (les 1ers de pool, Étape C) et le ' +
    'meilleur 2e (position 4, Étape B) — forment un classement ENTRE les pools. Il ne devient ' +
    'fiable qu\'une fois TOUTES les parties de pool de la classe jouées : avant, le meilleur 2e ' +
    'et l\'ordre des 1ers changent à chaque score entré.');
  addText(
    'Pour éviter d\'afficher un classement provisoire trompeur (surtout sur la page publique ' +
    'partagée sur Facebook), ces éléments RESTENT MASQUÉS tant que la classe n\'est pas ' +
    'terminée : la colonne « Avancement » des pools reste vide, et les sections « CLASSEMENT ' +
    'DES 1ers » (Étape C), « MEILLEUR 2e » (Étape B) et le récapitulatif « DEMI-FINALES » sont ' +
    'remplacés par un bandeau d\'attente indiquant la progression (parties jouées / total). ' +
    'Le classement INTERNE de chaque pool (rang, 1er/2e, V-D, RD), lui, s\'affiche normalement ' +
    'en direct.');
  addText(
    'Chaque classe est indépendante : la classe A révèle ses positions dès que ses 3 pools sont ' +
    'terminés, même si la classe B n\'a pas fini. Tout apparaît automatiquement au dernier score ' +
    'entré (si la mise à jour auto est activée) ; sinon lancer « Mettre à jour les classements ».');

  addBlank();
  addTitle('FORCER LE 2e D\'UN POOL (Note 5 — forfaits)', COLOR_SECTION);
  addText(
    'Cas rare réservé à l\'admin. La Note 5 de l\'Art. 42.11 exclut les parties gagnées par ' +
    'FORFAIT aux fins du « Meilleur 2e ». Si un forfait a faussé la 2e place d\'un pool, l\'admin ' +
    'peut désigner manuellement quelle équipe représente ce pool au « SECTION 5 — MEILLEUR 2e ' +
    '(Position 4) — Étape B ».');
  addText(
    'Comment : dans la section du pool (classement de gauche), COCHER la case de la colonne ' +
    '« Forcer 2e » (dernière colonne du tableau) à côté de l\'équipe choisie ; DÉCOCHER pour ' +
    'revenir au 2e automatique. Ce forçage a PRÉSÉANCE sur le 2e calculé automatiquement. Le ' +
    'classement se recalcule TOUT SEUL dès qu\'on coche/décoche (si la mise à jour auto est ' +
    'activée) ; sinon lancer « Mettre à jour les classements ». Un bandeau ℹ confirme le 2e ' +
    'forcé sous la section du pool.');
  addText(
    'Mettre « 2 » sur l\'équipe déjà 1re du pool est IGNORÉ (elle est déjà qualifiée comme 1re via ' +
    'l\'Étape C ; elle ne peut pas être aussi le meilleur 2e) : un avertissement ⚠ s\'affiche et le ' +
    '2e automatique est conservé. Le forçage ne change QUE le candidat au meilleur 2e — le ' +
    'classement du pool (rangs 1-2-3-4) et tous les ratios restent inchangés.');

  addBlank();
  addTitle('MISE À JOUR AUTOMATIQUE DES CLASSEMENTS', COLOR_SECTION);
  addText(
    'Les classements se mettent à jour TOUT SEULS, en direct et pour tout le monde, dès ' +
    'qu\'une partie est entrée au complet. Plusieurs personnes peuvent saisir des scores ' +
    'en même temps : chaque mise à jour est visible immédiatement par ceux qui regardent ' +
    'les onglets "Classements A / B". Aucune manipulation requise.');
  addText(
    'La mise à jour ne se déclenche que lorsque TOUTES les colonnes de saisie d\'une ' +
    'partie sont remplies (les 2 scores, Équipe Locale, Manches complètes, Retraits en fin, ' +
    'Type de fin) — sauf pour un "Forfait", où seuls les 2 scores et le Type suffisent, et ' +
    'pour une partie "Supplémentaires", où le "Pointage régl." (col. O) est aussi requis. ' +
    'Tant que la ligne n\'est pas complète, le classement ne bouge pas : c\'est voulu, pour ' +
    'qu\'il n\'affiche jamais un résultat à moitié entré.');
  addText(
    'IMPORTANT — pensez à remplir les Manches complètes (colonne K) même pour une partie ' +
    'normale : une journée de pluie peut écourter les parties (ex. 5 manches au lieu de ' +
    '6), et cette valeur change les ratios du classement.');
  addText(
    'Cas particulier : si vous EFFACEZ plusieurs cellules d\'un coup (correction en lot), ' +
    'ou en cas de doute, cliquez sur "🏆 Tournoi Baseball > Mettre à jour les classements" ' +
    'pour forcer un recalcul complet et propre.', true);

  addBlank();
  addTitle('📒 GRAND LIVRE DES MATCHS (AUDIT)', COLOR_SECTION);
  addText(
    'La feuille "Grand livre" liste chaque partie deux fois — une ligne par équipe — comme ' +
    'un compte de banque : les colonnes "Somme PP/PC/MO/MD" (EN GRAS) cumulent ' +
    'progressivement d\'un match à l\'autre pour une même équipe, alors que RD et RO sont ' +
    'recalculés à chaque ligne (jamais cumulés). Elle se reconstruit automatiquement en ' +
    'même temps que les classements (menu ou mise à jour auto). Les premières colonnes ' +
    '(Cl, Pl, #, J, H, T, Eq, Adv, Res, Sco) sont abrégées pour économiser l\'espace ; ' +
    'survolez l\'en-tête pour voir le nom complet.');
  addText(
    'Repères visuels : la colonne "Équipe" est verte et "Adversaire" rouge quand l\'équipe ' +
    'de la ligne gagne (l\'inverse si elle perd) ; "Loc/Vis" est bleu pâle (locale) ou ' +
    'orange pâle (visiteuse) ; RD et RO ont un dégradé de couleur (vert = bon, rouge = ' +
    'moins bon) ; des traits séparent les blocs (moyen = fin d\'équipe, épais = fin de ' +
    'pool, double = fin de classe).');
  addText(
    'Utile en cas de doute sur un classement : le "solde final" de chaque équipe (dernière ' +
    'ligne de son bloc) doit correspondre exactement à ses colonnes PP/PC/MO/MD dans le ' +
    'tableau de pool des Classements. Si ce n\'est pas le cas, ou si un écart apparaît en ' +
    'cours de tournoi, on peut retracer, ligne par ligne, à quel match précis le cumul ' +
    'dérape.');

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
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Générer les matchs',
    'Régénérer les matchs dans Résultats A/B à partir de la Configuration ? ' +
    'Les scores déjà saisis sont normalement conservés (appariés par pool + # de ' +
    'partie), mais les lignes seront réécrites. Continuer ?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) { return; }

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
      if (maxRowsEmpty > 1) { sheet.getRange(2, 1, maxRowsEmpty - 1, 20).clearContent(); }
      return;
    }

    // Conserve les scores / saisies manuelles déjà présents, appariés par clé
    // Pool + Partie # (robuste même si l'horaire est recollé dans un ordre différent).
    var existingByKey = {};
    var prevLast = sheet.getLastRow();
    if (prevLast >= 2) {
      var prevData = sheet.getRange(2, 1, prevLast - 1, 20).getValues();
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
        prev ? prev[12] : 6,  // M : Manches prévues (défaut 6)
        prev ? prev[13] : 'Normal', // N : Type de fin
        prev ? prev[14] : '', // O : Pointage régl. (suppl.)
        '',                // P : Gagnant
        '', '', '', ''     // Q..T : manches calculées
      ];
    });

    var maxRows = sheet.getMaxRows();
    if (maxRows > 1) { sheet.getRange(2, 1, maxRows - 1, 20).clearContent(); }
    sheet.getRange(2, 1, nRows, 20).setValues(rows);
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

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Matchs générés à partir de la Configuration (' + schedule.length + ' parties). ' +
    'Indiquez l\'Équipe Locale et les scores au moment de saisir chaque résultat, ' +
    'puis "Mettre à jour les classements".',
    'Tournoi Baseball', 6);
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

  var data = sheet.getRange(2, 1, last - 1, 20).getValues();
  data.forEach(function (row, idx) {
    var pool        = parseInt(row[0], 10);
    var partie      = row[1];
    var jour        = row[2];
    var heure       = row[3];
    var terrain     = String(row[4]).trim();
    var teamA       = String(row[5]).trim();
    var teamB       = String(row[6]).trim();
    var scoreA      = row[7];
    var scoreB      = row[8];
    var localSel    = String(row[9]).trim();
    var manches     = parseInt(row[10], 10);
    var retraits    = row[11] === '' ? 0 : parseInt(row[11], 10);
    var manchesPrevues = parseInt(row[12], 10);
    var type        = String(row[13]).trim() || 'Normal';
    var suppTieRaw  = row[14];   // O : Pointage régl. (suppl.) — seulement si Supplémentaires

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
    // Manches réglementaires prévues : défaut 6 si non saisi.
    if (isNaN(manchesPrevues) || manchesPrevues < 1) { manchesPrevues = TOTAL_INNINGS; }

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

    // Forfait : le score officiel est de 1 point par manche réglementaire prévue
    // (col. M) pour le gagnant et 0 pour le perdant (Art. 42.11). Le registraire
    // n'a qu'à désigner le gagnant en lui donnant un score plus élevé ; le système
    // normalise ensuite le pointage.
    if (type === 'Forfait') {
      var localeGagneForfait = (scoreLocal >= scoreVisiteur);
      scoreLocal    = localeGagneForfait ? manchesPrevues : 0;
      scoreVisiteur = localeGagneForfait ? 0 : manchesPrevues;
    }

    var inn = calculateInnings(scoreLocal, scoreVisiteur, manches, retraits, type, homeKnown, manchesPrevues);
    var winner = (type === 'Forfait')
      ? (scoreLocal >= scoreVisiteur ? local : visiteur)   // forfait : le score indique le gagnant
      : (scoreLocal > scoreVisiteur ? local : (scoreVisiteur > scoreLocal ? visiteur : ''));

    // ---- Note 4 (Art. 42.11) : manches supplémentaires ----
    // Pour le RATIO de bris d'égalité seulement, on exclut les manches
    // supplémentaires. Une partie ne va en supplémentaire que si elle est NULLE
    // au terme des manches réglementaires (Art. 42.4) : les deux équipes ont alors
    // marqué le même total réglementaire X = "Pointage régl." (col O), sur
    // manchesPrevues manches chacune. La fiche V-D, le gagnant et les PP/PC
    // affichés, eux, restent basés sur le score FINAL (suppl. incluses).
    var isSupp = (type === 'Supplémentaires');
    var suppTie = parseInt(suppTieRaw, 10);
    var suppNeedsTie = isSupp && (suppTieRaw === '' || isNaN(suppTie) || suppTie < 0);
    var regRsLocal, regRsVisiteur, regOffLocal, regDefLocal, regOffVisiteur, regDefVisiteur;
    if (isSupp && !suppNeedsTie) {
      regRsLocal = regRsVisiteur = suppTie;            // pointage réglementaire (nul) X
      regOffLocal = regDefLocal = manchesPrevues;      // manches régulières seulement
      regOffVisiteur = regDefVisiteur = manchesPrevues;
    } else {
      // Partie ordinaire (ou suppl. sans pointage régl. saisi → repli sans exclusion) :
      // base régulière = valeurs réelles de la partie.
      regRsLocal = scoreLocal; regRsVisiteur = scoreVisiteur;
      regOffLocal = inn.offLocal; regDefLocal = inn.defLocal;
      regOffVisiteur = inn.offVisiteur; regDefVisiteur = inn.defVisiteur;
    }

    games.push({
      pool: pool,
      partie: partie,
      rowIndex: idx + 2,
      jour: jour,
      heure: heure,
      terrain: terrain,
      local: local,
      visiteur: visiteur,
      localIsEq1: (local === teamA),   // la locale est-elle l'Équipe 1 (col F) ? — pour l'affichage Q..T
      homeKnown: homeKnown,
      scoreLocal: scoreLocal,
      scoreVisiteur: scoreVisiteur,
      manches: manches,
      retraits: retraits,
      manchesPrevues: manchesPrevues,
      type: type,
      winner: winner,
      offLocal: inn.offLocal,
      defLocal: inn.defLocal,
      offVisiteur: inn.offVisiteur,
      defVisiteur: inn.defVisiteur,
      // Champs « manches régulières seulement » pour les ratios (Note 4).
      regRsLocal: regRsLocal,
      regRsVisiteur: regRsVisiteur,
      regOffLocal: regOffLocal,
      regDefLocal: regDefLocal,
      regOffVisiteur: regOffVisiteur,
      regDefVisiteur: regDefVisiteur,
      suppNeedsTie: suppNeedsTie
    });
  });
  return games;
}

/**
 * Lit les parties d'une classe pour la VUE « Résultats » de la page publique —
 * affichage seulement, découplé du moteur de classement. Lit directement la feuille
 * Résultats (le Terrain, p. ex., n'est pas conservé par getGameResults). Retourne
 * toutes les parties générées (Éq.1 et Éq.2 présentes), jouées ou non.
 *
 * @param {string} classe  'A' ou 'B'
 * @return {Array} lignes { partie, classe, pool, jour, heure, terrain, eq1, eq2,
 *                          scoreA, scoreB, lastInn, type, played }
 */
function getMatchRows(classe) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_RESULTS[classe]);
  var rows = [];
  if (!sheet) { return rows; }
  var last = sheet.getLastRow();
  if (last < 2) { return rows; }

  var data = sheet.getRange(2, 1, last - 1, 15).getValues();          // colonnes A..O
  var disp = sheet.getRange(2, 1, last - 1, 15).getDisplayValues();   // Jour/Heure : texte affiché
  var tz   = ss.getSpreadsheetTimeZone();
  var isDate = function (v) { return Object.prototype.toString.call(v) === '[object Date]'; };
  data.forEach(function (r, i) {
    var eq1 = String(r[5]).trim();    // F : Équipe 1
    var eq2 = String(r[6]).trim();    // G : Équipe 2
    if (eq1 === '' || eq2 === '') { return; }   // ligne non générée

    var pool    = parseInt(r[0], 10); // A
    var partie  = r[1];               // B : Partie #
    // Jour/Heure : version COURTE, sans année (évite un retour à la ligne sur téléphone).
    // Si la cellule est une vraie date/heure → formatage explicite ; sinon on retire
    // simplement l'année (« 2026 ») du texte affiché.
    var jour  = isDate(r[2]) ? Utilities.formatDate(r[2], tz, 'd MMM')
                             : String(disp[i][2]).replace(/\b\d{4}\b/g, '').replace(/[,\-\/]\s*$/, '').replace(/^\s*[,\-\/]/, '').trim();
    var heure = isDate(r[3]) ? Utilities.formatDate(r[3], tz, 'HH:mm')
                             : String(disp[i][3]).trim();
    var terrain = String(r[4]).trim();// E
    var sA = r[7], sB = r[8];         // H, I
    var lastInn = r[10];              // K : Manches complètes (dernière manche jouée)
    var prevuesRaw = parseInt(r[12], 10); // M : Manches prévues (longueur réglementaire)
    var manchesPrevues = (isNaN(prevuesRaw) || prevuesRaw < 1) ? TOTAL_INNINGS : prevuesRaw;
    var type    = String(r[13]).trim(); // N : Type de fin

    var hasScore = (sA !== '' && sB !== '' && !isNaN(Number(sA)) && !isNaN(Number(sB)));
    var scoreA = hasScore ? Number(sA) : null;
    var scoreB = hasScore ? Number(sB) : null;

    // Forfait : pointage officiel = Manches prévues (col. M) au gagnant, 0 au perdant
    // (Art. 42.11) ; le registraire ne fait que désigner le gagnant par un score plus haut.
    if (hasScore && type === 'Forfait') {
      var prevues = parseInt(r[12], 10);
      if (isNaN(prevues) || prevues < 1) { prevues = TOTAL_INNINGS; }
      if (scoreA >= scoreB) { scoreA = prevues; scoreB = 0; }
      else { scoreA = 0; scoreB = prevues; }
    }

    rows.push({
      partie:  partie,
      classe:  classe,
      pool:    isNaN(pool) ? '' : pool,
      jour:    jour,
      heure:   heure,
      terrain: terrain,
      eq1: eq1, eq2: eq2,
      scoreA: scoreA, scoreB: scoreB,
      lastInn: (lastInn === '' || lastInn === null) ? null : lastInn,
      manchesPrevues: manchesPrevues,
      type: hasScore ? (type || 'Normal') : '',
      played: hasScore
    });
  });
  return rows;
}

/**
 * État d'avancement des parties de POOL d'une classe. Sert à ne révéler les
 * positions de demi-finale (1-2-3-4, un classement INTER-pools) qu'une fois toutes
 * les parties de pool de la classe jouées — avant cela, le meilleur 2e et l'ordre
 * des 1ers sont provisoires et changent à chaque score entré.
 *
 * Réutilise getMatchRows : sa marque `played` (les deux scores saisis) est
 * exactement le critère qui fait qu'une partie est comptée par le moteur
 * (getGameResults) — donc « complet » ici = « toutes les parties comptent ».
 *
 * @param {string} classe  'A' ou 'B'
 * @return {Object} { total, played, complete } — complete = total>0 && played===total
 */
function poolPlayCompletion(classe) {
  var rows   = getMatchRows(classe);
  var total  = rows.length;
  var played = rows.filter(function (r) { return r.played; }).length;
  return { total: total, played: played, complete: total > 0 && played === total };
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
 * @param {Object} v  valeurs des colonnes H..O : {scoreA, scoreB, local, manches,
 *                    retraits, type, suppTie}
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
  if (!filled(v.local) || !filled(v.manches) || !filled(v.retraits)) { return false; }

  // Supplémentaires : le Pointage régl. (col O) est aussi requis pour appliquer
  // la Note 4 (sinon le ratio inclurait les points des manches supplémentaires).
  if (type === 'Supplémentaires') { return filled(v.suppTie); }

  return true;
}

// ============================================================================
//  CALCUL DES MANCHES (fractions)
// ============================================================================

/**
 * Calcule les manches offensives/défensives pour la locale et la visiteuse.
 *
 * Règles :
 *  - Forfait : gagnant 0 off / regulation déf ; perdant regulation off / 0 déf (Art. 42.11).
 *  - Mercy : gagnant reçoit 6 manches défensives ; perdant reçoit les manches
 *            réellement jouées (avec fractions).
 *  - Victoire locale en bas (walk-off) : détectée si Score Local > Score Visiteur
 *    ET (manches < 6 OU retraits > 0). La locale n'a pas frappé en bas de la
 *    dernière manche complète -> fractions.
 *  - Sinon (partie normale complète, ou visiteur gagne) : symétrique = manches/manches.
 *
 * @param {boolean} [homeKnown=true]  si faux (équipe locale non indiquée), on force
 *                  un résultat symétrique : impossible de savoir qui a joué en bas.
 * @param {number}  [regulation=TOTAL_INNINGS]  nombre de manches réglementaires PRÉVUES
 *                  pour cette partie (6 en 13U normal, 5 lors d'une journée écourtée par la
 *                  pluie, etc.). Sert au crédit défensif du gagnant d'un Mercy / Forfait.
 * @return {{offLocal:number, defLocal:number, offVisiteur:number, defVisiteur:number}}
 */
function calculateInnings(scoreLocal, scoreVisiteur, manchesCompletes, retraitsEnFin, typeFin, homeKnown, regulation) {
  if (homeKnown === undefined) { homeKnown = true; }
  if (regulation === undefined || !regulation) { regulation = TOTAL_INNINGS; }
  var N = manchesCompletes;
  var H = retraitsEnFin;
  var frac = H / 3;                 // fraction de manche (0, 1/3, 2/3)
  var partial = (N - 1) + frac;     // ex: N=5, H=1 -> 4.333

  // -------- FORFAIT --------
  // Art. 42.11 : le gagnant reçoit le crédit de `regulation` manches DÉFENSIVES et
  // 0 manche OFFENSIVE ; le perdant 0 manche défensive et `regulation` manches
  // offensives. (Le pointage 1 pt/manche au gagnant est appliqué en amont, dans
  // getGameResults.)
  if (typeFin === 'Forfait') {
    if (scoreLocal >= scoreVisiteur) {
      // Locale gagne par forfait.
      return { offLocal: 0, defLocal: regulation, offVisiteur: regulation, defVisiteur: 0 };
    } else {
      // Visiteur gagne par forfait.
      return { offLocal: regulation, defLocal: 0, offVisiteur: 0, defVisiteur: regulation };
    }
  }

  // Détermine si la locale a gagné en n'ayant PAS complété sa dernière manche au bâton.
  // Dès que la locale gagne, c'est forcément le cas : soit elle menait déjà et n'est pas
  // retournée frapper en bas de la dernière manche, soit elle a marqué le point gagnant
  // avant ses 3 retraits (walk-off). Dans les deux cas → fractions (partial). On ne peut
  // le déterminer que si l'équipe locale a été indiquée (homeKnown = true).
  var localeWinsBottom = homeKnown && (scoreLocal > scoreVisiteur);

  // -------- MERCY --------
  if (typeFin === 'Mercy') {
    if (scoreLocal > scoreVisiteur) {
      // Locale gagne par mercy : elle n'a pas complété sa dernière manche au bâton
      // (localeWinsBottom toujours vrai ici quand homeKnown). Locale (gagnante) :
      // offensive partielle, défensive = manches réglementaires prévues. Visiteur
      // (perdant) : offensive = N, défensive = partielle.
      if (localeWinsBottom) {
        return {
          offLocal: partial, defLocal: regulation,
          offVisiteur: N,    defVisiteur: partial
        };
      } else {
        // Repli (équipe locale non indiquée) : pas de fraction possible.
        return {
          offLocal: N, defLocal: regulation,
          offVisiteur: N, defVisiteur: N
        };
      }
    } else {
      // Visiteur gagne par mercy : visiteur gagnant (def = réglementaire), locale perdante.
      return {
        offLocal: N, defLocal: N,
        offVisiteur: N, defVisiteur: regulation
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

    // 1) Écrit les colonnes calculées (P..T) dans la feuille Résultats.
    writeCalculatedResults(ss, classe, games);

    // 2) Construit le classement.
    buildStandingsSheet(ss, classe, games);
  });

  buildLedgerSheet(ss);

  SpreadsheetApp.getActiveSpreadsheet().toast('Classements recalculés.', 'Tournoi Baseball', 5);
}

/**
 * Écrit le gagnant et les manches calculées (P..T) dans la feuille Résultats.
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
 * Écrit (ou efface) les colonnes calculées P..T (16..20) d'UNE ligne de Résultats.
 * Source unique du mapping P..T, partagée par writeCalculatedResults (refresh
 * complet via le menu) et handleResultEdit (écriture ciblée de la ligne saisie).
 * Les MO/MD écrits ici sont les manches RÉELLES de la partie (suppl. incluses) —
 * c'est le journal de la partie telle que jouée ; l'exclusion Note 4 ne touche que
 * les ratios des Classements, pas ce bloc.
 *
 * IMPORTANT : les colonnes Q..T sont indexées par ÉQUIPE 1 / ÉQUIPE 2 (position
 * dans l'horaire, col F/G), PAS par locale/visiteuse. L'objet `game` raisonne en
 * locale/visiteuse ; on remappe donc via `localIsEq1` avant d'écrire, sinon Q/R et
 * S/T seraient intervertis dès que l'équipe locale est l'Équipe 2.
 *
 * @param {Sheet}  sheet     feuille Résultats A/B
 * @param {number} rowIndex  numéro de ligne (>= 2)
 * @param {Object} game      partie calculée, ou null/undefined pour effacer
 */
function writeRowCalc(sheet, rowIndex, game) {
  if (game) {
    var eq1Off = game.localIsEq1 ? game.offLocal    : game.offVisiteur;
    var eq1Def = game.localIsEq1 ? game.defLocal    : game.defVisiteur;
    var eq2Off = game.localIsEq1 ? game.offVisiteur : game.offLocal;
    var eq2Def = game.localIsEq1 ? game.defVisiteur : game.defLocal;
    sheet.getRange(rowIndex, 16).setValue(game.winner);            // P : Gagnant
    sheet.getRange(rowIndex, 17).setValue(formatFraction(eq1Off)); // Q : MO Éq.1
    sheet.getRange(rowIndex, 18).setValue(formatFraction(eq1Def)); // R : MD Éq.1
    sheet.getRange(rowIndex, 19).setValue(formatFraction(eq2Off)); // S : MO Éq.2
    sheet.getRange(rowIndex, 20).setValue(formatFraction(eq2Def)); // T : MD Éq.2
  } else {
    // Partie sans résultat : efface les calculs.
    sheet.getRange(rowIndex, 16, 1, 5).clearContent();
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
    rs: 0, ra: 0,            // points marqués / alloués AFFICHÉS (PP/PC, partie complète)
    rsNum: 0, raNum: 0,      // numérateurs des ratios de BRIS (manches régulières — Note 4)
    offInn: 0, defInn: 0,    // manches des ratios de BRIS (régulières ; exclusion forfait si demandé)
    rsRatio: 0, raRatio: 0,  // ratios de BRIS d'égalité (base régulière, Note 4)
    offInnFull: 0, defInnFull: 0,   // manches RÉELLES jouées (suppl. incluses) — tableau de pool
    rsRatioFull: 0, raRatioFull: 0  // ratios RÉELS (toutes manches jouées) — tableau de pool
  };

  // Replis défensifs : pour une partie ordinaire (et pour les objets construits à
  // la main par les tests), les champs « régulière » valent les valeurs réelles.
  function val(x, fallback) { return (x === undefined || x === null) ? fallback : x; }

  games.forEach(function (g) {
    var isLocal = (g.local === team);
    var isVis   = (g.visiteur === team);
    if (!isLocal && !isVis) { return; }

    s.pj++;
    if (g.winner === team) { s.v++; }
    else if (g.winner !== '' ) { s.d++; }

    var teamScore = isLocal ? g.scoreLocal : g.scoreVisiteur;
    var oppScore  = isLocal ? g.scoreVisiteur : g.scoreLocal;
    s.rs += teamScore;       // PP : points RÉELS (manches suppl. incluses)
    s.ra += oppScore;        // PC

    // Numérateurs des ratios : points des manches RÉGULIÈRES seulement (Note 4).
    // Pour toute partie non-supplémentaire, reg* == score réel → ratio inchangé.
    var regTeamRuns = isLocal ? val(g.regRsLocal, g.scoreLocal) : val(g.regRsVisiteur, g.scoreVisiteur);
    var regOppRuns  = isLocal ? val(g.regRsVisiteur, g.scoreVisiteur) : val(g.regRsLocal, g.scoreLocal);

    // Manches des ratios — base régulière (Note 4) ; peut exclure les forfaits.
    // Note 5 (Art. 42.11) : aux fins du « Meilleur deuxième » (Étapes B/C), une
    // partie gagnée par forfait n'est PAS comptabilisée — on exclut alors la partie
    // ENTIÈREMENT du ratio, autant les points (rsNum/raNum) que les manches
    // (offInn/defInn), sinon le numérateur garderait des points sans dénominateur.
    var skipRatio = (excludeForfaitRatios && g.type === 'Forfait');
    if (!skipRatio) {
      s.rsNum += regTeamRuns;
      s.raNum += regOppRuns;
      if (isLocal) {
        s.offInn += val(g.regOffLocal, g.offLocal);
        s.defInn += val(g.regDefLocal, g.defLocal);
        s.offInnFull += g.offLocal;       // manches réelles (suppl. incluses)
        s.defInnFull += g.defLocal;
      } else {
        s.offInn += val(g.regOffVisiteur, g.offVisiteur);
        s.defInn += val(g.regDefVisiteur, g.defVisiteur);
        s.offInnFull += g.offVisiteur;
        s.defInnFull += g.defVisiteur;
      }
    }
  });

  // Ratios : RA/DefInn (le plus bas gagne), RS/OffInn (le plus haut gagne).
  // Numérateurs en base régulière (Note 4), dénominateurs idem.
  s.raRatio = s.defInn > 0 ? (s.raNum / s.defInn) : Number.POSITIVE_INFINITY;
  s.rsRatio = s.offInn > 0 ? (s.rsNum / s.offInn) : 0;

  // Ratios RÉELS (toutes manches jouées, suppl. incluses) — affichés dans le
  // tableau de pool. Numérateurs = points réels (rs/ra), dénominateurs = manches
  // réelles. Diffèrent des ratios de bris seulement s'il y a des supplémentaires.
  s.raRatioFull = s.defInnFull > 0 ? (s.ra / s.defInnFull) : Number.POSITIVE_INFINITY;
  s.rsRatioFull = s.offInnFull > 0 ? (s.rs / s.offInnFull) : 0;
  return s;
}

/**
 * Classement d'un pool (Étape A). Trie les 4 équipes avec bris d'égalité.
 * @return {Array} stats triées (1er en tête) avec champ .rank
 */
function calculatePoolStandings(games, teams, forced) {
  // Stats globales du pool (pour affichage).
  var stats = teams.map(function (t) {
    return computeTeamStats(t, games, false);
  });

  // Tri avec bris d'égalité — Étape A utilise seulement les parties directes
  // entre équipes à égalité (géré dans tiebreaker via useAllGames=false).
  // `forced` = override « Forcer rang » de ce pool (Priorité 4, voir applyPriorities).
  var ordered = orderTeams(teams, games, false, forced || {});

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
function calculateStep(teams, games, useAllGames, forced) {
  return orderTeams(teams, games, useAllGames, forced || {});
}

/**
 * Ordonne une liste d'équipes en appliquant le bris d'égalité récursif.
 * Délègue à tiebreaker(), qui démarre une passe à la Priorité 1 et applique
 * les priorités de l'Art. 42.11 (en continuant après une séparation partielle,
 * Note 2).
 *
 * @param {Array}  teams
 * @param {Array}  games        parties disponibles
 * @param {boolean} useAllGames true = B/C (toutes parties de pool), false = A (direct)
 * @param {Object} [forced]     rangs forcés « Forcer rang » (Priorité 4), { team: number }
 * @return {Array} noms ordonnés
 */
function orderTeams(teams, games, useAllGames, forced) {
  if (teams.length <= 1) { return teams.slice(); }
  forced = forced || {};

  // Étapes B/C : un SEUL groupe, portée = toutes les parties impliquant une des
  // équipes (elles ne se sont jamais affrontées, cf. Note « étapes B et C » de
  // l'Art. 42.11). Comportement inchangé.
  if (useAllGames) {
    return tiebreaker(teams.slice(), games, true, forced);
  }

  // Étape A : structure à DEUX niveaux (Art. 42.11).
  //  1) Regrouper par fiche GLOBALE (V-D sur toutes les parties du pool) — c'est
  //     ce qui détermine QUI est à égalité, et l'ordre ENTRE les groupes.
  //  2) Pour CHAQUE groupe à égalité, départager sur la portée RESTREINTE aux
  //     parties jouées ENTRE ces équipes (tête-à-tête) : tiebreaker(group, …, false)
  //     pose relevantGames = headToHeadGames(group). La restriction de portée n'a
  //     lieu QU'ICI (premier découpage) ; la récursion Note 2 interne d'applyPriorities
  //     garde ensuite la portée du groupe (ce que l'exemple QC/RS/CN exige).
  var groups = groupByMetric(teams.slice(), function (t) {
    var st = computeTeamStats(t, games, false);
    return st.v - st.d;
  }, true);

  var ordered = [];
  var manual = false;
  groups.forEach(function (grp) {
    var sub = tiebreaker(grp, games, false, forced);
    if (sub.__needsManualCheck) { manual = true; }
    sub.forEach(function (t) { ordered.push(t); });
  });
  if (manual) { ordered.__needsManualCheck = true; }
  return ordered;
}

/**
 * Démarre une PASSE de bris d'égalité à la Priorité 1, Art. 42.11.
 *
 * Priorités :
 *   1. Fiche tête-à-tête (V-D dans les parties entre équipes à égalité)
 *   2. Ratio RA/DefInn (le plus bas gagne)
 *   3. Ratio RS/OffInn (le plus haut gagne)
 *   4. Manches en avance (avertissement : vérification manuelle requise)
 *
 * Fixe la PORTÉE des parties pour toute la passe, puis délègue à
 * applyPriorities() qui applique les priorités sur cette même portée.
 * - Étape A (useAllGames=false) : portée = parties jouées strictement entre
 *   les équipes à égalité (tête-à-tête).
 * - Étapes B/C (useAllGames=true) : portée = toutes les parties de pool
 *   impliquant l'une des équipes à égalité.
 *
 * @param {Array}  tiedTeams    équipes à départager
 * @param {Array}  games        parties disponibles
 * @param {boolean} useAllGames true = B/C, false = A (tête-à-tête direct)
 * @param {Object} [forced]     rangs forcés « Forcer rang » (Priorité 4), { team: number }
 * @return {Array} équipes ordonnées (meilleure en tête)
 */
function tiebreaker(tiedTeams, games, useAllGames, forced) {
  if (tiedTeams.length <= 1) { return tiedTeams.slice(); }

  // Portée fixée pour CETTE passe (Priorités 1 à 3 calculées dessus).
  var relevantGames = useAllGames
    ? games.filter(function (g) {
        return tiedTeams.indexOf(g.local) !== -1 || tiedTeams.indexOf(g.visiteur) !== -1;
      })
    : headToHeadGames(tiedTeams, games);

  return applyPriorities(tiedTeams, relevantGames, useAllGames, 1, forced || {});
}

/**
 * Applique les priorités à partir de `startP` sur une portée de parties FIXÉE
 * (`relevantGames`) — cœur de la Note 2 de l'Art. 42.11.
 *
 * Note 2 : « Lors d'une égalité multiple résolue en partie : on continue avec les
 * priorités restantes pour les équipes encore à égalité, avant de recommencer à
 * la priorité 1. »
 *
 * Concrètement :
 *  - Quand une priorité sépare le groupe partiellement, chaque sous-groupe
 *    encore à égalité CONTINUE à la priorité SUIVANTE (startP+1) sur la MÊME
 *    portée — on ne recommence PAS à la Priorité 1, et on ne re-restreint PAS
 *    la portée au sous-groupe.
 *  - Épuisement des priorités automatisables : si un sous-groupe (ou le groupe
 *    initial) épuise les priorités startP..3 (P1 fiche, P2 RA, P3 RS) sans se
 *    séparer, la priorité suivante est la PRIORITÉ 4 (« manches avec l'avance au
 *    pointage »), qui exige la feuille de pointage → NON automatisable → on lève
 *    le drapeau de vérification manuelle (__needsManualCheck) et on conserve un
 *    ordre alphabétique provisoire.
 *  - Le « recommencer à la priorité 1 » de la Note 2 n'intervient qu'APRÈS la
 *    Priorité 4 : il est donc hors de portée du code (résolu par le registraire via
 *    la colonne « Forcer rang » du tableau de bris d'égalité — resolveForcedRanks).
 *
 * Terminaison : la « continuation » réduit toujours strictement la taille du
 * groupe ; sinon on s'arrête au drapeau Priorité 4.
 *
 * @param {Array}   group        équipes à égalité (sous-ensemble de la passe)
 * @param {Array}   relevantGames portée fixée pour la passe (P1..P3 dessus)
 * @param {boolean} useAllGames  true = B/C, false = A
 * @param {number}  startP       priorité de départ (1..3)
 * @param {Object} [forced]      rangs forcés « Forcer rang » (Priorité 4), { team: number }
 * @return {Array} équipes ordonnées (meilleure en tête)
 */
function applyPriorities(group, relevantGames, useAllGames, startP, forced) {
  if (group.length <= 1) { return group.slice(); }
  forced = forced || {};

  // Métriques par priorité — TOUTES calculées sur la même `relevantGames`.
  //   P1 : différentiel V-D (plus haut = meilleur)
  //   P2 : ratio RA/DefInn  (plus bas  = meilleur)
  //   P3 : ratio RS/OffInn  (plus haut = meilleur)
  var priorities = [
    { metric: function (team) {
        var st = computeTeamStats(team, relevantGames, useAllGames);
        return st.v - st.d;
      }, descending: true },
    { metric: function (team) {
        return computeTeamStats(team, relevantGames, useAllGames).raRatio;
      }, descending: false },
    { metric: function (team) {
        return computeTeamStats(team, relevantGames, useAllGames).rsRatio;
      }, descending: true }
  ];

  for (var i = startP; i <= 3; i++) {
    var groups = groupByMetric(group, priorities[i - 1].metric, priorities[i - 1].descending);
    if (groups.length > 1) {
      // Séparation (au moins partielle) : on place les sous-groupes dans l'ordre
      // et on CONTINUE à la priorité suivante (i+1) sur la même portée pour ceux
      // encore à égalité (Note 2).
      var ordered = [];
      var manual = false;
      groups.forEach(function (sg) {
        var sub = applyPriorities(sg, relevantGames, useAllGames, i + 1, forced);
        if (sub.__needsManualCheck) { manual = true; }
        sub.forEach(function (t) { ordered.push(t); });
      });
      if (manual) { ordered.__needsManualCheck = true; }
      return ordered;
    }
  }

  // Priorités startP..3 (ratios automatisables) épuisées sans séparation.
  // La priorité suivante est la PRIORITÉ 4 (« manches avec l'avance au pointage »,
  // Art. 42.11) : elle exige la feuille de pointage -> non automatisable par le code.
  // Le registraire la calcule à la main (feuilles de pointage papier) et saisit
  // l'ordre obtenu via la colonne « Forcer rang » du tableau de bris d'égalité.
  // resolveForcedRanks applique cet override : si l'ordre saisi est strict et non
  // ambigu, le sous-groupe est RÉSOLU (pas de drapeau) ; sinon on conserve un ordre
  // provisoire et on lève __needsManualCheck. (Le « recommencer à la priorité 1 » de
  // la Note 2 vient APRÈS la Priorité 4 -> hors de portée du code.)
  var res = resolveForcedRanks(group, forced);
  if (!res.resolved) {
    Logger.log('AVERTISSEMENT : Priorité 4 atteinte pour : ' + group.join(', ') +
               '. Saisir l\'ordre dans la colonne « Forcer rang » du bris d\'égalité.');
    res.ordered.__needsManualCheck = true;
  }
  return res.ordered;
}

/**
 * Applique l'override « Forcer rang » (Priorité 4) à un sous-groupe d'équipes que
 * le moteur n'a pas pu départager. Source de vérité PARTAGÉE entre le moteur de tri
 * (applyPriorities) et le rendu (writeTiebreakTable) pour ne jamais diverger.
 *
 * Ordre : les équipes AVEC un numéro forcé d'abord (numéro croissant), puis celles
 * SANS numéro (alphabétique). Seul l'ordre RELATIF compte (« 1,3 » == « 2,3 »).
 *
 * Résolu (ordre strict, non ambigu) ssi : aucun numéro en double ET au plus UNE
 * équipe sans numéro (la seule sans numéro est forcément dernière). Sinon, l'ordre
 * retourné reste provisoire et l'appelant garde le drapeau de vérification manuelle.
 *
 * @param {Array}  teams   sous-groupe encore à égalité (>= 2 équipes en pratique)
 * @param {Object} forced  map { nomÉquipe: number } pour la portée courante
 * @return {{ordered: Array, resolved: boolean}}
 */
function resolveForcedRanks(teams, forced) {
  forced = forced || {};
  var ordered = teams.slice().sort(function (a, b) {
    var fa = forced[a], fb = forced[b];
    var hasA = typeof fa === 'number', hasB = typeof fb === 'number';
    if (hasA && hasB) { return fa - fb; }      // deux numéros : croissant
    if (hasA) { return -1; }                    // numéroté avant non-numéroté
    if (hasB) { return 1; }
    return a < b ? -1 : (a > b ? 1 : 0);        // deux sans numéro : alphabétique
  });

  var nums = teams.map(function (t) { return forced[t]; })
                  .filter(function (x) { return typeof x === 'number'; });
  var seen = {};
  var hasDup = false;
  nums.forEach(function (x) { if (seen[x]) { hasDup = true; } seen[x] = true; });
  var blanks = teams.length - nums.length;
  var resolved = !hasDup && blanks <= 1;

  return { ordered: ordered, resolved: resolved };
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
 * Détermine le représentant d'un pool pour le « Meilleur 2e » (Étape B), en tenant
 * compte d'un éventuel forçage manuel de l'admin (chiffre 2 dans la colonne « Forcer 2e »).
 *
 * Pourquoi : la Note 5 (Art. 42.11) exclut les parties gagnées par forfait aux fins du
 * meilleur deuxième, ce qui peut changer QUELLE équipe devrait représenter le pool. Plutôt
 * que de re-classer automatiquement chaque pool sans forfait (rare et risqué), l'admin force
 * le bon représentant. Ce forçage a préséance, SAUF s'il vise la 1re équipe (déjà qualifiée
 * via l'Étape C) — auquel cas il est ignoré avec avertissement. Fonction PURE (testable).
 *
 * @param {Array} standings    stats du pool triées, chaque élément ayant .team et .rank
 * @param {Array} markedTeams  noms des équipes du pool marquées « 2 » par l'admin
 * @return {{team:string, forced:boolean, warning:string}}
 *         team = représentant retenu ; forced = vrai si un forçage valide s'applique ;
 *         warning = message non vide si un forçage a été ignoré (1re équipe ou ambiguïté).
 */
function resolveSecondRepresentative(standings, markedTeams) {
  function teamAtRank(r) {
    for (var i = 0; i < standings.length; i++) {
      if (standings[i].rank === r) { return standings[i].team; }
    }
    return '';
  }
  var rank2 = teamAtRank(2);

  // Ne garder que les marques portant sur une équipe réellement dans ce pool.
  var marks = (markedTeams || []).filter(function (t) {
    return standings.some(function (s) { return s.team === t; });
  });
  if (marks.length === 0) {
    return { team: rank2, forced: false, warning: '' };
  }

  var rank1 = teamAtRank(1);
  var invalidFirst = marks.filter(function (t) { return t === rank1; });
  var valid        = marks.filter(function (t) { return t !== rank1; });

  if (valid.length === 1) {
    // Forçage valide. On signale tout de même une éventuelle marque parasite sur la 1re.
    var w = invalidFirst.length > 0
      ? 'Marque « 2 » ignorée sur la 1re équipe (' + invalidFirst.join(', ') + ').'
      : '';
    return { team: valid[0], forced: true, warning: w };
  }
  if (valid.length > 1) {
    return {
      team: rank2, forced: false,
      warning: 'Forçage du 2e ambigu (' + valid.join(', ') +
               ') — ignoré, 2e automatique conservé.'
    };
  }
  // Uniquement des marques sur la 1re équipe.
  return {
    team: rank2, forced: false,
    warning: 'Forçage du 2e sur la 1re équipe (' + invalidFirst.join(', ') +
             ') — ignoré : elle est déjà qualifiée comme 1re (Étape C).'
  };
}

/**
 * Relit, AVANT la reconstruction de la feuille Classements, les forçages du « 2e de pool »
 * saisis par l'admin (chiffre 2 en colonne 13 « Forcer 2e », à côté d'une équipe). Indexés
 * par NOM d'équipe (stable pendant un tournoi) pour survivre au sheet.clear() qui suit.
 *
 * @param {Sheet} sheet  feuille Classements existante (peut être vierge)
 * @return {Object} { nomÉquipe: true } pour chaque équipe marquée
 */
function readSecondOverrides(sheet) {
  var map = {};
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 13) { return map; }
  var values = sheet.getRange(1, 1, lastRow, 13).getValues();
  values.forEach(function (r) {
    var team = String(r[1]).trim();    // col B (index 1) = Équipe
    var mark = r[12];                   // col M (index 12) = Forcer 2e (case à cocher)
    // Case cochée (true) ; compat. ascendante avec un ancien « 2 » saisi à la main.
    var marked = (mark === true) || (String(mark).trim() === '2');
    if (team !== '' && marked) { map[team] = true; }
  });
  return map;
}

/**
 * Relit, AVANT la reconstruction de la feuille Classements, les rangs forcés saisis
 * par le registraire dans la colonne « Forcer rang » (col 24) des tableaux de bris
 * d'égalité — pour résoudre la Priorité 4 (Art. 42.11) calculée à la main.
 *
 * Chaque rang est rattaché à sa PORTÉE de bris (le tri concerné) en suivant les
 * titres de section de la colonne A, alignés ligne par ligne avec le bloc bris :
 *   « POOL p — … » -> scope "A"+p (Étape A du pool) ; « … Étape C » -> "C" ;
 *   « … Étape B » -> "B". Une équipe peut figurer dans 2 tableaux (son pool ET une
 * Étape B/C) ; le scope les distingue. Clés internes = nom d'équipe (stable pendant
 * le tournoi), comme readSecondOverrides — survit au sheet.clear() qui suit.
 *
 * @param {Sheet} sheet  feuille Classements existante (peut être vierge)
 * @return {Object} { scopeId: { nomÉquipe: number } }
 */
function readForcedRanks(sheet) {
  var map = {};
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 24) { return map; }   // pas de colonne « Forcer rang »
  var values = sheet.getRange(1, 1, lastRow, 24).getValues();
  var scope = null;
  values.forEach(function (r) {
    var aText = String(r[0] || '').trim();           // col A = titres de section
    if (aText) {
      var m = aText.match(/^POOL\s+(\d+)/);
      if (m) { scope = 'A' + m[1]; }
      else if (aText.indexOf('Étape C') !== -1) { scope = 'C'; }
      else if (aText.indexOf('Étape B') !== -1) { scope = 'B'; }
    }
    if (!scope) { return; }
    var team = String(r[14] || '').trim();           // col 15 = Équipe du bloc bris
    if (team === '') { return; }
    var raw = r[23];                                  // col 24 = Forcer rang
    var rank = Number(raw);
    if (raw !== '' && raw !== null && !isNaN(rank) && rank > 0) {
      if (!map[scope]) { map[scope] = {}; }
      map[scope][team] = rank;
    }
  });
  return map;
}

/**
 * Calcule le MODÈLE de classement d'une classe — purement en mémoire, AUCUNE écriture
 * de feuille. Rejoue exactement la même séquence de calcul que buildStandingsSheet
 * (mêmes fonctions pures : getGameResults, calculatePoolStandings, calculateStep,
 * resolveSecondRepresentative, computeTeamStats) et lit les mêmes forçages admin
 * (readSecondOverrides / readForcedRanks) — de sorte que l'affichage public (web app
 * doGet) ne peut pas diverger des feuilles Classements. Retourne un objet JSON-sérialisable.
 *
 * @param {Spreadsheet} ss
 * @param {string} classe  'A' ou 'B'
 * @return {Object} modèle { classe, pools[], firsts[], seconds[], semifinals, updatedAt }
 */
function computeStandingsModel(ss, classe) {
  var sheet = ss.getSheetByName(SHEET_STANDINGS[classe]);
  // Forçages admin (survivent au rebuild côté feuille ; ici on les lit pour rester aligné).
  var overrideByTeam = sheet ? readSecondOverrides(sheet) : {};
  var forcedRanks    = sheet ? readForcedRanks(sheet)    : {};

  var games       = getGameResults(classe);
  var teamsByPool = getTeams(classe);

  var firsts  = [];   // { team, pool } des 1ers
  var seconds = [];   // { team, pool, forced } des 2es (forçage Note 5 inclus)
  var poolStatsByTeam = {};
  var pools   = [];

  POOLS.forEach(function (p) {
    var teams     = teamsByPool[p] || [];
    var poolGames = games.filter(function (g) { return g.pool === p; });
    var forcedPool = forcedRanks['A' + p] || {};
    var standings = calculatePoolStandings(poolGames, teams, forcedPool);
    var markedInPool = teams.filter(function (t) { return overrideByTeam[t]; });
    var secondRep = resolveSecondRepresentative(standings, markedInPool);

    // Bandeaux ℹ/⚠ — mêmes règles que writePoolSection (Note 4 / supplémentaires).
    var suppGames    = poolGames.filter(gameIsSupp);
    var suppResolved = suppGames.filter(function (g) { return !g.suppNeedsTie; });
    var suppMissing  = suppGames.filter(function (g) { return g.suppNeedsTie; });

    pools.push({
      pool: p,
      // Affichage public épuré : ratios RÉELS (toutes manches), comme le tableau de pool.
      standings: standings.map(function (s) {
        return {
          rank: s.rank, team: s.team, pj: s.pj, v: s.v, d: s.d,
          rs: s.rs, ra: s.ra,                    // PP / PC (points réels)
          mo: formatFraction(s.offInnFull),      // manches offensives réelles
          md: formatFraction(s.defInnFull),      // manches défensives réelles
          rd: (s.defInnFull > 0 && isFinite(s.raRatioFull)) ? s.raRatioFull : null,
          ro: (s.offInnFull > 0) ? s.rsRatioFull : null
        };
      }),
      banners: {
        note4:        suppResolved.length > 0,
        suppMissing:  suppMissing.length > 0,
        forcedSecond: secondRep.forced,
        secondTeam:   secondRep.team,
        secondWarning: secondRep.warning || ''
      }
    });

    standings.forEach(function (s) {
      poolStatsByTeam[s.team] = s;
      if (s.rank === 1) { firsts.push({ team: s.team, pool: p }); }
    });
    seconds.push({ team: secondRep.team, pool: p, forced: secondRep.forced });
  });

  // Étapes C (1ers) et B (meilleur 2e) — mêmes appels que buildStandingsSheet.
  var orderedFirsts  = calculateStep(firsts.map(function (f) { return f.team; }),
                                     games, true, forcedRanks['C'] || {});
  var orderedSeconds = calculateStep(seconds.map(function (s) { return s.team; }),
                                     games, true, forcedRanks['B'] || {});

  var poolOf = {};
  firsts.forEach(function (f) { poolOf[f.team] = f.pool; });
  seconds.forEach(function (s) { poolOf[s.team] = s.pool; });

  function withPool(name) { return { team: name, pool: poolOf[name] || '' }; }

  // Les positions de demi-finale (classement INTER-pools) ne sont fiables qu'une fois
  // TOUTES les parties de pool de la classe jouées : avant cela on ne révèle aucun seed.
  var completion = poolPlayCompletion(classe);
  var reveal = completion.complete;

  // Rang FINAL (1-4) : seules les 4 équipes qualifiées en portent un. 1ers de pool
  // ordonnés par Étape C → 1/2/3 ; meilleur 2e (Étape B) → 4. Posé sur chaque ligne
  // de pool (s.seed) et utilisé par la colonne « Rang » + la carte « Meilleur 2e ».
  // Neutralisé (null) tant que les pools ne sont pas terminés.
  var seedByTeam = {};
  if (reveal) {
    orderedFirsts.forEach(function (name, i) { seedByTeam[name] = i + 1; });
    if (orderedSeconds[0]) { seedByTeam[orderedSeconds[0]] = 4; }
  }
  pools.forEach(function (pc) {
    pc.standings.forEach(function (s) { s.seed = seedByTeam[s.team] || null; });
  });

  // Entrées « Meilleur 2e » enrichies (V-D / RD réels, comme le tableau de pool) ;
  // l'ordre reste l'ordre officiel Étape B.
  var secondsCard = orderedSeconds.map(function (name, i) {
    var st = poolStatsByTeam[name] || {};
    return {
      team: name,
      pool: poolOf[name] || '',
      v: st.v || 0,
      d: st.d || 0,
      rd: (st.defInnFull > 0 && isFinite(st.raRatioFull)) ? st.raRatioFull : null,
      seed: (reveal && i === 0) ? 4 : null
    };
  });

  var p1 = orderedFirsts[0]  || '';
  var p2 = orderedFirsts[1]  || '';
  var p3 = orderedFirsts[2]  || '';
  var p4 = orderedSeconds[0] || '';   // meilleur 2e

  var tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();

  return {
    classe: classe,
    pools: pools,
    poolsComplete: reveal,
    poolPlayed: completion.played,
    poolTotal:  completion.total,
    firsts:  orderedFirsts.map(withPool),
    seconds: secondsCard,
    semifinals: {
      positions: [withPool(p1), withPool(p2), withPool(p3), withPool(p4)],
      demi1: { a: p1, b: p4 },
      demi2: { a: p2, b: p3 }
    },
    updatedAt: Utilities.formatDate(new Date(), tz, "d MMM yyyy 'à' HH'h'mm")
  };
}

/**
 * Construit la feuille Classements pour une classe : 3 pools + Étape C + Étape B.
 */
function buildStandingsSheet(ss, classe, games) {
  var sheet = ss.getSheetByName(SHEET_STANDINGS[classe]);
  if (!sheet) { sheet = getOrCreateSheet(ss, SHEET_STANDINGS[classe]); }

  // Forçages admin du « 2e de pool » (col 13) — relus AVANT l'effacement pour survivre
  // à la reconstruction ; réappliqués et réaffichés par writePoolSection.
  var overrideByTeam = readSecondOverrides(sheet);

  // Rangs forcés « Forcer rang » (col 24, Priorité 4) — relus AVANT l'effacement, par
  // portée de bris (A+pool / C / B) ; réappliqués par le moteur et réaffichés par
  // writeTiebreakTable.
  var forcedRanks = readForcedRanks(sheet);

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
  var poolData = [];   // données calculées par pool, écrites après le calcul des demi-finalistes

  // -------- SECTIONS 1-3 (calcul) : un classement par pool --------
  // En deux passes : la colonne « Avancement » doit identifier les 4 demi-finalistes
  // réels (Étapes B/C ci-dessous), qui dépendent des standings de TOUS les pools —
  // donc on calcule d'abord, on n'écrit qu'une fois ces 4 équipes connues.
  POOLS.forEach(function (p) {
    var teams = teamsByPool[p] || [];
    var poolGames = games.filter(function (g) { return g.pool === p; });

    var forcedPool = forcedRanks['A' + p] || {};
    var standings = calculatePoolStandings(poolGames, teams, forcedPool);

    // Représentant « 2e » de ce pool pour le Meilleur 2e (Étape B), forçage admin pris
    // en compte (Note 5 / forfaits) : par défaut le rang 2, sinon l'équipe forcée.
    var markedInPool = teams.filter(function (t) { return overrideByTeam[t]; });
    var secondRep = resolveSecondRepresentative(standings, markedInPool);

    poolData.push({ pool: p, standings: standings, poolGames: poolGames,
                     markedInPool: markedInPool, secondRep: secondRep, forcedPool: forcedPool });

    standings.forEach(function (s) {
      poolStatsByTeam[s.team] = s;
      if (s.rank === 1) { firsts.push({ team: s.team, pool: p }); }
    });
    seconds.push({ team: secondRep.team, pool: p, forced: secondRep.forced });
  });

  // -------- Étapes B/C : les 4 demi-finalistes --------
  var firstTeams = firsts.map(function (f) { return f.team; });
  var forcedC = forcedRanks['C'] || {};
  var orderedFirsts = calculateStep(firstTeams, games, true, forcedC);

  var secondTeams = seconds.map(function (s) { return s.team; });
  var forcedB = forcedRanks['B'] || {};
  var orderedSeconds = calculateStep(secondTeams, games, true, forcedB);

  // team -> numéro de demi-finale (1-4). Seules CES 4 équipes au total (toutes pools
  // confondues) avancent : les 1ers de pool (positions 1-3) et le seul meilleur 2e
  // retenu (position 4) — un 2e non retenu, un 3e ou un 4e ne reçoit rien.
  var seedByTeam = {};
  orderedFirsts.slice(0, 3).forEach(function (team, i) { seedByTeam[team] = i + 1; });
  if (orderedSeconds.length > 0) { seedByTeam[orderedSeconds[0]] = 4; }

  // Les positions de demi-finale (classement INTER-pools : Sections 4/5, récap, et la
  // colonne « Avancement ») ne sont fiables qu'une fois TOUTES les parties de pool de
  // la classe jouées. Tant que ce n'est pas le cas, on n'affiche pas les seeds ni les
  // sections inter-pools — seulement un bandeau d'attente. Les classements internes de
  // chaque pool, eux, restent affichés en direct.
  var completion = poolPlayCompletion(classe);
  var reveal = completion.complete;

  // -------- SECTIONS 1-3 (écriture) : un classement par pool --------
  poolData.forEach(function (pd) {
    row = writePoolSection(sheet, row, classe, pd.pool, pd.standings, pd.poolGames,
                           pd.markedInPool, pd.secondRep, pd.forcedPool,
                           reveal ? seedByTeam : {});
    row += 1;  // espace entre sections
  });

  if (reveal) {
    // -------- SECTION 4 : Étape C — classement des 1ers --------
    row = writeAdvancementSection(
      sheet, row, classe,
      'SECTION 4 — CLASSEMENT DES 1ers (Positions 1-2-3) — Étape C',
      orderedFirsts, games, poolStatsByTeam, firsts, 1, forcedC);

    row += 1;

    // -------- SECTION 5 : Étape B — meilleur 2e (position 4) --------
    row = writeAdvancementSection(
      sheet, row, classe,
      'SECTION 5 — MEILLEUR 2e (Position 4) — Étape B',
      orderedSeconds, games, poolStatsByTeam, seconds, 4, forcedB);

    row += 1;

    // -------- RÉCAPITULATIF DEMI-FINALES --------
    writeSemifinalSummary(sheet, row, classe, orderedFirsts, orderedSeconds);
  } else {
    // Pools non terminés : bandeau d'attente à la place des positions de demi-finale.
    writeSeedingPendingBanner(sheet, row, classe, completion);
  }

  // Largeurs de colonnes. 1-13 = classement de gauche ; 14 = espace ; 15-23 =
  // bloc « bris d'égalité » (Équipe, V-D, PP, PC, MO, MD, RD, RO, Critère décisif) ;
  // 24 = « Forcer rang » (saisie Priorité 4).
  var widths = [55, 170, 45, 45, 45, 80, 80, 75, 75, 105, 105, 120, 70,
                20, 160, 50, 50, 50, 60, 60, 65, 65, 175, 80];
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
      's\'est terminée hâtivement (walk-off / mercy). Inclut les manches supplémentaires. ' +
      'Dénominateur du ratio offensif RO.',
  9:  'MD — Manches DÉFENSIVES jouées (au champ), en fractions de tiers (⅓/⅔) si fin hâtive. ' +
      'Inclut les manches supplémentaires. Dénominateur du ratio défensif RD.',
  10: 'RD — Ratio DÉFENSIF = PC ÷ MD (points contre par manche défensive), sur TOUTES les ' +
      'manches jouées (supplémentaires incluses). Le plus BAS est le meilleur. 2e critère de ' +
      'bris d\'égalité (Art. 42.11), après la fiche tête-à-tête. Note : pour départager une ' +
      'égalité, le tableau de bris d\'égalité (à droite) recalcule ce ratio en EXCLUANT les ' +
      'supplémentaires (Note 4) ; sa valeur peut donc différer de celle affichée ici.',
  11: 'RO — Ratio OFFENSIF = PP ÷ MO (points pour par manche offensive), sur TOUTES les manches ' +
      'jouées (supplémentaires incluses). Le plus HAUT est le meilleur. 3e critère de bris ' +
      'd\'égalité (Art. 42.11). Le tableau de bris d\'égalité, lui, exclut les supplémentaires ' +
      '(Note 4).',
  12: 'AVANCEMENT — Numéro de demi-finale (1 à 4) UNIQUEMENT pour les 4 équipes qui s\'y ' +
      'qualifient réellement (les 3 gagnants de pool + le seul meilleur 2e retenu, Sections ' +
      '4-5 ci-dessous) ; vide pour toutes les autres (2e non retenu, 3e, 4e de pool). ' +
      'Reste vide TANT QUE toutes les parties de pool de la classe ne sont pas jouées : ' +
      'ce classement inter-pools est provisoire avant la fin des pools, donc masqué jusque-là.',
  13: 'FORCER 2e — Réservé à l\'admin. COCHER la case à côté d\'une équipe pour la désigner ' +
      'comme représentante de ce pool au « Meilleur 2e » (Étape B), à la place du 2e ' +
      'automatique ; DÉCOCHER pour revenir au 2e automatique. Utile quand une victoire par ' +
      'FORFAIT a faussé la 2e place : la Note 5 (Art. 42.11) exclut ces parties aux fins du ' +
      'meilleur 2e. Le forçage a PRÉSÉANCE. Cocher la case de la 1re équipe est IGNORÉ (elle ' +
      'est déjà qualifiée comme 1re, Étape C). Le classement se recalcule automatiquement dès ' +
      'qu\'on coche/décoche (si la mise à jour auto est activée) ; sinon lancer « Mettre à jour ' +
      'les classements ».'
};

var ADV_HEADER_NOTES = {
  1:  'POSITION — Rang inter-pool. Étape C : positions 1-2-3 (les 1ers de pool). ' +
      'Étape B : meilleur 2e = position 4. Ces sections n\'apparaissent qu\'une fois ' +
      'TOUTES les parties de pool de la classe jouées (avant, le classement inter-pools ' +
      'est provisoire) ; d\'ici là, un bandeau d\'attente les remplace.',
  3:  'POOL — Pool d\'origine de l\'équipe.',
  4:  'V — Victoires (sur TOUTES les parties de pool de l\'équipe).',
  5:  'D — Défaites (sur toutes les parties de pool).',
  6:  'RD — Ratio DÉFENSIF = PC / MD (le plus BAS est le meilleur). Calculé sur toutes les ' +
      'parties de pool (forfaits exclus des ratios).',
  7:  'RO — Ratio OFFENSIF = PP / MO (le plus HAUT est le meilleur). Calculé sur toutes les ' +
      'parties de pool (forfaits exclus des ratios).',
  8:  'PP — Points POUR (sur toutes les parties de pool).',
  9:  'PC — Points CONTRE (sur toutes les parties de pool).',
  10: 'NOTE — Avertissements : « 🔒 2e forcé par le registraire » si cette équipe a été désignée ' +
      'manuellement comme 2e de son pool (colonne « Forcer 2e », Note 5 / forfaits) au lieu du 2e ' +
      'automatique ; « ⚠ Vérif. manuelle (P4) » si la Priorité 4 (« manches en avance », non ' +
      'automatisée) est atteinte et pas encore résolue — saisir l\'ordre dans la colonne ' +
      '« Forcer rang » du tableau de bris d\'égalité à droite ; « 🔒 Rang forcé (P4) » une fois ' +
      'cet ordre saisi ; « Note 4 appliquée (suppl.) » si une partie de l\'équipe est allée en ' +
      'supplémentaires et que la Note 4 a été appliquée ; « pointage régl. manquant » si une ' +
      'partie supplémentaire n\'a pas son Pointage régl. (col. O) saisi.'
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

// Première colonne du bloc « bris d'égalité » écrit À DROITE des classements.
// (Col. 13 = mince séparateur du tableau de gauche ; col. 14 = espace.)
var TIEBREAK_START_COL = 15;
var TIEBREAK_NCOLS = 9;   // Équipe, V-D, PP, PC, MO, MD, RD, RO, Critère décisif
                          // (+ 1 colonne « Forcer rang » ajoutée par writeTiebreakTable)

// Info-bulle de la colonne « Critère décisif » du bloc bris.
var TIEBREAK_CRIT_NOTE =
  'CRITÈRE DÉCISIF — Critère de l\'Art. 42.11 qui classe cette équipe DEVANT celle ' +
  'de la ligne suivante du même groupe à égalité. Au sein d\'un groupe les équipes ' +
  'ont la MÊME fiche V-D (Priorité 1), donc le départage vient de : « RD » (ratio ' +
  'défensif PC/MD, le plus bas gagne — Priorité 2), « RO » (ratio offensif PP/MO, ' +
  'le plus haut gagne — Priorité 3) ou « Manuel (P4) » (manches avec l\'avance au ' +
  'pointage — non automatisable). En P4, calculez l\'ordre à la main (feuilles de ' +
  'pointage papier) puis saisissez-le dans la colonne « Forcer rang » : le critère ' +
  'passe alors à « 🔒 Forcé (P4) ». Les ratios sont en base RÉGULIÈRE : les manches ' +
  'supplémentaires sont exclues (Note 4).';

// Info-bulle de la colonne « Forcer rang » (saisie de la Priorité 4 par le registraire).
var FORCE_RANK_NOTE =
  'FORCER RANG (Priorité 4, Art. 42.11) — À remplir UNIQUEMENT quand le critère ' +
  'décisif affiche « ⚠ Manuel (P4) » : le système ne peut pas départager ces équipes ' +
  '(il faudrait les manches avec l\'avance au pointage, qu\'il n\'a pas). Calculez ' +
  'l\'ordre à la main sur les feuilles de pointage papier, puis saisissez un numéro ' +
  'd\'ordre (1 = meilleur, 2, 3 …) sur chaque équipe à égalité. Seul l\'ORDRE RELATIF ' +
  'compte (« 1,3 » équivaut à « 1,2 »). Le classement se résout dès qu\'aucun numéro ' +
  'n\'est en double et qu\'au plus une équipe reste sans numéro. Cette saisie a ' +
  'PRÉSÉANCE là où le système n\'a pas tranché (P4 seulement) ; un numéro sur une ' +
  'équipe déjà départagée par RD/RO est ignoré. Le recalcul est automatique (si la ' +
  'mise à jour auto est activée) ; sinon lancer « Mettre à jour les classements ».';

/**
 * Détermine le critère (Art. 42.11) qui classe l'équipe `hi` (mieux classée)
 * devant `lo`, en comparant leurs métriques dans l'ordre des priorités
 * automatisables — exactement comme groupByMetric dans applyPriorities (P2 ratio
 * défensif, puis P3 ratio offensif). Si ni l'un ni l'autre ne sépare : P4 manuelle.
 *
 * @param {Object} hi  stats de l'équipe la mieux classée (computeTeamStats)
 * @param {Object} lo  stats de l'équipe juste en dessous
 * @return {string} libellé du critère, avec les valeurs comparées
 */
function decisiveCriterion(hi, lo) {
  // P1 : fiche tête-à-tête (V-D dans les parties entre les équipes à égalité).
  // Dans un groupe affiché (regroupé par fiche GLOBALE), deux équipes peuvent
  // différer sur leur fiche tête-à-tête restreinte -> c'est alors le critère décisif.
  if ((hi.v - hi.d) !== (lo.v - lo.d)) {
    return 'Fiche ' + hi.v + '-' + hi.d + ' > ' + lo.v + '-' + lo.d;
  }
  // P2 : ratio défensif RA/DefInn — le plus BAS gagne.
  if (!approxEqual(hi.raRatio, lo.raRatio)) {
    return 'RD ' + round3(hi.raRatio) + ' < ' + round3(lo.raRatio);
  }
  // P3 : ratio offensif RS/OffInn — le plus HAUT gagne.
  if (!approxEqual(hi.rsRatio, lo.rsRatio)) {
    return 'RO ' + round3(hi.rsRatio) + ' > ' + round3(lo.rsRatio);
  }
  // Ratios épuisés sans départage -> Priorité 4 (manuelle).
  return '⚠ Manuel (P4)';
}

/**
 * Texte d'info-bulle (setNote) détaillant le CALCUL du ratio défensif d'une
 * équipe, pour survol de la cellule RD : « PC régl. ÷ MD = ratio ». Numérateur
 * et dénominateur en base RÉGULIÈRE (manches supplémentaires exclues, Note 4),
 * donc ce sont EXACTEMENT les valeurs utilisées par le bris d'égalité.
 * @param {Object} s  stats issues de computeTeamStats
 * @return {string}
 */
function rdCalcNote(s) {
  if (!(s.defInn > 0)) {
    return 'RD (ratio défensif) — indisponible : aucune manche défensive comptabilisée.';
  }
  return 'RD = points contre (manches régulières) ÷ manches défensives = ' +
    s.raNum + ' ÷ ' + formatFraction(s.defInn) + ' = ' + round3(s.raRatio) +
    '. Le plus BAS est le meilleur (Priorité 2, Art. 42.11). Manches ' +
    'supplémentaires exclues (Note 4).';
}

/**
 * Texte d'info-bulle (setNote) détaillant le CALCUL du ratio offensif d'une
 * équipe, pour survol de la cellule RO : « PP régl. ÷ MO = ratio » (base
 * régulière, Note 4). Voir rdCalcNote.
 * @param {Object} s  stats issues de computeTeamStats
 * @return {string}
 */
function roCalcNote(s) {
  if (!(s.offInn > 0)) {
    return 'RO (ratio offensif) — indisponible : aucune manche offensive comptabilisée.';
  }
  return 'RO = points pour (manches régulières) ÷ manches offensives = ' +
    s.rsNum + ' ÷ ' + formatFraction(s.offInn) + ' = ' + round3(s.rsRatio) +
    '. Le plus HAUT est le meilleur (Priorité 3, Art. 42.11). Manches ' +
    'supplémentaires exclues (Note 4).';
}

/**
 * Info-bulle (setNote) du CALCUL du ratio défensif RÉEL affiché dans le TABLEAU
 * DE POOL : « PC ÷ MD = ratio » sur TOUTES les manches jouées (supplémentaires
 * incluses). Contrairement à rdCalcNote (base régulière, Note 4, utilisé dans le
 * tableau de bris d'égalité), ici PC ÷ MD tombe juste.
 * @param {Object} s  stats issues de computeTeamStats
 * @return {string}
 */
function rdCalcNoteFull(s) {
  if (!(s.defInnFull > 0)) {
    return 'RD (ratio défensif) — indisponible : aucune manche défensive comptabilisée.';
  }
  return 'RD = points contre ÷ manches défensives jouées = ' +
    s.ra + ' ÷ ' + formatFraction(s.defInnFull) + ' = ' + round3(s.raRatioFull) +
    '. Toutes les manches jouées sont incluses (supplémentaires comprises). ' +
    'Le bris d\'égalité, lui, EXCLUT les supplémentaires (Note 4, Art. 42.11) — ' +
    'voir le tableau de bris d\'égalité à droite.';
}

/**
 * Info-bulle (setNote) du CALCUL du ratio offensif RÉEL affiché dans le tableau
 * de pool : « PP ÷ MO = ratio » sur toutes les manches jouées. Voir rdCalcNoteFull.
 * @param {Object} s  stats issues de computeTeamStats
 * @return {string}
 */
function roCalcNoteFull(s) {
  if (!(s.offInnFull > 0)) {
    return 'RO (ratio offensif) — indisponible : aucune manche offensive comptabilisée.';
  }
  return 'RO = points pour ÷ manches offensives jouées = ' +
    s.rs + ' ÷ ' + formatFraction(s.offInnFull) + ' = ' + round3(s.rsRatioFull) +
    '. Toutes les manches jouées sont incluses (supplémentaires comprises). ' +
    'Le bris d\'égalité, lui, EXCLUT les supplémentaires (Note 4, Art. 42.11).';
}

/**
 * Écrit, À DROITE d'une section de classement, le DÉTAIL DES BRIS D'ÉGALITÉ :
 * pour chaque groupe d'équipes à fiche V-D identique (sur la portée du bris),
 * montre les stats RÉELLEMENT utilisées par le moteur (RD/RO en base régulière —
 * manches supplémentaires exclues, Note 4) et le critère de l'Art. 42.11 qui a
 * départagé chaque équipe de celle classée juste au-dessus dans le même groupe.
 *
 * Pourquoi : le classement de gauche ne montre qu'un ordre trié ; quand deux
 * équipes ont la même fiche, on ne voit pas QUEL critère a tranché ni que les
 * supplémentaires sont exclues. Ce bloc rend le calcul transparent.
 *
 * Fidélité au moteur : la PORTÉE et les MÉTRIQUES sont calculées EXACTEMENT comme
 * dans tiebreaker()/applyPriorities() — même headToHeadGames / filtre « impliquant »,
 * et même appel computeTeamStats(team, portée, useAllGames). Les colonnes RD/RO —
 * ainsi que PP/PC/MO/MD — affichées ici sont en base RÉGULIÈRE (Note 4 : manches
 * supplémentaires exclues), si bien que RD = PC ÷ MD y tombe juste. Quand un pool a
 * une partie allée en supplémentaires, ces valeurs DIFFÈRENT de celles du tableau de
 * gauche (qui montre les ratios RÉELS, supplémentaires incluses) : c'est voulu, et le
 * bandeau « ℹ Manches supplémentaires exclues » sous ce tableau l'explique.
 *
 * @param {Sheet}   sheet
 * @param {number}  startRow     ligne du titre de la section de gauche (alignement)
 * @param {Array}   teams        toutes les équipes classées dans la section
 * @param {Array}   games        parties disponibles (pool pour A ; classe pour B/C)
 * @param {Array}   orderedNames ordre final résolu (noms d'équipes, meilleur en tête)
 * @param {boolean} useAllGames  true = Étapes B/C, false = Étape A (tête-à-tête)
 * @param {Object}  forced       rangs forcés « Forcer rang » (Priorité 4) de la portée
 * @return {number} dernière ligne écrite + 1
 */
function writeTiebreakTable(sheet, startRow, teams, games, orderedNames, useAllGames, forced) {
  forced = forced || {};
  var c0   = TIEBREAK_START_COL;
  var n    = TIEBREAK_NCOLS;        // 9 colonnes de stats
  var nAll = n + 1;                 // + colonne « Forcer rang » (Priorité 4)
  var fCol = c0 + n;                // colonne « Forcer rang » (= 24)
  var P4   = '⚠ Manuel (P4)';       // libellé renvoyé par decisiveCriterion en P4
  var row  = startRow;

  // Portée de REGROUPEMENT (identification des égalités) : fiche GLOBALE — le même
  // critère que le moteur (orderTeams) pour découper orderedNames en groupes.
  //  - Étape A : toutes les parties du pool passé -> fiche V-D globale du pool.
  //  - Étapes B/C : toutes les parties impliquant une des équipes.
  var groupScope = useAllGames
    ? games.filter(function (g) {
        return teams.indexOf(g.local) !== -1 || teams.indexOf(g.visiteur) !== -1;
      })
    : games;

  var groupStat = {};
  orderedNames.forEach(function (t) {
    groupStat[t] = computeTeamStats(t, groupScope, useAllGames);
  });
  function vd(t) { return groupStat[t].v - groupStat[t].d; }

  // Regroupe l'ordre final en « runs » de fiche GLOBALE identique. Ces équipes
  // sont forcément consécutives (le tri primaire du moteur est la fiche globale).
  var runs = [];
  orderedNames.forEach(function (t) {
    var last = runs[runs.length - 1];
    if (last && vd(last[0]) === vd(t)) { last.push(t); }
    else { runs.push([t]); }
  });
  var tieGroups = runs.filter(function (r) { return r.length >= 2; });
  var anySuppInScope = false;   // -> bandeau Note 4 (portées d'affichage réunies)

  // Titre du bloc.
  var scopeLabel = useAllGames ? 'toutes parties de pool' : 'tête-à-tête';
  sheet.getRange(row, c0, 1, nAll).merge();
  sheet.getRange(row, c0).setValue('BRIS D\'ÉGALITÉ (' + scopeLabel + ')')
    .setFontWeight('bold').setBackground(COLOR_SECTION)
    .setHorizontalAlignment('center');
  row++;

  // Pas d'égalité : message court (confirme que le moteur a vérifié).
  if (tieGroups.length === 0) {
    sheet.getRange(row, c0, 1, nAll).merge();
    sheet.getRange(row, c0)
      .setValue('Aucune égalité à départager — rangs établis par la fiche V-D.')
      .setWrap(true).setBackground(COLOR_CALC).setVerticalAlignment('top');
    return row + 1;
  }

  // En-têtes. PP/PC/MO/MD sont en base RÉGULIÈRE (Note 4) — comme RD/RO ici —
  // pour que RD = PC ÷ MD tombe juste dans ce tableau de bris. La 10e (« Forcer
  // rang ») est une saisie du registraire pour la Priorité 4.
  var headers = ['Équipe', 'V-D', 'PP', 'PC', 'MO', 'MD', 'RD', 'RO',
                 'Critère décisif', 'Forcer rang'];
  sheet.getRange(row, c0, 1, nAll).setValues([headers]);
  styleHeader(sheet.getRange(row, c0, 1, nAll));
  sheet.getRange(row, c0 + 8).setNote(TIEBREAK_CRIT_NOTE);
  sheet.getRange(row, fCol).setNote(FORCE_RANK_NOTE);
  row++;

  // Un sous-bloc par groupe à égalité (séparés par une ligne vide).
  tieGroups.forEach(function (grp, gi) {
    if (gi > 0) { row++; }   // espace entre groupes

    // Portée D'AFFICHAGE de CE groupe — identique à celle que le moteur a utilisée
    // pour départager ces équipes (Art. 42.11) :
    //  - Étape A : parties jouées STRICTEMENT entre les équipes de ce groupe (tête-à-tête).
    //  - Étapes B/C : toute la portée « impliquant » (Note 2, non re-restreinte).
    var scope = useAllGames ? groupScope : headToHeadGames(grp, games);
    if (scope.some(gameIsSupp)) { anySuppInScope = true; }
    var statByTeam = {};
    grp.forEach(function (t) {
      statByTeam[t] = computeTeamStats(t, scope, useAllGames);
    });

    // Détecte les SOUS-GROUPES Priorité 4 du groupe : suites d'équipes consécutives
    // dont le moteur n'a pu départager aucun couple (decisiveCriterion = P4). Ce sont
    // les seules lignes où « Forcer rang » est saisissable, et où son override agit.
    var runIdOf = {};      // équipe -> id de sous-groupe P4
    var p4runs = [];       // [[équipes], ...]
    grp.forEach(function (t, i) {
      if (i === 0) { return; }
      if (decisiveCriterion(statByTeam[grp[i - 1]], statByTeam[t]) === P4) {
        var prev = grp[i - 1];
        if (runIdOf[prev] === undefined) {
          runIdOf[prev] = p4runs.length;
          p4runs.push([prev]);
        }
        runIdOf[t] = runIdOf[prev];
        p4runs[runIdOf[t]].push(t);
      }
    });
    var runResolved = p4runs.map(function (rn) {
      return resolveForcedRanks(rn, forced).resolved;
    });

    grp.forEach(function (t, i) {
      var s  = statByTeam[t];
      var rd = (s.defInn > 0 && isFinite(s.raRatio)) ? s.raRatio.toFixed(3) : '—';
      var ro = (s.offInn > 0) ? s.rsRatio.toFixed(3) : '—';
      var crit = (i === 0) ? '—' : decisiveCriterion(statByTeam[grp[i - 1]], s);
      var inP4 = runIdOf[t] !== undefined;
      // En P4 résolue par override, on annonce « 🔒 Forcé (P4) » au lieu de « Manuel ».
      if (crit === P4 && inP4 && runResolved[runIdOf[t]]) { crit = '🔒 Forcé (P4)'; }
      // PP/PC/MO/MD en base RÉGULIÈRE (Note 4) : ce sont les chiffres exacts qui
      // produisent les ratios RD/RO du bris d'égalité.
      sheet.getRange(row, c0, 1, n).setValues([[
        t, s.v + '-' + s.d, s.rsNum, s.raNum,
        formatFraction(s.offInn), formatFraction(s.defInn),
        rd, ro, crit
      ]]);
      sheet.getRange(row, c0, 1, n).setBackground(COLOR_CALC).setWrap(true)
        .setVerticalAlignment('top');
      // Colonne « Forcer rang » : saisissable (jaune) uniquement sur les lignes P4 ;
      // pré-remplie avec la valeur relue pour survivre à la reconstruction.
      var fCell = sheet.getRange(row, fCol);
      if (inP4) {
        fCell.setBackground(COLOR_INPUT).setHorizontalAlignment('center')
             .setVerticalAlignment('top');
        fCell.setValue(typeof forced[t] === 'number' ? forced[t] : '');
      } else {
        fCell.setBackground(COLOR_CALC);
      }
      // Calcul détaillé au survol des cellules RD / RO.
      sheet.getRange(row, c0 + 6).setNote(rdCalcNote(s));
      sheet.getRange(row, c0 + 7).setNote(roCalcNote(s));
      row++;
    });
  });

  // Note 4 si une partie affichée (une des portées de groupe) est allée en supplémentaires.
  if (anySuppInScope) {
    sheet.getRange(row, c0, 1, nAll).merge();
    sheet.getRange(row, c0)
      .setValue('ℹ Manches supplémentaires exclues des ratios RD/RO (Note 4, Art. 42.11).')
      .setWrap(true).setBackground(COLOR_SECOND).setVerticalAlignment('top');
    row++;
  }

  return row;
}

/**
 * Écrit une section de classement de pool. Retourne la prochaine ligne libre.
 * @param {Array}  poolGames    parties du pool — pour signaler les manches supplémentaires.
 * @param {Array}  markedInPool noms des équipes du pool marquées « 2 » (forçage admin du 2e).
 * @param {Object} secondRep    résultat de resolveSecondRepresentative (team/forced/warning).
 * @param {Object} forced       rangs forcés « Forcer rang » de ce pool (Priorité 4).
 * @param {Object} seedByTeam   team -> numéro de demi-finale (1-4), pour les 4 SEULES
 *                              équipes (toutes pools confondues) qui s'y qualifient.
 */
function writePoolSection(sheet, startRow, classe, pool, standings, poolGames,
                          markedInPool, secondRep, forced, seedByTeam) {
  markedInPool = markedInPool || [];
  secondRep = secondRep || { team: '', forced: false, warning: '' };
  forced = forced || {};
  seedByTeam = seedByTeam || {};
  var row = startRow;

  // Titre de section.
  sheet.getRange(row, 1).setValue('POOL ' + pool + ' — CLASSE ' + classe);
  sheet.getRange(row, 1, 1, 13).merge();
  sheet.getRange(row, 1).setFontWeight('bold').setBackground(poolColor(pool))
    .setHorizontalAlignment('center');
  row++;

  // En-têtes (13 colonnes). La 13e (« Forcer 2e ») est une saisie admin (Note 5).
  var headers = ['Rang', 'Équipe', 'PJ', 'V', 'D', 'PP',
                 'PC', 'MO', 'MD',
                 'RD', 'RO', 'Avancement', 'Forcer 2e'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  styleHeader(sheet.getRange(row, 1, 1, headers.length));
  applyHeaderNotes(sheet, row, POOL_HEADER_NOTES);
  row++;

  // Lignes d'équipes.
  var firstTeamRow = row;
  standings.forEach(function (s) {
    // Demi-finale (1-4) SEULEMENT pour les 4 équipes qui s'y qualifient réellement
    // (1er de pool + le seul meilleur 2e retenu) — pas le rang de pool brut, qui
    // duplique la colonne « Rang » et n'indique pas la qualification.
    var advancement = seedByTeam[s.team] || '';
    // Tableau de pool : ratios RÉELS (toutes manches jouées, suppl. incluses).
    // L'exclusion des supplémentaires (Note 4) n'est appliquée que dans le tableau
    // de bris d'égalité à droite.
    var raRatioDisplay = s.defInnFull > 0 ? s.raRatioFull.toFixed(3) : '—';
    var rsRatioDisplay = s.offInnFull > 0 ? s.rsRatioFull.toFixed(3) : '—';
    var rowData = [
      s.rank, s.team, s.pj, s.v, s.d, s.rs, s.ra,
      formatFraction(s.offInnFull), formatFraction(s.defInnFull),
      raRatioDisplay, rsRatioDisplay, advancement,
      ''   // col 13 « Forcer 2e » : case à cocher insérée après la boucle
    ];
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);

    // Calcul détaillé au survol des cellules RD (col. 10) / RO (col. 11).
    sheet.getRange(row, 10).setNote(rdCalcNoteFull(s));
    sheet.getRange(row, 11).setNote(roCalcNoteFull(s));

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

  // Saisie admin « Forcer 2e » (col 13) : CASE À COCHER sur les lignes d'équipes.
  // Une case se coche/décoche d'un simple clic — facile à RETIRER, contrairement à
  // une liste déroulante à valeur unique (qui n'offrait aucune option « vide »).
  // La case de l'équipe forcée est rétablie à chaque reconstruction.
  var forceRange = sheet.getRange(firstTeamRow, 13, standings.length, 1);
  forceRange.insertCheckboxes();
  forceRange.setValues(standings.map(function (s) {
    return [markedInPool.indexOf(s.team) !== -1];   // true = forcée (case cochée)
  }));

  // Bandeau : forçage admin du 2e (Note 5 / forfaits).
  if (secondRep.forced) {
    sheet.getRange(row, 1, 1, 13).merge();
    sheet.getRange(row, 1)
      .setValue('ℹ 2e forcé par l\'admin pour le Meilleur 2e (Étape B) : ' + secondRep.team +
                ' — Note 5, Art. 42.11 (forfaits exclus du meilleur 2e).')
      .setWrap(true).setBackground(COLOR_SECOND).setVerticalAlignment('top');
    row++;
  }
  if (secondRep.warning) {
    sheet.getRange(row, 1, 1, 13).merge();
    sheet.getRange(row, 1)
      .setValue('⚠ ' + secondRep.warning)
      .setWrap(true).setBackground(COLOR_INPUT).setVerticalAlignment('top');
    row++;
  }

  // Note 4 (Art. 42.11) — parties allées en manches supplémentaires.
  var suppGames = (poolGames || []).filter(gameIsSupp);
  function gameLabel(g) {
    return g.local + ' vs ' + g.visiteur + (g.partie ? ' (partie #' + g.partie + ')' : '');
  }
  // (a) Parties résolues (Pointage régl. saisi). Le tableau de pool (à gauche) affiche
  // les ratios RÉELS, supplémentaires INCLUSES ; l'exclusion Note 4 ne s'applique qu'au
  // classement de bris d'égalité (à droite). Le bandeau l'explique pour éviter la confusion.
  var suppResolved = suppGames.filter(function (g) { return !g.suppNeedsTie; });
  if (suppResolved.length > 0) {
    sheet.getRange(row, 1, 1, 13).merge();
    sheet.getRange(row, 1)
      .setValue('ℹ Partie(s) en manches supplémentaires : ' +
                suppResolved.map(gameLabel).join(' ; ') + '. Dans ce tableau de pool, les ' +
                'ratios RD/RO incluent toutes les manches jouées (supplémentaires comprises). ' +
                'Pour départager les égalités, la Note 4 (Art. 42.11) exclut les manches ' +
                'supplémentaires : les ratios ajustés selon cette règle apparaissent dans le ' +
                'tableau de bris d\'égalité à droite.')
      .setWrap(true).setBackground(COLOR_SECOND).setVerticalAlignment('top');
    row++;
  }
  // (b) Parties supplémentaires sans Pointage régl. saisi : Note 4 NON applicable au bris
  // d'égalité tant que la donnée manque. Même structure que le bandeau (a) ci-dessus.
  var suppMissing = suppGames.filter(function (g) { return g.suppNeedsTie; });
  if (suppMissing.length > 0) {
    sheet.getRange(row, 1, 1, 13).merge();
    sheet.getRange(row, 1)
      .setValue('⚠ Partie(s) en manches supplémentaires SANS « Pointage régl. » (col. O) : ' +
                suppMissing.map(gameLabel).join(' ; ') + '. Comme toujours, ce tableau de pool ' +
                'inclut toutes les manches dans RD/RO ; mais le classement de bris d\'égalité ' +
                '(à droite) ne peut PAS encore exclure les manches supplémentaires (Note 4, ' +
                'Art. 42.11) tant que ce pointage manque. Saisissez le pointage réglementaire ' +
                '(nul) de ces parties pour l\'appliquer automatiquement.')
      .setWrap(true).setBackground(COLOR_INPUT).setVerticalAlignment('top');
    row++;
  }

  // Bloc « bris d'égalité » à droite (Étape A = tête-à-tête). Aligné sur le titre
  // de la section ; on retourne le max des deux hauteurs pour éviter tout
  // chevauchement vertical avec la section suivante.
  var orderedNames = standings.map(function (s) { return s.team; });
  var brisRow = writeTiebreakTable(sheet, startRow, orderedNames, poolGames,
                                   orderedNames, false, forced);

  return Math.max(row, brisRow);
}

/**
 * Écrit une section d'avancement (Étape B ou C). Retourne la prochaine ligne.
 * @param {number} basePosition  position de départ (1 pour Étape C, 4 pour Étape B)
 * @param {Object} forced        rangs forcés « Forcer rang » de cette portée (C ou B).
 */
function writeAdvancementSection(sheet, startRow, classe, title, orderedTeams,
                                 games, poolStatsByTeam, poolInfo, basePosition, forced) {
  forced = forced || {};
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

  // Map team -> pool d'origine, et équipes dont le 2e a été FORCÉ par le registraire
  // (Note 5 / forfaits) — seul l'Étape B (poolInfo = seconds) porte ce drapeau.
  var poolOf = {};
  var forcedTeam = {};
  poolInfo.forEach(function (info) {
    poolOf[info.team] = info.pool;
    if (info.forced) { forcedTeam[info.team] = true; }
  });

  // Détecte si Priorité 4 a été nécessaire.
  var needsManual = orderedTeams.__needsManualCheck === true;

  orderedTeams.forEach(function (team, idx) {
    // Recalcule les stats de cette équipe sur toutes ses parties de pool
    // (forfaits exclus des ratios pour B/C).
    var teamGames = games.filter(function (g) {
      return g.local === team || g.visiteur === team;
    });
    var st = computeTeamStats(team, teamGames, true);

    // Note : forçage admin du 2e (Note 5), Priorité 4 (à régler / résolue via « Forcer
    // rang ») et/ou marqueur manches supplémentaires (Note 4) — peuvent coexister.
    var noteParts = [];
    if (forcedTeam[team]) { noteParts.push('🔒 2e forcé par le registraire (Note 5)'); }
    if (needsManual) {
      noteParts.push('⚠ Vérif. manuelle (P4) — saisir « Forcer rang »');
    } else if (typeof forced[team] === 'number') {
      noteParts.push('🔒 Rang forcé (P4)');
    }
    if (teamGames.some(function (g) { return g.suppNeedsTie; })) {
      noteParts.push('⚠ Suppl. : pointage régl. manquant (col. O)');
    } else if (teamGames.some(gameIsSupp)) {
      noteParts.push('ℹ Note 4 appliquée (suppl.)');
    }
    var note = noteParts.join(' ');
    var rowData = [
      basePosition + idx, team, poolOf[team] || '',
      st.v, st.d,
      isFinite(st.raRatio) ? round3(st.raRatio) : '—',
      round3(st.rsRatio),
      st.rs, st.ra, note, ''
    ];
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);

    // Calcul détaillé au survol des cellules RD (col. 6) / RO (col. 7).
    sheet.getRange(row, 6).setNote(rdCalcNote(st));
    sheet.getRange(row, 7).setNote(roCalcNote(st));

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

  // Bloc « bris d'égalité » à droite (Étapes B/C = toutes les parties de pool de
  // chaque équipe). Aligné sur le titre ; max des deux hauteurs.
  var teamsArr = orderedTeams.slice();
  var brisRow = writeTiebreakTable(sheet, startRow, teamsArr, games, teamsArr, true, forced);

  return Math.max(row, brisRow);
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

/**
 * Bandeau affiché À LA PLACE des positions de demi-finale (Sections 4/5 + récap)
 * tant que toutes les parties de pool de la classe ne sont pas jouées. Les positions
 * 1-2-3-4 forment un classement inter-pools qui reste provisoire avant la fin des
 * pools ; on n'affiche donc qu'un rappel de la progression.
 *
 * @param {Sheet}  sheet
 * @param {number} startRow
 * @param {string} classe
 * @param {Object} completion  { total, played, complete } de poolPlayCompletion()
 * @return {number} ligne suivante
 */
function writeSeedingPendingBanner(sheet, startRow, classe, completion) {
  var row = startRow;
  sheet.getRange(row, 1).setValue(
    '⏳ Les positions de demi-finale (1-2-3-4) et la colonne « Avancement » ' +
    's\'afficheront lorsque TOUTES les parties des pools de la classe ' + classe +
    ' auront été jouées.  Parties jouées : ' + completion.played + ' / ' +
    completion.total + '.');
  sheet.getRange(row, 1, 1, 13).merge();
  sheet.getRange(row, 1).setFontWeight('bold').setBackground(COLOR_SECTION)
    .setHorizontalAlignment('center').setWrap(true);
  row++;
  return row;
}

// ============================================================================
//  GRAND LIVRE DES MATCHS (AUDIT)
// ============================================================================

// En-têtes du Grand livre : une ligne = une TRANSACTION (une équipe dans un
// match). 2 lignes par match (équipe locale + équipe visiteuse), 3 matchs par
// équipe, 24 équipes (2 classes x 3 pools x 4 équipes) → 72 lignes au total.
var LEDGER_HEADERS = [
  'Cl', 'Pl', '#', 'J', 'H', 'T',
  'Eq', 'Adv', 'Res', 'Sco', 'Loc/Vis',
  'Manches\ncomplètes', 'Retraits\nen\nfin', 'Manches\nprévues', 'Type\nde\nfin',
  'Pointage\nrégl.\n(suppl.)',
  'PC', 'Somme PC', 'MD', 'Somme MD', 'RD',
  'PP', 'Somme PP', 'MO', 'Somme MO', 'RO'
];

var LEDGER_HEADER_NOTES = {
  1:  'CL — abrégé de « Classe ». GRAND LIVRE — audit façon « compte ' +
      'bancaire » : chaque ligne est une TRANSACTION (une équipe dans un ' +
      'match, 2 lignes par match). Trié Classe → Pool → Équipe (les 3 ' +
      'matchs de chaque équipe regroupés) → Partie #. Les colonnes ' +
      '« Somme X » (EN GRAS) CUMULENT progressivement pour l\'équipe de la ' +
      'ligne (comme un solde qui augmente) ; le cumul final de chaque ' +
      'équipe (dernière ligne de son bloc) doit correspondre EXACTEMENT à ' +
      'PP/PC/MO/MD affichés dans le tableau de pool (feuille Classements). ' +
      'Équipe (vert) / Adversaire (rouge) indiquent la gagnante de la ' +
      'partie. Traits de séparation à 3 niveaux : trait moyen = fin ' +
      'd\'équipe, trait épais = fin de pool, double trait = fin de classe. ' +
      'Sert à retracer, en cas d\'écart, à quel match précis il apparaît.',
  2:  'PL — abrégé de « Pool ».',
  3:  '# — abrégé de « Partie # ».',
  4:  'J — abrégé de « Jour ».',
  5:  'H — abrégé de « Heure ».',
  6:  'T — abrégé de « Terrain ».',
  7:  'EQ — abrégé de « Équipe ». Équipe dont cette ligne retrace le cumul ' +
      '(vert si elle a gagné cette partie, rouge si elle l\'a perdue).',
  8:  'ADV — abrégé de « Adversaire ».',
  9:  'RES — abrégé de « Résultat ». Victoire / Défaite / Nul du point de ' +
      'vue de l\'équipe de cette ligne (colonne Eq).',
  10: 'SCO — abrégé de « Score ». Toujours « Eq-Adv » (l\'équipe de cette ' +
      'ligne en premier).',
  11: 'LOC/VIS — bleu pâle = Local, orange pâle = Visiteur. « Inconnu » (non ' +
      'coloré) si l\'équipe locale n\'a pas été précisée dans Résultats ' +
      '(colonne « Équipe Locale ») : dans ce cas les manches sont calculées ' +
      'de façon symétrique (aucune fraction ⅓/⅔), donc PP/PC/MO/MD restent ' +
      'exacts, mais on ne peut pas affirmer laquelle des deux équipes ' +
      'recevait réellement.',
  17: 'PC / MD / RD — PC et MD sont les valeurs RÉELLES de CETTE partie ' +
      'seulement (PC = points alloués ; MD = manches défensives, ' +
      'supplémentaires incluses). RD, lui, n\'est PAS cumulé : il est ' +
      'RECALCULÉ à chaque ligne = Somme PC ÷ Somme MD à ce point (même base ' +
      'réelle que le tableau de pool). Dégradé de couleur : vert (bas, ' +
      'meilleur) → rouge (haut, moins bon).',
  22: 'PP / MO / RO — mêmes principes que PC/MD/RD (colonne Q), côté ' +
      'offensif : PP et MO sont les valeurs réelles de CETTE partie ; RO est ' +
      'recalculé à chaque ligne = Somme PP ÷ Somme MO à ce point. Dégradé de ' +
      'couleur : rouge (bas, moins bon) → vert (haut, meilleur).'
};

/** Écrit les en-têtes, notes et largeurs de colonnes du Grand livre. */
function writeLedgerHeaders(sheet) {
  sheet.getRange(1, 1, 1, LEDGER_HEADERS.length).setValues([LEDGER_HEADERS]);
  styleHeader(sheet.getRange(1, 1, 1, LEDGER_HEADERS.length));
  applyHeaderNotes(sheet, 1, LEDGER_HEADER_NOTES);

  var widths = [32, 32, 40, 65, 55, 55, 170, 170, 70, 55, 80,
                70, 60, 70, 100, 80, 55, 80, 65, 80, 70,
                55, 80, 65, 80, 70];
  for (var c = 0; c < widths.length; c++) { sheet.setColumnWidth(c + 1, widths[c]); }
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(8);   // Cl..Adv restent visibles au défilement
  sheet.autoResizeRows(1, 1);  // ajuste la hauteur de l'en-tête aux libellés sur 2-3 lignes
}

/**
 * Structure vide du Grand livre (en-têtes seulement), pour rebuildSheets.
 * Rempli ensuite par buildLedgerSheet (menu, calculateStandings, ou live via
 * handleResultEdit/recalcStandingsOnly).
 */
function createLedgerSheet(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_LEDGER);
  sheet.clear();
  clearDataValidations(sheet);
  sheet.clearConditionalFormatRules();
  writeLedgerHeaders(sheet);
  sheet.getRange(2, 1).setValue(
    'Cliquez sur « Générer les matchs » puis « Mettre à jour les classements » ' +
    'pour remplir le grand livre.').setFontStyle('italic');
}

/**
 * Reconstruit ENTIÈREMENT la feuille Grand livre à partir des DEUX classes
 * (contrairement à buildStandingsSheet, appelée une fois par classe — le
 * grand livre les combine dans une seule feuille triée Classe → Pool →
 * Équipe → Partie #). Base de calcul RÉELLE (PP/PC/MO/MD réels, suppl.
 * incluses), identique à celle du tableau de pool (writePoolSection), pour
 * que le cumul final de chaque équipe corresponde aux classements officiels.
 */
function buildLedgerSheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, SHEET_LEDGER);
  sheet.clear();
  clearDataValidations(sheet);
  writeLedgerHeaders(sheet);

  var gamesByClasse = { A: getGameResults('A'), B: getGameResults('B') };
  var rows = buildLedgerRows(gamesByClasse);

  if (rows.length === 0) {
    sheet.clearConditionalFormatRules();
    sheet.getRange(2, 1).setValue('Aucun résultat saisi pour l\'instant.').setFontStyle('italic');
    return;
  }

  var values = rows.map(function (r) { return r.values; });
  sheet.getRange(2, 1, values.length, LEDGER_HEADERS.length).setValues(values);

  // Toute la zone est CALCULÉE (rien à saisir ici) : fond gris uniforme, avec
  // des bandes alternées par bloc équipe pour la lisibilité (3 lignes/équipe).
  var nRows = values.length;
  var fullRange = sheet.getRange(2, 1, nRows, LEDGER_HEADERS.length);
  fullRange.setBackground(COLOR_CALC);
  fullRange.setFontWeight('normal');

  // Alignement uniforme : centré partout, sauf Eq/Adv (colonnes 7-8, noms
  // d'équipes) laissées à gauche pour rester faciles à lire.
  fullRange.setHorizontalAlignment('center');
  sheet.getRange(2, 7, nRows, 2).setHorizontalAlignment('left');

  // Colonnes « Somme » (cumul) : seules colonnes en gras, pour les distinguer
  // d'un coup d'œil des valeurs PC/MD/PP/MO « de cette partie » à côté.
  [18, 20, 23, 25].forEach(function (col) {
    sheet.getRange(2, col, nRows, 1).setFontWeight('bold');
  });

  // RD et RO (colonnes 21 et 26) : encore plus mises en évidence (gras +
  // taille de police augmentée, l'API Sheets n'offrant pas de graisse au-delà
  // du gras) — ce sont les ratios décisifs des bris d'égalité.
  [21, 26].forEach(function (col) {
    var range = sheet.getRange(2, col, nRows, 1);
    range.setFontWeight('bold');
    range.setFontSize(12);
  });

  var band = false;
  var prevKey = null;
  rows.forEach(function (r, i) {
    var sheetRow = i + 2;
    if (r.key !== prevKey) { band = !band; prevKey = r.key; }
    if (band) {
      sheet.getRange(sheetRow, 1, 1, LEDGER_HEADERS.length).setBackground('#f5f5f5');
    }

    // Équipe (col 7) / Adversaire (col 8) : gagnante en vert, perdante en rouge.
    if (r.resultat === 'Victoire') {
      sheet.getRange(sheetRow, 7).setBackground(COLOR_WIN);
      sheet.getRange(sheetRow, 8).setBackground(COLOR_LOSS);
    } else if (r.resultat === 'Défaite') {
      sheet.getRange(sheetRow, 7).setBackground(COLOR_LOSS);
      sheet.getRange(sheetRow, 8).setBackground(COLOR_WIN);
    }

    // Loc/Vis (col 11) : bleu pâle si locale, orange pâle si visiteuse, neutre
    // (fond gris/bande déjà appliqué) si inconnue.
    if (r.locVis === 'Local') {
      sheet.getRange(sheetRow, 11).setBackground(COLOR_LOCAL);
    } else if (r.locVis === 'Visiteur') {
      sheet.getRange(sheetRow, 11).setBackground(COLOR_VISITOR);
    }

    // Bordure sous la dernière ligne de chaque bloc (équipe / pool / classe),
    // pour repérer chaque niveau de regroupement sans recourir au gras
    // (réservé aux colonnes Somme ci-dessus). Hiérarchie croissante :
    // équipe (trait moyen) < pool (trait épais) < classe (double trait).
    // Un bloc classe/pool est toujours aussi une fin de bloc équipe (tri
    // Classe → Pool → Équipe → Partie #), donc un seul niveau, le plus
    // englobant, suffit à choisir.
    var isLast = (i === rows.length - 1);
    var next = isLast ? null : rows[i + 1];
    var isLastOfTeam   = isLast || next.key !== r.key;
    var isLastOfPool   = isLast || next.classe !== r.classe || next.pool !== r.pool;
    var isLastOfClasse = isLast || next.classe !== r.classe;

    var rowRange = sheet.getRange(sheetRow, 1, 1, LEDGER_HEADERS.length);
    if (isLastOfClasse) {
      rowRange.setBorder(null, null, true, null, null, null,
                         '#000000', SpreadsheetApp.BorderStyle.DOUBLE);
    } else if (isLastOfPool) {
      rowRange.setBorder(null, null, true, null, null, null,
                         '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    } else if (isLastOfTeam) {
      rowRange.setBorder(null, null, true, null, null, null,
                         '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  });

  // Dégradé de couleur sur RD (col 21) et RO (col 26) : repère visuel des
  // équipes en difficulté défensive/offensive. RD (plus bas = meilleur) va du
  // vert (bas) au rouge (haut) ; RO (plus haut = meilleur) va du rouge (bas)
  // au vert (haut).
  sheet.clearConditionalFormatRules();
  var rdRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([sheet.getRange(2, 21, nRows, 1)])
    .setGradientMinpoint(COLOR_WIN)
    .setGradientMidpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(COLOR_LOSS)
    .build();
  var roRule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([sheet.getRange(2, 26, nRows, 1)])
    .setGradientMinpoint(COLOR_LOSS)
    .setGradientMidpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpoint(COLOR_WIN)
    .build();
  sheet.setConditionalFormatRules([rdRule, roRule]);
}

/**
 * Construit les lignes de transactions triées Classe → Pool → Équipe →
 * Partie #, avec cumul PP/PC/MO/MD par équipe calculé ligne par ligne (pas
 * via computeTeamStats, qui ne donne que le total final).
 * @return {Array<{key:string, values:Array}>}
 */
function buildLedgerRows(gamesByClasse) {
  // 1) Aplatit chaque partie en 2 transactions (équipe locale / visiteuse).
  var tx = [];
  CLASSES.forEach(function (classe) {
    (gamesByClasse[classe] || []).forEach(function (g) {
      tx.push(makeLedgerTx(classe, g, true));    // perspective équipe locale
      tx.push(makeLedgerTx(classe, g, false));   // perspective équipe visiteuse
    });
  });

  // 2) Tri Classe → Pool → Équipe (alpha) → Partie # (numérique, repli texte).
  tx.sort(function (a, b) {
    if (a.classe !== b.classe) { return a.classe.localeCompare(b.classe); }
    if (a.pool !== b.pool) { return a.pool - b.pool; }
    if (a.equipe !== b.equipe) { return a.equipe.localeCompare(b.equipe, 'fr'); }
    var pa = parseInt(a.partie, 10), pb = parseInt(b.partie, 10);
    if (isNaN(pa) || isNaN(pb)) { return String(a.partie).localeCompare(String(b.partie)); }
    return pa - pb;
  });

  // 3) Accumulation ligne par ligne par équipe (clé classe|pool|équipe) : les
  //    lignes d'une même équipe étant déjà contiguës après le tri, un simple
  //    reset au changement de clé suffit.
  var rows = [];
  var cum = null;
  var curKey = null;
  tx.forEach(function (t) {
    var key = t.classe + '|' + t.pool + '|' + t.equipe;
    if (key !== curKey) { cum = { pp: 0, pc: 0, mo: 0, md: 0 }; curKey = key; }

    cum.pp += t.ppRow;
    cum.pc += t.pcRow;
    cum.mo += t.moRow;
    cum.md += t.mdRow;

    var rd = cum.md > 0 ? (cum.pc / cum.md).toFixed(3) : '—';
    var ro = cum.mo > 0 ? (cum.pp / cum.mo).toFixed(3) : '—';

    rows.push({
      key: key,
      classe: t.classe,
      pool: t.pool,
      resultat: t.resultat,
      locVis: t.locVis,
      values: [
        t.classe, t.pool, t.partie, t.jour, t.heure, t.terrain,
        t.equipe, t.adversaire, t.resultat, t.score, t.locVis,
        t.manches, t.retraits, t.manchesPrevues, t.type, t.pointageRegl,
        t.pcRow, cum.pc, formatFraction(t.mdRow), formatFraction(cum.md), rd,
        t.ppRow, cum.pp, formatFraction(t.moRow), formatFraction(cum.mo), ro
      ]
    });
  });
  return rows;
}

/**
 * Construit UNE transaction (une équipe pour un match donné).
 * @param {string}  classe
 * @param {Object}  g          partie (retour de getGameResults)
 * @param {boolean} isLocal    perspective équipe locale (true) ou visiteuse (false)
 */
function makeLedgerTx(classe, g, isLocal) {
  var equipe     = isLocal ? g.local : g.visiteur;
  var adversaire = isLocal ? g.visiteur : g.local;
  var ppRow = isLocal ? g.scoreLocal : g.scoreVisiteur;   // PP réel de CETTE partie
  var pcRow = isLocal ? g.scoreVisiteur : g.scoreLocal;   // PC réel de CETTE partie
  var moRow = isLocal ? g.offLocal : g.offVisiteur;       // MO réel (suppl. incluses)
  var mdRow = isLocal ? g.defLocal : g.defVisiteur;       // MD réel (suppl. incluses)

  var resultat = (g.winner === equipe) ? 'Victoire'
               : (g.winner === '' ? 'Nul' : 'Défaite');

  var locVis = (g.homeKnown === false) ? 'Inconnu' : (isLocal ? 'Local' : 'Visiteur');

  // « Pointage régl. (suppl.) » : identique pour les 2 équipes (score nul
  // réglementaire X), uniquement si Type de fin = Supplémentaires ET résolu
  // (Note 4, Art. 42.11).
  var pointageRegl = (g.type === 'Supplémentaires' && !g.suppNeedsTie) ? g.regRsLocal : '';

  return {
    classe: classe, pool: g.pool, partie: g.partie,
    jour: g.jour, heure: g.heure, terrain: g.terrain,
    equipe: equipe, adversaire: adversaire,
    resultat: resultat,
    score: ppRow + '-' + pcRow,
    locVis: locVis,
    manches: g.manches, retraits: g.retraits, manchesPrevues: g.manchesPrevues,
    type: g.type, pointageRegl: pointageRegl,
    ppRow: ppRow, pcRow: pcRow, moRow: moRow, mdRow: mdRow
  };
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
    // Efface H, I (scores), J (Équipe Locale), K, L (manches/retraits),
    // O (Pointage régl. suppl.), P..T (calculs).
    sheet.getRange(2, 8, n, 2).clearContent();    // H, I
    sheet.getRange(2, 10, n, 1).clearContent();   // J : Équipe Locale
    sheet.getRange(2, 11, n, 2).clearContent();   // K, L
    sheet.getRange(2, 15, n, 1).clearContent();   // O : Pointage régl. (suppl.)
    sheet.getRange(2, 16, n, 5).clearContent();   // P..T : calculs
    // Réinitialise Manches prévues à 6 et Type de fin à "Normal".
    var regCol = [];
    var normalCol = [];
    for (var i = 0; i < n; i++) { regCol.push([6]); normalCol.push(['Normal']); }
    sheet.getRange(2, 13, n, 1).setValues(regCol);     // M : Manches prévues
    sheet.getRange(2, 14, n, 1).setValues(normalCol);  // N : Type de fin
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
  // (6e élément optionnel = Pointage régl. (suppl.), uniquement pour les parties
  //  "Supplémentaires" : pointage nul atteint à la fin du réglementaire.)
  // (Équipe 1 = équipe locale pour ces données fictives, écrite ci-dessous dans J.)
  // Les 18 lignes sont appliquées par position : Pool1 M1-M6, Pool2 M1-M6, Pool3 M1-M6
  // (correspond à l'ordre des matchs tel qu'écrit par "Générer les matchs").
  // "Manches prévues" (col M) est forcée à 6 à l'écriture (toutes ces parties fictives
  // sont prévues sur 6 manches) : les lignes Mercy (K=4/5) testent ainsi le crédit
  // défensif réglementaire de 6 manches au gagnant.

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
      [ 3,  4, 7, 0, 'Supplémentaires', 3], // M4 T1 vs T2 → T2 gagne (supp. 7e ; nul 3-3 après 6 → Note 4)
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
      [ 5,  4, 7, 1, 'Supplémentaires', 4], // M2 T0 vs T2 → T0 gagne (walk-off supp. 7e ; nul 4-4 après 6 → Note 4)
      [ 1,  2, 8, 0, 'Supplémentaires', 1], // M3 T0 vs T3 → T3 gagne (supp. 8e ; nul 1-1 après 6 → Note 4)
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

    // Écrit les scores (H, I) et paramètres de fin (K, L, M, N, O) en deux blocs.
    // Manches prévues (M) = 6 pour toutes les données fictives ; Pointage régl. (O)
    // seulement pour les parties "Supplémentaires" (6e élément des données), vide sinon.
    var scores = rows.map(function (r) { return [r[0], r[1]]; });
    var params = rows.map(function (r) {
      var tie = (r[4] === 'Supplémentaires' && r[5] != null) ? r[5] : '';
      return [r[2], r[3], 6, r[4], tie];
    });
    sheet.getRange(2, 8, n, 2).setValues(scores);   // H, I
    sheet.getRange(2, 11, n, 5).setValues(params);  // K, L, M, N, O

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

/**
 * Exporte les deux feuilles « Résultats » et les deux feuilles « Classements »
 * en TSV (4 fichiers) regroupés dans un seul ZIP, déposé dans le Google Drive
 * du propriétaire. Une boîte de dialogue affiche le lien vers le fichier.
 *
 * Les valeurs exportées sont les valeurs AFFICHÉES (getDisplayValues) — donc
 * les fractions de manches, scores, etc. tels que vus à l'écran. Les feuilles
 * absentes sont signalées dans le message final sans bloquer l'export.
 */
function exportSheetsToZip() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var names = [
    SHEET_RESULTS['A'], SHEET_RESULTS['B'],
    SHEET_STANDINGS['A'], SHEET_STANDINGS['B'],
    SHEET_LEDGER
  ];

  var blobs = [];
  var missing = [];
  names.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) { missing.push(name); return; }
    var fileName = name.replace(/[\/\\:*?"<>|]/g, '_') + '.tsv';
    // BOM UTF-8 en tête pour qu'Excel détecte correctement les accents.
    blobs.push(Utilities.newBlob('﻿' + sheetToTsv(sheet),
                                 'text/tab-separated-values', fileName));
  });

  if (blobs.length === 0) {
    ui.alert('Export TSV',
             'Aucune des feuilles à exporter n\'existe encore.\n\n' +
             'Lancez d\'abord « Initialiser les feuilles » puis « Générer les matchs ».',
             ui.ButtonSet.OK);
    return;
  }

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(),
                                   'yyyy-MM-dd_HH-mm');
  var zipName = 'TournoiBaseball_export_' + stamp + '.zip';
  var zipBlob = Utilities.zip(blobs, zipName);
  var file = DriveApp.createFile(zipBlob);

  var msg = 'ZIP créé dans votre Google Drive :\n\n' + file.getName() +
            '\n\nLien : ' + file.getUrl();
  if (missing.length) {
    msg += '\n\n⚠ Feuilles introuvables (non exportées) : ' + missing.join(', ');
  }
  ui.alert('Export TSV', msg, ui.ButtonSet.OK);
}

/**
 * Convertit toute la plage de données d'une feuille en TSV à partir des valeurs
 * affichées. Les tabulations et sauts de ligne contenus dans une cellule sont
 * remplacés par une espace pour ne pas corrompre la structure du TSV.
 */
function sheetToTsv(sheet) {
  var values = sheet.getDataRange().getDisplayValues();
  return values.map(function (row) {
    return row.map(function (cell) {
      return String(cell).replace(/[\t\r\n]+/g, ' ');
    }).join('\t');
  }).join('\n');
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
    SHEET_LEDGER
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

// ============================================================================
//  AFFICHAGE PUBLIC — WEB APP (publication Facebook)
// ============================================================================
//
// Page web publiée (doGet) destinée aux responsables des communications : ils
// l'ouvrent sur leur téléphone pendant le tournoi pour screenshoter proprement
// les classements (ou partager le lien) sur la page Facebook du tournoi.
//
// Données : computeStandingsModel(ss, classe) — MÊME moteur de calcul que les
// feuilles Classements (aucun recalcul parallèle), forçages admin inclus. La page
// est épurée : on n'affiche QUE Rang / Équipe / V-D / RD + les qualifiés en demi-
// finale, plus les bandeaux « ℹ ». Tout l'admin (Avancement, Forcer 2e, Notes,
// Critère décisif, bris d'égalité) est volontairement masqué.
//
// DÉPLOIEMENT (distinct du simple collage de code) — à refaire après chaque
// changement du code de la page :
//   Apps Script → Déployer → Nouveau déploiement → Application Web
//     • « Exécuter en tant que » : moi (propriétaire) — pour lire les feuilles
//     • « Qui a accès » : tous les utilisateurs disposant du lien (lecture seule)
//   Copier l'URL /exec → c'est le lien à donner aux communications.
//   (Menu « 📱 Lien affichage public » → affiche cette URL.)

/**
 * Point d'entrée de la web app. Calcule les modèles des classes A et B et retourne
 * la page HTML d'affichage public.
 */
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Parties des deux classes mélangées, triées par numéro de partie (vue « Résultats »).
  var matches = getMatchRows('A').concat(getMatchRows('B'));
  matches.sort(function (a, b) {
    var na = parseInt(a.partie, 10), nb = parseInt(b.partie, 10);
    if (isNaN(na) || isNaN(nb)) { return String(a.partie).localeCompare(String(b.partie)); }
    return na - nb;
  });
  var data = {
    A: computeStandingsModel(ss, 'A'),
    B: computeStandingsModel(ss, 'B'),
    matches: matches,
    version: APP_VERSION
  };
  return HtmlService.createHtmlOutput(renderPublicHtml_(data))
    .setTitle('Tournoi provincial de baseball 13U 2026 de Rimouski')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Affiche, dans une alerte, l'URL de la web app déployée (à partager aux
 * communications), ou un rappel des étapes de déploiement si elle ne l'est pas.
 */
function showPublicUrl() {
  var ui = SpreadsheetApp.getUi();
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (err) { url = ''; }
  if (url) {
    ui.alert('📱 Affichage public',
      'Lien à donner aux responsables des communications.\n' +
      'À ouvrir sur un téléphone pour publier les classements sur Facebook :\n\n' + url,
      ui.ButtonSet.OK);
  } else {
    ui.alert('📱 Affichage public — à déployer',
      "La web app n'est pas encore déployée.\n\n" +
      "Dans l'éditeur Apps Script :\n" +
      "Déployer → Nouveau déploiement → Application Web\n" +
      "  • « Exécuter en tant que » : moi\n" +
      "  • « Qui a accès » : tous les utilisateurs disposant du lien\n\n" +
      "Copiez l'URL /exec, puis relancez ce menu pour la retrouver.",
      ui.ButtonSet.OK);
  }
}

/**
 * Construit la page HTML complète (CSS + JS + données embarquées en JSON).
 * Le JSON est injecté à la place du repère /*__DATA__*\/ ; tous les « < » du JSON
 * sont échappés en < pour qu'aucun nom d'équipe ne puisse fermer la balise script.
 * @param {Object} data  { A: modèleA, B: modèleB }
 * @return {string} HTML
 */
function renderPublicHtml_(data) {
  var json = JSON.stringify(data).replace(/</g, '\\u003c');
  // Forme fonctionnelle pour le bloc règles : neutralise les motifs « $ » de
  // String.replace (le texte des règles contient « 100 $ »).
  return PUBLIC_HTML_TEMPLATE_
    .replace('/*__DATA__*/', 'window.DATA = ' + json + ';')
    .replace('/*__RULES__*/', function(){ return RULES_HTML_; });
}

// Gabarit de la page publique. Le code client n'utilise NI backticks NI « ${ } »
// pour ne pas interférer avec ce template literal serveur ; les données sont
// injectées via le repère /*__DATA__*/ par renderPublicHtml_.
// Contenu statique des règles (13U classes A et B), généré depuis
// Regles13U_2026.md (revalidé vs ReglesRegie2026.pdf). Injecté dans la page via
// le repère /*__RULES__*/ par renderPublicHtml_. N’utilise NI backtick NI « ${ } »
// afin de ne pas interférer avec les template literals serveur.
var RULES_HTML_ = `
<div class="rules-title">Règles de Régie 2026 — Division 13U (classes A et B)</div>
<div class="rules-sub">Baseball Québec | Référence superviseur de tournoi</div>
<section>
<h2>1. Terrain et équipement</h2>
<h3>Dimensions du terrain (Art. 104.3)</h3>
<div class="rwrap"><table><thead><tr><th>Élément</th><th>Mesure</th></tr></thead><tbody><tr><td>Distance entre les buts</td><td>70 pi (21,34 m)</td></tr><tr><td>Marbre → plaque du lanceur</td><td>48 pi (14,63 m)</td></tr><tr><td>Clôture extérieure (coin)</td><td>210 pi (64 m) idéalement</td></tr><tr><td>Clôture extérieure (centre)</td><td>240 pi (73 m) idéalement</td></tr></tbody></table></div>
<div class="callout"><p><strong>Note :</strong> Le monticule est <strong>facultatif</strong> en 13U (Art. 107.3).</p></div>
<h3>Équipement (Art. 103.x)</h3>
<div class="rwrap"><table><thead><tr><th>Élément</th><th>13U B</th><th>13U A</th></tr></thead><tbody><tr><td>Crampons métalliques</td><td>❌ Non</td><td>❌ Non</td></tr><tr><td>Bâton — limite DROP</td><td>Aucune limite</td><td>Aucune limite</td></tr><tr><td>Certification bâton requise</td><td>USSSA, USA ou BBCOR</td><td>USSSA, USA ou BBCOR</td></tr><tr><td>Balle</td><td>9 pouces de circonférence</td><td>9 po</td></tr><tr><td>Casque</td><td>2 oreilles obligatoire</td><td>2 oreilles</td></tr><tr><td>Coussin double 1er but</td><td>✅ Oui</td><td>✅ Oui</td></tr><tr><td>Rectangle du frappeur</td><td>✅ Oui</td><td>✅ Oui</td></tr></tbody></table></div>
<div class="callout"><p><strong>Bâton illégal :</strong> Frappeur retiré pour action irrégulière. Aucun joueur ne peut avancer, mais les retraits effectués sont maintenus.</p></div>
<div class="callout"><p><strong>Entraîneur ≤ 18 ans sur les buts :</strong> doit porter un casque à 2 oreilles (Art. 103.1c).</p></div>
</section>
<section>
<h2>2. Durée des parties</h2>
<h3>Saison régulière (Art. 107.1)</h3>
<ul><li><strong>6 manches</strong> à finir (si conditions le permettent)</li><li><strong>Partie réglementaire :</strong> 3 ½ manches si l'équipe locale est en avance</li><li><strong>Forfait :</strong> 15 minutes après l'heure prévue sans le nombre minimum de joueurs (9)</li></ul>
<h3>Compétitions provinciales (Art. 42.3)</h3>
<div class="rwrap"><table><thead><tr><th>Classe</th><th>Limite de temps</th></tr></thead><tbody><tr><td>13U A</td><td>Toute manche débutant après <strong>80 minutes (1h20)</strong> est déclarée dernière manche (ouverte)</td></tr><tr><td>13U B</td><td>Toute manche débutant après <strong>70 minutes (1h10)</strong> est déclarée dernière manche (ouverte)</td></tr></tbody></table></div>
<ul><li>La durée est mesurée à partir du « Play Ball » et se termine à la fin de la manche offensive locale (3e retrait ou limite de points atteinte)</li><li>Tout arrêt de la partie de <strong>plus de 10 minutes</strong> n'est <strong>pas</strong> comptabilisé dans la durée</li><li>En cas d'<strong>égalité au terme du temps</strong>, la partie se poursuit — toutes les manches subséquentes sont <strong>ouvertes</strong> (sans limite de points)</li></ul>
<h3>Manches supplémentaires (Art. 42.4)</h3>
<p>Si égalité au terme des manches réglementaires ou du temps alloué :</p>
<ul><li>Chaque équipe débute la manche supplémentaire avec <strong>un coureur au 1er but et un coureur au 2e but, aucun retrait</strong></li><li>L'ordre : si le 6e frappeur était le dernier de la manche précédente → 7e frappeur au bâton, 5e frappeur au 2e but, 6e frappeur au 1er but</li><li>S'applique pour <strong>toutes les parties</strong>, incluant les médailles</li></ul>
</section>
<section>
<h2>3. Règles de jeu — par classe</h2>
<h3>3.1 Coureurs sur les buts</h3>
<div class="rwrap"><table><thead><tr><th>Situation</th><th>13U B</th><th>13U A</th></tr></thead><tbody><tr><td>Quitter son but</td><td>Seulement quand la balle traverse le marbre ou est frappée</td><td>Librement (règles BCan)</td></tr><tr><td>Vol de but</td><td>✅ Permis</td><td>✅ Permis</td></tr><tr><td>Vol du marbre</td><td>❌ Non permis</td><td>✅ Permis</td></tr><tr><td>Atteindre le marbre</td><td>Seulement si balle frappée ou forcé</td><td>Règles BCan</td></tr><tr><td>Chandelle intérieure</td><td>❌ Ne s'applique pas</td><td>✅ Oui</td></tr><tr><td>Amorti (coup retenu)</td><td>❌ Non (balle morte + prise)</td><td>✅ Permis</td></tr></tbody></table></div>
<div class="callout"><p><strong>13U B :</strong> Lorsque le lanceur reçoit la balle et prend place sur sa plaque et que le receveur est en position, les coureurs doivent retourner à leurs buts.</p></div>
<h3>3.2 Frappeur — 3e prise échappée</h3>
<div class="rwrap"><table><thead><tr><th>Classe</th><th>Règle</th></tr></thead><tbody><tr><td>13U B</td><td>Frappeur retiré automatiquement — balle au jeu — coureurs peuvent avancer <strong>sauf au marbre</strong> et lors d'écarts de pointage</td></tr><tr><td>13U A</td><td>Règle officielle de baseball (BCan) s'applique</td></tr></tbody></table></div>
<h3>3.3 Feinte irrégulière</h3>
<div class="rwrap"><table><thead><tr><th>Classe</th><th>Règle</th></tr></thead><tbody><tr><td>13U B</td><td>Aucune feinte irrégulière appelée — balle morte, coureurs retournent à leur but</td></tr><tr><td>13U A</td><td>Règle officielle de baseball (BCan) s'applique</td></tr></tbody></table></div>
</section>
<section>
<h2>4. Équité de jeu</h2>
<h3>4.1 Rôle offensif (Art. 103.14)</h3>
<p>S'applique aux <strong>deux classes</strong> (B, A) :</p>
<ul><li>Tous les joueurs sont inscrits à l'ordre des frappeurs et frappent à leur tour</li><li>Joueur arrivant en cours de partie → inscrit au <strong>dernier rang</strong></li><li>Joueur quittant en cours de partie → retiré de l'ordre (<strong>aucun retrait automatique</strong>)</li><li>Si un joueur ne peut compléter sa présence au bâton : remplacé par le frappeur suivant <strong>avec le même compte</strong></li><li>Si un coureur ne peut compléter sa présence sur les buts : remplacé par le <strong>dernier retrait</strong> (peut avancer seulement si balle frappée ou forcé)</li><li>Si un joueur <strong>saute son tour</strong> au bâton mais demeure dans la partie : <strong>1 retrait comptabilisé</strong></li></ul>
<h4>Points maximum par manche</h4>
<div class="rwrap"><table><thead><tr><th>Classe</th><th>Limite</th></tr></thead><tbody><tr><td>13U B</td><td><strong>3 points</strong> maximum par manche</td></tr><tr><td>13U A</td><td><strong>5 points</strong> maximum par manche</td></tr></tbody></table></div>
<div class="callout"><p><strong>Circuit par-dessus les clôtures :</strong> L'équipe est créditée de <strong>tous ses points</strong>, même si la limite est dépassée.</p></div>
<h4>Manche ouverte (saison régulière seulement — 13U B et A)</h4>
<ul><li>Lors d'une manche ouverte, l'équipe <strong>visiteuse</strong> peut prendre un maximum de <strong>10 points d'avance</strong></li><li>Une fois cet écart atteint, la demi-manche visiteuse prend fin</li><li>L'équipe receveuse est alors limitée à tenter de faire une <strong>partie nulle</strong></li></ul>
<h3>4.2 Rôle défensif (Art. 103.14b — 13U B et A)</h3>
<ul><li>Les joueurs sur le banc à la <strong>1re manche</strong> se joignent sur base <strong>rotative</strong> à compter de la 2e manche</li><li>Un joueur <strong>ne peut passer 2 manches consécutives</strong> sur le banc</li><li>Ne peut retourner au banc avant que <strong>tous les autres joueurs aient passé une manche complète</strong> sur le banc</li><li><strong>Exceptions :</strong> lanceur actuellement au monticule, blessure, balle à la tête</li><li>L'entraîneur-chef <strong>doit aviser le marqueur</strong> de l'identité des joueurs sur le banc à chaque manche</li></ul>
<p><strong>Sanctions :</strong></p>
<ol><li>1re intervention : temps d'arrêt, avertissement à l'entraîneur-chef</li><li>2e intervention : expulsion de l'entraîneur-chef (suspension automatique)</li></ol>
<h3>4.3 Avance sur les buts avec écart de 5 points (Art. 103.14c)</h3>
<div class="rwrap"><table><thead><tr><th>Classe</th><th>Application</th></tr></thead><tbody><tr><td>13U B</td><td>❌ Non applicable</td></tr><tr><td>13U A</td><td>✅ Sous conditions</td></tr></tbody></table></div>
<p><strong>Lorsqu'une équipe a 5 points d'avance ou plus :</strong></p>
<ul><li>Le coureur de cette équipe peut atteindre le but suivant <strong>seulement si la balle est frappée ou s'il est forcé d'avancer</strong></li><li><strong>Sans avance sur les buts :</strong> infraction → balle morte, coureur retourné à son but</li><li><strong>Avec avance sur les buts :</strong> infraction → balle morte, coureur retiré (si plusieurs fautifs : seul le plus près du marbre est retiré)</li></ul>
<h3>4.4 Retour au jeu (Art. 103.14d)</h3>
<div class="rwrap"><table><thead><tr><th>Classe</th><th>Règle</th></tr></thead><tbody><tr><td>13U B</td><td>✅ Tout joueur peut revenir en tout temps à n'importe quelle position, <strong>sauf lanceur s'il a déjà lancé</strong></td></tr><tr><td>13U A</td><td>✅ Idem</td></tr></tbody></table></div>
</section>
<section>
<h2>5. Règlement du lanceur</h2>
<h3>5.1 Visites au monticule (Art. 107.2)</h3>
<ul><li>2e visite dans la <strong>même manche</strong> : le lanceur doit être retiré (peut jouer à une autre position)</li><li>Un lanceur retiré du monticule <strong>ne peut revenir</strong> à titre de lanceur dans cette partie</li></ul>
<h3>5.2 Buts sur balle intentionnels</h3>
<div class="rwrap"><table><thead><tr><th>Classe</th><th>Règle</th></tr></thead><tbody><tr><td>13U B</td><td>❌ Interdits</td></tr><tr><td>13U A</td><td>✅ Permis — l'entraîneur avise l'arbitre, balle morte, frappeur au 1er but automatiquement</td></tr></tbody></table></div>
<h3>5.3 Saison régulière — manches lancées</h3>
<div class="rwrap"><table><thead><tr><th>Période</th><th>Maximum / jour</th><th>Maximum / 7 jours</th></tr></thead><tbody><tr><td>Mai – Juin</td><td>2 manches</td><td>4 manches</td></tr><tr><td>Juillet – Septembre</td><td>3 manches</td><td>6 manches</td></tr></tbody></table></div>
<div class="callout"><ul><li>Interdit de lancer <strong>3 journées consécutives</strong></li><li>Une journée de repos = <strong>journée complète de calendrier</strong></li><li>Une balle lancée dans une manche compte pour une manche lancée</li><li>Les manches en compétition provinciale <strong>ne sont pas comptabilisées</strong> dans le maximum de 7 jours</li></ul></div>
<h3>5.4 Compétition provinciale — lancers et repos</h3>
<div class="rwrap"><table><thead><tr><th>Lancers dans la journée</th><th>Repos requis</th></tr></thead><tbody><tr><td>1 – 40 lancers</td><td>Aucun repos</td></tr><tr><td>41 – 55 lancers</td><td><strong>1 journée</strong> de repos</td></tr><tr><td>56 – 70 lancers</td><td><strong>2 journées</strong> de repos</td></tr><tr><td>71 – 85 lancers</td><td><strong>3 journées</strong> de repos</td></tr></tbody></table></div>
<div class="callout"><p><strong>Maximum journalier : 85 lancers</strong></p></div>
<p><strong>Règles spéciales en compétition provinciale :</strong></p>
<ul><li>Un lanceur peut lancer une 2e journée consécutive <strong>s'il a effectué moins de 41 lancers</strong> la journée précédente</li><li>Le cumulatif maximal sur 2 jours est de <strong>85 lancers</strong></li><li><strong>Interdit de lancer 3 journées consécutives</strong>, peu importe le nombre de lancers</li><li>Un entraîneur peut effectuer une <strong>2e visite</strong> au lanceur pour le retirer, sur le même frappeur dans la même manche (compétition provinciale seulement)</li><li>Un registre officiel des lancers est tenu par une personne assignée (source officielle)</li></ul>
<p><strong>Règle du frappeur en cours :</strong> Un lanceur ayant atteint une étape (40, 55 ou 70 lancers) peut <strong>terminer sa présence face au frappeur courant</strong> sans pénalité additionnelle s'il n'affronte pas d'autres frappeurs. On indique alors l'étape franchie au registre.</p>
<p><strong>Pénalité :</strong> Un lanceur inéligible doit être remplacé immédiatement. En compétition provinciale, l'entraîneur-chef <strong>et</strong> le lanceur sont expulsés (suspension automatique). Toutes les actions effectuées par le lanceur demeurent valides.</p>
</section>
<section>
<h2>6. Différence de pointage</h2>
<h3>Saison régulière (Art. 103.13)</h3>
<p>La règle de <strong>10 points d'écart</strong> s'applique à toutes les classes 13U :</p>
<ul><li>Après la <strong>4e manche</strong> complète</li><li>Ou après <strong>3 manches et demie</strong> si l'équipe locale est en avance</li></ul>
<h3>Compétitions provinciales (Art. 42.3d) — règles supplémentaires</h3>
<p>En tournoi/championnat, <strong>deux</strong> règles d'écart s'appliquent en séquence :</p>
<div class="rwrap"><table><thead><tr><th>Écart</th><th>Après la manche</th></tr></thead><tbody><tr><td><strong>15 points</strong></td><td>3e manche complète (ou 2½ si locale en avance)</td></tr><tr><td><strong>10 points</strong></td><td>4e manche complète (ou 3½ si locale en avance)</td></tr></tbody></table></div>
<div class="callout warn"><p>⚠️ <strong>Règle additionnelle en compétition :</strong> La partie se termine <strong>immédiatement</strong> à partir du moment où une équipe, en raison de la <strong>limitation des points par manche</strong>, ne peut mathématiquement plus remporter la victoire.</p></div>
</section>
<section>
<h2>7. Règles spécifiques aux compétitions provinciales</h2>
<h3>Forfait (Art. 42.5)</h3>
<ul><li>Chaque équipe doit se présenter au terrain <strong>au moins 1 heure avant</strong> la partie</li><li>Si une équipe n'aligne pas 9 joueurs dans les <strong>15 minutes</strong> suivant l'heure prévue → <strong>forfait</strong></li><li><strong>Sanction additionnelle au forfait :</strong> tous les lanceurs de la <strong>partie suivante</strong> se voient créditer au registre le nombre de lancers correspondant à la <strong>première étape</strong> de leur division (40 lancers pour 13U), à la date précédant celle de la partie suivante</li></ul>
<h3>Réservistes en compétition (Art. 48.1)</h3>
<div class="rwrap"><table><thead><tr><th>Classe</th><th>Règle</th></tr></thead><tbody><tr><td>13U A</td><td>Un réserviste <strong>ne peut pas</strong> évoluer à la position de lanceur</td></tr><tr><td>13U B</td><td>Un réserviste <strong>ne peut pas</strong> évoluer à la position de lanceur — et <strong>ne peut pas</strong> utiliser un réserviste de classe A</td></tr></tbody></table></div>
<h3>Équité de jeu — ajout de joueurs en compétition (Art. 48.2)</h3>
<ul><li>Il <strong>n'est pas permis d'ajouter un joueur en cours de partie</strong> (exception : poursuite d'une partie suspendue)</li><li>Un joueur peut être inscrit à l'ordre des frappeurs même s'il n'est pas encore présent. À son tour au bâton : soit le <strong>retirer définitivement</strong>, soit <strong>prendre un retrait</strong> à chaque occasion où il est absent</li><li>Une équipe peut continuer même si elle descend sous 9 joueurs <strong>pour cause de blessure</strong> (retrait automatique pour le frappeur absent)</li><li>Si l'équipe ne peut plus aligner <strong>7 joueurs en défensive</strong> → forfait</li></ul>
<h3>Détermination receveur/visiteur en championnat (Art. 40.4)</h3>
<div class="rwrap"><table><thead><tr><th>Classe</th><th>Procédure</th></tr></thead><tbody><tr><td>13U A</td><td>1re partie : déterminée par le classement de la saison précédente. Toutes les autres parties : <strong>tirage au sort 30 min avant</strong></td></tr><tr><td>13U B</td><td><strong>Tirage au sort 30 min avant</strong>, pour <strong>toutes</strong> les parties</td></tr></tbody></table></div>
<div class="callout"><p>Si une équipe est absente au tirage au sort, elle <strong>perd son tirage</strong>.</p></div>
<h3>Protêt en compétition (Art. 42.6)</h3>
<ul><li>L'avis de protêt doit être donné à l'arbitre <strong>au moment de l'infraction</strong>, avant que le jeu ne reprenne</li><li>Accompagné d'un montant de <strong>100 $</strong> comptant</li><li>Le comité de protêt se réunit <strong>immédiatement</strong> — décision définitive et sans appel</li><li>Remboursement si gain de cause</li></ul>
<h3>Bris d'égalité (Art. 42.11)</h3>
<div class="callout warn"><p>⚠️ Les parties nulles sont <strong>interdites</strong> dans les tournois sanctionnés. En cas d'égalité au terme du temps, la procédure de manches supplémentaires s'applique.</p></div>
<h4>Étape A — Égalité dans une même section</h4>
<p>Les priorités sont appliquées dans l'ordre suivant :</p>
<p><strong>Priorité 1 — Fiche tête-à-tête</strong></p>
<p>L'équipe avec la meilleure fiche victoires-défaites dans les parties <strong>entre les équipes à égalité</strong> est placée plus haut.</p>
<p><strong>Priorité 2 — Ratio de points accordés (défense)</strong></p>
<p>Si l'égalité persiste : proportion du nombre de <strong>points accordés</strong> par rapport aux <strong>manches défensives</strong>, pour les parties entre les équipes à égalité.</p>
<ul><li>Plus le ratio est <em>bas</em>, meilleur est le classement</li><li>Les manches sont calculées en <strong>fractions</strong> (ex. : 4⅓ manches)</li></ul>
<p><strong>Priorité 3 — Ratio de points marqués (offensive)</strong></p>
<p>Si l'égalité persiste : proportion du nombre de <strong>points marqués</strong> par rapport aux <strong>manches offensives</strong>, pour les parties entre les équipes à égalité.</p>
<ul><li>Plus le ratio est <em>élevé</em>, meilleur est le classement</li></ul>
<p><strong>Priorité 4 — Manches avec l'avance</strong></p>
<p>Si l'égalité persiste : l'équipe qui a cumulé le plus grand nombre de <strong>manches complètes en avance</strong> au pointage. Un point est accordé à la fin de chaque manche complète à l'équipe qui mène.</p>
<div class="callout"><p><strong>Notes importantes :</strong></p><ul><li>Aux étapes B et C, <strong>toutes</strong> les parties de la ronde préliminaire sont prises en compte (pas seulement les confrontations directes)</li><li>Lors d'une égalité multiple résolue en partie : on continue avec les priorités restantes pour les équipes encore à égalité, avant de recommencer à la priorité 1</li><li><strong>Mercy rule :</strong> l'équipe gagnante reçoit le crédit de <strong>6 manches défensives</strong> (13U = 6 manches); l'équipe perdante reçoit le crédit du nombre de manches réellement jouées</li><li><strong>Manches supplémentaires :</strong> seuls les points des manches régulières comptent dans les ratios — ne pas inclure les statistiques des manches supplémentaires</li><li><strong>Forfait :</strong> les parties non disputées par forfait ne sont <strong>pas</strong> comptabilisées aux fins du « Meilleur deuxième »</li></ul></div>
<h4>Étape B — Meilleur deuxième</h4>
<p>Les équipes ayant terminé <strong>au 2e rang</strong> dans chaque section passent par le même bris d'égalité (Étape A, mais en utilisant <strong>toutes</strong> les parties de la ronde préliminaire) pour déterminer le meilleur deuxième.</p>
<h4>Étape C — Positions 1 à 3</h4>
<p>Définies selon le même bris d'égalité que l'Étape A, en utilisant toutes les parties de la ronde préliminaire.</p>
</section>
<section>
<h2>8. Contact et sécurité</h2>
<h3>Contact coureur/défenseur (Art. 103.10)</h3>
<ul><li>Les coureurs <strong>doivent glisser ou tenter d'éviter</strong> le joueur défensif</li><li>Contact malicieux (force excessive, intention de blesser) → <strong>expulsion automatique</strong>, qu'il soit offensif ou défensif</li></ul>
<h3>Joueur atteint à la tête (Art. 103.24a)</h3>
<p><strong>Frappeur :</strong></p>
<ul><li>Atteint à la tête par un lancer → droit au 1er but, mais <strong>remplacé par le dernier retrait</strong> comme coureur</li><li>Peut retourner en défensive après la demi-manche s'il est apte</li><li>Au prochain tour au bâton : doit se présenter ou être substitué définitivement (Option 1) ou sauter son tour avec 1 retrait comptabilisé (Option 2)</li></ul>
<p><strong>Joueur défensif :</strong></p>
<ul><li>Atteint à la tête → doit être retiré pour la reste de la demi-manche défensive</li><li>Un lanceur retiré peut revenir au monticule s'il est éligible</li></ul>
</section>
<section>
<h2>9. Règles procédurales</h2>
<h3>Visite aux arbitres (Art. 103.21)</h3>
<ul><li><strong>Seul l'entraîneur-chef</strong> peut rendre visite à un arbitre</li><li>Raisons permises : changement de lanceur, substitution, dépôt de protêt, explication d'une règle</li><li>Joueur ou entraîneur-adjoint qui conteste une décision de jugement → <strong>expulsion automatique</strong></li><li>Exception : les entraîneurs-adjoints peuvent faire la rencontre d'avant-partie au marbre et les visites au monticule</li></ul>
<h3>Règle des 60 secondes (Art. 103.23)</h3>
<ul><li>Toutes les classes (A, B) : changements défensifs/offensifs dans un délai de <strong>60 secondes</strong></li><li>L'arbitre des buts chronomètre à partir du dernier retrait</li><li>À l'expiration : l'officiel du marbre appelle <strong>un dernier lancer</strong> au lanceur</li><li><strong>5 lancers de réchauffement</strong> accordés (sans limite de temps) <strong>seulement</strong> lors de :</li><li>(a) début de chaque demi-manche de la <strong>1re manche</strong></li><li>(b) <strong>substitution de lanceur</strong> en cours de manche</li></ul>
<h3>Coussin double au 1er but (Art. 103.11)</h3>
<ul><li>La partie <strong>blanche</strong> = en jeu pour les règles de bonne/fausse balle</li><li>La partie <strong>orange</strong> = utilisée par le frappeur-coureur sur un jeu au 1er but (ou 3e prise échappée)</li><li>Après avoir dépassé le 1er but, le coureur peut se diriger vers le 2e <strong>sans retoucher la partie blanche</strong></li></ul>
<h3>Coureur de courtoisie (Art. 103.16)</h3>
<ul><li><strong>Aucun coureur de courtoisie</strong> n'est permis (même pour le receveur) en 13U A et B</li><li>S'il est utilisé par erreur, la situation est corrigée dès qu'elle est constatée, <strong>sans autre sanction</strong></li></ul>
<h3>Uniforme (Art. 103.3)</h3>
<ul><li>Uniforme complet obligatoire pour joueurs et entraîneurs</li><li>Short, camisole et sandales : <strong>interdits</strong></li></ul>
</section>
<section>
<h2>10. Résumé rapide par classe</h2>
<div class="rwrap"><table><thead><tr><th>Règle</th><th>13U B</th><th>13U A</th></tr></thead><tbody><tr><td>Crampons métalliques</td><td>❌</td><td>❌</td></tr><tr><td>DROP bâton</td><td>Aucune limite</td><td>Aucune limite</td></tr><tr><td>Balles</td><td>9 po</td><td>9 po</td></tr><tr><td>Manches / partie</td><td>6</td><td>6</td></tr><tr><td>Points max / manche</td><td>3</td><td>5</td></tr><tr><td>Équité défensive</td><td>✅</td><td>✅</td></tr><tr><td>Équité offensive</td><td>✅</td><td>✅</td></tr><tr><td>Restriction avance 5 pts</td><td>❌</td><td>✅ sous conditions</td></tr><tr><td>Coureurs libres</td><td>❌ (balle au marbre)</td><td>✅</td></tr><tr><td>Amorti permis</td><td>❌</td><td>✅</td></tr><tr><td>Chandelle intérieure</td><td>❌</td><td>✅</td></tr><tr><td>3e prise échappée</td><td>Retiré auto</td><td>Règles BCan</td></tr><tr><td>Feinte irrégulière</td><td>Non appelée</td><td>Règles BCan</td></tr><tr><td>BSB intentionnel</td><td>❌</td><td>✅</td></tr><tr><td>Retour au jeu</td><td>✅ (103.14d)</td><td>✅ (103.14d)</td></tr><tr><td>Coureur de courtoisie (103.16)</td><td>❌</td><td>❌</td></tr><tr><td>Réserviste comme lanceur (comp.)</td><td>❌</td><td>❌</td></tr><tr><td>Réserviste de classe A (comp.)</td><td>❌</td><td>✅</td></tr><tr><td>Manches lanceur (mai-juin)</td><td>2/jour — 4/sem</td><td>2/jour — 4/sem</td></tr><tr><td>Manches lanceur (juil-sept)</td><td>3/jour — 6/sem</td><td>3/jour — 6/sem</td></tr><tr><td>Lancers compétition (repos 1j)</td><td>41–55</td><td>41–55</td></tr><tr><td>Lancers max / jour (compétition)</td><td>85</td><td>85</td></tr><tr><td>Écart 15 pts (comp. – 3e manche)</td><td>✅</td><td>✅</td></tr><tr><td>Écart 10 pts (saison – 4e manche)</td><td>✅</td><td>✅</td></tr><tr><td>Limite de temps (tournoi)</td><td><strong>70 min</strong></td><td><strong>80 min</strong></td></tr><tr><td>Manches supplémentaires</td><td>✅ (1er/2e buts)</td><td>✅</td></tr></tbody></table></div>
<div class="rules-source">Source : Règlements de Régie Baseball Québec, saison 2026</div>
</section>
`;

var PUBLIC_HTML_TEMPLATE_ = `<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<style>
  :root{
    --pool:#ffe0b2;
    --first:#c8e6c9; --second:#bbdefb; --dark:#37474f; --line:#e0e0e0;
  }
  *{box-sizing:border-box;}
  body{
    margin:0; background:#eceff1; color:#263238;
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;
    -webkit-text-size-adjust:100%;
  }
  #app{max-width:520px; margin:0 auto; padding:12px 12px 32px;}
  header{text-align:center; padding:8px 0 4px;}
  h1{font-size:20px; margin:0; letter-spacing:.2px;}
  #subtitle{font-size:14px; color:#546e7a; margin-top:2px;}
  .controls{display:flex; flex-direction:column; gap:8px; margin:12px 0;}
  .seg{display:flex; gap:6px; flex-wrap:wrap; justify-content:center;}
  .seg button{
    flex:1 1 auto; min-width:64px; padding:9px 10px; font-size:14px;
    border:1px solid #b0bec5; background:#fff; color:#37474f; border-radius:999px;
    cursor:pointer; -webkit-tap-highlight-color:transparent;
  }
  .seg button.active{background:var(--dark); color:#fff; border-color:var(--dark); font-weight:600;}
  .card{
    background:#fff; border-radius:14px; overflow:hidden; margin:0 0 16px;
    box-shadow:0 1px 4px rgba(0,0,0,.12);
  }
  .card-head{
    padding:11px 14px; font-weight:700; font-size:16px; letter-spacing:.3px;
  }
  .pool1,.pool2,.pool3{background:var(--pool);}
  .head-dark{background:var(--dark); color:#fff;}
  table{width:100%; border-collapse:collapse;}
  th,td{padding:9px 8px; font-size:15px; text-align:left;}
  th{font-size:12px; text-transform:uppercase; letter-spacing:.4px; color:#607d8b; border-bottom:1px solid var(--line);}
  td.rank{width:30px; text-align:center; font-weight:700; color:#455a64;}
  td.team{font-weight:600;}
  td.vd{width:50px; text-align:center; white-space:nowrap;}
  td.num{text-align:center; color:#546e7a; font-variant-numeric:tabular-nums; white-space:nowrap;}
  td.seed{width:42px; text-align:center; font-weight:700; color:#1b5e20;}
  th.vd,th.num,th.seed{text-align:center;}
  td.vd,td.num,td.seed,th.vd,th.num,th.seed{padding-left:5px; padding-right:5px;}
  /* Mode « toutes les colonnes » : plus compact pour faire tenir toutes les colonnes. */
  table.full th,table.full td{font-size:13px; padding-top:7px; padding-bottom:7px;}
  table.full td.vd,table.full td.num,table.full td.seed,table.full th.vd,table.full th.num,table.full th.seed{padding-left:3px; padding-right:3px;}
  table.full td.team{font-size:13px;}
  tr.first td{background:var(--first);}
  tr.second td{background:var(--second);}
  tr+tr td{border-top:1px solid var(--line);}
  .note{padding:8px 14px; font-size:12.5px; color:#37474f; background:#f1f8ff; border-top:1px solid var(--line);}
  .qual{padding:6px 0;}
  .qrow{display:flex; justify-content:space-between; padding:8px 14px; border-bottom:1px solid var(--line); font-size:15px;}
  .qrow:last-child{border-bottom:none;}
  .qpos{color:#607d8b;}
  .qteam{font-weight:700;}
  .matchups{padding:10px 14px; background:#fafafa; border-top:1px solid var(--line);}
  .mu{font-size:15px; font-weight:600; padding:3px 0;}
  /* Vue « Résultats » : table large, scroll horizontal de secours sur petit écran. */
  .mwrap{overflow-x:auto;}
  table.mtable th,table.mtable td{font-size:12.5px; padding:7px 5px; white-space:nowrap;}
  table.mtable td.c{text-align:center;}
  table.mtable td.t{font-weight:600; white-space:normal;}
  table.mtable td.win{font-weight:800; color:#1b5e20;}
  table.mtable tr.pending td{color:#9e9e9e;}
  footer{text-align:center; color:#78909c; font-size:12px; margin-top:6px; line-height:1.5;}
  /* ---- Règles du tournoi (panneau superposé, ouvert depuis le pied de page) ---- */
  #rules-link{margin-top:14px; text-align:center;}
  #btn-rules{
    width:100%; padding:12px 14px; font-size:15px; font-weight:600;
    border:1px solid #b0bec5; background:#fff; color:#37474f; border-radius:12px;
    cursor:pointer; -webkit-tap-highlight-color:transparent;
  }
  #btn-rules:active{background:#eceff1;}
  #rules{display:none; padding:4px 0 24px;}
  #btn-rules-back{
    position:sticky; top:8px; z-index:5; margin:4px 0 12px;
    padding:10px 14px; font-size:14px; font-weight:600;
    border:1px solid var(--dark); background:var(--dark); color:#fff;
    border-radius:999px; cursor:pointer; -webkit-tap-highlight-color:transparent;
  }
  .rules-title{font-size:19px; font-weight:800; text-align:center; margin:4px 0 2px;}
  .rules-sub{font-size:13px; color:#546e7a; text-align:center; margin-bottom:8px;}
  #rules section{
    background:#fff; border-radius:14px; box-shadow:0 1px 4px rgba(0,0,0,.12);
    padding:12px 14px 16px; margin:0 0 14px;
  }
  #rules h2{font-size:17px; margin:2px 0 10px; padding-bottom:6px; border-bottom:2px solid var(--pool);}
  #rules h3{font-size:15px; margin:14px 0 6px; color:#37474f;}
  #rules h4{font-size:13.5px; margin:12px 0 5px; color:#546e7a; text-transform:uppercase; letter-spacing:.3px;}
  #rules p{font-size:14px; line-height:1.5; margin:6px 0;}
  #rules ul,#rules ol{margin:6px 0; padding-left:20px;}
  #rules li{font-size:14px; line-height:1.45; margin:3px 0;}
  #rules .rwrap{overflow-x:auto; margin:8px 0;}
  #rules table{width:100%; border-collapse:collapse; font-size:13px;}
  #rules th,#rules td{padding:7px 8px; text-align:left; border:1px solid var(--line); vertical-align:top;}
  #rules th{background:#eceff1; font-size:12px; text-transform:uppercase; letter-spacing:.3px; color:#455a64;}
  #rules .callout{
    background:#f1f8ff; border-left:4px solid #1e88e5; border-radius:0 8px 8px 0;
    padding:8px 12px; margin:8px 0; font-size:13.5px;
  }
  #rules .callout.warn{background:#fff8e1; border-left-color:#f9a825;}
  #rules .callout p{margin:4px 0; font-size:13.5px;}
  #rules .callout ul{margin:4px 0;}
  #rules .rules-source{font-size:12px; color:#78909c; text-align:center; margin-top:10px;}
  .rules-disclaimer{
    background:#ffebee; border:1px solid #ef9a9a; border-left:5px solid #c62828;
    border-radius:0 10px 10px 0; padding:11px 14px; margin:0 0 16px;
    color:#b71c1c; font-size:13.5px; line-height:1.5;
  }
  .rules-disclaimer strong{display:inline; font-weight:800;}
  .rules-disclaimer strong:first-child{display:block; font-size:14.5px; margin-bottom:4px;}
</style>
</head>
<body>
<div id="app">
  <header>
    <h1>🏆 TOURNOI PROVINCIAL DE BASEBALL 13U 2026 DE RIMOUSKI</h1>
    <div id="subtitle"></div>
  </header>
  <div id="main-view">
  <div class="controls">
    <div class="seg" id="seg-mode">
      <button data-mode="standings">Classements</button>
      <button data-mode="matches">Résultats</button>
    </div>
    <div class="seg" id="seg-classe">
      <button data-classe="A">Classe A</button>
      <button data-classe="B">Classe B</button>
    </div>
    <div class="seg" id="seg-view">
      <button data-view="1">Pool 1</button>
      <button data-view="2">Pool 2</button>
      <button data-view="3">Pool 3</button>
      <button data-view="all">Tout</button>
    </div>
    <div class="seg" id="seg-cols">
      <button data-cols="simple">Simple</button>
      <button data-cols="full">Toutes les colonnes</button>
    </div>
  </div>
  <div id="content"></div>
  <footer id="footer"></footer>
  <div id="rules-link"><button type="button" id="btn-rules">📋 Règles du tournoi (13U)</button></div>
  </div>
  <div id="rules">
  <button type="button" id="btn-rules-back">← Retour aux classements</button>
  <div class="rules-disclaimer">
  <strong>⚠️ Document de référence non officiel</strong>
  Ce résumé des règlements 13U a été préparé par le comité organisateur du tournoi pour faciliter la consultation. Il ne remplace pas les <a href="https://media.publicationsports.com/289/1f1291c1-2d55-674c-8998-0242ac120003" target="_blank">Règlements de Régie de Baseball Québec 2026</a>: <strong>en cas de divergence, la version officielle prévaut.</strong>
  </div>
  /*__RULES__*/
  </div>
</div>
<script>
/*__DATA__*/
(function(){
  var DATA = window.DATA || {A:null,B:null,matches:[]};
  var state = {mode:'standings', classe:'A', view:'all', cols:'simple'};

  function el(tag, cls, txt){
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (txt !== undefined && txt !== null) e.textContent = txt;
    return e;
  }
  function fmt3(x){ return (x === null || x === undefined) ? '—' : Number(x).toFixed(3); }
  function vd(s){ return s.v + '-' + s.d; }

  // Définition des colonnes statistiques. Vue « Simple » = essentiel ;
  // vue « Toutes les colonnes » = tableau de pool complet (plus serré).
  var COLDEF = {
    vd: {label:'V-D', cls:'vd',  get:function(s){ return vd(s); }},
    pp: {label:'PP',  cls:'num', get:function(s){ return s.rs; }},
    pc: {label:'PC',  cls:'num', get:function(s){ return s.ra; }},
    mo: {label:'MO',  cls:'num', get:function(s){ return s.mo; }},
    md: {label:'MD',  cls:'num', get:function(s){ return s.md; }},
    rd: {label:'RD',  cls:'num', get:function(s){ return fmt3(s.rd); }},
    ro: {label:'RO',  cls:'num', get:function(s){ return fmt3(s.ro); }}
  };
  var COLS_SIMPLE = ['vd','pc','md','rd'];
  var COLS_FULL   = ['vd','pp','pc','mo','md','rd','ro'];

  function poolCard(pc, full){
    var cols = full ? COLS_FULL : COLS_SIMPLE;
    var card = el('div','card');
    card.appendChild(el('div','card-head pool'+pc.pool, 'POOL '+pc.pool));
    var table = el('table');
    if (full) table.className = 'full';
    var hr = el('tr');
    hr.appendChild(el('th',null,'#'));
    hr.appendChild(el('th',null,'Équipe'));
    cols.forEach(function(k){ hr.appendChild(el('th', COLDEF[k].cls, COLDEF[k].label)); });
    hr.appendChild(el('th','seed','Rang'));   // rang final (1-4), dernière colonne
    table.appendChild(hr);
    pc.standings.forEach(function(s){
      var tr = el('tr', s.rank === 1 ? 'first' : (s.rank === 2 ? 'second' : ''));
      tr.appendChild(el('td','rank', s.rank));
      tr.appendChild(el('td','team', s.team));
      cols.forEach(function(k){ tr.appendChild(el('td', COLDEF[k].cls, COLDEF[k].get(s))); });
      tr.appendChild(el('td','seed', s.seed || ''));
      table.appendChild(tr);
    });
    card.appendChild(table);
    // Note 4 (exclusion des suppl.) n'est PAS pertinente ici : la page publique
    // n'affiche que les ratios RD/RO du tableau de pool (suppl. INCLUSES) et masque
    // les tableaux de bris d'égalité — donc aucun bandeau « exclues » à montrer.
    if (pc.banners.forcedSecond){
      card.appendChild(el('div','note','ℹ 2e de pool désigné par le registraire : '+pc.banners.secondTeam+'.'));
    }
    return card;
  }

  function secondCard(model){
    var seconds = model.seconds || [];
    var card = el('div','card');
    card.appendChild(el('div','card-head pool1','🥈 MEILLEUR 2e'));
    var table = el('table');
    var hr = el('tr');
    hr.appendChild(el('th','seed','Rang'));
    hr.appendChild(el('th',null,'Équipe'));
    hr.appendChild(el('th','num','Pool'));
    hr.appendChild(el('th','vd','V-D'));
    hr.appendChild(el('th','num','RD'));
    table.appendChild(hr);
    seconds.forEach(function(s){
      var tr = el('tr', s.seed ? 'second' : '');
      tr.appendChild(el('td','seed', s.seed || ''));
      tr.appendChild(el('td','team', s.team));
      tr.appendChild(el('td','num', s.pool));
      tr.appendChild(el('td','vd', vd(s)));
      tr.appendChild(el('td','num', fmt3(s.rd)));
      table.appendChild(tr);
    });
    card.appendChild(table);
    return card;
  }

  function semiCard(model){
    var card = el('div','card');
    card.appendChild(el('div','card-head head-dark','✅ DEMI-FINALES'));
    var sf = model.semifinals;
    var labels = ['1re place','2e place','3e place','Meilleur 2e'];
    var q = el('div','qual');
    sf.positions.forEach(function(p, i){
      var row = el('div','qrow');
      row.appendChild(el('span','qpos', labels[i]));
      row.appendChild(el('span','qteam', p.team || '—'));
      q.appendChild(row);
    });
    card.appendChild(q);
    var mu = el('div','matchups');
    mu.appendChild(el('div','mu','DF1 :  '+(sf.demi1.a || '—')+'   vs   '+(sf.demi1.b || '—')));
    mu.appendChild(el('div','mu','DF2 :  '+(sf.demi2.a || '—')+'   vs   '+(sf.demi2.b || '—')));
    card.appendChild(mu);
    return card;
  }

  // Bloc « Meilleur 2e » + « Demi-finales » — affiché SEULEMENT quand toutes les
  // parties de pool de la classe sont jouées (positions inter-pools fiables). Avant
  // cela, une note d'attente avec la progression (les tableaux de pool, eux, restent).
  function appendSemifinalBlock(content, model){
    if (model.poolsComplete){
      content.appendChild(secondCard(model));
      content.appendChild(semiCard(model));
    } else {
      content.appendChild(el('div','note',
        '⏳ Les positions de demi-finale (1-2-3-4) s\\'afficheront lorsque toutes les ' +
        'parties des pools de la classe ' + model.classe + ' auront été jouées (' +
        (model.poolPlayed || 0) + '/' + (model.poolTotal || 0) + ').'));
    }
  }

  function teamPts(name, pts){ return name + (pts === null || pts === undefined ? '' : ' ('+pts+')'); }
  function finLabel(t){ return !t ? '' : (t === 'Supplémentaires' ? 'Suppl.' : t); }
  // Manche affichée : n° de la dernière manche jouée ; pour une partie allée en
  // supplémentaire, on ajoute entre parenthèses le nombre de manches supplémentaires
  // (ex. « 8 (2) » = terminée en 8e, soit 2 manches au-delà des 6 réglementaires).
  function mancheLabel(m){
    if (m.lastInn === null || m.lastInn === undefined || m.lastInn === '') { return ''; }
    if (m.type === 'Supplémentaires'){
      var supp = Number(m.lastInn) - Number(m.manchesPrevues);
      if (supp > 0){ return m.lastInn + ' (' + supp + ')'; }
    }
    return String(m.lastInn);
  }

  function matchesView(){
    var matches = DATA.matches || [];
    var card = el('div','card');
    card.appendChild(el('div','card-head head-dark','📋 RÉSULTATS DES PARTIES'));
    if (!matches.length){
      card.appendChild(el('div','note','Aucune partie générée pour le moment.'));
      return card;
    }
    var wrap = el('div','mwrap');
    var table = el('table','mtable');
    var hr = el('tr');
    ['#','Date/Heure','CL','PO','TR','Éq.1 (Pts)','Éq.2 (Pts)','M','FIN'].forEach(function(h){
      hr.appendChild(el('th', null, h));
    });
    table.appendChild(hr);
    matches.forEach(function(m){
      var tr = el('tr', m.played ? '' : 'pending');
      var win1 = m.played && m.scoreA > m.scoreB;
      var win2 = m.played && m.scoreB > m.scoreA;
      var dh = [m.jour, m.heure].filter(function(x){ return x; }).join(' ');
      tr.appendChild(el('td','c', m.partie));
      tr.appendChild(el('td','c', dh || '—'));
      tr.appendChild(el('td','c', m.classe));
      tr.appendChild(el('td','c', m.pool));
      tr.appendChild(el('td','c', m.terrain || '—'));
      tr.appendChild(el('td', win1 ? 't win' : 't', teamPts(m.eq1, m.scoreA)));
      tr.appendChild(el('td', win2 ? 't win' : 't', teamPts(m.eq2, m.scoreB)));
      tr.appendChild(el('td','c', mancheLabel(m)));
      tr.appendChild(el('td','c', finLabel(m.type)));
      table.appendChild(tr);
    });
    wrap.appendChild(table);
    card.appendChild(wrap);
    return card;
  }

  function render(){
    var inMatches = (state.mode === 'matches');
    document.getElementById('seg-classe').style.display = inMatches ? 'none' : 'flex';
    document.getElementById('seg-view').style.display   = inMatches ? 'none' : 'flex';
    document.getElementById('seg-cols').style.display   = inMatches ? 'none' : 'flex';

    var content = document.getElementById('content');
    content.textContent = '';
    var foot = document.getElementById('footer');
    foot.textContent = '';

    if (inMatches){
      document.getElementById('subtitle').textContent = 'Résultats des parties — Classes A et B';
      content.appendChild(matchesView());
      foot.appendChild(el('div', null,
        'CL = classe (A / B) · PO = pool · TR = terrain · ' +
        'M = n° de la dernière manche jouée ; entre parenthèses = manches supplémentaires (ex. « 8 (2) ») · ' +
        'FIN = type de fin (Normal / Mercy / Forfait / Suppl.).'));
    } else {
      var model = DATA[state.classe];
      document.getElementById('subtitle').textContent = 'Classe ' + state.classe +
        (state.view === 'all' ? '' : ' — Pool ' + state.view);
      var full = (state.cols === 'full');
      if (!model){
        content.appendChild(el('div','note','Aucune donnée disponible pour cette classe.'));
      } else if (state.view === 'all'){
        model.pools.forEach(function(pc){ content.appendChild(poolCard(pc, full)); });
        appendSemifinalBlock(content, model);
      } else {
        var pc = null;
        model.pools.forEach(function(x){ if (String(x.pool) === state.view) pc = x; });
        if (pc) content.appendChild(poolCard(pc, full));
        else content.appendChild(el('div','note','Pool introuvable.'));
        appendSemifinalBlock(content, model);
      }
      foot.appendChild(el('div', null, full
        ? 'PP/PC = points pour/contre · MO/MD = manches off./déf. · RD = PC÷MD · RO = PP÷MO.'
        : 'PC = points contre · MD = manches défensives · RD = PC ÷ MD (plus bas = mieux).'));
      if (model.poolsComplete){
        foot.appendChild(el('div', null, 'Rang = classement final pour les demi-finales : 1-2-3 = 1ers de pool, 4 = meilleur 2e.'));
      }
    }

    var anyModel = DATA.A || DATA.B;
    foot.appendChild(el('div', null, anyModel ? ('Mis à jour : ' + (anyModel.updatedAt || '')) : ''));
    foot.appendChild(el('div', null, 'Rechargez la page pour voir les derniers résultats.'));
    foot.appendChild(el('div', null, DATA.version ? ('Version ' + DATA.version) : ''));

    document.querySelectorAll('#seg-mode button').forEach(function(b){
      b.classList.toggle('active', b.getAttribute('data-mode') === state.mode);
    });
    document.querySelectorAll('#seg-classe button').forEach(function(b){
      b.classList.toggle('active', b.getAttribute('data-classe') === state.classe);
    });
    document.querySelectorAll('#seg-view button').forEach(function(b){
      b.classList.toggle('active', b.getAttribute('data-view') === state.view);
    });
    document.querySelectorAll('#seg-cols button').forEach(function(b){
      b.classList.toggle('active', b.getAttribute('data-cols') === state.cols);
    });
  }

  document.querySelectorAll('#seg-mode button').forEach(function(b){
    b.addEventListener('click', function(){ state.mode = b.getAttribute('data-mode'); render(); });
  });
  document.querySelectorAll('#seg-classe button').forEach(function(b){
    b.addEventListener('click', function(){ state.classe = b.getAttribute('data-classe'); render(); });
  });
  document.querySelectorAll('#seg-view button').forEach(function(b){
    b.addEventListener('click', function(){ state.view = b.getAttribute('data-view'); render(); });
  });
  document.querySelectorAll('#seg-cols button').forEach(function(b){
    b.addEventListener('click', function(){ state.cols = b.getAttribute('data-cols'); render(); });
  });

  render();

  // Panneau « Règles du tournoi » : superpose les règles sur toute la page.
  // (La page ne se rafraîchit plus automatiquement : recharger manuellement pour
  // voir les derniers résultats.)
  function openRules(){
    document.getElementById('main-view').style.display = 'none';
    // Masque le sous-titre (ex. « Classe B — Pool 2 ») : hors contexte sur la
    // page des règles, il pourrait porter à confusion.
    document.getElementById('subtitle').style.display = 'none';
    document.getElementById('rules').style.display = 'block';
    window.scrollTo(0, 0);
  }
  function closeRules(){
    document.getElementById('rules').style.display = 'none';
    document.getElementById('subtitle').style.display = '';
    document.getElementById('main-view').style.display = '';
    window.scrollTo(0, 0);
  }
  document.getElementById('btn-rules').addEventListener('click', openRules);
  document.getElementById('btn-rules-back').addEventListener('click', closeRules);
})();
</script>
</body>
</html>`;
