# 🏆 Tournoi Baseball 13U — Gestion via Google Sheets

Un unique script **Google Apps Script** (`TournoiBaseball_Script.gs`) qui transforme un
Google Sheet en système complet de gestion d'un tournoi de baseball **13U** (Baseball
Québec) : horaire, saisie des résultats, calcul des manches, classements de pool et
**bris d'égalité conformes à l'Art. 42.11**, jusqu'aux demi-finales.

Il n'y a **ni build, ni gestionnaire de paquets, ni framework** : le fichier `.gs` se
colle tel quel dans l'éditeur Apps Script lié à une feuille Google.

## Fonctionnalités

- **Horaire collé, pas de saisie d'équipes** — on colle l'horaire officiel de Baseball
  Québec dans l'onglet `Configuration` ; tout le reste se régénère à partir de là.
- **Saisie des résultats** dans `Résultats A / B` avec scores, équipe locale, manches,
  type de fin (Normal / Mercy / Forfait).
- **Calcul des manches fractionnaires** (⅓, ⅔) pour les fins hâtives (walk-off, mercy),
  nécessaire à la justesse des ratios de bris d'égalité.
- **Classements automatiques** par pool + Étapes B/C + récapitulatif des demi-finales.
- **Bris d'égalité récursif à trois niveaux** (Art. 42.11) : tête-à-tête → ratio
  RA/manches défensives → ratio RS/manches offensives → manches en avance
  (vérification manuelle signalée).
- **Mise à jour des classements EN DIRECT et multi-postes** — dès qu'une partie est
  entrée au complet, le classement se met à jour pour tous, y compris les responsables
  qui regardent les onglets `Classements`. Conçu pour 2-3 personnes saisissant en
  parallèle (verrou anti-chevauchement).
- **Simulateur de résultats** pour tester tous les cas (Normal/Mercy/Forfait/manches
  supplémentaires/walk-off + égalités à chaque niveau de priorité).

## Installation / déploiement

1. Ouvrir la feuille Google cible → **Extensions › Apps Script**.
2. Coller **tout** le contenu de `TournoiBaseball_Script.gs` dans `Code.gs`, **Enregistrer**.
3. Recharger la feuille de calcul (un menu **🏆 Tournoi Baseball** apparaît).
4. Menu › **Initialiser les feuilles** (crée/réinitialise tous les onglets).
5. Menu › **⚡ Activer la mise à jour auto** *(une seule fois après chaque collage de
   code)* — enregistre le déclencheur installable qui rafraîchit les classements en
   direct. La première fois, Google demandera d'autoriser le script (consentement
   unique du propriétaire).

> ⚠️ « Initialiser les feuilles » **efface et régénère tous les onglets** — ne pas le
> relancer une fois des scores saisis.

## Utilisation annuelle

1. Coller l'horaire de l'onglet `Horaire globalArbitre` du fichier Excel fourni par
   Baseball Québec dans l'onglet **`Configuration`**, à partir de la cellule **A2**.
2. Menu › **Générer les matchs** (répartit chaque match dans `Résultats A` ou `B`
   selon la colonne `# pool`).
3. Saisir les résultats dans `Résultats A / B`. Le classement se met à jour
   automatiquement une fois chaque partie **complètement** entrée.
4. En cas de doute ou de correction en lot : menu › **Mettre à jour les classements**
   (recalcul complet et autoritaire).

L'onglet **`Aide`**, généré automatiquement, explique chaque colonne et la logique des
fractions de manches — pas besoin de consulter ce README pendant la saisie.

## Menu « 🏆 Tournoi Baseball »

| Item | Rôle |
|------|------|
| Initialiser les feuilles | Crée/réinitialise tous les onglets |
| Générer les matchs | Remplit `Résultats A/B` depuis `Configuration` |
| Mettre à jour les classements | Recalcul complet (manuel / filet de sécurité) |
| ⚡ Activer la mise à jour auto | Enregistre le déclencheur live (à faire une fois) |
| Effacer les résultats | Vide les scores saisis |
| 🧪 Simuler résultats de match | Injecte des scores de test sur l'horaire généré |

## Tests hors-ligne

Aucun runtime Apps Script hors de Google, mais deux vérifications locales existent :

```bash
# Vérification de syntaxe JavaScript
cp TournoiBaseball_Script.gs /tmp/check.js && node --check /tmp/check.js

# Logique pure de bris d'égalité (Art. 42.11) + règle de complétude de ligne
node tests/test_tiebreaker.js
```

Le test charge les vraies fonctions de classement du `.gs` (par `eval`) et vérifie qu'un
changement de score peut faire **basculer** l'ordre d'un pool via les ratios, sans
changer les fiches victoires-défaites.

## Fichiers du dépôt

| Fichier | Rôle |
|---------|------|
| `TournoiBaseball_Script.gs` | Le script complet (le seul code) |
| `tests/test_tiebreaker.js` | Test hors-ligne du bris d'égalité + `isRowComplete` |
| `CLAUDE.md` | Notes d'architecture pour le développement |
| `Regles13U_2026.md` / `ReglesRegie2026.pdf` | Règles officielles Baseball Québec |
| `Horaire-tournoi-2026.xlsx` | Format d'horaire officiel (référence) |
| `Équipes.txt` | Notes de travail |

## Licence

[MIT](LICENSE) © 2026 Daniel Tremblay
