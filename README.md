# Boussole-BV

Version actuelle : **1.2** (`1.2.0` au format semver dans les fichiers npm).

Boussole-BV est une application web statique d'aide à la cotation Budget Vert des dépenses d'une collectivité.

L'application permet d'importer un fichier tableur, de qualifier les lignes de dépense selon des règles intégrées, puis de produire un fichier enrichi avec les informations de cotation.

## Fonctionnalités

- Import de fichiers tableurs compatibles avec SheetJS.
- Qualification des dépenses selon les règles Budget Vert intégrées à l'application.
- Consultation des résultats par ligne.
- Export ou sauvegarde du fichier coté.
- Génération d'une version portable de l'application.

## Validation et calculs (version 1.2)

- Seuls les fichiers ODS contenant un onglet et au moins une dépense sont acceptés.
- Le fichier doit respecter les 85 colonnes du modèle. Les en-têtes Nature,
  Rubrique, Libellé, Montant de la réalisation et les en-têtes techniques des axes,
  des critères et de version sont contrôlés à leur position exacte. Les autres
  en-têtes peuvent être personnalisés ; les colonnes ne doivent pas être déplacées.
- Un montant renseigné mais invalide entraîne le refus du fichier avec indication
  des lignes concernées. Les nombres français (`1 234,56 €`), les nombres à
  séparateurs internationaux (`1,234.56`) et les négatifs (`-100`, `(100)`) sont
  acceptés. Un séparateur unique est décimal : `1,234` signifie `1.234`.
- Les lignes entièrement vides sont ignorées. Les dépenses sans champ obligatoire
  restent visibles comme données incomplètes ; un montant absent est exclu de l'analyse.
- Une cotation complète exige un domaine, un type et exactement une catégorie
  par axe. Les cotations partielles et incohérentes ont leurs propres statuts et
  filtres. Les indicateurs autorisés sont vide, 0 et 1. Les incohérences peuvent
  être corrigées en appliquant une nouvelle cotation complète.
- Les totaux sont nets. Les graphiques et classements des opérations utilisent
  les valeurs absolues : une dépense de 100 € et une régularisation de -20 €
  représentent un total net de 80 € et un volume graphique de 120 €.
- Un import refusé préserve le fichier en cours, ses cotations et sa sélection.
- L'export conserve les informations de cellules que SheetJS sait restituer
  lorsque les coordonnées du fichier source n'ont pas changé. La fidélité des
  styles ODS n'est pas garantie. Après suppression de lignes vides, le tableau
  est reconstruit pour éviter des décalages. Les noms d'onglets sont limités à
  31 caractères ; le nom du fichier conserve l'identité complète.

## Vérifications

### Filtres et expressions régulières

Chaque filtre de colonne dispose d'un bouton `.*` pour activer les expressions
régulières. La recherche simple reste le mode par défaut. Les deux modes ignorent
la casse ; plusieurs colonnes se combinent avec un ET et avec le filtre de statut.
Saisissez les expressions sans délimiteurs `/…/` : `^21` (commence par 21),
`2026$` (se termine par 2026), `travaux|rénovation` (l'un ou l'autre),
`\.` (un point littéral). Une info-bulle dans l'interface rappelle ces exemples.
Une expression invalide est signalée sous le champ et ne renvoie aucun résultat.
Un champ vide ne filtre pas. Le bouton de remise à zéro et un nouvel import réussi
effacent les recherches et désactivent les modes regex.

### Commandes

Avec Node.js 18 ou plus récent :

```bash
npm test
npm run build
npm run export:portable
```

## Prérequis

- Node.js
- npm
- Python 3 uniquement pour régénérer les règles depuis un fichier source externe

## Installation

```bash
npm install
```

## Lancement en développement

```bash
npm run dev
```

L'application est ensuite accessible depuis l'adresse affichée par Vite dans le terminal.

## Build production

```bash
npm run build
```

Le build est généré dans le dossier `dist/`.

## Prévisualisation du build

```bash
npm run preview
```

## Export portable

```bash
npm run export:portable
```

Cette commande génère une version portable dans le dossier `export-portable/`.

## Export pour GitHub Pages

```bash
npm run export:github-pages
```

Cette commande vérifie les tests et la compilation, puis génère le dossier
`export-github-pages/`. Il contient les sources, le logo, le modèle ODS, les tests,
les fichiers npm et le workflow de déploiement automatique GitHub Pages.

Le guide **[DEPLOIEMENT-GITHUB-PAGES.md](DEPLOIEMENT-GITHUB-PAGES.md)** décrit la
création du dépôt, l'activation de Pages et l'envoi des fichiers. Le contenu du
dossier exporté doit être placé à la racine du dépôt GitHub. Dans les paramètres
Pages, choisissez **GitHub Actions** comme source.

L'export est régénéré à chaque exécution : effectuez vos modifications dans le
projet source ou une copie de travail. Un dossier exporté initialisé comme dépôt
Git ne sera pas écrasé.

## Structure du projet

- `src/` : code source TypeScript et styles de l'application.
- `src/data/rules.json` : règles Budget Vert intégrées à l'application.
- `public/` : fichiers statiques servis par Vite.
- `scripts/` : scripts utilitaires, notamment pour générer les règles.
- `index.html` : point d'entrée HTML.
- `vite.config.ts` : configuration Vite.
- `tsconfig.json` : configuration TypeScript.

## Régénération des règles

Le script `scripts/extract-rules.py` permet de régénérer `src/data/rules.json` à partir d'un fichier source externe.

Les fichiers métier sources ne sont pas inclus dans ce dépôt. Pour régénérer les règles, placez le fichier source attendu à l'emplacement requis par le script, puis lancez :

```bash
python3 scripts/extract-rules.py
```

## Dépôt GitHub

Le dépôt du projet est prévu à l'adresse suivante :

```text
https://github.com/isabellerougeblayac/Boussole-BV
```

## Licence

Copyright © 2026 isabellerougeblayac.

Ce logiciel est mis à disposition sous la licence
PolyForm Noncommercial 1.0.0.

Toute utilisation commerciale nécessite l'obtention préalable
d'une licence commerciale distincte auprès de l'auteur.

Contact : aucun contact fourni
