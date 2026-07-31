# Boussole-BV

Boussole-BV est une application web statique d'aide à la cotation Budget Vert des dépenses d'une collectivité.

L'application permet d'importer un fichier tableur, de qualifier les lignes de dépense selon des règles intégrées, puis de produire un fichier enrichi avec les informations de cotation.

## Fonctionnalités

- Import de fichiers tableurs compatibles avec SheetJS.
- Qualification des dépenses selon les règles Budget Vert intégrées à l'application.
- Consultation des résultats par ligne.
- Export ou sauvegarde du fichier coté.
- Génération d'une version portable de l'application.

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
