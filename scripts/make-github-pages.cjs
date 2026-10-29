const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const destination = path.join(root, 'export-github-pages');
const markerName = '.bbv-export.json';
const markerKind = 'boussole-budget-vert-github-pages';
const entries = [
  '.github', '.gitignore', 'src', 'public', 'scripts', 'tests',
  'index.html', 'package.json', 'package-lock.json', 'tsconfig.json',
  'vite.config.ts', 'logo-afigese-fond-blanc-1536x561.png',
  'README.md', 'DEPLOIEMENT-GITHUB-PAGES.md', 'LICENSE',
];

for (const entry of entries) {
  if (!fs.existsSync(path.join(root, entry))) throw new Error(`Fichier requis absent : ${entry}`);
}

if (fs.existsSync(destination)) {
  const marker = path.join(destination, markerName);
  if (fs.lstatSync(destination).isSymbolicLink() || !fs.existsSync(marker) ||
      JSON.parse(fs.readFileSync(marker, 'utf8')).kind !== markerKind ||
      fs.existsSync(path.join(destination, '.git'))) {
    throw new Error('Le dossier export-github-pages contient un dossier non généré ou un dépôt Git. Déplacez-le avant de régénérer l’export.');
  }
}

const temporary = fs.mkdtempSync(path.join(root, '.github-pages-export-'));
try {
  for (const entry of entries) {
    fs.cpSync(path.join(root, entry), path.join(temporary, entry), { recursive: true });
  }
  const { version } = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  fs.writeFileSync(path.join(temporary, markerName), JSON.stringify({ kind: markerKind, version }, null, 2) + '\n');
  fs.rmSync(destination, { recursive: true, force: true });
  fs.renameSync(temporary, destination);
  console.log(`Export GitHub Pages ${version} prêt : ${destination}`);
  console.log('Consultez DEPLOIEMENT-GITHUB-PAGES.md dans ce dossier pour la mise en ligne.');
} finally {
  fs.rmSync(temporary, { recursive: true, force: true });
}
