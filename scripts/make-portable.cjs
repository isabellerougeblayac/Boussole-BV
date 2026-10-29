const fs = require('node:fs');
const path = require('node:path');

const exportDir = path.resolve(__dirname, '..', 'export-portable');
const indexPath = path.join(exportDir, 'index.html');

let html = fs.readFileSync(indexPath, 'utf8');

html = html
  .replace(/<script\s+type="module"\s+crossorigin\s+src="([^"]+)"><\/script>/g, '<script defer src="$1"></script>')
  .replace(/<script\s+type="module"\s+src="([^"]+)"><\/script>/g, '<script defer src="$1"></script>')
  .replace(/\s+crossorigin(?=[\s>])/g, '');

fs.writeFileSync(indexPath, html);
