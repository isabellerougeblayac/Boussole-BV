import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import XLSX from 'xlsx';
import { transformWithEsbuild } from 'vite';

const source = fs.readFileSync(new URL('../src/spreadsheet.ts', import.meta.url), 'utf8');
const { code } = await transformWithEsbuild(source, 'spreadsheet.ts', { loader: 'ts', format: 'esm' });
const { appVersion, parseAmount, qualificationStatus, validateHeaders, safeSheetName, ratingColumns } = await import(`data:text/javascript;base64,${Buffer.from(code).toString('base64')}`);

const model = XLSX.read(fs.readFileSync(new URL('../public/NOM_DE_COLLECTIVITE-BBV-modele.ods', import.meta.url)));
const header = XLSX.utils.sheet_to_json(model.Sheets[model.SheetNames[0]], { header: 1 })[0];
const expense = () => {
  const row = Array(85).fill('');
  row[0] = '2151'; row[1] = '01'; row[9] = 'Travaux'; row[10] = 100;
  return row;
};

test('la version applicative correspond aux métadonnées npm et HTML', () => {
  const pkg = JSON.parse(fs.readFileSync(new URL('../package.json', import.meta.url), 'utf8'));
  const lock = JSON.parse(fs.readFileSync(new URL('../package-lock.json', import.meta.url), 'utf8'));
  assert.equal(pkg.version, `${appVersion}.0`);
  assert.equal(lock.version, pkg.version);
  assert.equal(lock.packages[''].version, pkg.version);
  assert.ok(fs.readFileSync(new URL('../index.html', import.meta.url), 'utf8').includes(`name="application-version" content="${appVersion}"`));
});

test('les montants usuels sont interprétés sans perdre les signes', () => {
  for (const [input, expected] of [[0, 0], ['0', 0], ['1 234,56 €', 1234.56], ['1\u202f234,56', 1234.56], ['1,234.56', 1234.56], ['1.234,56', 1234.56], ['(100)', -100], ['-20,50', -20.5], ['−20', -20], ['1,234', 1.234]]) {
    assert.equal(parseAmount(input), expected, String(input));
  }
});

test('les montants malformés ne sont jamais convertis en zéro', () => {
  for (const input of ['', 'abc', '12 euros', '1 2', '1,2.3', '(+100)', '--10', true, Infinity, NaN]) assert.equal(parseAmount(input), null, String(input));
});

test('le modèle fourni est accepté et les déplacements sont détectés', () => {
  assert.deepEqual(validateHeaders(header), []);
  const moved = [...header];
  [moved[13], moved[14]] = [moved[14], moved[13]];
  assert.equal(validateHeaders(moved).length, 2);
  assert.ok(validateHeaders(header.slice(0, 49)).length);
  const custom = [...header]; custom[2] = 'Mon service';
  assert.deepEqual(validateHeaders(custom), []);
});

test('une cotation complète requiert les six axes et les métadonnées', () => {
  const row = expense();
  assert.equal(qualificationStatus(row), 'unrated');
  row[11] = 'Voirie';
  assert.equal(qualificationStatus(row), 'partial');
  row[12] = 'Construction'; row[13] = 1;
  assert.equal(qualificationStatus(row), 'partial');
  for (const columns of Object.values(ratingColumns)) row[columns.favorable] = 1;
  assert.equal(qualificationStatus(row), 'rated');
  row[14] = 1;
  assert.equal(qualificationStatus(row), 'inconsistent');
  row[14] = 'invalide';
  assert.equal(qualificationStatus(row), 'inconsistent');
  row[10] = 'abc';
  assert.equal(qualificationStatus(row), 'invalid');
});

test('un montant nul est valide et un montant absent ne l’est pas', () => {
  const row = expense(); row[10] = 0;
  assert.equal(qualificationStatus(row), 'unrated');
  row[10] = '';
  assert.equal(qualificationStatus(row), 'invalid');
});

test('les noms longs ou interdits peuvent être écrits en ODS', () => {
  const name = safeSheetName('2026', 'Une collectivité / très longue : avec [caractères]');
  assert.ok(name.length <= 31);
  assert.doesNotMatch(name, /[\\/?:*\[\]]/);
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([header, expense()]), name);
  const roundTrip = XLSX.read(XLSX.write(book, { type: 'buffer', bookType: 'ods' }));
  assert.equal(roundTrip.SheetNames[0], name);
  const rows = XLSX.utils.sheet_to_json(roundTrip.Sheets[name], { header: 1 });
  assert.deepEqual(validateHeaders(rows[0]), []);
  assert.equal(rows[1][10], 100);
});
