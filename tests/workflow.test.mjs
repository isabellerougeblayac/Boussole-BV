import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
import XLSX from 'xlsx';
import { build } from 'vite';

const result = await build({
  configFile: false,
  logLevel: 'silent',
  plugins: [{ name: 'expose-workflow-for-tests', transform(source, id) {
    if (id.endsWith('/src/main.ts')) return `${source}\nglobalThis.__bbv = { state, loadSpreadsheet, applyQualification, exportWorkbook, buildAxisAnalysis, renderAnalysis, filteredRows, renderTable };`;
  } }],
  build: { write: false, minify: false, lib: {
    entry: fileURLToPath(new URL('../src/main.ts', import.meta.url)), name: 'BBV', formats: ['iife'],
  } },
});
const code = (Array.isArray(result) ? result[0] : result).output.find((item) => item.type === 'chunk').code;
const rules = JSON.parse(fs.readFileSync(new URL('../src/data/rules.json', import.meta.url), 'utf8'));
const fixture = fs.readFileSync(new URL('../public/NOM_DE_COLLECTIVITE-BBV-modele.ods', import.meta.url));
const model = XLSX.read(fixture);
const headers = XLSX.utils.sheet_to_json(model.Sheets[model.SheetNames[0]], { header: 1 })[0];
const expense = (amount = 100) => {
  const row = Array(85).fill('');
  row[0] = '2151'; row[1] = '01'; row[5] = 'Travaux'; row[9] = 'Travaux'; row[10] = amount;
  return row;
};
function app() {
  const context = vm.createContext({
    console, TextDecoder, TextEncoder, Uint8Array, ArrayBuffer, structuredClone,
    document: { querySelector: (selector) => selector === '#app' ? { innerHTML: '' } : null, querySelectorAll: () => [] },
    window: { localStorage: { getItem: () => null, setItem: () => {} } },
  });
  vm.runInContext(code, context);
  return context.__bbv;
}
function file(rows, names = ['2026-Collectivité']) {
  const workbook = XLSX.utils.book_new();
  for (const name of names) XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name);
  const bytes = XLSX.write(workbook, { bookType: 'ods', type: 'array' });
  return { name: 'depenses.ods', arrayBuffer: async () => bytes };
}

test('un ODS valide se charge et les erreurs préservent le travail courant', async () => {
  const api = app();
  await api.loadSpreadsheet(file([headers, expense('1 234,56 €')]));
  assert.equal(api.state.rows[1][10], 1234.56);
  assert.equal(api.state.year, '2026');
  api.state.selectedRows.add(1);
  api.state.domain = 'Voirie';
  const previous = api.state.rows;
  for (const rejected of [
    file([headers, expense('abc')]),
    file([headers, expense()], ['Un', 'Deux']),
    file([headers.slice(0, 49), expense()]),
    file([headers]),
    file([[...headers, 'Colonne en trop'], expense()]),
    { name: 'cassé.ods', arrayBuffer: async () => new ArrayBuffer(8) },
    { name: 'erreur.ods', arrayBuffer: async () => { throw new Error('lecture impossible'); } },
  ]) {
    await api.loadSpreadsheet(rejected);
    assert.match(api.state.status, /Import refusé/);
    assert.equal(api.state.rows, previous);
    assert.equal(api.state.selectedRows.has(1), true);
    assert.equal(api.state.domain, 'Voirie');
    assert.equal(api.state.loading, false);
  }
});

test('les lignes vides sont ignorées et les montants absents restent explicitement exclus', async () => {
  const api = app();
  await api.loadSpreadsheet(file([headers, expense(0), [], expense('')]));
  assert.equal(api.state.rows.length, 3);
  api.state.selectedRows.add(1); api.state.selectedRows.add(2);
  const html = api.renderAnalysis();
  assert.match(html, /1 ligne\(s\) sans montant valide exclue/);
  assert.match(html, /Montant net sélectionné<\/span><strong>0/);
});

test('les cotations conflictuelles sont signalées puis réparées par application', async () => {
  const api = app();
  const row = expense(); row[13] = 1; row[14] = 1;
  await api.loadSpreadsheet(file([headers, row]));
  assert.match(api.state.status, /1 incohérente/);
  const first = rules.rules[0];
  api.state.domain = first.domain;
  api.state.expenseType = first.expenseType;
  for (const axis of rules.axes) api.state.selectedRuleByAxis[axis.id] = 0;
  api.state.selectedRows.add(1);
  api.applyQualification();
  assert.equal(api.state.rows[1][84], '1.2');
  for (let axis = 0; axis < 6; axis++) {
    assert.equal(api.state.rows[1].slice(13 + axis * 6, 19 + axis * 6).filter((value) => value === 1).length, 1);
    assert.ok(api.state.rows[1][78 + axis]);
  }
});

test('les graphiques utilisent un volume absolu et la synthèse garde le total net', async () => {
  const api = app();
  const positive = expense(100); positive[13] = 1;
  const negative = expense(-20); negative[16] = 1;
  await api.loadSpreadsheet(file([headers, positive, negative]));
  api.state.selectedRows.add(1); api.state.selectedRows.add(2);
  const analysis = api.buildAxisAnalysis(api.state.rows.slice(1), rules.axes[0]);
  assert.equal(analysis.total, 120);
  assert.ok(Math.abs(analysis.ratings.reduce((sum, rating) => sum + rating.percent, 0) - 100) < 1e-10);
  assert.ok(analysis.ratings.every((rating) => rating.percent >= 0 && rating.percent <= 100));
  assert.equal(analysis.ratings.find((rating) => rating.label === 'Défavorable').operations[0].amount, 20);
  const html = api.renderAnalysis();
  assert.match(html, /Montant net sélectionné<\/span><strong>80/);
  assert.match(html, /valeurs absolues/);
});

test('un échec d’export est affiché et conserve les données', async () => {
  const api = app();
  await api.loadSpreadsheet(file([headers, expense()]));
  const previous = api.state.rows;
  // Aucun téléchargement disponible dans cet environnement DOM simulé.
  api.exportWorkbook();
  assert.match(api.state.status, /Export impossible/);
  assert.equal(api.state.rows, previous);
  assert.equal(api.state.sheetName, '2026-Collectivité');
});

test('les filtres simples traitent les caractères spéciaux comme du texte', async () => {
  const api = app();
  const row = expense(); row[9] = 'Travaux (phase 1)';
  await api.loadSpreadsheet(file([headers, row, expense()]));
  api.state.filters[9] = '(';
  assert.equal(api.filteredRows().length, 1);
  api.state.filters[9] = 'TRAVAUX';
  assert.equal(api.filteredRows().length, 2);
});

test('les regex se combinent entre colonnes, avec les statuts et la sélection', async () => {
  const api = app();
  const first = expense(); first[0] = '2151'; first[9] = 'Travaux';
  const second = expense(); second[0] = '2182'; second[9] = 'Rénovation';
  const third = expense(); third[0] = '6151'; third[9] = 'Entretien';
  await api.loadSpreadsheet(file([headers, first, second, third]));
  api.state.filters[0] = '^21[0-9]+$'; api.state.filterRegex[0] = true;
  api.state.filters[9] = 'TRAVAUX|rénovation'; api.state.filterRegex[9] = true;
  assert.equal(api.filteredRows().length, 2);
  api.filteredRows().forEach(({ index }) => api.state.selectedRows.add(index));
  assert.deepEqual([...api.state.selectedRows], [1, 2]);
  api.state.statusFilter = 'rated';
  assert.equal(api.filteredRows().length, 0);
  api.state.statusFilter = 'unrated';
  assert.equal(api.filteredRows().length, 2);
  api.state.filters[9] = 'tion$';
  assert.equal(api.filteredRows().length, 1);
});

test('les regex invalides sont signalées sans exception et les modes se réinitialisent à l’import', async () => {
  const api = app();
  await api.loadSpreadsheet(file([headers, expense()]));
  api.state.filters[9] = '['; api.state.filterRegex[9] = true;
  assert.equal(api.filteredRows().length, 0);
  const html = api.renderTable();
  assert.match(html, /Expression régulière invalide/);
  assert.match(html, /aria-invalid="true" aria-describedby="filter-error-9"/);
  assert.match(html, /id="regex-help" role="tooltip"/);
  api.state.filters[9] = '';
  assert.equal(api.filteredRows().length, 1);
  await api.loadSpreadsheet(file([headers, expense()]));
  assert.equal(api.state.filterRegex.some(Boolean), false);
  assert.equal(api.state.filters.some(Boolean), false);
});
