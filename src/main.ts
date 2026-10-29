import './styles.css';
import * as XLSX from 'xlsx';
import rulesData from './data/rules.json';
import afigeseLogo from '../logo-afigese-fond-blanc-1536x561.png';
import { compileFilter } from './filters';
import { appVersion, ratingColumns, criterionIdColumns, exportVersionColumn, normalize,
  parseAmount, requiredIssues, qualificationIssues, qualificationStatus, validateHeaders, safeSheetName,
  type Row, type RowStatus, type RatingBucket } from './spreadsheet';

type Axis = { id: string; key: string; label: string };
type Rule = {
  axisId: string;
  axisKey: string;
  axisLabel: string;
  domain: string;
  domainKey: string;
  expenseType: string;
  expenseTypeKey: string;
  criteria: string;
  rating: string;
  criterionId: string;
};
type RulesData = { axes: Axis[]; domains: string[]; typesByDomain: Record<string, string[]>; rules: Rule[] };
type ColumnDefinition = { id: string; label: string; kind: 'action' | 'status' | 'source' | 'bbv' | 'axis'; sourceIndex?: number; axisId?: string };
type AnalysisRating = { label: string; className: string; color: string };
type OperationAmount = { label: string; amount: number };
type RatingAnalysis = AnalysisRating & { amount: number; percent: number; operations: OperationAmount[] };
type AxisAnalysis = { axis: Axis; total: number; ratings: RatingAnalysis[] };

const data = rulesData as RulesData;
const app = document.querySelector<HTMLDivElement>('#app');
const templateFileName = 'NOM_DE_COLLECTIVITE-BBV-modele.ods';
const templateDownloadHref = `${import.meta.env.BASE_URL}${templateFileName}`;
const sourceCodeUrl = 'https://github.com/isabellerougeblayac/Boussole-BV';

const analysisRatings: AnalysisRating[] = [
  { label: 'Très favorable', className: 'good', color: '#4f7a32' },
  { label: 'Favorable sous conditions', className: 'condition', color: '#92d050' },
  { label: 'Neutre', className: 'neutral', color: '#bfbfbf' },
  { label: 'Défavorable', className: 'bad', color: '#cc6600' },
  { label: 'Non cotée', className: 'unknown', color: '#eadcf4' },
  { label: 'Hors périmètre', className: 'out-of-scope', color: '#7b6f63' },
];

const sourceHeaders = [
  'Nature',
  'Rubrique',
  'Champ 1',
  'Champ 2',
  'Champ 3',
  'Champ 4',
  'Champ 5',
  'Champ 6',
  'Champ 7',
  'Libellé',
  'Montant de la réalisation',
];

const tableColumns: ColumnDefinition[] = [
  { id: 'action', label: 'Action', kind: 'action' },
  { id: 'status', label: 'Statut', kind: 'status' },
  ...sourceHeaders.map((label, sourceIndex) => ({ id: `source-${sourceIndex}`, label, kind: 'source' as const, sourceIndex })),
  { id: 'bbv-domain', label: 'Domaine BBV', kind: 'bbv' },
  { id: 'bbv-type', label: 'Type BBV', kind: 'bbv' },
  ...data.axes.map((axis) => ({ id: `axis-${axis.id}`, label: axis.key, kind: 'axis' as const, axisId: axis.id })),
];

const visibleColumnsStorageKey = 'bbv.visibleColumns';
const columnChooserOpenStorageKey = 'bbv.columnChooserOpen';
const allColumnIds = tableColumns.map((column) => column.id);

function loadVisibleColumns(): Set<string> {
  try {
    const stored = window.localStorage.getItem(visibleColumnsStorageKey);
    if (!stored) return new Set(allColumnIds);
    const parsed = JSON.parse(stored) as unknown;
    if (!Array.isArray(parsed)) return new Set(allColumnIds);
    const validIds = parsed.filter((id): id is string => typeof id === 'string' && allColumnIds.includes(id));
    return validIds.length ? new Set(validIds) : new Set(allColumnIds);
  } catch {
    return new Set(allColumnIds);
  }
}

function loadColumnChooserOpen(): boolean {
  try {
    return window.localStorage.getItem(columnChooserOpenStorageKey) === 'true';
  } catch {
    return false;
  }
}

function saveVisibleColumns(): void {
  try {
    window.localStorage.setItem(visibleColumnsStorageKey, JSON.stringify([...state.visibleColumns]));
  } catch {
    // L'application reste utilisable si le navigateur bloque le stockage local.
  }
}

function saveColumnChooserOpen(): void {
  try {
    window.localStorage.setItem(columnChooserOpenStorageKey, String(state.columnChooserOpen));
  } catch {
    // L'application reste utilisable si le navigateur bloque le stockage local.
  }
}


const state = {
  domain: '',
  expenseType: '',
  selectedRuleByAxis: {} as Record<string, number>,
  workbook: null as XLSX.WorkBook | null,
  sheetName: '',
  rows: [] as Row[],
  columnLabels: [] as string[],
  filters: Array(11).fill('') as string[],
  filterRegex: Array(11).fill(false) as boolean[],
  statusFilter: 'all' as 'all' | RowStatus,
  visibleColumns: loadVisibleColumns(),
  columnChooserOpen: loadColumnChooserOpen(),
  selectedRows: new Set<number>(),
  year: '',
  community: '',
  status: '',
  activeTool: 'single' as 'single' | 'file' | 'help',
  loading: false,
};

function escapeHtml(value: unknown): string {
  return normalize(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function escapeHtmlPreserveLines(value: unknown): string {
  return String(value ?? '')
    .replace(/\xa0/g, ' ')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function ratingClass(rating: string): string {
  const lower = rating.toLocaleLowerCase('fr-FR');
  if (lower.includes('hors périmètre')) return 'out-of-scope';
  if (lower.includes('défavorable')) return 'bad';
  if (lower.includes('condition')) return 'condition';
  if (lower.includes('très favorable') || lower === 'favorable') return 'good';
  if (lower.includes('neutre')) return 'neutral';
  return 'unknown';
}

function ratingBucket(rating: string): RatingBucket {
  const lower = rating.toLocaleLowerCase('fr-FR');
  if (lower.includes('hors périmètre')) return 'outOfScope';
  if (lower.includes('défavorable')) return 'unfavorable';
  if (lower.includes('condition')) return 'favorableCondition';
  if (lower.includes('très favorable') || lower === 'favorable') return 'favorable';
  if (lower.includes('neutre')) return 'neutral';
  return 'unrated';
}

function bucketLabel(bucket: string): string {
  return {
    favorable: 'Très favorable',
    favorableCondition: 'Favorable sous conditions',
    neutral: 'Neutre',
    unfavorable: 'Défavorable',
    unrated: 'Non cotée',
    outOfScope: 'Hors périmètre',
  }[bucket] ?? '';
}

function rowAxisRating(row: Row, axisId: string): string {
  const columns = ratingColumns[axisId];
  const values = Object.values(columns).map((column) => normalize(row[column]));
  if (values.some((value) => !['', '0', '1'].includes(value)) || values.filter((value) => value === '1').length > 1) return 'Incohérente';
  const found = Object.entries(columns).find(([, column]) => normalize(row[column]) === '1');
  return found ? bucketLabel(found[0]) : '';
}

function rowHasRating(row: Row): boolean {
  return rowStatus(row) === 'rated';
}

function rowStatus(row: Row): RowStatus {
  return qualificationStatus(row);
}

function rowStatusLabel(status: RowStatus): string {
  return { rated: 'Cotée', partial: 'Cotation partielle', inconsistent: 'Cotation incohérente', invalid: 'Données incomplètes', unrated: 'À coter' }[status];
}

function amountForRow(row: Row): number {
  const amount = parseAmount(row[10]);
  if (amount === null) throw new Error('Montant absent ou invalide.');
  return amount;
}

function formatEuro(value: number): string {
  return new Intl.NumberFormat('fr-FR', { style: 'currency', currency: 'EUR', maximumFractionDigits: 2 }).format(value);
}

function formatDonutAmount(value: number): string {
  const formatter = new Intl.NumberFormat('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  if (Math.abs(value) >= 1_000_000) return `${formatter.format(value / 1_000_000)} M€`;
  if (Math.abs(value) >= 1_000) return `${formatter.format(value / 1_000)} k€`;
  return formatEuro(value);
}

function formatPercent(value: number): string {
  return new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 1 }).format(value);
}

function selectedAnalysisRows(): Row[] {
  return [...state.selectedRows]
    .map((index) => state.rows[index])
    .filter((row): row is Row => Boolean(row) && parseAmount(row[10]) !== null);
}

function analysisRatingForRow(row: Row, axisId: string): string {
  const rating = rowAxisRating(row, axisId);
  return rating === 'Incohérente' ? 'Non cotée' : rating || 'Non cotée';
}

function topOperations(rows: Row[], axisId: string, rating: string, total: number): OperationAmount[] {
  const amounts = new Map<string, number>();
  rows.forEach((row) => {
    if (analysisRatingForRow(row, axisId) !== rating) return;
    const operation = normalize(row[5]) || 'Opération non renseignée';
    amounts.set(operation, (amounts.get(operation) ?? 0) + Math.abs(amountForRow(row)));
  });

  const sorted = [...amounts.entries()]
    .map(([label, amount]) => ({ label, amount }))
    .filter((item) => item.amount > 0)
    .sort((a, b) => b.amount - a.amount);
  const top = sorted.slice(0, 5);
  const other = Math.max(0, total - top.reduce((sum, item) => sum + item.amount, 0));
  return other > 0 ? [...top, { label: 'Autres', amount: other }] : top;
}

function buildAxisAnalysis(rows: Row[], axis: Axis): AxisAnalysis {
  const total = rows.reduce((sum, row) => sum + Math.abs(amountForRow(row)), 0);
  const ratings = analysisRatings.map((rating) => {
    const amount = rows
      .filter((row) => analysisRatingForRow(row, axis.id) === rating.label)
      .reduce((sum, row) => sum + Math.abs(amountForRow(row)), 0);
    return {
      ...rating,
      amount,
      percent: total > 0 ? (amount / total) * 100 : 0,
      operations: topOperations(rows, axis.id, rating.label, amount),
    };
  });
  return { axis, total, ratings };
}

function polarPoint(center: number, radius: number, angle: number): { x: number; y: number } {
  const radians = ((angle - 90) * Math.PI) / 180;
  return { x: center + radius * Math.cos(radians), y: center + radius * Math.sin(radians) };
}

function donutPath(startAngle: number, endAngle: number, outerRadius = 90, innerRadius = 54): string {
  const center = 100;
  const largeArc = endAngle - startAngle > 180 ? 1 : 0;
  const outerStart = polarPoint(center, outerRadius, endAngle);
  const outerEnd = polarPoint(center, outerRadius, startAngle);
  const innerStart = polarPoint(center, innerRadius, startAngle);
  const innerEnd = polarPoint(center, innerRadius, endAngle);
  return [
    `M ${outerStart.x} ${outerStart.y}`,
    `A ${outerRadius} ${outerRadius} 0 ${largeArc} 0 ${outerEnd.x} ${outerEnd.y}`,
    `L ${innerStart.x} ${innerStart.y}`,
    `A ${innerRadius} ${innerRadius} 0 ${largeArc} 1 ${innerEnd.x} ${innerEnd.y}`,
    'Z',
  ].join(' ');
}

function visibleTableColumns(): ColumnDefinition[] {
  return tableColumns.filter((column) => state.visibleColumns.has(column.id));
}

function uploadedColumnLabel(index: number, fallback: string): string {
  return normalize(state.columnLabels[index]) || fallback;
}

function columnLabel(column: ColumnDefinition): string {
  if (column.kind === 'source') return uploadedColumnLabel(column.sourceIndex ?? 0, column.label);
  if (column.id === 'bbv-domain') return uploadedColumnLabel(11, column.label);
  if (column.id === 'bbv-type') return uploadedColumnLabel(12, column.label);
  return column.label;
}

function setVisibleColumns(ids: string[]): void {
  state.visibleColumns = new Set(ids);
  saveVisibleColumns();
  render();
}

function renderColumnCell(column: ColumnDefinition, row: Row, index: number, status: RowStatus): string {
  if (column.kind === 'action') {
    return `<td class="action-cell"><input type="checkbox" data-row="${index}" ${state.selectedRows.has(index) ? 'checked' : ''} aria-label="Sélectionner la ligne ${index + 1}" /><button type="button" class="small" data-quote-row="${index}">Coter cette ligne</button></td>`;
  }
  if (column.kind === 'status') return `<td><span class="status-badge ${status}">${rowStatusLabel(status)}</span><small class="row-issues">${escapeHtml([...validateRequired(row), ...qualificationIssues(row)].join(' · '))}</small></td>`;
  if (column.kind === 'source') return `<td>${escapeHtml(row[column.sourceIndex ?? 0])}</td>`;
  if (column.id === 'bbv-domain') return `<td>${escapeHtml(row[11])}</td>`;
  if (column.id === 'bbv-type') return `<td>${escapeHtml(row[12])}</td>`;
  const rating = column.axisId ? rowAxisRating(row, column.axisId) : '';
  return `<td>${rating ? `<span class="pill ${ratingClass(rating)}">${escapeHtml(rating)}</span>` : ''}</td>`;
}

function renderColumnFilterCell(column: ColumnDefinition): string {
  if (column.kind !== 'source') return '<th></th>';
  const index = column.sourceIndex ?? 0;
  const regexMode = state.filterRegex[index];
  const { error } = compileFilter(state.filters[index], regexMode);
  const label = escapeHtml(columnLabel(column));
  return `<th>
    <div class="filter-control">
      <input class="filter" data-filter="${index}" value="${escapeHtmlPreserveLines(state.filters[index])}" aria-label="Filtrer ${label}${regexMode ? ' par expression régulière' : ''}" aria-invalid="${Boolean(error)}" ${error ? `aria-describedby="filter-error-${index}"` : ''} />
      <button type="button" class="small regex-toggle ${regexMode ? 'active' : 'secondary'}" data-regex-toggle="${index}" aria-pressed="${regexMode}" aria-label="Expressions régulières pour ${label}" title="${regexMode ? 'Désactiver' : 'Activer'} les expressions régulières">.*</button>
    </div>
    ${error ? `<small class="filter-error" id="filter-error-${index}" role="status">${escapeHtml(error)}</small>` : ''}
  </th>`;
}

function renderColumnChooser(): string {
  return `
    <details class="column-chooser" ${state.columnChooserOpen ? 'open' : ''}>
      <summary>Colonnes visibles</summary>
      <div class="column-actions">
        <button type="button" class="small" id="show-all-columns">Tout afficher</button>
      </div>
      <div class="column-grid">
        ${tableColumns.map((column) => `
          <label>
            <input type="checkbox" data-column-toggle="${column.id}" ${state.visibleColumns.has(column.id) ? 'checked' : ''} />
            ${escapeHtml(columnLabel(column))}
          </label>
        `).join('')}
      </div>
    </details>
  `;
}

function currentRules(): Rule[] {
  return data.rules.filter((rule) => rule.domain === state.domain && rule.expenseType === state.expenseType);
}

function selectedRules(): Rule[] {
  return data.axes
    .map((axis) => currentRules().filter((rule) => rule.axisId === axis.id)[state.selectedRuleByAxis[axis.id] ?? -1])
    .filter(Boolean);
}

function canApplyQualification(): boolean {
  return Boolean(state.domain && state.expenseType && selectedRules().length === data.axes.length);
}

function applyDisabledReason(): string {
  if (!state.rows.length) return 'Téléversez un fichier ODS.';
  if (!state.domain || !state.expenseType) return 'Sélectionnez un domaine et un type de dépense.';
  if (selectedRules().length !== data.axes.length) return 'Sélectionnez un critère pour chacun des 6 axes.';
  if (!state.selectedRows.size) return 'Sélectionnez au moins une ligne de dépense.';
  return '';
}

function parseSheetIdentity(sheetName: string): void {
  const parts = sheetName.split('-').filter(Boolean);
  state.year = parts[0] || new Date().getFullYear().toString();
  state.community = parts.slice(1).join('-') || 'NOM_DE_COLLECTIVITE';
}

function setStatus(message: string): void {
  state.status = message;
  render();
}

function resetQualificationForm(): void {
  state.domain = '';
  state.expenseType = '';
  state.selectedRuleByAxis = {};
}

function filteredRows(): Array<{ row: Row; index: number }> {
  const filters = state.filters.map((filter, column) => compileFilter(filter, state.filterRegex[column]));
  return state.rows
    .slice(1)
    .map((row, offset) => ({ row, index: offset + 1 }))
    .filter(({ row }) => {
      const matchesText = filters.every((filter, column) => filter.matches(row[column]));
      const status = rowStatus(row);
      const matchesStatus = state.statusFilter === 'all' || state.statusFilter === status;
      return matchesText && matchesStatus;
    });
}

function hasActiveColumnFilters(): boolean {
  return state.filters.some((filter) => Boolean(normalize(filter))) || state.filterRegex.some(Boolean);
}

function ensureWidth(row: Row, width: number): Row {
  while (row.length < width) row.push('');
  return row;
}

function validateRequired(row: Row): string[] {
  return requiredIssues(row);
}

function applyQualification(): void {
  const rules = selectedRules();
  if (!canApplyQualification()) {
    setStatus('Sélectionnez un critère pour chacun des 6 axes avant application.');
    return;
  }
  if (!state.selectedRows.size) {
    setStatus('Sélectionnez au moins une ligne de dépense.');
    return;
  }

  let skipped = 0;
  state.selectedRows.forEach((rowIndex) => {
    const row = ensureWidth(state.rows[rowIndex], 85);
    if (validateRequired(row).length) {
      skipped += 1;
      return;
    }
    row[11] = state.domain;
    row[12] = state.expenseType;
    for (const axis of data.axes) {
      const columns = ratingColumns[axis.id];
      Object.values(columns).forEach((column) => (row[column] = ''));
      row[criterionIdColumns[axis.id]] = '';
      const rule = rules.find((item) => item.axisId === axis.id);
      if (rule) {
        row[columns[ratingBucket(rule.rating)]] = 1;
        row[criterionIdColumns[axis.id]] = rule.criterionId;
      }
    }
    row[exportVersionColumn] = appVersion;
  });
  const message = skipped ? `${state.selectedRows.size - skipped} ligne(s) cotée(s), ${skipped} ignorée(s) car incomplète(s).` : `${state.selectedRows.size} ligne(s) cotée(s).`;
  resetQualificationForm();
  setStatus(message);
}

function quoteCurrentLine(rowIndex: number): void {
  state.selectedRows.clear();
  state.selectedRows.add(rowIndex);
  resetQualificationForm();
  setStatus('Ligne sélectionnée. Choisissez les critères puis appliquez la cotation.');
  document.querySelector('#file-qualification-title, #questionnaire-title')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function exportWorkbook(): void {
  if (!state.rows.length) {
    setStatus('Téléversez un fichier ODS avant export.');
    return;
  }
  try {
    const baseName = `${state.year || 'ANNEE'}-${state.community || 'NOM_DE_COLLECTIVITE'}`.replace(/[^\p{L}\p{N}_-]+/gu, '_');
    const workbook = XLSX.utils.book_new();
    const original = state.workbook?.Sheets[state.sheetName];
    const sheet: XLSX.WorkSheet = original ? structuredClone(original) : {};
    const updated = XLSX.utils.aoa_to_sheet(state.rows);
    for (const [address, cell] of Object.entries(updated)) {
      if (address.startsWith('!')) continue;
      if (sheet[address]?.v === cell.v && sheet[address]?.t === cell.t) continue;
      sheet[address] = { ...sheet[address], ...cell };
      delete sheet[address].f;
      delete sheet[address].w;
    }
    sheet['!ref'] = updated['!ref'];
    XLSX.utils.book_append_sheet(workbook, sheet, safeSheetName(state.year, state.community));
    XLSX.writeFile(workbook, `${baseName}-BBV.ods`, { bookType: 'ods', cellStyles: true });
    setStatus('Fichier ODS généré. Vérifiez son téléchargement dans votre navigateur.');
  } catch (error) {
    setStatus(`Export impossible : ${error instanceof Error ? error.message : String(error)}. Les données restent disponibles, vous pouvez réessayer.`);
  }
}

async function loadSpreadsheet(file: File): Promise<void> {
  if (state.loading) return;
  state.loading = true;
  setStatus(`Vérification de ${file.name}…`);
  try {
    if (!/\.ods$/i.test(file.name)) throw new Error('Sélectionnez un fichier au format ODS.');
    const buffer = await file.arrayBuffer();
    // SheetJS reconnaît aussi d'autres formats : vérifier le type ODS réel.
    const archive = XLSX.CFB.read(new Uint8Array(buffer), { type: 'array' });
    const mime = XLSX.CFB.find(archive, 'mimetype');
    if (!mime || new TextDecoder().decode(mime.content).trim() !== 'application/vnd.oasis.opendocument.spreadsheet') {
      throw new Error('Le contenu du fichier ne correspond pas à un classeur ODS.');
    }
    const workbook = XLSX.read(buffer, { type: 'array', cellStyles: true });
    if (workbook.SheetNames.length !== 1) throw new Error(`Un seul onglet est attendu ; ${workbook.SheetNames.length} détecté(s).`);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json<Row>(sheet, { header: 1, defval: '', blankrows: true });
    const issues = validateHeaders(rows[0] ?? []);
    if (issues.length) throw new Error(`Structure incompatible : ${issues.slice(0, 5).join(' ; ')}${issues.length > 5 ? ` ; et ${issues.length - 5} autre(s) erreur(s)` : ''}. Utilisez le modèle fourni.`);
    const invalidAmounts: number[] = [];
    rows.slice(1).forEach((row, offset) => {
      if (row.slice(85).some((value) => normalize(value))) throw new Error(`Ligne ${offset + 2} : données au-delà des 85 colonnes du modèle.`);
      if (normalize(row[10]) && parseAmount(row[10]) === null) invalidAmounts.push(offset + 2);
    });
    if (invalidAmounts.length) throw new Error(`Montant invalide à la/aux ligne(s) ${invalidAmounts.slice(0, 10).join(', ')}${invalidAmounts.length > 10 ? '…' : ''}. Corrigez ces montants dans le tableur avant import.`);
    // Ignore les lignes entièrement vides, mais conserve les lignes incomplètes.
    const expenses = rows.slice(1).filter((row) => row.some((value) => normalize(value)));
    if (!expenses.length) throw new Error('Le fichier ne contient aucune ligne de dépense.');
    const importedRows = [rows[0], ...expenses];
    // Les valeurs acceptées sont converties en nombres pour assurer un aller-retour cohérent.
    importedRows.slice(1).forEach((row) => {
      if (normalize(row[10])) row[10] = parseAmount(row[10]);
    });
    // Validation terminée : remplacer l'état seulement à partir d'ici.
    state.workbook = workbook;
    state.sheetName = sheetName;
    // La suppression des lignes vides modifie les coordonnées du classeur source.
    // Reconstruire dans ce cas pour ne pas conserver des cellules décalées.
    if (expenses.length !== rows.length - 1) state.workbook = null;
    state.rows = importedRows;
    state.columnLabels = rows[0].map((value) => normalize(value));
    state.selectedRows.clear();
    state.filters = Array(11).fill('');
    state.filterRegex = Array(11).fill(false);
    state.statusFilter = 'all';
    resetQualificationForm();
    parseSheetIdentity(sheetName);
    const partial = expenses.filter((row) => rowStatus(row) === 'partial').length;
    const inconsistent = expenses.filter((row) => rowStatus(row) === 'inconsistent').length;
    state.status = `${expenses.length} dépense(s) chargée(s) depuis ${file.name}. ${partial} cotation(s) partielle(s), ${inconsistent} incohérente(s).`;
  } catch (error) {
    state.status = `Import refusé : ${error instanceof Error ? error.message : String(error)} ${state.rows.length ? 'Le fichier précédemment chargé et sa cotation sont conservés.' : 'Aucun fichier n’a été chargé.'}`;
  } finally {
    state.loading = false;
    render();
  }
}

function renderCriteria(rule: Rule, ruleIndex: number, axisId: string): string {
  return `
    <label class="criterion-card">
      <input type="radio" name="axis-${axisId}" value="${ruleIndex}" ${state.selectedRuleByAxis[axisId] === ruleIndex ? 'checked' : ''} />
      <span>
        <span class="criterion-meta"><span class="criterion-id">${escapeHtml(rule.criterionId)}</span><strong class="pill ${ratingClass(rule.rating)}">${escapeHtml(rule.rating)}</strong></span>
        <span class="criterion-text">${escapeHtmlPreserveLines(rule.criteria)}</span>
      </span>
    </label>
  `;
}

function renderQuestionnaireContent(): string {
  const types = state.domain ? data.typesByDomain[state.domain] ?? [] : [];
  const rules = currentRules();
  return `
    <div class="form-grid">
      <label>Domaine
        <select id="domain-select">
          <option value="">Sélectionner un domaine</option>
          ${data.domains.map((domain) => `<option value="${escapeHtml(domain)}" ${domain === state.domain ? 'selected' : ''}>${escapeHtml(domain)}</option>`).join('')}
        </select>
      </label>
      <label>Type de dépense
        <select id="type-select" ${state.domain ? '' : 'disabled'}>
          <option value="">Sélectionner un type</option>
          ${types.map((type) => `<option value="${escapeHtml(type)}" ${type === state.expenseType ? 'selected' : ''}>${escapeHtml(type)}</option>`).join('')}
        </select>
      </label>
    </div>
    ${state.domain && state.expenseType ? data.axes.map((axis) => {
      const axisRules = rules.filter((rule) => rule.axisId === axis.id);
      return `
        <fieldset class="axis-block">
          <legend>${escapeHtml(axis.key)} <span>${escapeHtml(axis.label)}</span></legend>
          <div class="criteria-list">
            ${axisRules.map((rule, index) => renderCriteria(rule, index, axis.id)).join('')}
          </div>
        </fieldset>
      `;
    }).join('') : '<p class="muted">Choisissez un domaine et un type de dépense pour afficher les critères de cotation.</p>'}
  `;
}

function renderQuestionnaire(): string {
  return `
    <section class="panel" aria-labelledby="questionnaire-title">
      <div class="panel-heading">
        <p class="eyebrow">Outil 1</p>
        <h2 id="questionnaire-title">Qualifier une dépense</h2>
      </div>
      ${renderQuestionnaireContent()}
    </section>
  `;
}

function renderSummary(): string {
  const rules = selectedRules();
  return `
    <aside class="summary" aria-labelledby="summary-title">
      <h2 id="summary-title">Synthèse de cotation</h2>
      <p>${escapeHtml(state.domain || 'Domaine non sélectionné')}</p>
      <p>${escapeHtml(state.expenseType || 'Type de dépense non sélectionné')}</p>
      <div class="summary-list">
        ${data.axes.map((axis) => {
          const rule = rules.find((item) => item.axisId === axis.id);
          return `<div><span>${escapeHtml(`Axe ${Number(axis.id)} - ${axis.key}`)}</span><strong class="pill ${rule ? ratingClass(rule.rating) : 'unknown'}">${escapeHtml(rule?.rating ?? 'À renseigner')}</strong></div>`;
        }).join('')}
      </div>
    </aside>
  `;
}

function renderTable(): string {
  if (!state.rows.length) return '<p class="muted">Aucun fichier chargé. Téléversez un fichier ODS de dépenses pour visualiser les lignes.</p>';
  const rows = filteredRows();
  const columns = visibleTableColumns();
  return `
    ${renderColumnChooser()}
    <div class="filter-help tooltip-wrap">
      <button type="button" class="small tooltip-trigger" aria-describedby="regex-help">Filtres : aide sur les expressions régulières <span aria-hidden="true">ⓘ</span></button>
      <span class="tooltip" id="regex-help" role="tooltip">Par défaut, les filtres recherchent un texte simple. Activez le bouton .* d'une colonne pour utiliser une expression régulière, sans tenir compte des majuscules/minuscules.

Exemples (à saisir sans barres obliques) :
^21 : commence par 21
2026$ : se termine par 2026
travaux|rénovation : contient l'un des deux mots
^21[0-9]+$ : commence par 21, suivi uniquement de chiffres
\\. : recherche un point littéral

Les filtres de plusieurs colonnes se combinent avec un ET. Un champ vide ne filtre pas. Une expression invalide affiche une erreur et aucun résultat jusqu'à sa correction. « Vider les filtres de colonnes » efface les recherches et désactive les modes regex.</span>
    </div>
    <div class="status-filters" aria-label="Filtrer par statut">
      ${[
        ['all', 'Toutes'],
        ['unrated', 'À coter'],
        ['rated', 'Cotées'],
        ['partial', 'Cotations partielles'],
        ['inconsistent', 'Cotations incohérentes'],
        ['invalid', 'Incomplètes'],
      ].map(([value, label]) => `<button type="button" class="status-filter ${state.statusFilter === value ? 'active' : 'secondary'}" data-status-filter="${value}">${label}</button>`).join('')}
      <button id="clear-column-filters" type="button" class="secondary" ${hasActiveColumnFilters() ? '' : 'disabled'}>Vider les filtres de colonnes</button>
    </div>
    <div class="table-actions">
      <button id="select-visible" type="button">Sélectionner les lignes filtrées</button>
      <button id="clear-selection" type="button" class="secondary">Vider la sélection</button>
      <span>${state.selectedRows.size} ligne(s) sélectionnée(s), ${rows.length} visible(s)</span>
    </div>
    <div class="table-wrap" tabindex="0" aria-label="Dépenses chargées">
      <table>
        <thead>
          <tr>${columns.map((column) => `<th>${escapeHtml(columnLabel(column))}</th>`).join('')}</tr>
          <tr>${columns.map(renderColumnFilterCell).join('')}</tr>
        </thead>
        <tbody>
          ${rows.map(({ row, index }) => {
            const status = rowStatus(row);
            return `
              <tr class="row-${status}">
                ${columns.map((column) => renderColumnCell(column, row, index, status)).join('')}
              </tr>
            `;
          }).join('')}
        </tbody>
      </table>
    </div>
  `;
}

function renderDonut(axisAnalysis: AxisAnalysis): string {
  let currentAngle = 0;
  const slices = axisAnalysis.ratings
    .filter((rating) => rating.amount > 0 && rating.percent > 0)
    .map((rating) => {
      const startAngle = currentAngle;
      const endAngle = currentAngle + (rating.percent / 100) * 360;
      const textPoint = polarPoint(100, 72, startAngle + (endAngle - startAngle) / 2);
      currentAngle = endAngle;
      const fullCircle = rating.percent > 99.99;
      const path = fullCircle
        ? `${donutPath(0, 180, 90, 54)} ${donutPath(180, 359.99, 90, 54)}`
        : donutPath(startAngle, endAngle, 90, 54);
      return `
        <path d="${path}" fill="${rating.color}"></path>
        ${rating.percent >= 8 ? `<text x="${textPoint.x}" y="${textPoint.y}" class="donut-percent">${formatPercent(rating.percent)}%</text>` : ''}
      `;
    })
    .join('');

  return `
    <svg class="donut-chart" viewBox="0 0 200 200" role="img" aria-label="Répartition des cotations ${escapeHtml(axisAnalysis.axis.key)}">
      ${slices || '<circle cx="100" cy="100" r="90" fill="#eee4f4"></circle><circle cx="100" cy="100" r="54" fill="#fffdf7"></circle>'}
      <circle cx="100" cy="100" r="54" fill="#fffdf7"></circle>
      <text x="100" y="96" class="donut-total">${escapeHtml(formatDonutAmount(axisAnalysis.total))}</text>
      <text x="100" y="114" class="donut-caption">volume absolu</text>
    </svg>
  `;
}

function renderOperationBreakdown(rating: RatingAnalysis, axisTotal: number): string {
  if (!rating.amount) return '';
  return `
    <div class="operation-breakdown">
      <h4><span class="legend-dot" style="background:${rating.color}"></span>${escapeHtml(rating.label)} <small>${escapeHtml(formatEuro(rating.amount))} · ${formatPercent(rating.percent)}%</small></h4>
      <ul>
        ${rating.operations.map((operation) => {
          const percent = axisTotal > 0 ? (operation.amount / axisTotal) * 100 : 0;
          return `<li><span>${escapeHtml(operation.label)}</span><strong>${escapeHtml(formatEuro(operation.amount))}</strong><em>${formatPercent(percent)}%</em></li>`;
        }).join('')}
      </ul>
    </div>
  `;
}

function renderAnalysis(): string {
  const rows = selectedAnalysisRows();
  if (!rows.length) {
    return `
      <section class="panel analysis" aria-labelledby="analysis-title">
        <div class="panel-heading">
          <p class="eyebrow">Étape 2.3</p>
          <h2 id="analysis-title">Analyse</h2>
        </div>
        <p class="muted">${state.selectedRows.size ? `${state.selectedRows.size} ligne(s) sélectionnée(s) sans montant valide : aucune analyse disponible.` : "Sélectionnez les lignes pour afficher l'analyse."}</p>
      </section>
    `;
  }

  const ratedRows = rows.filter(rowHasRating);
  const totalAmount = rows.reduce((sum, row) => sum + amountForRow(row), 0);
  const ratedAmount = ratedRows.reduce((sum, row) => sum + amountForRow(row), 0);
  const positiveAmount = rows.reduce((sum, row) => sum + Math.max(0, amountForRow(row)), 0);
  const negativeAmount = rows.reduce((sum, row) => sum + Math.min(0, amountForRow(row)), 0);
  const excludedCount = state.selectedRows.size - rows.length;
  const partialCount = rows.filter((row) => rowStatus(row) === 'partial').length;
  const inconsistentCount = rows.filter((row) => rowStatus(row) === 'inconsistent').length;
  const axesAnalysis = data.axes.map((axis) => buildAxisAnalysis(rows, axis));
  const year = normalize(state.year) || 'Année non renseignée';
  const community = normalize(state.community) || 'Collectivité non renseignée';

  return `
    <section class="panel analysis" aria-labelledby="analysis-title">
      <div class="panel-heading analysis-heading">
        <div>
          <p class="eyebrow">Étape 2.3</p>
          <h2 id="analysis-title">Analyse</h2>
        </div>
        <button id="print-analysis" type="button" class="secondary">Exporter l'analyse PDF</button>
      </div>
      <div class="print-title">
        <h1>Analyse BBV</h1>
        <p>${escapeHtml(year)} · ${escapeHtml(community)}</p>
      </div>
      <div class="analysis-summary">
        <div><span>Année</span><strong>${escapeHtml(year)}</strong></div>
        <div><span>Collectivité</span><strong>${escapeHtml(community)}</strong></div>
        <div><span>Lignes sélectionnées</span><strong>${rows.length}</strong></div>
        <div><span>Cotations complètes</span><strong>${ratedRows.length}</strong></div>
        <div><span>Autres lignes</span><strong>${rows.length - ratedRows.length}</strong></div>
        <div><span>Cotations partielles</span><strong>${partialCount}</strong></div>
        <div><span>Cotations incohérentes</span><strong>${inconsistentCount}</strong></div>
        <div><span>Montant net sélectionné</span><strong>${escapeHtml(formatEuro(totalAmount))}</strong></div>
        <div><span>Montant net entièrement coté</span><strong>${escapeHtml(formatEuro(ratedAmount))}</strong></div>
        <div><span>Montants positifs</span><strong>${escapeHtml(formatEuro(positiveAmount))}</strong></div>
        <div><span>Montants négatifs</span><strong>${escapeHtml(formatEuro(negativeAmount))}</strong></div>
      </div>
      <p class="hint">Les graphiques et les opérations représentent les valeurs absolues des montants (positifs et négatifs), sans compensation. Les totaux nets sont affichés ci-dessus. Les axes incohérents sont regroupés dans « Non cotée ». ${excludedCount ? `${excludedCount} ligne(s) sans montant valide exclue(s) de l'analyse.` : ''}</p>
      <div class="donut-grid">
        ${axesAnalysis.map((axisAnalysis) => `
          <article class="donut-card">
            <h3>Axe ${Number(axisAnalysis.axis.id)} - ${escapeHtml(axisAnalysis.axis.key)}</h3>
            <p>${escapeHtml(axisAnalysis.axis.label)}</p>
            ${renderDonut(axisAnalysis)}
            <div class="donut-legend">
              ${axisAnalysis.ratings.map((rating) => `
                <div><span class="legend-dot" style="background:${rating.color}"></span><span>${escapeHtml(rating.label)}</span><strong>${formatPercent(rating.percent)}%</strong></div>
              `).join('')}
            </div>
            <div class="operation-groups">
              ${axisAnalysis.ratings.map((rating) => renderOperationBreakdown(rating, axisAnalysis.total)).join('') || '<p class="muted">Aucun montant à analyser pour cet axe.</p>'}
            </div>
          </article>
        `).join('')}
      </div>
    </section>
  `;
}

function renderFiles(): string {
  return `
    <section class="panel" aria-labelledby="files-title">
      <div class="panel-heading">
        <p class="eyebrow">Étape 2.1</p>
        <h2 id="files-title">Charger le fichier des dépenses</h2>
      </div>
      <p class="muted">Téléchargez le modèle ou téléversez un fichier ODS contenant un seul onglet de dépenses.</p>
      <div class="file-grid">
        <span class="tooltip-wrap template-help">
          <a id="download-template" class="button" href="${templateDownloadHref}" download="${templateFileName}" aria-describedby="template-columns-help">Télécharger le modèle ODS <span class="info-badge" aria-hidden="true">i</span></a>
          <span class="tooltip" id="template-columns-help" role="tooltip">Les colonnes rouges ne doivent pas être renommées. Vous pouvez renommer les autres colonnes, mais vous ne devez supprimer aucune colonne. L'ordre et le nombre de colonnes doivent rester inchangés pour permettre la qualification et l'export du fichier.</span>
        </span>
        <label class="file-input secondary">Téléverser un fichier ODS
          <input id="ods-upload" type="file" accept=".ods,application/vnd.oasis.opendocument.spreadsheet" ${state.loading ? 'disabled' : ''} />
        </label>
      </div>
      <div class="form-grid compact">
        <label>Année
          <input id="year-input" value="${escapeHtml(state.year)}" placeholder="2026" />
        </label>
        <label>Collectivité
          <input id="community-input" value="${escapeHtml(state.community)}" placeholder="NOM_DE_COLLECTIVITE" />
        </label>
      </div>
      <p class="status" role="status">${escapeHtml(state.status)}</p>
    </section>
    <section class="panel" aria-labelledby="file-qualification-title">
      <div class="panel-heading">
        <p class="eyebrow">Étape 2.2</p>
        <h2 id="file-qualification-title">Qualifier chaque ligne</h2>
      </div>
      <p class="muted">Visualisez les lignes de dépenses, choisissez une cotation, puis appliquez-la à une ou plusieurs lignes sélectionnées.</p>
      ${renderTable()}
      ${renderQuestionnaireContent()}
      <div class="apply-bar">
        <button id="apply-qualification" type="button" ${canApplyQualification() && state.selectedRows.size ? '' : 'disabled'}>Appliquer la cotation aux lignes sélectionnées</button>
        <button id="export-ods" type="button" class="secondary" ${state.rows.length ? '' : 'disabled'}>Télécharger le fichier coté</button>
      </div>
      <p class="hint">${escapeHtml(applyDisabledReason() || 'Prêt à appliquer la cotation aux lignes sélectionnées.')}</p>
    </section>
    ${renderAnalysis()}
  `;
}

function renderToolChooser(): string {
  return `
    <nav class="tool-chooser" aria-label="Choisir un outil">
      <button type="button" class="tool-card ${state.activeTool === 'single' ? 'active' : ''}" data-tool="single" aria-pressed="${state.activeTool === 'single'}">
        <span class="eyebrow">Outil 1</span>
        <strong>Qualifier une dépense</strong>
        <span>Obtenir une cotation pour une dépense isolée.</span>
      </button>
      <button type="button" class="tool-card ${state.activeTool === 'file' ? 'active' : ''}" data-tool="file" aria-pressed="${state.activeTool === 'file'}">
        <span class="eyebrow">Outil 2</span>
        <strong>Qualifier un fichier de dépenses</strong>
        <span>Charger un fichier, visualiser les lignes, les qualifier et les analyser.</span>
      </button>
    </nav>
  `;
}

function renderHelpPage(): string {
  return `
    <section class="help-page" aria-labelledby="help-title">
      <div class="help-hero panel">
        <div>
          <p class="eyebrow">Guide utilisateur</p>
          <h2 id="help-title">Comment ça marche ?</h2>
          <p>La Boussole Budget Vert est un outil d'aide à l'analyse des données budgétaires des collectivités au regard de l'impact environnemental. Il est composé de deux fonctions : qualifier une dépense ou qualifier et analyser un fichier de dépenses.</p>
        </div>
        <div class="help-actions" aria-label="Accès rapides">
          <button type="button" data-tool="single">Qualifier une dépense</button>
          <button type="button" class="secondary" data-tool="file">Qualifier un fichier</button>
        </div>
      </div>

      <section class="panel" aria-labelledby="uses-title">
        <div class="panel-heading">
          <p class="eyebrow">Deux usages possibles</p>
          <h2 id="uses-title">Choisir le bon parcours</h2>
        </div>
        <div class="use-grid">
          <article class="use-card">
            <span class="use-number">1</span>
            <h3>Qualifier une dépense</h3>
            <p>Qualifier une dépense en sélectionnant un domaine et un type de dépense par l'intermédiaire d'un arbre de décision suivant les 6 axes.</p>
          </article>
          <article class="use-card">
            <span class="use-number">2</span>
            <h3>Qualifier et analyser un fichier</h3>
            <p>Qualifier et analyser un fichier de dépenses en téléversant un fichier ODS, puis en appliquant une cotation à une ou plusieurs lignes sélectionnées.</p>
          </article>
        </div>
      </section>

      <section class="panel" aria-labelledby="axes-title">
        <div class="panel-heading">
          <p class="eyebrow">Les 6 axes environnementaux</p>
          <h2 id="axes-title">Une lecture multicritère des dépenses</h2>
        </div>
        <div class="axis-guide-grid">
          ${data.axes.map((axis) => `
            <article class="axis-guide-card">
              <span>Axe ${Number(axis.id)}</span>
              <strong>${escapeHtml(axis.key)}</strong>
              <p>${escapeHtml(axis.label)}</p>
            </article>
          `).join('')}
        </div>
      </section>

      <section class="panel" aria-labelledby="steps-title">
        <div class="panel-heading">
          <p class="eyebrow">Parcours fichier</p>
          <h2 id="steps-title">Qualifier un fichier en 7 étapes</h2>
        </div>
        <ol class="step-path">
          <li><span>1</span><p>Télécharger le modèle ODS.</p></li>
          <li><span>2</span><p>Remplir uniquement les lignes des colonnes bleues.</p></li>
          <li><span>3</span><p>Téléverser le fichier sur la page.</p></li>
          <li><span>4</span><p>Qualifier une ou plusieurs lignes en sélectionnant la cotation pour chaque axe, puis cliquer sur le bouton <strong>Appliquer les cotations aux lignes sélectionnées</strong>.</p></li>
          <li><span>5</span><p>Analyser toutes les lignes ou seulement celles d'un service en faisant une sélection. Les filtres distinguent les cotations complètes, partielles et incohérentes. Les totaux sont nets ; les graphiques utilisent les valeurs absolues des montants pour représenter aussi les régularisations négatives.</p></li>
          <li><span>6</span><p>Imprimer l'analyse.</p></li>
          <li><span>7</span><p>Cliquer sur <strong>Télécharger le fichier coté</strong>.</p></li>
        </ol>
      </section>

      <section class="privacy-callout" aria-labelledby="privacy-title">
        <div>
          <p class="eyebrow">Données personnelles</p>
          <h2 id="privacy-title">Vos données restent locales</h2>
          <p>Attention ! Aucune donnée n'est transmise au serveur, elle reste en local dans votre navigateur. Avant de fermer l'onglet, il est important d'exporter le fichier pour l'importer lors de la prochaine connexion.</p>
        </div>
      </section>

      <section class="panel" aria-labelledby="resources-title">
        <div class="panel-heading">
          <p class="eyebrow">Ressources officielles</p>
          <h2 id="resources-title">Comprendre le cadre du Budget Vert</h2>
        </div>
        <p>La France a mis en place depuis 2021 une « budgétisation verte » qui classe les dépenses budgétaires et fiscales selon leur impact environnemental, afin de mieux intégrer les enjeux écologiques dans les politiques publiques. Les collectivités territoriales jouent un rôle central dans la transition écologique puisqu'elles gèrent environ 20 % du parc immobilier public et supportent près de 60 % de l'investissement public civil.</p>
        <p>La loi de finances pour 2024 introduit une « annexe environnementale des collectivités locales » qui mesure l'impact des budgets locaux sur la transition écologique. Cette annexe permet de valoriser les investissements des collectivités qui ont un impact positif sur l'environnement. Cet outil facilite ainsi la planification écologique à l'échelle du territoire national.</p>
        <p>Cet outil d'aide à la décision est réalisé à partir des guides méthodologiques consultables à l'adresse suivante.</p>
        <p>Il existe également une foire aux questions.</p>
        <div class="resource-links">
          <a class="button" href="https://www.collectivites-locales.gouv.fr/gerer-les-finances-publiques-locales/budget/budget-vert-des-collectivites" target="_blank" rel="noopener noreferrer">Guides méthodologiques officiels</a>
          <a class="button secondary-link" href="https://www.collectivites-locales.gouv.fr/FAQ/budget-vert-des-collectivites-foire-aux-questions" target="_blank" rel="noopener noreferrer">Foire aux questions</a>
        </div>
      </section>

      <section class="panel libre-section" aria-labelledby="license-title">
        <div class="panel-heading">
          <p class="eyebrow">Licence et logiciel libre</p>
          <h2 id="license-title">La licence de Boussole Budget Vert</h2>
        </div>
        <p>La licence de cette application est <a href="https://polyformproject.org/licenses/noncommercial/1.0.0" target="_blank" rel="noopener noreferrer">PolyForm Noncommercial 1.0.0</a>. Toute utilisation commerciale nécessite l'obtention préalable d'une licence commerciale distincte auprès de l'auteur.</p>
        <p>La licence PolyForm Noncommercial n'est pas une licence libre, car elle restreint les usages commerciaux. Le code source est toutefois accessible afin de permettre à chacun de l'examiner et d'en vérifier le fonctionnement.</p>
        <h3>Qu'est-ce qu'un logiciel libre ?</h3>
        <p>Le logiciel libre a été créé au début des années 80 par Richard Stallman, un informaticien américain, qui cherchait à réparer son imprimante dont le fabricant refusait de partager le logiciel permettant de l'améliorer.</p>
        <p>Le logiciel libre respecte <strong>quatre libertés</strong> :</p>
        <ul class="free-liberties">
          <li>la liberté de faire <strong>fonctionner</strong> le programme comme vous voulez, pour n'importe quel usage ;</li>
          <li>la liberté d'<strong>étudier</strong> le fonctionnement du programme, et de le modifier pour qu'il effectue vos tâches informatiques comme vous le souhaitez ; l'accès au code source est une condition nécessaire ;</li>
          <li>la liberté de <strong>redistribuer</strong> des copies, donc d'aider les autres ;</li>
          <li>la liberté de <strong>distribuer</strong> aux autres des <strong>copies de vos versions modifiées</strong> ; en faisant cela, vous donnez à toute la communauté une possibilité de profiter de vos changements ; l'accès au code source est une condition nécessaire.</li>
        </ul>
        <h3>Pourquoi ?</h3>
        <div class="video-card">
          <video controls preload="metadata" src="https://peertube.lyceeconnecte.fr/download/web-videos/ff570cd0-4dbb-4ed2-9d01-85445051b79c-1080.mp4"></video>
          <p>Vidéo - Pourquoi utiliser un logiciel libre ? Créée par Free Software Foundation Europe (FSFE).</p>
        </div>
        <div class="free-grid">
          <article><h3>Interopérabilité et pérennité</h3><p>Ne pas être prisonnier d'un format de fichier ou d'un logiciel. Exemple : on ne peut pas répondre à une personne qui partage votre vidéo YouTube sur Instagram, et pourtant on peut tous s'envoyer des mails sans avoir le même gestionnaire.</p></article>
          <article><h3>Protection de la vie privée</h3><p>La vie privée est un enjeu collectif et ne concerne pas seulement les usages individuels.</p></article>
          <article><h3>Indépendance technologique</h3><p>L'argent public doit être utilisé pour créer des communs numériques.</p></article>
          <article><h3>Sécurité renforcée</h3><p>Le code source est ouvert, ce qui permet une vérification constante par la communauté.</p></article>
          <article><h3>Support communautaire</h3><p>Une large communauté d'utilisateurs et de développeurs peut offrir un support. Les administrations de la France, des Pays-Bas et de l'Allemagne collaborent déjà ensemble sur plusieurs logiciels de <a href="https://lasuite.numerique.gouv.fr/" target="_blank" rel="noopener noreferrer">La Suite Numérique</a>.</p></article>
        </div>
      </section>
    </section>
  `;
}

function renderActiveTool(): string {
  if (state.activeTool === 'single') return renderQuestionnaire();
  if (state.activeTool === 'file') return renderFiles();
  return renderHelpPage();
}

function renderMainContent(): string {
  if (state.activeTool === 'help') return renderActiveTool();
  return `
    <div class="layout">
      <div>${renderActiveTool()}</div>
      ${renderSummary()}
    </div>
  `;
}

function renderFooter(): string {
  return `
    <footer class="site-footer">
      <div class="footer-brand">
        <strong>Boussole Budget Vert</strong>
        <span>Version ${appVersion}</span>
      </div>
      <p>Crédits : Boussole Budget Vert</p>
      <p class="disclaimer">Cet outil a été conçu à partir de l’analyse des guides méthodologiques officiels relatifs au Budget Vert. Malgré le soin apporté à sa réalisation, des erreurs ou omissions peuvent subsister. Les guides méthodologiques officiels demeurent les seules références exhaustives et opposables. L’utilisateur reste seul responsable de la qualification des dépenses qu’il réalise à l’aide de cet outil.</p>
      <p class="license-line">
        <span>Licence :</span>
        <a href="https://polyformproject.org/licenses/noncommercial/1.0.0" target="_blank" rel="noopener noreferrer" aria-label="Licence PolyForm Noncommercial 1.0.0">
          PolyForm Noncommercial 1.0.0
        </a>
      </p>
    </footer>
  `;
}

function render(): void {
  if (!app) return;
  app.innerHTML = `
    <header class="hero">
      <div class="header-actions">
        <button type="button" class="header-link" data-tool="help">Comment ça marche ?</button>
        <a class="header-link" href="${sourceCodeUrl}" target="_blank" rel="noopener noreferrer" aria-label="Voir le code source du projet sur GitHub">
          <span aria-hidden="true">GitHub</span>
          Code source
        </a>
      </div>
      <div class="hero-copy">
        <p class="eyebrow app-version">Version ${appVersion}</p>
        <h1>Boussole Budget Vert</h1>
        <p>Outil d'aide à l'analyse des données budgétaires des collectivités au regard de l'impact environnemental.</p>
      </div>
      <a class="afigese-card" href="https://afigese.fr/" target="_blank" rel="noopener noreferrer" aria-label="Outil créé par l’Afigese — visiter afigese.fr (nouvel onglet)">
        <img src="${afigeseLogo}" alt="Afigese" width="1536" height="561" />
        <span>Outil créé par l’Afigese</span>
      </a>
    </header>
    <main>
      ${renderToolChooser()}
      ${renderMainContent()}
    </main>
    ${renderFooter()}
  `;
  bindEvents();
}

function renderAndRestoreFilter(column: number, caret: number | null): void {
  render();
  requestAnimationFrame(() => {
    const input = document.querySelector<HTMLInputElement>(`.filter[data-filter="${column}"]`);
    input?.focus();
    if (caret !== null) input?.setSelectionRange(caret, caret);
  });
}

function bindEvents(): void {
  document.querySelectorAll<HTMLButtonElement>('[data-tool]').forEach((button) => {
    button.addEventListener('click', () => {
      state.activeTool = (button.dataset.tool as typeof state.activeTool) || 'single';
      render();
    });
  });
  document.querySelector<HTMLSelectElement>('#domain-select')?.addEventListener('change', (event) => {
    state.domain = (event.target as HTMLSelectElement).value;
    state.expenseType = '';
    state.selectedRuleByAxis = {};
    render();
  });
  document.querySelector<HTMLSelectElement>('#type-select')?.addEventListener('change', (event) => {
    state.expenseType = (event.target as HTMLSelectElement).value;
    state.selectedRuleByAxis = {};
    render();
  });
  document.querySelectorAll<HTMLInputElement>('input[type="radio"][name^="axis-"]').forEach((input) => {
    input.addEventListener('change', () => {
      const axisId = input.name.replace('axis-', '');
      state.selectedRuleByAxis[axisId] = Number(input.value);
      render();
    });
  });
  document.querySelectorAll<HTMLButtonElement>('.tooltip-trigger').forEach((button) => {
    button.addEventListener('click', (event) => {
      event.preventDefault();
      event.stopPropagation();
      button.focus();
    });
  });
  document.querySelector<HTMLInputElement>('#ods-upload')?.addEventListener('change', (event) => {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (file) void loadSpreadsheet(file);
  });
  document.querySelectorAll<HTMLInputElement>('.filter').forEach((input) => {
    input.addEventListener('input', () => {
      const column = Number(input.dataset.filter);
      state.filters[column] = input.value;
      renderAndRestoreFilter(column, input.selectionStart);
    });
  });
  document.querySelector<HTMLButtonElement>('#clear-column-filters')?.addEventListener('click', () => {
    state.filters = Array(11).fill('');
    state.filterRegex = Array(11).fill(false);
    render();
  });
  document.querySelectorAll<HTMLButtonElement>('[data-regex-toggle]').forEach((button) => {
    button.addEventListener('click', () => {
      const column = Number(button.dataset.regexToggle);
      state.filterRegex[column] = !state.filterRegex[column];
      render();
      document.querySelector<HTMLButtonElement>(`[data-regex-toggle="${column}"]`)?.focus();
    });
  });
  document.querySelectorAll<HTMLButtonElement>('[data-status-filter]').forEach((button) => {
    button.addEventListener('click', () => {
      state.statusFilter = (button.dataset.statusFilter as typeof state.statusFilter) || 'all';
      render();
    });
  });
  document.querySelectorAll<HTMLInputElement>('[data-column-toggle]').forEach((input) => {
    input.addEventListener('change', () => {
      const columnId = input.dataset.columnToggle;
      if (!columnId) return;
      if (input.checked) state.visibleColumns.add(columnId);
      else state.visibleColumns.delete(columnId);
      if (!state.visibleColumns.size) state.visibleColumns.add('action');
      saveVisibleColumns();
      render();
    });
  });
  document.querySelector<HTMLDetailsElement>('.column-chooser')?.addEventListener('toggle', (event) => {
    state.columnChooserOpen = (event.currentTarget as HTMLDetailsElement).open;
    saveColumnChooserOpen();
  });
  document.querySelector<HTMLButtonElement>('#show-all-columns')?.addEventListener('click', () => {
    setVisibleColumns(tableColumns.map((column) => column.id));
  });
  document.querySelectorAll<HTMLInputElement>('input[data-row]').forEach((input) => {
    input.addEventListener('change', () => {
      const rowIndex = Number(input.dataset.row);
      if (input.checked) state.selectedRows.add(rowIndex);
      else state.selectedRows.delete(rowIndex);
      resetQualificationForm();
      render();
    });
  });
  document.querySelector<HTMLButtonElement>('#select-visible')?.addEventListener('click', () => {
    filteredRows().forEach(({ index }) => state.selectedRows.add(index));
    resetQualificationForm();
    render();
  });
  document.querySelector<HTMLButtonElement>('#clear-selection')?.addEventListener('click', () => {
    state.selectedRows.clear();
    resetQualificationForm();
    render();
  });
  document.querySelectorAll<HTMLButtonElement>('[data-quote-row]').forEach((button) => {
    button.addEventListener('click', () => quoteCurrentLine(Number(button.dataset.quoteRow)));
  });
  document.querySelector<HTMLInputElement>('#year-input')?.addEventListener('input', (event) => {
    state.year = (event.target as HTMLInputElement).value;
  });
  document.querySelector<HTMLInputElement>('#community-input')?.addEventListener('input', (event) => {
    state.community = (event.target as HTMLInputElement).value;
  });
  document.querySelector<HTMLButtonElement>('#print-analysis')?.addEventListener('click', () => window.print());
  document.querySelector<HTMLButtonElement>('#apply-qualification')?.addEventListener('click', applyQualification);
  document.querySelector<HTMLButtonElement>('#export-ods')?.addEventListener('click', exportWorkbook);
}

render();
