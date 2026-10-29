export type Row = Array<string | number | boolean | null | undefined>;
export type RowStatus = 'rated' | 'partial' | 'inconsistent' | 'unrated' | 'invalid';
export type RatingBucket = 'favorable' | 'favorableCondition' | 'neutral' | 'unfavorable' | 'unrated' | 'outOfScope';

export const appVersion = '1.2';
export const ratingColumns: Record<string, Record<RatingBucket, number>> = Object.fromEntries(
  ['01', '02', '03', '04', '05', '06'].map((id, index) => [id, {
    favorable: 13 + index * 6, favorableCondition: 14 + index * 6,
    neutral: 15 + index * 6, unfavorable: 16 + index * 6,
    unrated: 17 + index * 6, outOfScope: 18 + index * 6,
  }]),
);
export const criterionIdColumns: Record<string, number> = { '01': 78, '02': 79, '03': 80, '04': 81, '05': 82, '06': 83 };
export const exportVersionColumn = 84;

export function normalize(value: unknown): string {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

// Accepte les conventions françaises et internationales sans deviner les
// séparateurs ambigus (par exemple « 1,234 »). Une virgule seule est décimale.
export function parseAmount(value: unknown): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (typeof value !== 'string') return null;
  let text = value.trim().replace(/\u2212/g, '-');
  let negative = false;
  if (text.startsWith('(') && text.endsWith(')')) {
    negative = true;
    text = text.slice(1, -1).trim();
  }
  text = text.replace(/^€\s*|\s*€$/g, '').trim();
  if (negative && /^[+-]/.test(text)) return null;
  // Les espaces ne sont autorisés que comme séparateurs de milliers.
  if (/\s/.test(text)) {
    if (!/^[+-]?\d{1,3}(?:[\s\u00a0\u202f]\d{3})+(?:[.,]\d+)?$/.test(text)) return null;
    text = text.replace(/\s/g, '');
  }
  if (text.includes(',') && text.includes('.')) {
    if (/^[+-]?\d{1,3}(?:,\d{3})+\.\d+$/.test(text)) text = text.replace(/,/g, '');
    else if (/^[+-]?\d{1,3}(?:\.\d{3})+,\d+$/.test(text)) text = text.replace(/\./g, '').replace(',', '.');
    else return null;
  } else {
    text = text.replace(',', '.');
  }
  if (!/^[+-]?\d+(?:\.\d+)?$/.test(text)) return null;
  const amount = Number(text) * (negative ? -1 : 1);
  return Number.isFinite(amount) ? amount : null;
}

export function requiredIssues(row: Row): string[] {
  const issues: string[] = [];
  for (const [index, label] of [[0, 'Nature'], [1, 'Rubrique'], [9, 'Libellé'], [10, 'Montant']] as const) {
    if (!normalize(row[index])) issues.push(`${label} manquant`);
  }
  if (normalize(row[10]) && parseAmount(row[10]) === null) issues.push('Montant invalide');
  return issues;
}

export function qualificationIssues(row: Row): string[] {
  const issues: string[] = [];
  for (const [axis, columns] of Object.entries(ratingColumns)) {
    const values = Object.values(columns).map((column) => normalize(row[column]));
    if (values.some((value) => !['', '0', '1'].includes(value))) issues.push(`Axe ${axis} : indicateur différent de 0 ou 1`);
    if (values.filter((value) => value === '1').length > 1) issues.push(`Axe ${axis} : plusieurs cotations`);
  }
  return issues;
}

export function qualificationStatus(row: Row): RowStatus {
  if (requiredIssues(row).length) return 'invalid';
  if (qualificationIssues(row).length) return 'inconsistent';
  const axisCount = Object.values(ratingColumns).filter((columns) => Object.values(columns).some((column) => normalize(row[column]) === '1')).length;
  if (normalize(row[11]) && normalize(row[12]) && axisCount === 6) return 'rated';
  const hasMetadata = normalize(row[11]) || normalize(row[12]) || Object.values(criterionIdColumns).some((column) => normalize(row[column]));
  return axisCount || hasMetadata ? 'partial' : 'unrated';
}

export function validateHeaders(header: Row): string[] {
  const expected: Record<number, string> = {
    0: 'Nature', 1: 'Rubrique', 9: 'Libellé', 10: 'Montant de la réalisation',
    84: 'VERSION_BBV',
  };
  const suffixes = ['FAVORABLE', 'FAVORABLE_SS_COND', 'NEUTRE', 'DEFAVORABLE', 'NONCOTE', 'HORS_PERIMETRE'];
  for (const [axis, columns] of Object.entries(ratingColumns)) {
    Object.values(columns).forEach((column, index) => { expected[column] = `${axis}_${suffixes[index]}`; });
    expected[criterionIdColumns[axis]] = `${axis}_ID_CRITERE`;
  }
  const issues = Object.entries(expected)
    .filter(([index, label]) => normalize(header[Number(index)]) !== label)
    .map(([index, label]) => `Colonne ${Number(index) + 1} : en-tête attendu « ${label} »`);
  if (header.slice(85).some((value) => normalize(value))) issues.push('Colonnes supplémentaires après VERSION_BBV');
  return issues;
}

export function safeSheetName(year: string, community: string): string {
  return `${year || 'ANNEE'}-${community || 'NOM_DE_COLLECTIVITE'}`
    .replace(/[\\/?:*\[\]\u0000-\u001f]/g, '_').replace(/^'+|'+$/g, '').slice(0, 31).replace(/'+$/g, '') || 'BBV';
}
