import { normalize } from './spreadsheet';

export type CompiledFilter = { matches: (value: unknown) => boolean; error: string };

export function compileFilter(text: string, regexMode: boolean): CompiledFilter {
  if (regexMode && text !== '') {
    try {
      const expression = new RegExp(text, 'i');
      return { matches: (value) => expression.test(normalize(value)), error: '' };
    } catch {
      return { matches: () => false, error: 'Expression régulière invalide. Vérifiez les parenthèses, crochets et caractères spéciaux.' };
    }
  }
  const search = normalize(text).toLocaleLowerCase('fr-FR');
  return { matches: (value) => normalize(value).toLocaleLowerCase('fr-FR').includes(search), error: '' };
}
