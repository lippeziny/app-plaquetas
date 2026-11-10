/** Shared text helpers kept separate from the legacy XLSX parser so the mobile app
 * does not bundle the full SheetJS runtime just to normalize search strings. */
export function normalizeSearch(value: string): string {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}
