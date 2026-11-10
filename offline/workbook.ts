import * as XLSX from 'xlsx';

export type CatalogItem = { code: string; name: string; quantity?: string; sheetName?: string };
export type WorkbookPreview = {
  headers: string[];
  rows: string[][];
  columns: { index: number; label: string }[];
  suggested: { codeColumn: number; nameColumn: number; quantityColumn: number };
};
export type ParsedWorkbook = { sheetNames: string[]; workbook?: XLSX.WorkBook; sheets?: Record<string, string[][]> };

function workbookRows(workbook: ParsedWorkbook, sheetName: string): string[][] {
  if (workbook.workbook) {
    const worksheet = workbook.workbook.Sheets[sheetName];
    if (!worksheet) throw new Error(`A aba "${sheetName}" não foi localizada na planilha.`);
    const usedRange = worksheet['!ref'];
    if (usedRange) {
      const range = XLSX.utils.decode_range(usedRange);
      if (range.e.r > 250000) throw new Error('A aba selecionada ultrapassa 250 mil linhas. Filtre ou divida a planilha antes de importar no celular.');
    }
    const rows = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, raw: false, defval: '' });
    return rows.map(row => row.map(cellText));
  }
  const rows = workbook.sheets?.[sheetName];
  if (!rows) throw new Error(`A aba "${sheetName}" não foi localizada na planilha.`);
  return rows;
}

function cellText(value: unknown): string {
  if (value == null) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'object') {
    const item = value as Record<string, unknown>;
    if ('result' in item) return cellText(item.result);
    if (Array.isArray(item.richText)) return item.richText.map(part => String((part as any).text || '')).join('');
    if (item.text) return String(item.text);
    if (item.hyperlink) return String(item.text || item.hyperlink);
  }
  return String(value).trim();
}
function normalize(value: string): string {
  return value.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
}
function suggest(headers: string[], terms: string[], fallback: number): number {
  const index = headers.findIndex(header => terms.some(term => normalize(header).includes(term)));
  return index >= 0 ? index + 1 : fallback;
}

export function parseWorkbookBytes(bytes: Uint8Array): ParsedWorkbook {
  if (!bytes?.byteLength) throw new Error('O arquivo Excel está vazio ou não pôde ser lido.');
  // Lê bytes diretamente, evitando uma string Base64 que duplica o consumo de memória.
  const workbook = XLSX.read(bytes, { type: 'array', cellDates: true, dense: true });
  if (!workbook.SheetNames?.length) throw new Error('Nenhuma aba foi encontrada no arquivo Excel.');
  // Mantém o workbook e converte somente a aba selecionada pelo usuário.
  return { sheetNames: [...workbook.SheetNames], workbook };
}

export function parseWorkbookBase64(base64: string): ParsedWorkbook {
  if (!base64) throw new Error('O arquivo Excel está vazio ou não pôde ser lido.');
  const workbook = XLSX.read(base64, { type: 'base64', cellDates: true, dense: true });
  if (!workbook.SheetNames?.length) throw new Error('Nenhuma aba foi encontrada no arquivo Excel.');
  return { sheetNames: [...workbook.SheetNames], workbook };
}

export function previewWorkbookSheet(workbook: ParsedWorkbook, sheetName: string, limit = 5): WorkbookPreview {
  const rows = workbookRows(workbook, sheetName);
  const headers = (rows[0] || []).map((value, index) => value || `Coluna ${index + 1}`);
  const sample = rows.slice(1).filter(row => row.some(Boolean)).slice(0, limit);
  return {
    headers,
    rows: sample,
    columns: headers.map((label, index) => ({ index: index + 1, label })),
    suggested: {
      codeColumn: suggest(headers, ['n do item', 'no item', 'codigo', 'material', 'item'], 1),
      nameColumn: suggest(headers, ['descricao', 'nome', 'item'], 2),
      quantityColumn: suggest(headers, ['em estoque', 'estoque', 'quantidade', 'qde'], 6),
    },
  };
}

export function importWorkbookSheet(
  workbook: ParsedWorkbook,
  sheetName: string,
  mapping: { codeColumn: number; nameColumn: number; quantityColumn: number },
  sourceFile: string,
): { items: CatalogItem[]; skipped: number } {
  const rows = workbookRows(workbook, sheetName);
  const codeIndex = Number(mapping.codeColumn) - 1;
  const nameIndex = Number(mapping.nameColumn) - 1;
  const quantityIndex = Number(mapping.quantityColumn) - 1;
  if (codeIndex < 0 || nameIndex < 0) throw new Error('Selecione as colunas de código e nome antes de importar.');
  const items: CatalogItem[] = [];
  let skipped = 0;
  for (const row of rows.slice(1)) {
    const code = cellText(row[codeIndex]).replace(/\s+/g, ' ').trim();
    const name = cellText(row[nameIndex]).replace(/\s+/g, ' ').trim();
    const quantity = mapping.quantityColumn > 0 ? cellText(row[quantityIndex]) : '';
    if (!code || !name) { skipped++; continue; }
    items.push({ code, name, quantity, sheetName });
    if (items.length > 200000) throw new Error('A planilha tem mais de 200.000 linhas válidas. Divida a base em arquivos menores para importar no celular.');
  }
  return { items, skipped };
}

export function normalizeSearch(value: string): string {
  return normalize(value.trim());
}
