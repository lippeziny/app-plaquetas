import { SaxesParser } from 'saxes';
import type { RawImportRow, SharedStringRow } from './catalog-db';

export type SheetReference = { name: string; relationId: string; path?: string };
export type XmlCell = { value: string; type: string; style: string };
export type XmlPreview = {
  rows: Array<{ rowNumber: number; cells: Map<number, XmlCell> }>;
  totalRows?: number;
  maxColumn: number;
  sharedIds: Set<number>;
};

const BUILTIN_NUMBER_FORMATS: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
  18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
  22: 'm/d/yy h:mm', 37: '#,##0;(#,##0)', 38: '#,##0;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  49: '@',
};

export function parseStylesXml(xml: string): string[] {
  const parser = new SaxesParser({ xmlns: false });
  const custom = new Map<number, string>();
  const formats: string[] = [];
  let inCellXfs = false;
  parser.on('opentag', (tag: any) => {
    const a = attrs(tag);
    if (tag.name === 'numFmt' && a.numFmtId !== undefined && a.formatCode !== undefined) {
      custom.set(Number(a.numFmtId), a.formatCode);
    } else if (tag.name === 'cellXfs') {
      inCellXfs = true;
    } else if (tag.name === 'xf' && inCellXfs) {
      const id = Number(a.numFmtId || 0);
      formats.push(custom.get(id) || BUILTIN_NUMBER_FORMATS[id] || 'General');
    }
  });
  parser.on('closetag', (tag: any) => { if (closeTagName(tag) === 'cellXfs') inCellXfs = false; });
  parser.write(xml).close();
  return formats;
}

export function formatExcelValue(value: string, type: string, format = 'General'): string {
  const baseType = type.split(':')[0];
  if (baseType !== 'n' || !value || !/^[-+]?\d+$/.test(value)) return value.trim();
  // Preserve custom identifiers such as 000000, 000-000 and (000) 000-0000.
  const section = String(format || 'General').split(';')[0]
    .replace(/\[[^\]]*\]/g, '')
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/[_*]./g, '')
    .trim();
  if (!section || !/^[0\-(). +]+$/.test(section) || !section.includes('0')) return value.trim();
  const negative = value.startsWith('-');
  let digits = value.replace(/^[+-]/, '');
  let output = '';
  for (let index = section.length - 1; index >= 0; index--) {
    const char = section[index];
    if (char === '0') {
      output = (digits.length ? digits.slice(-1) : '0') + output;
      if (digits.length) digits = digits.slice(0, -1);
    } else output = char + output;
  }
  if (digits) output = digits + output;
  return `${negative ? '-' : ''}${output}`;
}

function closeTagName(tag: any): string {
  // Saxes 6 emits the complete tag object for closetag events.
  // Accept strings as well to keep the parser helper tolerant to alternate SAX adapters.
  return typeof tag === 'string' ? tag : String(tag?.name || '');
}

function attrs(tag: any): Record<string, string> {
  const result: Record<string, string> = {};
  for (const [key, value] of Object.entries(tag?.attributes || {})) {
    result[key] = typeof value === 'string' ? value : String((value as any)?.value || '');
  }
  return result;
}

export function parseWorkbookXml(xml: string): SheetReference[] {
  const parser = new SaxesParser({ xmlns: false });
  const sheets: SheetReference[] = [];
  parser.on('opentag', (tag: any) => {
    if (tag.name !== 'sheet') return;
    const a = attrs(tag);
    if (a.name) sheets.push({ name: a.name, relationId: a['r:id'] || a.id || '' });
  });
  parser.write(xml).close();
  return sheets;
}

export function parseWorkbookRelationshipsXml(xml: string): Record<string, string> {
  const parser = new SaxesParser({ xmlns: false });
  const map: Record<string, string> = {};
  parser.on('opentag', (tag: any) => {
    if (tag.name !== 'Relationship') return;
    const a = attrs(tag);
    if (a.Id && a.Target && a.TargetMode !== 'External') map[a.Id] = a.Target;
  });
  parser.write(xml).close();
  return map;
}

export function resolveWorksheetPath(target: string): string {
  const parts: string[] = [];
  const normalized = target.replace(/\\/g, '/').replace(/^\//, '').startsWith('xl/')
    ? target.replace(/\\/g, '/').replace(/^\//, '')
    : `xl/${target.replace(/\\/g, '/').replace(/^\//, '')}`;
  for (const part of normalized.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') { parts.pop(); continue; }
    parts.push(part);
  }
  const path = parts.join('/');
  if (!path.startsWith('xl/worksheets/') || path.includes('..')) throw new Error('A planilha contém uma referência inválida para a aba selecionada.');
  return path;
}

function columnNumber(reference: string): number {
  const letters = (reference.match(/^[A-Z]+/i)?.[0] || '').toUpperCase();
  let value = 0;
  for (const letter of letters) value = value * 26 + letter.charCodeAt(0) - 64;
  return value;
}

function rowNumber(reference: string, fallback: number): number {
  const digits = reference.match(/\d+/)?.[0];
  return digits ? Number(digits) : fallback;
}

export async function parseSharedStrings(
  chunks: AsyncIterable<string>,
  writeBatch: (rows: SharedStringRow[]) => Promise<void>,
  shouldCancel?: () => boolean,
  onBytes?: (bytes: number) => void,
  wantedIds?: ArrayLike<number>,
): Promise<{ count: number }> {
  // wantedIds must be sorted and unique (SQLite returns it that way). This lets the
  // parser skip text for unused strings and stop as soon as the last referenced ID is read.
  if (wantedIds && wantedIds.length === 0) return { count: 0 };
  const parser = new SaxesParser({ xmlns: false });
  let index = 0;
  let wantedIndex = 0;
  let captureCurrent = !wantedIds;
  let foundAll = false;
  let insideSi = false;
  let insideText = false;
  let value = '';
  let batch: SharedStringRow[] = [];
  let count = 0;
  parser.on('opentag', (tag: any) => {
    if (tag.name === 'si') {
      insideSi = true;
      captureCurrent = !wantedIds || wantedIds[wantedIndex] === index;
      value = '';
    } else if (tag.name === 't' && insideSi && captureCurrent) insideText = true;
  });
  parser.on('text', (text: string) => { if (insideSi && insideText && captureCurrent) value += text; });
  parser.on('cdata', (text: string) => { if (insideSi && insideText && captureCurrent) value += text; });
  parser.on('closetag', (tag: any) => {
    const name = closeTagName(tag);
    if (name === 't') insideText = false;
    if (name === 'si') {
      if (captureCurrent) {
        batch.push({ id: index, value });
        count++;
        wantedIndex++;
        if (wantedIds && wantedIndex >= wantedIds.length) foundAll = true;
      }
      index++;
      insideSi = false;
      captureCurrent = false;
      value = '';
    }
  });
  for await (const chunk of chunks) {
    if (shouldCancel?.()) throw new Error('IMPORT_CANCELLED');
    parser.write(chunk);
    onBytes?.(chunk.length);
    if (batch.length >= 1000) { const pending = batch; batch = []; await writeBatch(pending); }
    if (foundAll) break;
  }
  // When all requested IDs have been read, the XML stream is intentionally cut
  // short while the root <sst> remains open; do not call close() in that case.
  if (!foundAll) parser.close();
  if (batch.length) await writeBatch(batch);
  return { count };
}

export async function readPreviewRows(
  chunks: AsyncIterable<string>,
  maxRows = 6,
  shouldCancel?: () => boolean,
  onBytes?: (bytes: number) => void,
): Promise<XmlPreview> {
  const parser = new SaxesParser({ xmlns: false });
  const rows: XmlPreview['rows'] = [];
  const sharedIds = new Set<number>();
  let totalRows: number | undefined;
  let maxColumn = 0;
  let currentRow: { rowNumber: number; cells: Map<number, XmlCell> } | null = null;
  let currentCell: { col: number; type: string; style: string; value: string } | null = null;
  let capture: 'v' | 't' | null = null;
  parser.on('opentag', (tag: any) => {
    const a = attrs(tag);
    if (tag.name === 'dimension' && a.ref) {
      const end = a.ref.split(':').pop() || a.ref;
      totalRows = rowNumber(end, 0) || undefined;
    } else if (tag.name === 'row') {
      currentRow = { rowNumber: Number(a.r) || rows.length + 1, cells: new Map() };
    } else if (tag.name === 'c' && currentRow) {
      currentCell = { col: columnNumber(a.r || ''), type: a.t || 'n', style: a.s || '0', value: '' };
      maxColumn = Math.max(maxColumn, currentCell.col);
    } else if (currentCell && (tag.name === 'v' || tag.name === 't')) capture = tag.name;
  });
  parser.on('text', (text: string) => { if (currentCell && capture) currentCell.value += text; });
  parser.on('cdata', (text: string) => { if (currentCell && capture) currentCell.value += text; });
  parser.on('closetag', (tag: any) => {
    const name = closeTagName(tag);
    if ((name === 'v' || name === 't') && capture === name) capture = null;
    if (name === 'c' && currentRow && currentCell) {
      currentRow.cells.set(currentCell.col, { value: currentCell.value, type: currentCell.type, style: currentCell.style });
      if (currentCell.type === 's' && currentCell.value !== '') sharedIds.add(Number(currentCell.value));
      currentCell = null;
    }
    if (name === 'row' && currentRow) {
      if (rows.length < maxRows) rows.push(currentRow);
      currentRow = null;
    }
  });
  let stoppedEarly = false;
  for await (const chunk of chunks) {
    if (shouldCancel?.()) throw new Error('IMPORT_CANCELLED');
    parser.write(chunk);
    onBytes?.(chunk.length);
    if (rows.length >= maxRows) { stoppedEarly = true; break; }
  }
  // The parser may be inside an unfinished XML document when we intentionally stop at the sample.
  if (!stoppedEarly) parser.close();
  return { rows, totalRows, maxColumn: Math.min(maxColumn, 64), sharedIds };
}

export async function readSelectedSharedStrings(
  chunks: AsyncIterable<string>,
  wanted: Set<number>,
  shouldCancel?: () => boolean,
): Promise<Map<number, string>> {
  const parser = new SaxesParser({ xmlns: false });
  const values = new Map<number, string>();
  let index = 0;
  let insideSi = false;
  let insideText = false;
  let value = '';
  parser.on('opentag', (tag: any) => {
    if (tag.name === 'si') { insideSi = true; value = ''; }
    else if (tag.name === 't' && insideSi) insideText = true;
  });
  parser.on('text', (text: string) => { if (insideSi && insideText && wanted.has(index)) value += text; });
  parser.on('cdata', (text: string) => { if (insideSi && insideText && wanted.has(index)) value += text; });
  parser.on('closetag', (tag: any) => {
    const name = closeTagName(tag);
    if (name === 't') insideText = false;
    if (name === 'si') { if (wanted.has(index)) values.set(index, value); index++; insideSi = false; value = ''; }
  });
  const maxWanted = wanted.size ? Math.max(...wanted) : -1;
  let reachedLastWanted = false;
  for await (const chunk of chunks) {
    if (shouldCancel?.()) throw new Error('IMPORT_CANCELLED');
    if (maxWanted >= 0 && index > maxWanted) { reachedLastWanted = true; break; }
    parser.write(chunk);
  }
  if (!reachedLastWanted) parser.close();
  return values;
}

export async function parseWorksheetToImportRows(
  chunks: AsyncIterable<string>,
  mapping: { codeColumn: number; nameColumn: number; quantityColumn: number; headerRow?: number },
  writeBatch: (rows: RawImportRow[]) => Promise<void>,
  shouldCancel?: () => boolean,
  onProgress?: (progress: { bytesRead: number; totalRows?: number; rowsRead: number }) => void,
): Promise<{ totalRows?: number; rowsRead: number }> {
  const parser = new SaxesParser({ xmlns: false });
  const targetColumns = new Set([mapping.codeColumn, mapping.nameColumn, mapping.quantityColumn].filter(column => column > 0));
  let totalRows: number | undefined;
  let rowsRead = 0;
  let bytesRead = 0;
  const headerRow = Math.max(1, Math.floor(mapping.headerRow || 1));
  let currentRow: { rowNumber: number; cells: Map<number, { value: string; type: string }> } | null = null;
  let currentCell: { col: number; type: string; value: string } | null = null;
  let capture: 'v' | 't' | null = null;
  let batch: RawImportRow[] = [];
  parser.on('opentag', (tag: any) => {
    const a = attrs(tag);
    if (tag.name === 'dimension' && a.ref) {
      const end = a.ref.split(':').pop() || a.ref;
      totalRows = rowNumber(end, 0) || undefined;
    } else if (tag.name === 'row') {
      currentRow = { rowNumber: Number(a.r) || rowsRead + 1, cells: new Map() };
    } else if (tag.name === 'c' && currentRow) {
      const col = columnNumber(a.r || '');
      currentCell = targetColumns.has(col) ? { col, type: `${a.t || 'n'}:${a.s || '0'}`, value: '' } : null;
    } else if (currentCell && (tag.name === 'v' || tag.name === 't')) capture = tag.name;
  });
  parser.on('text', (text: string) => { if (currentCell && capture) currentCell.value += text; });
  parser.on('cdata', (text: string) => { if (currentCell && capture) currentCell.value += text; });
  parser.on('closetag', (tag: any) => {
    const name = closeTagName(tag);
    if ((name === 'v' || name === 't') && capture === name) capture = null;
    if (name === 'c' && currentRow && currentCell) {
      currentRow.cells.set(currentCell.col, { value: currentCell.value, type: currentCell.type });
      currentCell = null;
    }
    if (name === 'row' && currentRow) {
      // Ignore preamble rows and the detected header row. This supports workbooks
      // where a title or report heading appears before the actual column names.
      if (currentRow.rowNumber <= headerRow) { currentRow = null; return; }
      const pick = (column: number) => currentRow?.cells.get(column) || { value: '', type: '' };
      const code = pick(mapping.codeColumn);
      const itemName = pick(mapping.nameColumn);
      const quantity = pick(mapping.quantityColumn);
      // Blank rows are counted for progress but never written to SQLite.
      if (code.value !== '' || itemName.value !== '') {
        batch.push({
          rowNumber: currentRow.rowNumber,
          codeRaw: code.value,
          codeType: code.type,
          nameRaw: itemName.value,
          nameType: itemName.type,
          quantityRaw: quantity.value,
          quantityType: quantity.type,
        });
      }
      rowsRead++;
      currentRow = null;
    }
  });
  for await (const chunk of chunks) {
    if (shouldCancel?.()) throw new Error('IMPORT_CANCELLED');
    parser.write(chunk);
    bytesRead += chunk.length;
    if (batch.length >= 1000) { const pending = batch; batch = []; await writeBatch(pending); }
    onProgress?.({ bytesRead, totalRows, rowsRead });
  }
  parser.close();
  if (batch.length) await writeBatch(batch);
  onProgress?.({ bytesRead, totalRows, rowsRead });
  return { totalRows, rowsRead };
}

export function resolvePreviewRows(preview: XmlPreview, shared: Map<number, string>, columns = 16, styleFormats: string[] = []): string[][] {
  return preview.rows.map(row => {
    const cells: string[] = [];
    for (let column = 1; column <= columns; column++) {
      const cell = row.cells.get(column);
      if (!cell) { cells.push(''); continue; }
      if (cell.type === 's') cells.push(shared.get(Number(cell.value)) || '');
      else cells.push(formatExcelValue(cell.value, `${cell.type}:${cell.style}`, styleFormats[Number(cell.style)] || 'General'));
    }
    return cells;
  });
}
