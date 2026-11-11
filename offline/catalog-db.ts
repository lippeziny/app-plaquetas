import { Platform } from 'react-native';
import * as SQLite from 'expo-sqlite';
import type { SQLiteDatabase } from 'expo-sqlite';
import type { CatalogItem } from './workbook';
import { formatExcelValue } from './xlsx-xml';

export type RawImportRow = {
  rowNumber: number;
  codeRaw: string;
  codeType: string;
  nameRaw: string;
  nameType: string;
  quantityRaw: string;
  quantityType: string;
};
export type SharedStringRow = { id: number; value: string };
export type CatalogRow = CatalogItem & { id?: number; code_norm?: string; name_norm?: string };
export type FinalizeProgress = { rowsRead: number; totalRows: number; saved: number; skipped: number };

const DB_NAME = 'plaquetas-catalog.db';
const WEB_KEY = 'plaquetas-catalog-v2';
let dbPromise: Promise<SQLiteDatabase> | null = null;

function normalize(value: string): string {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

async function getDb(): Promise<SQLiteDatabase> {
  if (!dbPromise) {
    dbPromise = (async () => {
      const db = await SQLite.openDatabaseAsync(DB_NAME);
      await db.execAsync(`
        PRAGMA journal_mode = WAL;
        PRAGMA synchronous = NORMAL;
        PRAGMA temp_store = FILE;
        PRAGMA cache_size = -8192;
        PRAGMA busy_timeout = 5000;
        CREATE TABLE IF NOT EXISTS catalog (
          id INTEGER PRIMARY KEY,
          code TEXT NOT NULL,
          name TEXT NOT NULL,
          quantity TEXT NOT NULL DEFAULT '',
          sheet_name TEXT NOT NULL DEFAULT '',
          code_norm TEXT NOT NULL,
          name_norm TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_catalog_code_norm ON catalog(code_norm);
        CREATE INDEX IF NOT EXISTS idx_catalog_name_norm ON catalog(name_norm);
        CREATE TABLE IF NOT EXISTS catalog_import_shared (
          sid INTEGER PRIMARY KEY,
          value TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS catalog_import_rows (
          row_num INTEGER PRIMARY KEY,
          code_raw TEXT NOT NULL,
          code_type TEXT NOT NULL,
          name_raw TEXT NOT NULL,
          name_type TEXT NOT NULL,
          quantity_raw TEXT NOT NULL,
          quantity_type TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_import_code_ref ON catalog_import_rows(code_type, code_raw);
        CREATE INDEX IF NOT EXISTS idx_import_name_ref ON catalog_import_rows(name_type, name_raw);
        CREATE INDEX IF NOT EXISTS idx_import_quantity_ref ON catalog_import_rows(quantity_type, quantity_raw);
        CREATE TABLE IF NOT EXISTS catalog_import_stage (
          row_num INTEGER PRIMARY KEY,
          code TEXT NOT NULL,
          name TEXT NOT NULL,
          quantity TEXT NOT NULL,
          sheet_name TEXT NOT NULL,
          code_norm TEXT NOT NULL,
          name_norm TEXT NOT NULL
        );
      `);
      return db;
    })().catch(error => {
      dbPromise = null;
      throw error;
    });
  }
  return dbPromise;
}

export async function initializeCatalogStore(): Promise<void> {
  if (Platform.OS !== 'web') await getDb();
}

function getWebCatalog(): CatalogRow[] {
  try { return JSON.parse(localStorage.getItem(WEB_KEY) || '[]') as CatalogRow[]; }
  catch { return []; }
}

export async function getCatalogCount(): Promise<number> {
  if (Platform.OS === 'web') return getWebCatalog().length;
  const db = await getDb();
  const row = await db.getFirstAsync<{ total: number }>('SELECT COUNT(*) AS total FROM catalog');
  return Number(row?.total || 0);
}

async function insertMany<T>(table: string, columns: string[], rows: T[], map: (value: T) => unknown[]): Promise<void> {
  if (!rows.length) return;
  const db = await getDb();
  const width = columns.length;
  const perBatch = Math.max(1, Math.floor(900 / width));
  await db.withTransactionAsync(async () => {
    for (let start = 0; start < rows.length; start += perBatch) {
      const batch = rows.slice(start, start + perBatch);
      const values = batch.map(map);
      const placeholders = values.map(() => `(${columns.map(() => '?').join(',')})`).join(',');
      const params = values.flat();
      const sql = `INSERT OR REPLACE INTO ${table} (${columns.join(',')}) VALUES ${placeholders}`;
      await db.runAsync(sql, ...params as any[]);
    }
  });
}

export async function replaceCatalogFromItems(items: CatalogItem[]): Promise<void> {
  if (Platform.OS === 'web') {
    localStorage.setItem(WEB_KEY, JSON.stringify(items));
    return;
  }
  const db = await getDb();
  await db.execAsync('DELETE FROM catalog;');
  const normalized = items.map((item, index) => ({
    id: index + 1,
    code: item.code,
    name: item.name,
    quantity: item.quantity || '',
    sheetName: item.sheetName || '',
    codeNorm: normalize(item.code),
    nameNorm: normalize(item.name),
  }));
  await insertMany('catalog', ['id', 'code', 'name', 'quantity', 'sheet_name', 'code_norm', 'name_norm'], normalized,
    row => [row.id, row.code, row.name, row.quantity, row.sheetName, row.codeNorm, row.nameNorm]);
}

export async function searchCatalogSuggestions(
  query: string,
  field: 'code' | 'name',
  limit = 8,
): Promise<CatalogItem[]> {
  const normalized = normalize(query);
  if (!normalized) return [];
  const target = field === 'code' ? 'code_norm' : 'name_norm';
  if (Platform.OS === 'web') {
    return getWebCatalog()
      .filter(item => normalize(item[field]).startsWith(normalized))
      .sort((a, b) => normalize(a[field]).localeCompare(normalize(b[field])))
      .slice(0, limit);
  }
  const db = await getDb();
  return db.getAllAsync<CatalogRow>(
    `SELECT code, name, quantity, sheet_name AS sheetName FROM catalog WHERE ${target} LIKE ? ORDER BY ${target}, id LIMIT ?`,
    `${normalized}%`, limit,
  );
}

/** Loads a bounded page from the local catalog. Name filtering is accent-insensitive
 * and supports text anywhere in the product description (e.g. "calca"). */
export async function getCatalogPage(
  query = '',
  limit = 60,
  offset = 0,
): Promise<CatalogItem[]> {
  const normalized = normalize(query);
  const safeLimit = Math.max(1, Math.min(100, Math.floor(limit)));
  const safeOffset = Math.max(0, Math.floor(offset));
  if (Platform.OS === 'web') {
    const items = getWebCatalog();
    const filtered = normalized
      ? items.filter(item => normalize(item.name).includes(normalized))
      : items;
    filtered.sort((a, b) => normalize(a.name).localeCompare(normalize(b.name)));
    return filtered.slice(safeOffset, safeOffset + safeLimit);
  }
  const db = await getDb();
  if (normalized) {
    // Escaping prevents %, _ and \\ typed by the user from becoming SQL wildcards.
    const escaped = normalized.replace(/[\\%_]/g, '\\$&');
    return db.getAllAsync<CatalogRow>(
      `SELECT code, name, quantity, sheet_name AS sheetName
       FROM catalog WHERE name_norm LIKE ? ESCAPE '\\'
       ORDER BY CASE WHEN name_norm LIKE ? ESCAPE '\\' THEN 0 ELSE 1 END, name_norm, id
       LIMIT ? OFFSET ?`,
      `%${escaped}%`, `${escaped}%`, safeLimit, safeOffset,
    );
  }
  return db.getAllAsync<CatalogRow>(
    'SELECT code, name, quantity, sheet_name AS sheetName FROM catalog ORDER BY name_norm, id LIMIT ? OFFSET ?',
    safeLimit, safeOffset,
  );
}

export async function searchCatalogNames(query: string, limit = 8): Promise<CatalogItem[]> {
  const normalized = normalize(query);
  if (!normalized) return [];
  if (Platform.OS === 'web') return getWebCatalog().filter(item => normalize(item.name).includes(normalized)).slice(0, limit);
  const db = await getDb();
  let rows = await db.getAllAsync<CatalogRow>(
    'SELECT code, name, quantity, sheet_name AS sheetName FROM catalog WHERE name_norm LIKE ? ORDER BY name_norm LIMIT ?',
    `${normalized}%`, limit,
  );
  if (rows.length < limit) {
    const extras = await db.getAllAsync<CatalogRow>(
      'SELECT code, name, quantity, sheet_name AS sheetName FROM catalog WHERE name_norm LIKE ? AND name_norm NOT LIKE ? ORDER BY name_norm LIMIT ?',
      `%${normalized}%`, `${normalized}%`, limit - rows.length,
    );
    rows = [...rows, ...extras];
  }
  return rows;
}

export async function findCatalogItems(values: string[], field: 'code' | 'name'): Promise<Map<string, CatalogItem>> {
  const unique = [...new Set(values.map(normalize).filter(Boolean))];
  const found = new Map<string, CatalogItem>();
  if (!unique.length) return found;
  if (Platform.OS === 'web') {
    const items = getWebCatalog();
    const target = field === 'code' ? 'code' : 'name';
    for (const item of items) {
      const key = normalize(item[target]);
      if (unique.includes(key) && !found.has(key)) found.set(key, item);
    }
    return found;
  }
  const db = await getDb();
  const column = field === 'code' ? 'code_norm' : 'name_norm';
  for (let start = 0; start < unique.length; start += 700) {
    const batch = unique.slice(start, start + 700);
    const placeholders = batch.map(() => '?').join(',');
    const rows = await db.getAllAsync<CatalogRow>(
      `SELECT code, name, quantity, sheet_name AS sheetName, ${column} FROM catalog WHERE ${column} IN (${placeholders}) ORDER BY id`,
      ...batch,
    );
    for (const row of rows) {
      const key = String(row[column] || '');
      if (!found.has(key)) found.set(key, { code: row.code, name: row.name, quantity: row.quantity, sheetName: row.sheetName });
    }
  }
  return found;
}

export async function findCatalogByCode(value: string): Promise<CatalogItem | undefined> {
  const key = normalize(value);
  if (!key) return undefined;
  if (Platform.OS === 'web') {
    const items = getWebCatalog();
    return items.find(item => normalize(item.code) === key) || items.find(item => normalize(item.code).startsWith(key));
  }
  const db = await getDb();
  const exact = await db.getFirstAsync<CatalogRow>(
    'SELECT code, name, quantity, sheet_name AS sheetName FROM catalog WHERE code_norm = ? ORDER BY id LIMIT 1', key,
  );
  if (exact) return exact;
  return (await db.getFirstAsync<CatalogRow>(
    'SELECT code, name, quantity, sheet_name AS sheetName FROM catalog WHERE code_norm LIKE ? ORDER BY id LIMIT 1', `${key}%`,
  )) || undefined;
}

export async function clearImportStaging(): Promise<void> {
  if (Platform.OS === 'web') return;
  const db = await getDb();
  await db.execAsync('DELETE FROM catalog_import_shared; DELETE FROM catalog_import_rows; DELETE FROM catalog_import_stage;');
}

export async function getNeededSharedStringIds(shouldCancel?: () => boolean): Promise<Uint32Array> {
  const db = await getDb();
  // UNION deduplicates and sorts IDs on SQLite's disk-backed temp store. getEachAsync
  // streams the result to JS instead of materializing a potentially huge result array.
  const sql = `
    SELECT sid FROM (
      SELECT CAST(code_raw AS INTEGER) AS sid FROM catalog_import_rows
        WHERE code_type >= 's:' AND code_type < 's;' AND code_raw <> '' AND code_raw NOT GLOB '*[^0-9]*'
      UNION
      SELECT CAST(name_raw AS INTEGER) AS sid FROM catalog_import_rows
        WHERE name_type >= 's:' AND name_type < 's;' AND name_raw <> '' AND name_raw NOT GLOB '*[^0-9]*'
      UNION
      SELECT CAST(quantity_raw AS INTEGER) AS sid FROM catalog_import_rows
        WHERE quantity_type >= 's:' AND quantity_type < 's;' AND quantity_raw <> '' AND quantity_raw NOT GLOB '*[^0-9]*'
    ) WHERE sid >= 0 ORDER BY sid`;
  let ids = new Uint32Array(4096);
  let length = 0;
  for await (const row of db.getEachAsync<{ sid: number }>(sql)) {
    if ((length & 1023) === 0 && shouldCancel?.()) throw new Error('IMPORT_CANCELLED');
    const sid = Number(row.sid);
    if (!Number.isSafeInteger(sid) || sid < 0 || sid > 0xffffffff) continue;
    if (length === ids.length) {
      const expanded = new Uint32Array(ids.length * 2);
      expanded.set(ids);
      ids = expanded;
    }
    ids[length++] = sid;
  }
  // A typed array uses four bytes per shared-string ID instead of a JS number array.
  return ids.subarray(0, length);
}

export async function insertSharedStringsBatch(rows: SharedStringRow[]): Promise<void> {
  if (!rows.length) return;
  const db = await getDb();
  // The XML parser now emits only IDs referenced by the selected worksheet columns.
  // Simple batched inserts avoid a repeated EXISTS lookup for every shared-string batch.
  const perBatch = 450; // 900 bind parameters, below SQLite's common 999-variable limit.
  await db.withTransactionAsync(async () => {
    for (let start = 0; start < rows.length; start += perBatch) {
      const batch = rows.slice(start, start + perBatch);
      const values = batch.map(() => '(?, ?)').join(',');
      const params = batch.flatMap(row => [row.id, row.value]);
      await db.runAsync(`INSERT OR REPLACE INTO catalog_import_shared (sid, value) VALUES ${values}`, ...params as any[]);
    }
  });
}

export async function insertImportRowsBatch(rows: RawImportRow[]): Promise<void> {
  await insertMany('catalog_import_rows',
    ['row_num', 'code_raw', 'code_type', 'name_raw', 'name_type', 'quantity_raw', 'quantity_type'],
    rows,
    row => [row.rowNumber, row.codeRaw, row.codeType, row.nameRaw, row.nameType, row.quantityRaw, row.quantityType]);
}

export async function finalizeStagedCatalog(
  sheetName: string,
  styleFormats: string[] = [],
  onProgress?: (progress: FinalizeProgress) => void,
  shouldCancel?: () => boolean,
): Promise<{ total: number; skipped: number; sourceRows: number }> {
  const db = await getDb();
  const source = await db.getFirstAsync<{ total: number }>('SELECT COUNT(*) AS total FROM catalog_import_rows');
  const sourceRows = Number(source?.total || 0);
  let lastRow = 0;
  let saved = 0;
  let skipped = 0;
  let processedRows = 0;
  try {
    while (true) {
      if (shouldCancel?.()) throw new Error('IMPORT_CANCELLED');
      const rows = await db.getAllAsync<any>(`
        SELECT r.row_num AS rowNumber,
          CASE WHEN (r.code_type >= 's:' AND r.code_type < 's;') THEN COALESCE(sc.value, '') ELSE r.code_raw END AS code, r.code_type AS codeType,
          CASE WHEN (r.name_type >= 's:' AND r.name_type < 's;') THEN COALESCE(sn.value, '') ELSE r.name_raw END AS name, r.name_type AS nameType,
          CASE WHEN (r.quantity_type >= 's:' AND r.quantity_type < 's;') THEN COALESCE(sq.value, '') ELSE r.quantity_raw END AS quantity, r.quantity_type AS quantityType
        FROM catalog_import_rows r
        LEFT JOIN catalog_import_shared sc ON sc.sid = CAST(r.code_raw AS INTEGER)
        LEFT JOIN catalog_import_shared sn ON sn.sid = CAST(r.name_raw AS INTEGER)
        LEFT JOIN catalog_import_shared sq ON sq.sid = CAST(r.quantity_raw AS INTEGER)
        WHERE r.row_num > ? ORDER BY r.row_num LIMIT 2000
      `, lastRow);
      if (!rows.length) break;
      const staged: Array<[number, string, string, string, string, string, string]> = [];
      for (const row of rows) {
        lastRow = Number(row.rowNumber);
        processedRows++;
        const codeStyle = Number(String(row.codeType || '').split(':')[1] || 0);
        const nameStyle = Number(String(row.nameType || '').split(':')[1] || 0);
        const quantityStyle = Number(String(row.quantityType || '').split(':')[1] || 0);
        const code = formatExcelValue(String(row.code || ''), String(row.codeType || ''), styleFormats[codeStyle] || 'General').replace(/\s+/g, ' ').trim();
        const name = formatExcelValue(String(row.name || ''), String(row.nameType || ''), styleFormats[nameStyle] || 'General').replace(/\s+/g, ' ').trim();
        const quantity = formatExcelValue(String(row.quantity || ''), String(row.quantityType || ''), styleFormats[quantityStyle] || 'General').trim();
        if (!code || !name) { skipped++; continue; }
        staged.push([lastRow, code, name, quantity, sheetName, normalize(code), normalize(name)]);
        saved++;
      }
      await db.withTransactionAsync(async () => {
        for (let start = 0; start < staged.length; start += 125) {
          const batch = staged.slice(start, start + 125);
          const placeholders = batch.map(() => '(?, ?, ?, ?, ?, ?, ?)').join(',');
          await db.runAsync(`INSERT OR REPLACE INTO catalog_import_stage (row_num, code, name, quantity, sheet_name, code_norm, name_norm) VALUES ${placeholders}`,
            ...batch.flat() as any[]);
        }
        // Free raw rows as soon as their resolved values have been copied to staging.
        // This keeps peak disk usage closer to (shared strings + staging), not raw + shared + staging.
        await db.runAsync('DELETE FROM catalog_import_rows WHERE row_num <= ?', lastRow);
      });
      onProgress?.({ rowsRead: processedRows, totalRows: sourceRows, saved, skipped });
    }

  const stageCount = await db.getFirstAsync<{ total: number }>('SELECT COUNT(*) AS total FROM catalog_import_stage');
  const total = Number(stageCount?.total || 0);
  await db.withTransactionAsync(async () => {
    await db.execAsync('DELETE FROM catalog;');
    await db.execAsync(`INSERT INTO catalog (id, code, name, quantity, sheet_name, code_norm, name_norm)
      SELECT row_num, code, name, quantity, sheet_name, code_norm, name_norm FROM catalog_import_stage ORDER BY row_num;`);
    await db.execAsync('DELETE FROM catalog_import_stage; DELETE FROM catalog_import_rows; DELETE FROM catalog_import_shared;');
  });
  onProgress?.({ rowsRead: sourceRows, totalRows: sourceRows, saved: total, skipped });
  return { total, skipped, sourceRows };
  } catch (error) {
    await db.execAsync('DELETE FROM catalog_import_stage;').catch(() => {});
    throw error;
  }
}
