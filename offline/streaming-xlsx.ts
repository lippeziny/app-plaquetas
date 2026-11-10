import { File as ExpoFile, FileMode } from 'expo-file-system';
import * as FileSystem from 'expo-file-system/legacy';
import { Utf8StreamDecoder } from './utf8-stream';
import {
  clearImportStaging,
  finalizeStagedCatalog,
  getNeededSharedStringIds,
  insertImportRowsBatch,
  insertSharedStringsBatch,
  type FinalizeProgress,
} from './catalog-db';
import {
  parseSharedStrings,
  parseWorkbookRelationshipsXml,
  parseWorkbookXml,
  parseWorksheetToImportRows,
  readPreviewRows,
  readSelectedSharedStrings,
  resolvePreviewRows,
  resolveWorksheetPath,
  parseStylesXml,
} from './xlsx-xml';

declare const require: (name: string) => any;

export type WorkbookFileSource = {
  uri: string;
  name: string;
  directory: string;
  sheetNames: string[];
  sheetPaths: Record<string, string>;
  sharedStringsPath?: string;
  stylesPath?: string;
  extracted: string[];
  sizeBytes: number;
};
export type StreamingWorkbookPreview = {
  headers: string[];
  rows: string[][];
  columns: { index: number; label: string }[];
  suggested: { codeColumn: number; nameColumn: number; quantityColumn: number };
};
export type ImportProgress = {
  stage: string;
  percent: number;
  rowsProcessed?: number;
  totalRows?: number;
  bytesProcessed?: number;
  totalBytes?: number;
  saved?: number;
  skipped?: number;
};

const XML_CHUNK_SIZE = 256 * 1024;


function zipApi() { return require('react-native-zip-archive') as typeof import('react-native-zip-archive'); }
function joinUri(directory: string, path: string): string { return `${directory.replace(/\/?$/, '/')}${path}`; }
function zipInputUri(uri: string): string {
  // react-native-zip-archive expects a decoded local path/URI; DocumentPicker can return %20 etc.
  try { return decodeURIComponent(uri); } catch { return decodeURI(uri); }
}

async function* xmlChunks(
  uri: string,
  signal?: AbortSignal,
  onBytesRead?: (bytesRead: number, totalBytes: number) => void,
): AsyncGenerator<string> {
  const file = new ExpoFile(uri);
  const handle = file.open(FileMode.ReadOnly);
  const decoder = new Utf8StreamDecoder();
  let bytesRead = 0;
  const totalBytes = Number(handle.size ?? file.size ?? 0);
  try {
    while ((handle.offset ?? 0) < (handle.size ?? 0)) {
      if (signal?.aborted) throw new Error('IMPORT_CANCELLED');
      const bytes = await handle.readBytes(XML_CHUNK_SIZE);
      if (!bytes.length) break;
      bytesRead += bytes.length;
      onBytesRead?.(bytesRead, totalBytes);
      const text = decoder.decode(bytes, false);
      if (text) yield text;
    }
    const tail = decoder.decode(new Uint8Array(0), true);
    if (tail) yield tail;
  } finally {
    handle.close();
  }
}

function normalizeRelationshipTarget(target: string): string {
  const clean = target.replace(/\\/g, '/');
  if (/^https?:/i.test(clean) || clean.startsWith('//')) throw new Error('O Excel referencia uma aba externa, que não pode ser importada offline.');
  return clean.startsWith('/') || clean.startsWith('xl/') ? clean.replace(/^\//, '') : `xl/${clean.replace(/^\//, '')}`;
}

// Default mapping follows the stock layout requested for this app: A=code, B=name, D=quantity.
// The UI still allows the user to remap columns for other workbook formats.

async function extractEntries(source: WorkbookFileSource, paths: string[], onProgress?: (value: number) => void, signal?: AbortSignal) {
  const unique = [...new Set(paths.map(path => path.replace(/^\//, '')))].filter(Boolean).filter(path => !source.extracted.includes(path));
  if (!unique.length) return;
  if (signal?.aborted) throw new Error('IMPORT_CANCELLED');
  const zip = zipApi();
  const subscription = zip.subscribe?.(({ progress }: { progress: number }) => onProgress?.(progress));
  try {
    await zip.unzip(source.uri, source.directory, { entries: unique, signal });
    source.extracted.push(...unique);
  } finally { subscription?.remove?.(); }
}

export async function inspectXlsxFile(
  uri: string,
  name: string,
  onProgress?: (progress: ImportProgress) => void,
  signal?: AbortSignal,
): Promise<WorkbookFileSource> {
  if (!uri) throw new Error('O arquivo selecionado não possui um caminho válido.');
  const zip = zipApi();
  onProgress?.({ stage: 'Verificando o arquivo Excel', percent: 0.02 });
  const archiveUri = zipInputUri(uri);
  let entries: any[];
  try { entries = await zip.listContents(archiveUri); }
  catch { throw new Error('Não consegui abrir esse arquivo como Excel .xlsx. Confira se ele não está protegido por senha e se foi salvo no formato .xlsx.'); }
  const fileEntries = entries.filter((entry: any) => !entry.isDirectory);
  const findEntry = (path: string) => fileEntries.find((entry: any) => entry.path.replace(/^\//, '') === path);
  const workbookEntry = findEntry('xl/workbook.xml');
  const relsEntry = findEntry('xl/_rels/workbook.xml.rels');
  if (!workbookEntry || !relsEntry) throw new Error('Este arquivo não parece ser uma planilha XLSX válida. Salve novamente no formato .xlsx e tente outra vez.');
  if (workbookEntry.isEncrypted || relsEntry.isEncrypted) throw new Error('A planilha está protegida por senha. Remova a proteção do arquivo e tente novamente.');
  const directory = `${FileSystem.cacheDirectory || FileSystem.documentDirectory}xlsx-import-${Date.now()}-${Math.random().toString(36).slice(2, 7)}/`;
  if (!directory.startsWith('file://')) throw new Error('O armazenamento temporário do aparelho não está disponível.');
  await FileSystem.makeDirectoryAsync(directory, { intermediates: true });
  const source: WorkbookFileSource = {
    uri: archiveUri, name, directory, sheetNames: [], sheetPaths: {},
    sharedStringsPath: findEntry('xl/sharedStrings.xml')?.path.replace(/^\//, ''),
    stylesPath: findEntry('xl/styles.xml')?.path.replace(/^\//, ''),
    extracted: [], sizeBytes: Number(new ExpoFile(uri).size || 0),
  };
  try {
    onProgress?.({ stage: 'Lendo a estrutura da planilha', percent: 0.05 });
    await extractEntries(source, ['xl/workbook.xml', 'xl/_rels/workbook.xml.rels'], progress => {
      onProgress?.({ stage: 'Preparando a planilha', percent: 0.05 + progress * 0.08 });
    }, signal);
    const workbookXml = await new ExpoFile(joinUri(directory, 'xl/workbook.xml')).text();
    const relsXml = await new ExpoFile(joinUri(directory, 'xl/_rels/workbook.xml.rels')).text();
    const sheets = parseWorkbookXml(workbookXml);
    const relationships = parseWorkbookRelationshipsXml(relsXml);
    for (const sheet of sheets) {
      const target = relationships[sheet.relationId];
      if (!target) continue;
      const path = resolveWorksheetPath(normalizeRelationshipTarget(target));
      if (findEntry(path)) source.sheetPaths[sheet.name] = path;
    }
    source.sheetNames = Object.keys(source.sheetPaths);
    if (!source.sheetNames.length) throw new Error('Não encontrei abas válidas nesta planilha.');
    return source;
  } catch (error) {
    await cleanupWorkbookFile(source).catch(() => {});
    if (signal?.aborted || (error as any)?.code === 'ERR_CANCELLED') throw new Error('IMPORT_CANCELLED');
    throw error;
  }
}

export async function previewXlsxSheet(
  source: WorkbookFileSource,
  sheetName: string,
  onProgress?: (progress: ImportProgress) => void,
  signal?: AbortSignal,
): Promise<StreamingWorkbookPreview> {
  const sheetPath = source.sheetPaths[sheetName];
  if (!sheetPath) throw new Error('A aba selecionada não foi encontrada dentro do Excel.');
  // Do not extract sharedStrings.xml until the preview proves it is needed.
  // Some workbooks use inline strings and can skip that potentially huge entry entirely.
  const wanted = [sheetPath, source.stylesPath || ''];
  await extractEntries(source, wanted, progress => onProgress?.({ stage: 'Preparando prévia da aba', percent: 0.14 + progress * 0.06 }), signal);
  const parsedRows = await readPreviewRows(xmlChunks(joinUri(source.directory, sheetPath), signal), 6, () => !!signal?.aborted);
  let shared = new Map<number, string>();
  if (source.sharedStringsPath && parsedRows.sharedIds.size) {
    await extractEntries(source, [source.sharedStringsPath], progress => onProgress?.({ stage: 'Preparando os nomes da prévia', percent: 0.20 + progress * 0.01 }), signal);
    shared = await readSelectedSharedStrings(xmlChunks(joinUri(source.directory, source.sharedStringsPath), signal), parsedRows.sharedIds, () => !!signal?.aborted);
  }
  const styleFormats = source.stylesPath
    ? parseStylesXml(await new ExpoFile(joinUri(source.directory, source.stylesPath)).text())
    : [];
  const rows = resolvePreviewRows(parsedRows, shared, Math.max(4, Math.min(parsedRows.maxColumn || 4, 32)), styleFormats);
  const headers = (rows[0] || []).map((cell, index) => cell || `Coluna ${index + 1}`);
  const sample = rows.slice(1, 6);
  const columns = headers.map((label, index) => ({ index: index + 1, label }));
  return {
    headers, rows: sample, columns,
    suggested: { codeColumn: 1, nameColumn: 2, quantityColumn: 4 },
  };
}

export async function importXlsxSheetToCatalog(
  source: WorkbookFileSource,
  sheetName: string,
  mapping: { codeColumn: number; nameColumn: number; quantityColumn: number },
  onProgress: (progress: ImportProgress) => void,
  signal?: AbortSignal,
): Promise<{ total: number; skipped: number; sourceRows: number }> {
  const sheetPath = source.sheetPaths[sheetName];
  if (!sheetPath) throw new Error('A aba selecionada não foi encontrada dentro do Excel.');
  if (!mapping.codeColumn || !mapping.nameColumn) throw new Error('Selecione as colunas de código e nome antes de importar.');
  await clearImportStaging();
  try {
    // The worksheet and styles are small compared with sharedStrings.xml. Extract
    // shared strings only after the selected columns reveal which string IDs are used.
    const entries = [sheetPath, source.stylesPath || ''];
    onProgress({ stage: 'Extraindo somente os dados necessários', percent: 0.03, bytesProcessed: 0, totalBytes: source.sizeBytes });
    await extractEntries(source, entries, progress => onProgress({
      stage: 'Extraindo a aba selecionada', percent: 0.03 + progress * 0.12,
      bytesProcessed: Math.round(source.sizeBytes * progress), totalBytes: source.sizeBytes,
    }), signal);

    const sheetFile = new ExpoFile(joinUri(source.directory, sheetPath));
    const sheetSize = Number(sheetFile.size || 0);
    let lastReport = 0;
    let actualSheetBytesRead = 0;
    onProgress({ stage: 'Lendo produtos e quantidades', percent: 0.16, rowsProcessed: 0, totalRows: undefined, bytesProcessed: 0, totalBytes: sheetSize });
    const parsed = await parseWorksheetToImportRows(
      xmlChunks(sheetFile.uri, signal, bytesRead => { actualSheetBytesRead = bytesRead; }), mapping, insertImportRowsBatch, () => !!signal?.aborted,
      progress => {
        const now = Date.now();
        if (now - lastReport < 180 && progress.bytesRead < sheetSize) return;
        lastReport = now;
        const fraction = sheetSize
          ? Math.min(1, actualSheetBytesRead / sheetSize)
          : (progress.totalRows ? Math.min(1, progress.rowsRead / Math.max(1, progress.totalRows - 1)) : 0);
        onProgress({
          stage: 'Extraindo código, descrição e quantidade',
          percent: 0.16 + fraction * 0.52,
          rowsProcessed: progress.rowsRead,
          totalRows: progress.totalRows ? Math.max(0, progress.totalRows - 1) : undefined,
          bytesProcessed: actualSheetBytesRead,
          totalBytes: sheetSize,
        });
      },
    );

    const styleFormats = source.stylesPath
      ? parseStylesXml(await new ExpoFile(joinUri(source.directory, source.stylesPath)).text())
      : [];

    // Only keep shared strings referenced by the selected columns on the selected sheet.
    // This is much smaller than importing the workbook's global sharedStrings table wholesale.
    {
      onProgress({ stage: 'Identificando os textos realmente usados', percent: 0.685, rowsProcessed: parsed.rowsRead, totalRows: parsed.totalRows });
      const neededStringIds = await getNeededSharedStringIds(() => !!signal?.aborted);
      if (neededStringIds.length) {
        onProgress({ stage: `Preparando ${neededStringIds.length.toLocaleString('pt-BR')} textos necessários`, percent: 0.685, rowsProcessed: parsed.rowsRead, totalRows: parsed.totalRows });
        if (!source.sharedStringsPath) {
          throw new Error('A planilha referencia textos compartilhados, mas não contém o catálogo de strings do Excel.');
        }
        await extractEntries(source, [source.sharedStringsPath], progress => onProgress({
          stage: 'Extraindo os textos referenciados', percent: 0.685 + progress * 0.015,
        }), signal);
        const sharedFile = new ExpoFile(joinUri(source.directory, source.sharedStringsPath));
        const sharedSize = Number(sharedFile.size || 0);
        onProgress({ stage: `Resolvendo ${neededStringIds.length.toLocaleString('pt-BR')} textos usados`, percent: 0.70, bytesProcessed: 0, totalBytes: sharedSize });
        let lastSharedReport = 0;
        let actualSharedBytesRead = 0;
        await parseSharedStrings(
          xmlChunks(sharedFile.uri, signal, bytesRead => { actualSharedBytesRead = bytesRead; }), insertSharedStringsBatch, () => !!signal?.aborted,
          () => {
            const now = Date.now();
            if (now - lastSharedReport < 220 && actualSharedBytesRead < sharedSize) return;
            lastSharedReport = now;
            onProgress({ stage: 'Resolvendo códigos e descrições', percent: 0.70 + (sharedSize ? Math.min(1, actualSharedBytesRead / sharedSize) * 0.16 : 0.08), bytesProcessed: actualSharedBytesRead, totalBytes: sharedSize });
          },
          neededStringIds,
        );
      }
    }

    const finalized = await finalizeStagedCatalog(sheetName, styleFormats, (progress: FinalizeProgress) => {
      const fraction = progress.totalRows ? progress.rowsRead / progress.totalRows : 1;
      onProgress({ stage: 'Salvando a base de produtos no aparelho', percent: 0.87 + fraction * 0.12, rowsProcessed: progress.rowsRead, totalRows: progress.totalRows, saved: progress.saved, skipped: progress.skipped });
    }, () => !!signal?.aborted);
    onProgress({ stage: 'Concluído', percent: 1, rowsProcessed: parsed.rowsRead, totalRows: parsed.totalRows, saved: finalized.total, skipped: finalized.skipped });
    return finalized;
  } catch (error) {
    await clearImportStaging().catch(() => {});
    if (signal?.aborted || (error as any)?.code === 'ERR_CANCELLED' || (error as Error)?.message === 'IMPORT_CANCELLED') throw new Error('IMPORT_CANCELLED');
    throw error;
  }
}

export async function cleanupWorkbookFile(source?: WorkbookFileSource | null): Promise<void> {
  if (!source) return;
  if (source.directory) await FileSystem.deleteAsync(source.directory, { idempotent: true }).catch(() => {});
  // DocumentPicker copies selected files into cache when copyToCacheDirectory=true.
  // Remove that temporary copy after import; never delete a user document outside the app cache.
  const cacheDirectory = FileSystem.cacheDirectory || '';
  if (cacheDirectory && source.uri.startsWith(cacheDirectory)) {
    await FileSystem.deleteAsync(source.uri, { idempotent: true }).catch(() => {});
  }
}
