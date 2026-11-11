import { Platform } from 'react-native';
import * as FileSystem from 'expo-file-system/legacy';
import type { CatalogItem } from './workbook';
import type { SlidePositionSettings } from './pptx';

export type LocalDatabaseInfo = { total: number; sheetName: string; sourceFile: string; ready: boolean };
export type StoredStockColumnMapping = { codeColumn: number; nameColumn: number; quantityColumn: number; headerRow: number };
export type StoredAppData = {
  catalog: CatalogItem[];
  database: LocalDatabaseInfo;
  counts: CatalogItem[];
  templateUri?: string;
  templateBase64?: string;
  templateName?: string;
  templateAdjustments?: Record<string, SlidePositionSettings>;
  themeId?: string;
  /** Persistent copy of the user's source XLSX, kept in the app's private documents directory. */
  stockWorkbookUri?: string;
  stockWorkbookName?: string;
  selectedSheetName?: string;
  stockColumnMappings?: Record<string, StoredStockColumnMapping>;
};

export const EMPTY_LOCAL_DATA: StoredAppData = {
  catalog: [],
  database: { total: 0, sheetName: '', sourceFile: '', ready: false },
  counts: [],
  templateAdjustments: {},
  themeId: 'azul',
};

const STORAGE_KEY = 'plaquetas-by-lippe-local-v1';
const FILE_NAME = 'plaquetas-by-lippe-local-v1.json';

export async function readLocalData(): Promise<StoredAppData> {
  try {
    let raw = '';
    if (Platform.OS === 'web') {
      raw = typeof localStorage !== 'undefined' ? localStorage.getItem(STORAGE_KEY) || '' : '';
    } else {
      const directory = FileSystem.documentDirectory;
      if (!directory) return { ...EMPTY_LOCAL_DATA };
      const path = directory + FILE_NAME;
      const info = await FileSystem.getInfoAsync(path);
      if (!info.exists) return { ...EMPTY_LOCAL_DATA };
      raw = await FileSystem.readAsStringAsync(path);
    }
    if (!raw) return { ...EMPTY_LOCAL_DATA };
    const parsed = JSON.parse(raw) as Partial<StoredAppData>;
    return {
      ...EMPTY_LOCAL_DATA,
      ...parsed,
      catalog: Array.isArray(parsed.catalog) ? parsed.catalog : [],
      counts: Array.isArray(parsed.counts) ? parsed.counts : [],
      database: parsed.database || EMPTY_LOCAL_DATA.database,
      templateAdjustments: parsed.templateAdjustments && typeof parsed.templateAdjustments === 'object' ? parsed.templateAdjustments : {},
      themeId: typeof parsed.themeId === 'string' ? parsed.themeId : 'azul',
      stockColumnMappings: parsed.stockColumnMappings && typeof parsed.stockColumnMappings === 'object' ? parsed.stockColumnMappings : {},
    };
  } catch (error) {
    console.warn('Não foi possível carregar os dados locais do Plaquetas:', error);
    return { ...EMPTY_LOCAL_DATA };
  }
}

export async function writeLocalData(data: StoredAppData): Promise<void> {
  const raw = JSON.stringify(data);
  if (Platform.OS === 'web') {
    if (typeof localStorage === 'undefined') throw new Error('O armazenamento local do navegador não está disponível.');
    localStorage.setItem(STORAGE_KEY, raw);
    return;
  }
  const directory = FileSystem.documentDirectory;
  if (!directory) throw new Error('O armazenamento interno do aparelho não está disponível.');
  await FileSystem.writeAsStringAsync(directory + FILE_NAME, raw);
}
