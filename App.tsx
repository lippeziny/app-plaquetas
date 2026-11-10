import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  ActivityIndicator,
  Alert,
  Platform,
  Pressable,
  ScrollView,
  StatusBar,
  StyleSheet,
  Text,
  TextInput,
  View,
  useWindowDimensions,
} from 'react-native';
import { Picker } from '@react-native-picker/picker';
import * as DocumentPicker from 'expo-document-picker';
import { CameraView, useCameraPermissions } from 'expo-camera';
import * as FileSystem from 'expo-file-system/legacy';
import { File as ExpoFile } from 'expo-file-system';
import * as Sharing from 'expo-sharing';
import { Asset } from 'expo-asset';
import * as Print from 'expo-print';
import { applyTemplateFieldMapping, buildPptx, type SlidePositionSettings, type TemplateVisual } from './offline/pptx';
import TemplatePreview from './offline/TemplatePreview';
import { buildLabelsPdfHtml } from './offline/pdf';
import { importWorkbookSheet, normalizeSearch, parseWorkbookBytes, previewWorkbookSheet, type ParsedWorkbook } from './offline/workbook';
import { EMPTY_LOCAL_DATA, readLocalData, writeLocalData, type StoredAppData } from './offline/storage';
import { Ionicons } from '@expo/vector-icons';

type TemplateField = { id: string; shape: string; sample: string };
type TemplateSlide = {
  index: number;
  capacity: number;
  widthEmu: number;
  heightEmu: number;
  orientation: 'portrait' | 'landscape';
  nameFields: TemplateField[];
  codeFields: TemplateField[];
  shapes: { id: string; shape: string; text: string; field: string | null }[];
  visuals: TemplateVisual[];
};
type TemplateInfo = { templateId: string; name: string; slides: TemplateSlide[] };
type CatalogItem = { code: string; name: string; quantity?: string; sheetName?: string };
type LabelItem = CatalogItem & { localId: string; sourceInput?: string };
type WorkbookPreview = { headers: string[]; rows: string[][]; columns: { index: number; label: string }[]; suggested: { codeColumn: number; nameColumn: number; quantityColumn: number } };
type ExportResult = { fileName: string; format: string; pages: number; capacity: number; printedItems: number; remainingItems: number };

declare const require: (path: string) => any;
const DEFAULT_TEMPLATE_ASSET = require('./assets/default-template.pptx');
const GREEN = '#17734F';
const INK = '#17251F';
const MUTED = '#75827B';
const LINE = '#E2E9E4';

type AppTheme = { id: string; name: string; primary: string; onPrimary: string; primaryContainer: string; onPrimaryContainer: string; secondary: string; tertiary: string; background: string; surface: string; surfaceContainer: string; surfaceVariant: string; outline: string; };
const APP_THEMES: AppTheme[] = [
  { id: 'azul', name: 'Azul', primary: '#315B9A', onPrimary: '#FFFFFF', primaryContainer: '#D7E3FF', onPrimaryContainer: '#001B3F', secondary: '#565F71', tertiary: '#715573', background: '#F7F8FF', surface: '#FFFFFF', surfaceContainer: '#EEF1F8', surfaceVariant: '#E1E5EF', outline: '#747984' },
  { id: 'verde', name: 'Verde', primary: '#386A4A', onPrimary: '#FFFFFF', primaryContainer: '#C9EBD0', onPrimaryContainer: '#00210F', secondary: '#506352', tertiary: '#496879', background: '#F5FAF5', surface: '#FFFFFF', surfaceContainer: '#EAF2E9', surfaceVariant: '#DDE7DC', outline: '#727970' },
  { id: 'roxo', name: 'Roxo', primary: '#65558F', onPrimary: '#FFFFFF', primaryContainer: '#EADDFF', onPrimaryContainer: '#210D4A', secondary: '#625B71', tertiary: '#7D5260', background: '#FBF7FF', surface: '#FFFFFF', surfaceContainer: '#F1ECF8', surfaceVariant: '#E7E0F0', outline: '#797583' },
  { id: 'laranja', name: 'Âmbar', primary: '#805600', onPrimary: '#FFFFFF', primaryContainer: '#FFDEA5', onPrimaryContainer: '#291800', secondary: '#6B5D40', tertiary: '#8B4B3A', background: '#FFFAF2', surface: '#FFFFFF', surfaceContainer: '#F6EFE2', surfaceVariant: '#EDE2CF', outline: '#817668' },
  { id: 'turquesa', name: 'Turquesa', primary: '#006A69', onPrimary: '#FFFFFF', primaryContainer: '#9CF1EE', onPrimaryContainer: '#00201F', secondary: '#4A6363', tertiary: '#4C5E7F', background: '#F2FAF9', surface: '#FFFFFF', surfaceContainer: '#E6F2F1', surfaceVariant: '#D9E7E5', outline: '#6F7978' },
];
const EMPTY_POSITION_SETTINGS: SlidePositionSettings = { global: {}, slots: {} };

function showMessage(message: string) {
  if (Platform.OS === 'web') {
    if (typeof window !== 'undefined') window.alert(message);
    else Alert.alert('Plaquetas', message);
  } else Alert.alert('Plaquetas', message);
}
function makeId() { return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`; }
function splitValues(text: string) { return text.split(/[\n,;\t]+/).map(v => v.trim()).filter(Boolean); }
function parseShapeIds(value: string) { return value.split(/[\s,;]+/).map(v => v.trim()).filter(Boolean); }
function bytesToBase64(bytes: Uint8Array): string {
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  let output = '';
  for (let index = 0; index < bytes.length; index += 3) {
    const a = bytes[index];
    const hasB = index + 1 < bytes.length;
    const hasC = index + 2 < bytes.length;
    const b = hasB ? bytes[index + 1] : 0;
    const c = hasC ? bytes[index + 2] : 0;
    output += alphabet[a >> 2];
    output += alphabet[((a & 3) << 4) | (b >> 4)];
    output += hasB ? alphabet[((b & 15) << 2) | (c >> 6)] : '=';
    output += hasC ? alphabet[c & 63] : '=';
  }
  return output;
}
async function readPickedFileBytes(asset: any): Promise<Uint8Array> {
  if (Platform.OS === 'web' && asset.file?.arrayBuffer) {
    return new Uint8Array(await asset.file.arrayBuffer());
  }
  // Expo File.bytes() retorna Uint8Array sem criar uma string Base64 gigante
  // pela ponte nativa do Android.
  const file = new ExpoFile(asset.uri);
  return await file.bytes();
}
async function readTemplateAssetBase64(): Promise<string> {
  const asset = Asset.fromModule(DEFAULT_TEMPLATE_ASSET);
  await asset.downloadAsync();
  const uri = asset.localUri || asset.uri;
  if (Platform.OS === 'web') {
    const response = await fetch(uri);
    if (!response.ok) throw new Error('Não foi possível carregar o modelo PowerPoint incluído no aplicativo.');
    return bytesToBase64(new Uint8Array(await response.arrayBuffer()));
  }
  return FileSystem.readAsStringAsync(uri, { encoding: FileSystem.EncodingType.Base64 });
}
async function persistPickedTemplate(asset: any, base64: string, fileName: string): Promise<Partial<StoredAppData>> {
  if (Platform.OS === 'web') return { templateBase64: base64, templateUri: undefined, templateName: fileName };
  const root = FileSystem.documentDirectory;
  if (!root) throw new Error('O armazenamento interno do aparelho não está disponível.');
  const directory = `${root}templates/`;
  await FileSystem.makeDirectoryAsync(directory, { intermediates: true }).catch(() => {});
  const safeName = `modelo-${Date.now()}.pptx`;
  const destination = directory + safeName;
  await FileSystem.copyAsync({ from: asset.uri, to: destination });
  return { templateUri: destination, templateBase64: undefined, templateName: fileName };
}
async function saveAndShareBase64(base64: string, fileName: string, mimeType: string) {
  if (Platform.OS === 'web') {
    const uri = `data:${mimeType};base64,${base64}`;
    const anchor = document.createElement('a');
    anchor.href = uri;
    anchor.download = fileName;
    anchor.click();
    return;
  }
  const directory = FileSystem.cacheDirectory || FileSystem.documentDirectory;
  if (!directory) throw new Error('Não foi possível acessar o armazenamento temporário do aparelho.');
  const uri = `${directory}${fileName}`;
  await FileSystem.writeAsStringAsync(uri, base64, { encoding: FileSystem.EncodingType.Base64 });
  if (await Sharing.isAvailableAsync()) {
    await Sharing.shareAsync(uri, { mimeType, dialogTitle: `Compartilhar ${fileName}` });
  } else {
    showMessage(`Arquivo criado em: ${uri}`);
  }
}

async function chooseFile(type: string) {
  const result = await DocumentPicker.getDocumentAsync({ type, copyToCacheDirectory: true, multiple: false });
  if (result.canceled || !result.assets?.length) return null;
  return result.assets[0];
}

function SectionTitle({ eyebrow, title, subtitle, accent = GREEN, titleColor = INK }: { eyebrow: string; title: string; subtitle?: string; accent?: string; titleColor?: string }) {
  return <View style={styles.sectionTitle}>
    <Text style={[styles.eyebrow, { color: accent }]}>{eyebrow}</Text>
    <Text style={[styles.sectionHeading, { color: titleColor }]}>{title}</Text>
    {!!subtitle && <Text style={styles.sectionSub}>{subtitle}</Text>}
  </View>;
}
function ActionButton({ title, onPress, icon, secondary = false, disabled = false, small = false, primaryColor = GREEN, secondaryBackground, textColor }: { title: string; onPress: () => void; icon?: any; secondary?: boolean; disabled?: boolean; small?: boolean; primaryColor?: string; secondaryBackground?: string; textColor?: string }) {
  return <Pressable onPress={onPress} disabled={disabled} style={({ pressed }) => [styles.button, secondary ? styles.buttonSecondary : styles.buttonPrimary, !secondary && { backgroundColor: primaryColor }, secondary && { backgroundColor: secondaryBackground || `${primaryColor}14`, borderColor: `${primaryColor}45` }, small && styles.buttonSmall, disabled && styles.buttonDisabled, pressed && !disabled && styles.buttonPressed]}>
    {!!icon && <Ionicons name={icon} size={small ? 15 : 17} color={secondary ? primaryColor : (textColor || '#FFFFFF')} />}
    <Text style={[styles.buttonText, secondary && { color: primaryColor }, !secondary && textColor ? { color: textColor } : null, small && styles.buttonTextSmall]}>{title}</Text>
  </Pressable>;
}
function Card({ children, style }: { children: React.ReactNode; style?: any }) {
  return <View style={[styles.card, style]}>{children}</View>;
}

function CameraScanner({ onScan, onClose, accent = GREEN }: { onScan: (data: string) => void; onClose: () => void; accent?: string }) {
  const [cameraPermission, requestCameraPermission] = useCameraPermissions();
  return (
    <View style={styles.cameraContainer}>
      {!cameraPermission?.granted ? (
        <View style={styles.cameraPermission}>
          <Ionicons name="camera-outline" size={28} color={accent} />
          <Text style={styles.helper}>A câmera precisa de permissão.</Text>
          <ActionButton title="Permitir câmera" onPress={async () => { const result = await requestCameraPermission(); if (!result.granted) showMessage('A permissão da câmera não foi concedida.'); }} primaryColor={accent} />
        </View>
      ) : (
        <CameraView style={styles.camera} facing="back" onBarcodeScanned={({ data }) => onScan(data)} barcodeScannerSettings={{ barcodeTypes: ['qr', 'ean13', 'ean8', 'code128', 'code39', 'upc_a', 'upc_e'] }} />
      )}
      <ActionButton title="Fechar leitor" onPress={onClose} secondary small primaryColor={accent} />
    </View>
  );
}

function AppScreen() {
  const { width } = useWindowDimensions();
  const [localReady, setLocalReady] = useState(false);
  const [busy, setBusy] = useState('');
  const [tab, setTab] = useState<'labels' | 'count'>('labels');
  const [workflowStep, setWorkflowStep] = useState<'data' | 'items' | 'layout'>('data');
  const [themeId, setThemeId] = useState('azul');
  const [themePickerOpen, setThemePickerOpen] = useState(false);
  const [templateAdjustments, setTemplateAdjustments] = useState<Record<string, SlidePositionSettings>>({});
  const [adjustScope, setAdjustScope] = useState<'global' | 'single'>('global');
  const [adjustField, setAdjustField] = useState<'code' | 'name'>('name');
  const [adjustSlotIndex, setAdjustSlotIndex] = useState(0);
  const [adjustStep, setAdjustStep] = useState(1);
  const [database, setDatabase] = useState({ total: 0, sheetName: '', sourceFile: '', ready: false });
  const [databaseFile, setDatabaseFile] = useState('');
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [sheetName, setSheetName] = useState('');
  const [workbookPreview, setWorkbookPreview] = useState<WorkbookPreview | null>(null);
  const [workbookData, setWorkbookData] = useState<ParsedWorkbook | null>(null);
  const [codeColumn, setCodeColumn] = useState(1);
  const [nameColumn, setNameColumn] = useState(2);
  const [quantityColumn, setQuantityColumn] = useState(6);
  const [catalog, setCatalog] = useState<CatalogItem[]>([]);
  const [template, setTemplate] = useState<TemplateInfo | null>(null);
  const [templateBase64, setTemplateBase64] = useState('');
  const [templateSlide, setTemplateSlide] = useState(0);
  const [customCodeIds, setCustomCodeIds] = useState('');
  const [customNameIds, setCustomNameIds] = useState('');
  const [templateError, setTemplateError] = useState('');
  const [inputMode, setInputMode] = useState<'code' | 'name'>('code');
  const [inputText, setInputText] = useState('');
  const [suggestions, setSuggestions] = useState<CatalogItem[]>([]);
  const [labels, setLabels] = useState<LabelItem[]>([]);
  const [activeLabelIndex, setActiveLabelIndex] = useState(0);
  const [compactPreview, setCompactPreview] = useState(false);
  const [requestedPages, setRequestedPages] = useState('');
  const [lastExport, setLastExport] = useState<ExportResult | null>(null);
  const [scanEnabled, setScanEnabled] = useState(false);
  const [scanCode, setScanCode] = useState('');
  const [scanName, setScanName] = useState('');
  const [scanStep, setScanStep] = useState<1 | 10>(1);
  const [counts, setCounts] = useState<CatalogItem[]>([]);
  const lastScan = useRef({ code: '', at: 0 });
  const localDataRef = useRef<StoredAppData>(EMPTY_LOCAL_DATA);

  const persistPatch = useCallback(async (patch: Partial<StoredAppData>) => {
    const next = { ...localDataRef.current, ...patch };
    localDataRef.current = next;
    await writeLocalData(next);
  }, []);

  const selectedSlide = template?.slides?.[templateSlide] || template?.slides?.[0];
  const manualCapacity = Math.min(customCodeIds.split(/[\s,;]+/).filter(Boolean).length, customNameIds.split(/[\s,;]+/).filter(Boolean).length);
  const capacity = selectedSlide ? (selectedSlide.capacity || manualCapacity || 0) : 0;
  const requiredPages = Math.max(1, Math.ceil(labels.length / Math.max(capacity, 1)));
  const activeLabel = labels[activeLabelIndex] || labels[0];
  const validCustomFields = customCodeIds.trim().length > 0 && customNameIds.trim().length > 0;
  const theme = APP_THEMES.find(option => option.id === themeId) || APP_THEMES[0];
  const adjustmentKey = `${template?.templateId || 'local'}:${template?.name || 'modelo-padrao'}:${templateSlide}`;
  const storedAdjustments = templateAdjustments[adjustmentKey] || EMPTY_POSITION_SETTINGS;
  const activeAdjustments: SlidePositionSettings = { global: storedAdjustments.global || {}, slots: storedAdjustments.slots || {} };
  const selectedFieldDelta = adjustScope === 'global'
    ? activeAdjustments.global[adjustField]
    : activeAdjustments.slots[String(adjustSlotIndex)]?.[adjustField];
  const previewSlide = useMemo(() => {
    if (!selectedSlide) return undefined;
    if (!validCustomFields) return selectedSlide;
    return applyTemplateFieldMapping(selectedSlide, parseShapeIds(customCodeIds), parseShapeIds(customNameIds));
  }, [selectedSlide, validCustomFields, customCodeIds, customNameIds]);

  const selectTheme = async (id: string) => {
    setThemeId(id);
    setThemePickerOpen(false);
    await persistPatch({ themeId: id });
  };
  const updateTemplatePosition = async (dx: number, dy: number) => {
    const current = templateAdjustments[adjustmentKey] || EMPTY_POSITION_SETTINGS;
    const target = adjustScope === 'global'
      ? { ...(current.global[adjustField] || { xMm: 0, yMm: 0 }), xMm: (current.global[adjustField]?.xMm || 0) + dx, yMm: (current.global[adjustField]?.yMm || 0) + dy }
      : { ...(current.slots[String(adjustSlotIndex)]?.[adjustField] || { xMm: 0, yMm: 0 }), xMm: (current.slots[String(adjustSlotIndex)]?.[adjustField]?.xMm || 0) + dx, yMm: (current.slots[String(adjustSlotIndex)]?.[adjustField]?.yMm || 0) + dy };
    const next: SlidePositionSettings = adjustScope === 'global'
      ? { ...current, global: { ...current.global, [adjustField]: target } }
      : { ...current, slots: { ...current.slots, [String(adjustSlotIndex)]: { ...current.slots[String(adjustSlotIndex)], [adjustField]: target } } };
    const all = { ...templateAdjustments, [adjustmentKey]: next };
    setTemplateAdjustments(all);
    await persistPatch({ templateAdjustments: all });
  };
  const resetTemplatePositions = async () => {
    const all = { ...templateAdjustments };
    delete all[adjustmentKey];
    setTemplateAdjustments(all);
    await persistPatch({ templateAdjustments: all });
  };

  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        const saved = await readLocalData();
        let base64 = saved.templateBase64 || '';
        if (!base64 && saved.templateUri) {
          try { base64 = await FileSystem.readAsStringAsync(saved.templateUri, { encoding: FileSystem.EncodingType.Base64 }); }
          catch (error) { console.warn('Modelo local salvo não pôde ser reaberto. Será usado o modelo padrão.', error); }
        }
        if (!base64) base64 = await readTemplateAssetBase64();
        const info = await (await import('./offline/pptx')).inspectPptx(base64, saved.templateName || 'PLAQUETAS_BY_FILLIPE.PPTX');
        if (!alive) return;
        localDataRef.current = saved;
        setCatalog(saved.catalog);
        setDatabase(saved.database);
        setTemplateAdjustments(saved.templateAdjustments || {});
        setThemeId(APP_THEMES.some(theme => theme.id === saved.themeId) ? (saved.themeId as string) : 'azul');
        setDatabaseFile(saved.database.sourceFile || '');
        setCounts(saved.counts);
        setTemplate(info);
        setTemplateBase64(base64);
        setLocalReady(true);
        if (saved.templateUri && !saved.templateBase64 && !saved.templateName) {
          await persistPatch({ templateName: info.name });
        }
      } catch (error: any) {
        if (alive) showMessage(`Não foi possível preparar o aplicativo offline: ${error.message || error}`);
      }
    })();
    return () => { alive = false; };
  }, [persistPatch]);

  useEffect(() => {
    if (inputMode !== 'name' || inputText.trim().length < 2 || !localReady) { setSuggestions([]); return; }
    const query = normalizeSearch(inputText.split(/[\n,;]+/).pop() || '');
    if (!query) { setSuggestions([]); return; }
    setSuggestions(catalog.filter(item => normalizeSearch(item.name).includes(query)).slice(0, 8));
  }, [inputText, inputMode, catalog, localReady]);

  const refreshCounts = useCallback(async () => {
    const stored = await readLocalData();
    localDataRef.current = stored;
    setCounts(stored.counts || []);
  }, []);

  const loadSheetPreview = useCallback(async (name: string) => {
    if (!workbookData || !name) return;
    try {
      const preview = previewWorkbookSheet(workbookData, name, 5);
      setWorkbookPreview(preview);
      setCodeColumn(preview.suggested.codeColumn || 1);
      setNameColumn(preview.suggested.nameColumn || 2);
      setQuantityColumn(preview.suggested.quantityColumn || 0);
    } catch (error: any) { showMessage(error.message); }
  }, [workbookData]);

  const onChooseDatabase = async () => {
    try {
      const asset = await chooseFile('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      if (!asset) return;
      const fileSize = Number(asset.size || 0);
      const maxImportBytes = 50 * 1024 * 1024;
      if (fileSize > maxImportBytes) {
        return showMessage(`Esta planilha tem ${(fileSize / 1024 / 1024).toFixed(1)} MB. Para evitar falta de memória no celular, reduza o arquivo para até 50 MB: mantenha apenas as abas necessárias e remova imagens/objetos incorporados. Depois tente novamente.`);
      }
      setBusy('Lendo planilha no aparelho...');
      const bytes = await readPickedFileBytes(asset);
      const parsed = parseWorkbookBytes(bytes);
      setWorkbookData(parsed);
      setDatabaseFile(asset.name || 'planilha.xlsx');
      setSheetNames(parsed.sheetNames);
      const first = parsed.sheetNames[0] || '';
      setSheetName(first);
      const preview = previewWorkbookSheet(parsed, first, 5);
      setWorkbookPreview(preview);
      setCodeColumn(preview.suggested.codeColumn || 1);
      setNameColumn(preview.suggested.nameColumn || 2);
      setQuantityColumn(preview.suggested.quantityColumn || 0);
    } catch (error: any) { showMessage(error.message || 'Não foi possível ler essa planilha.'); }
    finally { setBusy(''); }
  };

  const onImportDatabase = async () => {
    if (!workbookData || !sheetName) return showMessage('Selecione uma planilha e uma aba primeiro.');
    if (!codeColumn || !nameColumn) return showMessage('Selecione as colunas de código e nome antes de importar.');
    setBusy('Salvando a base no aparelho...');
    try {
      const result = importWorkbookSheet(workbookData, sheetName, { codeColumn, nameColumn, quantityColumn }, databaseFile || 'planilha.xlsx');
      const meta = { total: result.items.length, ready: result.items.length > 0, sheetName, sourceFile: databaseFile || 'planilha.xlsx' };
      await persistPatch({ catalog: result.items, database: meta });
      setCatalog(result.items);
      setDatabase(meta);
      setWorkbookData(null);
      setSheetNames([]);
      setWorkbookPreview(null);
      setWorkflowStep('items');
      showMessage(`${result.items.length.toLocaleString('pt-BR')} itens importados da aba ${sheetName} e salvos neste aparelho.${result.skipped ? ` ${result.skipped} linha(s) sem código ou nome foram ignoradas.` : ''}`);
    } catch (error: any) { showMessage(error.message); }
    finally { setBusy(''); }
  };

  const onChooseTemplate = async () => {
    try {
      const asset = await chooseFile('application/vnd.openxmlformats-officedocument.presentationml.presentation');
      if (!asset) return;
      setBusy('Analisando modelo PowerPoint localmente...');
      const base64 = await readPickedFileBase64(asset);
      const { inspectPptx } = await import('./offline/pptx');
      const result = await inspectPptx(base64, asset.name || 'Modelo personalizado.pptx');
      const savedTemplate = await persistPickedTemplate(asset, base64, asset.name || 'Modelo personalizado.pptx');
      await persistPatch(savedTemplate);
      setTemplate(result);
      setTemplateBase64(base64);
      setTemplateSlide(0);
      setCustomCodeIds('');
      setCustomNameIds('');
      const detected = (result.slides?.[0]?.capacity || 0) > 0;
      setTemplateError(detected ? '' : 'Não identifiquei automaticamente as caixas de código e nome. Informe os IDs das caixas no mapeamento avançado.');
    } catch (error: any) { showMessage(error.message || 'Não foi possível abrir esse modelo PowerPoint.'); }
    finally { setBusy(''); }
  };

  const addItems = async () => {
    const values = splitValues(inputText);
    if (!values.length) return showMessage('Cole ou digite ao menos um código ou nome.');
    if (values.length > 1000) return showMessage('Adicione até 1.000 itens por vez.');
    setBusy('Conferindo itens na base local...');
    try {
      const entries = values.map(input => {
        const key = normalizeSearch(input);
        const item = inputMode === 'name'
          ? catalog.find(row => normalizeSearch(row.name) === key)
          : catalog.find(row => normalizeSearch(row.code) === key);
        return { input, item: item || null };
      });
      const mapped: LabelItem[] = entries.map(entry => ({
        localId: makeId(),
        code: entry.item?.code || (inputMode === 'code' ? entry.input : ''),
        name: entry.item?.name || (inputMode === 'name' ? entry.input : ''),
        quantity: entry.item?.quantity || '',
        sheetName: entry.item?.sheetName || database.sheetName,
        sourceInput: entry.input,
      }));
      setLabels(previous => [...previous, ...mapped]);
      setActiveLabelIndex(labels.length);
      setInputText('');
      setSuggestions([]);
      const missing = entries.filter(entry => !entry.item).length;
      if (missing) showMessage(`${mapped.length - missing} item(ns) encontrado(s) na base local. ${missing} precisa(m) de revisão manual.`);
    } catch (error: any) { showMessage(error.message); }
    finally { setBusy(''); }
  };

  const chooseSuggestion = (item: CatalogItem) => {
    const before = inputText.split(/[\n,;]+/).slice(0, -1).join('\n');
    setInputText(`${before ? `${before}\n` : ''}${item.name}`);
    setSuggestions([]);
  };

  const updateLabel = (localId: string, field: 'code' | 'name', value: string) => {
    setLabels(previous => previous.map(item => item.localId === localId ? { ...item, [field]: value } : item));
  };
  const removeLabel = (localId: string) => {
    setLabels(previous => previous.filter(item => item.localId !== localId));
    setActiveLabelIndex(0);
  };
  const addShapeId = (field: 'code' | 'name', id: string) => {
    const current = (field === 'code' ? customCodeIds : customNameIds).split(/[\s,;]+/).filter(Boolean);
    if (!current.includes(String(id))) current.push(String(id));
    const value = current.join(', ');
    if (field === 'code') setCustomCodeIds(value); else setCustomNameIds(value);
  };
  const useFields = () => {
    if (!validCustomFields) return showMessage('Informe a lista de IDs das caixas de código e de nome.');
    if (!selectedSlide) return showMessage('Carregue um modelo PowerPoint primeiro.');
    const known = new Set(selectedSlide.shapes.map(shape => String(shape.id)));
    const codeIds = parseShapeIds(customCodeIds);
    const nameIds = parseShapeIds(customNameIds);
    const unknown = [...codeIds, ...nameIds].filter(id => !known.has(id));
    if (unknown.length) return showMessage(`Estes IDs não foram encontrados no slide: ${[...new Set(unknown)].join(', ')}.`);
    if (new Set(codeIds).size !== codeIds.length || new Set(nameIds).size !== nameIds.length) return showMessage('Remova os IDs repetidos do mapeamento.');
    if (codeIds.some(id => nameIds.includes(id))) return showMessage('Uma mesma caixa não pode ser usada como código e nome.');
    setTemplateError('Mapeamento manual aplicado. A prévia já usa os campos selecionados.');
  };

  const exportFile = async (format: 'pptx' | 'pdf') => {
    if (!labels.length) return showMessage('Adicione itens antes de gerar as plaquetas.');
    if (!template || !templateBase64) return showMessage('Nenhum modelo PowerPoint está configurado.');
    const fields = validCustomFields ? { codeShapeIds: parseShapeIds(customCodeIds), nameShapeIds: parseShapeIds(customNameIds) } : {};
    setBusy(format === 'pdf' ? 'Montando PDF neste aparelho...' : 'Montando PowerPoint neste aparelho...');
    try {
      const items = labels.map(({ code, name }) => ({ code, name }));
      const minimumPages = Math.max(1, Math.ceil(items.length / Math.max(capacity, 1)));
      const pageCount = Math.max(minimumPages, Math.min(100, Number(requestedPages) || minimumPages));
      const slide = selectedSlide;
      if (!slide) throw new Error('O slide/modelo escolhido não existe no arquivo.');
      if (format === 'pptx') {
        const result = await buildPptx(templateBase64, templateSlide, items, requestedPages ? Number(requestedPages) : undefined, fields, activeAdjustments);
        const fileName = `plaquetas-by-lippe-${Date.now()}.pptx`;
        setLastExport({ fileName, format: 'pptx', pages: result.pages, capacity: result.capacity, printedItems: result.printedItems, remainingItems: result.remainingItems });
        await saveAndShareBase64(result.base64, fileName, 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
        if (result.remainingItems > 0) showMessage(`${result.printedItems} plaquetas incluídas. ${result.remainingItems} ficaram fora do total de folhas escolhido.`);
      } else {
        const html = buildLabelsPdfHtml(items, { pages: pageCount, capacity: slide.capacity || manualCapacity || 8, orientation: slide.orientation, fileTitle: 'Plaquetas by Lippe', adjustments: activeAdjustments });
        if (Platform.OS === 'web') {
          const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
          const url = URL.createObjectURL(blob);
          window.open(url, '_blank');
          setLastExport({ fileName: 'plaquetas-by-lippe.html (imprimir como PDF)', format: 'pdf', pages: pageCount, capacity: slide.capacity || manualCapacity, printedItems: Math.min(items.length, pageCount * capacity), remainingItems: Math.max(0, items.length - pageCount * capacity) });
          showMessage('A visualização foi aberta. Use a opção Imprimir do navegador e escolha Salvar como PDF.');
        } else {
          const pdf = await Print.printToFileAsync({ html, width: slide.orientation === 'portrait' ? 595 : 842, height: slide.orientation === 'portrait' ? 842 : 595, margins: { left: 0, top: 0, right: 0, bottom: 0 } });
          const fileName = `plaquetas-by-lippe-${Date.now()}.pdf`;
          setLastExport({ fileName, format: 'pdf', pages: pageCount, capacity: slide.capacity || manualCapacity, printedItems: Math.min(items.length, pageCount * capacity), remainingItems: Math.max(0, items.length - pageCount * capacity) });
          if (await Sharing.isAvailableAsync()) await Sharing.shareAsync(pdf.uri, { mimeType: 'application/pdf', dialogTitle: `Compartilhar ${fileName}` });
          else showMessage(`PDF criado: ${pdf.uri}`);
        }
      }
    } catch (error: any) { showMessage(error.message || 'Não foi possível gerar o arquivo.'); }
    finally { setBusy(''); }
  };

  const onScan = async (data: string) => {
    const now = Date.now();
    if (!data || (lastScan.current.code === data && now - lastScan.current.at < 1800)) return;
    lastScan.current = { code: data, at: now };
    setScanCode(data);
    const key = normalizeSearch(data);
    const found = catalog.find(item => normalizeSearch(item.code) === key) || catalog.find(item => normalizeSearch(item.code).includes(key));
    setScanName(found?.name || 'Item não encontrado na base local');
  };

  const incrementCount = async () => {
    const code = scanCode.trim();
    if (!code) return showMessage('Leia um código ou digite o código do produto.');
    try {
      setBusy('Salvando contagem local...');
      const key = normalizeSearch(code);
      const existing = counts.find(item => normalizeSearch(item.code) === key);
      const found = catalog.find(item => normalizeSearch(item.code) === key);
      const current = Number(String(existing?.quantity || '0').replace(',', '.')) || 0;
      const nextItem: CatalogItem = {
        code,
        name: found?.name || scanName || existing?.name || 'Item sem descrição',
        quantity: String(current + scanStep),
      };
      const nextCounts = existing
        ? counts.map(item => normalizeSearch(item.code) === key ? nextItem : item)
        : [...counts, nextItem];
      nextCounts.sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));
      await persistPatch({ counts: nextCounts });
      setCounts(nextCounts);
    } catch (error: any) { showMessage(error.message || 'Não foi possível salvar a contagem local.'); }
    finally { setBusy(''); }
  };

  const pageLabel = useMemo(() => `${requiredPages} ${requiredPages === 1 ? 'folha' : 'folhas'}`, [requiredPages]);

  return <View style={[styles.screen, { backgroundColor: theme.background }]}>
    <StatusBar barStyle="dark-content" backgroundColor={theme.background} />
    <View style={[styles.topBar, { backgroundColor: theme.surface, borderBottomColor: theme.surfaceVariant }]}>
      <View style={[styles.brandMark, { backgroundColor: theme.primary, borderRadius: 15 }]}><Ionicons name="pricetag" size={22} color={theme.onPrimary} /></View>
      <View style={{ flex: 1 }}>
        <Text style={styles.brandTitle}>PLAQUETAS <Text style={styles.brandMuted}>by lippe</Text></Text>
        <Text style={styles.brandSubtitle}>SEU ESTOQUE · NO PRÓPRIO APARELHO</Text>
      </View>
      <View style={[styles.offlineTag, { backgroundColor: theme.primaryContainer }]}><Ionicons name="phone-portrait-outline" size={13} color={theme.primary} /><Text style={[styles.offlineTagText, { color: theme.primary }]}>OFFLINE</Text></View>
      <Pressable onPress={() => setThemePickerOpen(value => !value)} style={[styles.themeOpenButton, { backgroundColor: theme.surfaceContainer, borderColor: theme.surfaceVariant }]}>
        <Ionicons name="color-palette-outline" size={20} color={theme.primary} />
      </Pressable>
    </View>
    {themePickerOpen && <View style={[styles.themePickerPanel, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}>
      <View style={{ flex: 1, gap: 2 }}><Text style={styles.fieldLabel}>Aparência</Text><Text style={styles.helper}>Escolha uma paleta clara Material You</Text></View>
      <View style={styles.themeSwatches}>{APP_THEMES.map(option => <Pressable key={option.id} onPress={() => selectTheme(option.id)} accessibilityLabel={`Tema ${option.name}`} style={[styles.themeSwatch, { backgroundColor: option.primary, borderColor: themeId === option.id ? INK : option.primary }]}>{themeId === option.id ? <Ionicons name="checkmark" size={17} color="#FFFFFF" /> : null}</Pressable>)}</View>
    </View>}
    <View style={[styles.tabBar, { backgroundColor: theme.surface, borderBottomColor: theme.surfaceVariant }]}>
      <Pressable onPress={() => setTab('labels')} style={[styles.tab, tab === 'labels' && { backgroundColor: theme.primaryContainer, borderRadius: 18 }]}>
        <Ionicons name="pricetags-outline" size={18} color={tab === 'labels' ? theme.primary : MUTED} />
        <Text style={[styles.tabText, tab === 'labels' && { color: theme.primary }]}>Criar plaquetas</Text>
      </Pressable>
      <Pressable onPress={() => setTab('count')} style={[styles.tab, tab === 'count' && { backgroundColor: theme.primaryContainer, borderRadius: 18 }]}>
        <Ionicons name="scan-outline" size={18} color={tab === 'count' ? theme.primary : MUTED} />
        <Text style={[styles.tabText, tab === 'count' && { color: theme.primary }]}>Contagem</Text>
      </Pressable>
    </View>

    {tab === 'labels' ? <ScrollView contentContainerStyle={[styles.content, { paddingBottom: 36 }]} keyboardShouldPersistTaps="handled">
      <View style={[styles.heroRow, { backgroundColor: theme.primaryContainer, borderRadius: 24, padding: 18, marginBottom: 14 }]}>
        <View style={{ flex: 1, gap: 5 }}>
          <Text style={[styles.heroEyebrow, { color: theme.primary }]}>CRIADOR DE PLAQUETAS · 100% LOCAL</Text>
          <Text style={[styles.heroTitle, { color: theme.onPrimaryContainer }]}>Vamos preparar suas plaquetas.</Text>
          <Text style={[styles.heroText, { color: theme.onPrimaryContainer, opacity: 0.78 }]}>Importe a base, escolha os itens, confira o modelo e exporte. Tudo funciona sem servidor.</Text>
        </View>
        <View style={[styles.heroIcon, { backgroundColor: `${theme.surface}B8` }]}><Ionicons name="pricetags" size={32} color={theme.primary} /></View>
      </View>

      <View style={styles.statRow}>
        <View style={[styles.statCard, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}><Text style={styles.statLabel}>NA BASE</Text><Text style={[styles.statValue, { color: theme.primary }]}>{database.total.toLocaleString('pt-BR')}</Text><Text style={styles.statFoot}>{database.ready ? 'Materiais salvos' : 'Importe o Excel'}</Text></View>
        <View style={[styles.statCard, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}><Text style={styles.statLabel}>NA LISTA</Text><Text style={[styles.statValue, { color: theme.primary }]}>{labels.length}</Text><Text style={styles.statFoot}>Plaquetas preparadas</Text></View>
        <View style={[styles.statCard, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}><Text style={styles.statLabel}>FOLHAS</Text><Text style={[styles.statValue, { color: theme.primary }]}>{labels.length ? requiredPages : '—'}</Text><Text style={styles.statFoot}>{capacity} por folha</Text></View>
      </View>

      <View style={[styles.workflowTabs, { backgroundColor: theme.surfaceContainer, borderColor: theme.surfaceVariant }]}>
        {([
          { key: 'data', number: '1', title: 'Base', icon: 'file-tray-full-outline' },
          { key: 'items', number: '2', title: 'Itens', icon: 'list-outline' },
          { key: 'layout', number: '3', title: 'Modelo', icon: 'options-outline' },
        ] as const).map(step => {
          const active = workflowStep === step.key;
          const done = (step.key === 'data' && database.ready) || (step.key === 'items' && labels.length > 0);
          return <Pressable key={step.key} onPress={() => setWorkflowStep(step.key)} style={[styles.workflowStep, active && { backgroundColor: theme.surface, borderColor: theme.primary }, !active && { borderColor: 'transparent' }]}>
            <View style={[styles.workflowNumber, { backgroundColor: active ? theme.primary : done ? theme.primaryContainer : theme.surfaceVariant }]}><Text style={{ color: active ? theme.onPrimary : theme.primary, fontSize: 12, fontWeight: '800' }}>{done && !active ? '✓' : step.number}</Text></View>
            <Text style={[styles.workflowStepText, { color: active ? theme.primary : MUTED }]}>{step.title}</Text>
          </Pressable>;
        })}
      </View>

      {workflowStep === 'data' && <Card style={[styles.materialCard, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}>
        <SectionTitle eyebrow="PASSO 1 DE 3" title="Sua base de materiais" subtitle="A planilha fica armazenada apenas neste aparelho. Não é enviada para nenhum servidor." accent={theme.primary} />
        <View style={[styles.materialFileRow, { backgroundColor: theme.surfaceContainer, borderColor: theme.surfaceVariant }]}>
          <View style={[styles.materialIconCircle, { backgroundColor: theme.primaryContainer }]}><Ionicons name="document-text-outline" size={24} color={theme.primary} /></View>
          <View style={{ flex: 1, gap: 3 }}>
            <Text style={styles.uploadTitle}>{databaseFile || database.sourceFile || 'Nenhuma planilha selecionada'}</Text>
            <Text style={styles.helper}>{database.ready ? `${database.total.toLocaleString('pt-BR')} registros · ${database.sheetName}` : 'Planilha Excel .xlsx com uma ou mais abas'}</Text>
          </View>
        </View>
        <ActionButton title={database.ready ? 'Trocar planilha Excel' : 'Selecionar planilha Excel'} onPress={onChooseDatabase} icon="folder-open-outline" primaryColor={theme.primary} disabled={!!busy} />
        {sheetNames.length > 0 && <View style={styles.mappingBlock}>
          <Text style={styles.fieldLabel}>1. Escolha a aba da planilha</Text>
          <View style={[styles.pickerWrap, { backgroundColor: theme.surfaceContainer, borderColor: theme.surfaceVariant }]}><Picker selectedValue={sheetName} onValueChange={(value) => { setSheetName(value); loadSheetPreview(value); }} style={styles.picker}><Picker.Item label="Selecione a aba" value="" />{sheetNames.map(name => <Picker.Item key={name} label={name} value={name} />)}</Picker></View>
          {workbookPreview && <>
            <Text style={styles.fieldLabel}>2. Confira as colunas</Text>
            <View style={[styles.mappingRow, width < 550 && styles.mappingRowMobile]}>
              <View style={styles.mappingField}><Text style={styles.miniLabel}>CÓDIGO</Text><View style={[styles.pickerWrap, { backgroundColor: theme.surfaceContainer, borderColor: theme.surfaceVariant }]}><Picker selectedValue={codeColumn} onValueChange={setCodeColumn} style={styles.picker}>{workbookPreview.columns.map(c => <Picker.Item key={c.index} label={`${c.index}. ${c.label}`} value={c.index} />)}</Picker></View></View>
              <View style={styles.mappingField}><Text style={styles.miniLabel}>NOME / DESCRIÇÃO</Text><View style={[styles.pickerWrap, { backgroundColor: theme.surfaceContainer, borderColor: theme.surfaceVariant }]}><Picker selectedValue={nameColumn} onValueChange={setNameColumn} style={styles.picker}>{workbookPreview.columns.map(c => <Picker.Item key={c.index} label={`${c.index}. ${c.label}`} value={c.index} />)}</Picker></View></View>
              <View style={styles.mappingField}><Text style={styles.miniLabel}>QUANTIDADE</Text><View style={[styles.pickerWrap, { backgroundColor: theme.surfaceContainer, borderColor: theme.surfaceVariant }]}><Picker selectedValue={quantityColumn} onValueChange={setQuantityColumn} style={styles.picker}><Picker.Item label="Não importar" value={0} />{workbookPreview.columns.map(c => <Picker.Item key={c.index} label={`${c.index}. ${c.label}`} value={c.index} />)}</Picker></View></View>
            </View>
            <View style={[styles.previewTable, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}>
              <Text style={styles.miniLabel}>AMOSTRA DOS DADOS</Text>
              <ScrollView horizontal showsHorizontalScrollIndicator><View>
                <View style={styles.tableRow}>{workbookPreview.headers.slice(0, 7).map((h, i) => <Text key={i} style={[styles.tableHead, { color: theme.primary, backgroundColor: theme.primaryContainer }]}>{h}</Text>)}</View>
                {workbookPreview.rows.slice(0, 3).map((row, ri) => <View key={ri} style={styles.tableRow}>{row.slice(0, 7).map((cell, ci) => <Text key={ci} style={styles.tableCell}>{cell || '—'}</Text>)}</View>)}
              </View></ScrollView>
            </View>
            <ActionButton title="Salvar base neste aparelho" onPress={onImportDatabase} icon="save-outline" primaryColor={theme.primary} disabled={!!busy} />
          </>}
        </View>}
        {database.ready && <View style={[styles.inlineSuccess, { backgroundColor: theme.primaryContainer }]}><Ionicons name="checkmark-circle" size={18} color={theme.primary} /><Text style={[styles.inlineSuccessText, { color: theme.onPrimaryContainer }]}>Base salva localmente e disponível mesmo sem internet.</Text></View>}
        <View style={styles.stepFooter}>
          <View style={{ flex: 1 }}><Text style={styles.helper}>Você pode continuar sem importar uma planilha e digitar os itens manualmente.</Text></View>
          <ActionButton title="Continuar" onPress={() => setWorkflowStep('items')} icon="arrow-forward" primaryColor={theme.primary} />
        </View>
      </Card>}

      {workflowStep === 'items' && <Card style={[styles.materialCard, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}>
        <SectionTitle eyebrow="PASSO 2 DE 3" title="Quais plaquetas você precisa?" subtitle="Digite ou cole uma lista. O app procura os códigos e nomes na base salva no celular." accent={theme.primary} />
        <View style={[styles.segmented, { backgroundColor: theme.surfaceContainer }]}>
          <Pressable onPress={() => { setInputMode('code'); setSuggestions([]); }} style={[styles.segment, inputMode === 'code' && { backgroundColor: theme.surface, borderRadius: 12 }]}><Text style={[styles.segmentText, inputMode === 'code' && { color: theme.primary }]}>Por código</Text></Pressable>
          <Pressable onPress={() => { setInputMode('name'); setSuggestions([]); }} style={[styles.segment, inputMode === 'name' && { backgroundColor: theme.surface, borderRadius: 12 }]}><Text style={[styles.segmentText, inputMode === 'name' && { color: theme.primary }]}>Por nome</Text></Pressable>
        </View>
        <TextInput value={inputText} onChangeText={setInputText} multiline textAlignVertical="top" placeholder={inputMode === 'code' ? 'Cole os códigos aqui, um por linha…\nEx.: 76120200487\n76150100651' : 'Digite parte do nome do material…'} style={[styles.textInput, styles.listInput, { backgroundColor: theme.surfaceContainer, borderColor: theme.surfaceVariant, minHeight: 126, borderRadius: 16 }]} />
        {inputMode === 'name' && suggestions.length > 0 && <View style={[styles.suggestionPanel, { borderColor: theme.surfaceVariant, backgroundColor: theme.surface }]}>
          {suggestions.slice(0, 6).map((item, i) => <Pressable key={`${item.code}-${i}`} onPress={() => chooseSuggestion(item)} style={styles.suggestionRow}>
            <View style={{ flex: 1 }}><Text style={styles.suggestionName}>{item.name}</Text><Text style={styles.helper}>Código: {item.code} · Estoque: {item.quantity || '—'}</Text></View><Ionicons name="add-circle" size={21} color={theme.primary} />
          </Pressable>)}
        </View>}
        <ActionButton title="Adicionar à lista" onPress={addItems} icon="add" primaryColor={theme.primary} disabled={!!busy || !localReady} />
        <View style={[styles.listSummary, { backgroundColor: theme.surfaceContainer }]}><Text style={[styles.listSummaryTitle, { color: theme.primary }]}>{labels.length} plaqueta(s) na lista</Text><Text style={styles.helper}>{labels.length ? `Serão necessárias ${requiredPages} folha(s), com ${capacity} posições por folha.` : 'Ainda não há itens. Cole uma lista acima para começar.'}</Text></View>
        {labels.length === 0 ? <View style={styles.emptyState}><Ionicons name="documents-outline" size={30} color={theme.outline} /><Text style={styles.emptyTitle}>Sua lista está vazia</Text><Text style={styles.helper}>Depois de adicionar, você poderá editar cada código e descrição.</Text></View> : <View style={styles.itemList}>
          {labels.map((item, index) => <View key={item.localId} style={[styles.itemEditor, { backgroundColor: theme.surfaceContainer, borderColor: theme.surfaceVariant, borderRadius: 16 }]}>
            <View style={styles.itemEditorTop}><Text style={[styles.itemNumber, { color: theme.primary }]}>PLAQUETA {String(index + 1).padStart(2, '0')}</Text><View style={{ flex: 1 }} /><Pressable onPress={() => removeLabel(item.localId)} hitSlop={8}><Ionicons name="trash-outline" size={18} color="#B3261E" /></Pressable></View>
            <Text style={styles.miniLabel}>CÓDIGO</Text><TextInput value={item.code} onChangeText={v => updateLabel(item.localId, 'code', v)} style={[styles.editInput, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]} />
            <Text style={styles.miniLabel}>NOME DO MATERIAL</Text><TextInput value={item.name} onChangeText={v => updateLabel(item.localId, 'name', v)} style={[styles.editInput, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]} />
            {!item.code || !item.name ? <Text style={styles.warningText}>Preencha o código e o nome antes de exportar.</Text> : null}
            <Pressable onPress={() => { setAdjustSlotIndex(index % Math.max(1, capacity)); setWorkflowStep('layout'); }} style={styles.showPreviewLink}><Ionicons name="scan-outline" size={15} color={theme.primary} /><Text style={[styles.showPreviewText, { color: theme.primary }]}>Conferir posição na prévia</Text></Pressable>
          </View>)}
        </View>}
        <View style={styles.stepFooter}>
          <ActionButton title="Voltar" onPress={() => setWorkflowStep('data')} secondary primaryColor={theme.primary} />
          <ActionButton title="Continuar para o modelo" onPress={() => setWorkflowStep('layout')} icon="arrow-forward" primaryColor={theme.primary} />
        </View>
      </Card>}

      {workflowStep === 'layout' && <Card style={[styles.materialCard, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}>
        <SectionTitle eyebrow="PASSO 3 DE 3" title="Confira e ajuste seu modelo" subtitle="Esta prévia é montada a partir do PowerPoint carregado. Toque no código ou nome para selecionar o campo e veja as mudanças na hora." accent={theme.primary} />
        <View style={[styles.materialFileRow, { backgroundColor: theme.surfaceContainer, borderColor: theme.surfaceVariant }]}>
          <View style={[styles.materialIconCircle, { backgroundColor: theme.primaryContainer }]}><Ionicons name="document-text-outline" size={23} color={theme.primary} /></View>
          <View style={{ flex: 1, gap: 3 }}><Text style={styles.uploadTitle}>{template?.name || 'Carregando modelo padrão...'}</Text><Text style={styles.helper}>{template?.slides?.length || 0} modelo(s) · {capacity} posições por folha</Text></View>
          <ActionButton title="Trocar" onPress={onChooseTemplate} secondary small primaryColor={theme.primary} disabled={!!busy} />
        </View>
        {template?.slides?.length ? <>
          <Text style={styles.fieldLabel}>Escolha o tamanho de plaqueta</Text>
          <View style={styles.sizeChoices}>{template.slides.map((slide, index) => <Pressable key={slide.index} onPress={() => { setTemplateSlide(index); setAdjustSlotIndex(0); }} style={[styles.sizeChoice, { borderColor: templateSlide === index ? theme.primary : theme.surfaceVariant, backgroundColor: templateSlide === index ? theme.primaryContainer : theme.surface }]}>
            <Ionicons name={templateSlide === index ? 'radio-button-on' : 'radio-button-off'} color={templateSlide === index ? theme.primary : MUTED} size={17} />
            <View style={{ flex: 1 }}><Text style={[styles.sizeChoiceTitle, { color: templateSlide === index ? theme.onPrimaryContainer : INK }]}>{index === 0 ? 'Plaqueta maior' : index === 1 ? 'Plaqueta menor' : `Modelo ${index + 1}`}</Text><Text style={styles.sizeChoiceMeta}>{slide.capacity} por folha · {slide.orientation === 'portrait' ? 'Retrato' : 'Paisagem'}</Text></View>
          </Pressable>)}</View>
          {!!templateError && <Text style={styles.warningText}>{templateError}</Text>}
          {selectedSlide && selectedSlide.capacity === 0 && <View style={styles.advancedMap}>
            <Text style={styles.fieldLabel}>Mapeamento manual</Text><Text style={styles.helper}>Use apenas se o app não reconheceu automaticamente os campos. Informe os IDs das caixas de código e nome.</Text>
            <Text style={styles.miniLabel}>IDs DO CÓDIGO</Text><TextInput value={customCodeIds} onChangeText={setCustomCodeIds} style={styles.textInput} placeholder="12, 26, 41…" />
            <Text style={styles.miniLabel}>IDs DO NOME</Text><TextInput value={customNameIds} onChangeText={setCustomNameIds} style={styles.textInput} placeholder="15, 29, 44…" />
            <ActionButton title="Aplicar mapeamento" onPress={useFields} secondary primaryColor={theme.primary} small />
          </View>}
        </> : <Text style={styles.helper}>Preparando o modelo padrão incluído no aplicativo…</Text>}

        {selectedSlide && previewSlide && <>
          <View style={styles.previewHeader}>
            <View style={{ flex: 1, gap: 2 }}><Text style={styles.fieldLabel}>Prévia em tempo real</Text><Text style={styles.helper}>Folha inteira · toque nos campos para selecionar</Text></View>
            <View style={[styles.previewLiveTag, { backgroundColor: theme.primaryContainer }]}><View style={[styles.previewLiveDot, { backgroundColor: theme.primary }]} /><Text style={[styles.previewLiveText, { color: theme.onPrimaryContainer }]}>AO VIVO</Text></View>
          </View>
          {previewSlide.visuals?.length ? <TemplatePreview
            slide={previewSlide}
            labels={labels.slice(0, Math.max(1, previewSlide.capacity || capacity))}
            adjustments={activeAdjustments}
            activeSlot={Math.min(adjustSlotIndex, Math.max(0, previewSlide.capacity - 1))}
            activeField={adjustField}
            accent={theme.primary}
            onSelectField={(slot, field) => { setAdjustSlotIndex(slot); setAdjustField(field); }}
          /> : <View style={styles.previewFallback}><Ionicons name="image-outline" size={24} color={theme.outline} /><Text style={styles.helper}>Este modelo não possui uma prévia vetorial disponível. A exportação PowerPoint continuará usando o arquivo original.</Text></View>}
          <View style={[styles.adjustPanel, { backgroundColor: theme.surfaceContainer, borderColor: theme.surfaceVariant }]}>
            <View style={{ gap: 3 }}><Text style={[styles.adjustTitle, { color: theme.onPrimaryContainer }]}>Ajustar posição do texto</Text><Text style={styles.helper}>Cada toque move o campo selecionado e atualiza a prévia imediatamente.</Text></View>
            <View style={styles.adjustSegmentRow}>
              <Pressable onPress={() => setAdjustScope('global')} style={[styles.adjustSegment, { backgroundColor: adjustScope === 'global' ? theme.primary : theme.surface, borderColor: adjustScope === 'global' ? theme.primary : theme.surfaceVariant }]}><Text style={{ color: adjustScope === 'global' ? theme.onPrimary : INK, fontWeight: '700', fontSize: 12 }}>Todas as plaquetas</Text></Pressable>
              <Pressable onPress={() => setAdjustScope('single')} style={[styles.adjustSegment, { backgroundColor: adjustScope === 'single' ? theme.primary : theme.surface, borderColor: adjustScope === 'single' ? theme.primary : theme.surfaceVariant }]}><Text style={{ color: adjustScope === 'single' ? theme.onPrimary : INK, fontWeight: '700', fontSize: 12 }}>Só uma plaqueta</Text></Pressable>
            </View>
            {adjustScope === 'single' && <View style={styles.slotPickerRow}><Text style={styles.fieldLabel}>Plaqueta</Text><View style={[styles.pickerWrap, { width: 130, backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}><Picker selectedValue={Math.min(adjustSlotIndex, Math.max(0, capacity - 1))} onValueChange={value => setAdjustSlotIndex(Number(value))} style={styles.picker}>{Array.from({ length: Math.max(1, capacity) }, (_, i) => <Picker.Item key={i} label={`Plaqueta ${i + 1}`} value={i} />)}</Picker></View></View>}
            <Text style={styles.fieldLabel}>Qual texto deseja mover?</Text>
            <View style={styles.adjustSegmentRow}>
              <Pressable onPress={() => setAdjustField('code')} style={[styles.adjustSegment, { backgroundColor: adjustField === 'code' ? theme.primaryContainer : theme.surface, borderColor: adjustField === 'code' ? theme.primary : theme.surfaceVariant }]}><Text style={{ color: adjustField === 'code' ? theme.primary : INK, fontWeight: '700', fontSize: 12 }}>Código</Text></Pressable>
              <Pressable onPress={() => setAdjustField('name')} style={[styles.adjustSegment, { backgroundColor: adjustField === 'name' ? theme.primaryContainer : theme.surface, borderColor: adjustField === 'name' ? theme.primary : theme.surfaceVariant }]}><Text style={{ color: adjustField === 'name' ? theme.primary : INK, fontWeight: '700', fontSize: 12 }}>Nome do item</Text></Pressable>
            </View>
            <View style={styles.adjustStepRow}><Text style={styles.fieldLabel}>Precisão do movimento</Text>{[0.5, 1, 2, 5].map(step => <Pressable key={step} onPress={() => setAdjustStep(step)} style={[styles.adjustStepChip, { backgroundColor: adjustStep === step ? theme.primaryContainer : theme.surface, borderColor: adjustStep === step ? theme.primary : theme.surfaceVariant }]}><Text style={{ color: adjustStep === step ? theme.primary : MUTED, fontSize: 11, fontWeight: '800' }}>{step} mm</Text></Pressable>)}</View>
            <View style={styles.dpadWrap}>
              <View style={styles.dpadColumn}>
                <Pressable onPress={() => updateTemplatePosition(0, -adjustStep)} style={[styles.dpadButton, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}><Ionicons name="chevron-up" size={24} color={theme.primary} /></Pressable>
                <View style={styles.dpadMiddle}>
                  <Pressable onPress={() => updateTemplatePosition(-adjustStep, 0)} style={[styles.dpadButton, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}><Ionicons name="chevron-back" size={24} color={theme.primary} /></Pressable>
                  <View style={[styles.dpadCenter, { backgroundColor: theme.primaryContainer }]}><Ionicons name="move" size={18} color={theme.primary} /></View>
                  <Pressable onPress={() => updateTemplatePosition(adjustStep, 0)} style={[styles.dpadButton, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}><Ionicons name="chevron-forward" size={24} color={theme.primary} /></Pressable>
                </View>
                <Pressable onPress={() => updateTemplatePosition(0, adjustStep)} style={[styles.dpadButton, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}><Ionicons name="chevron-down" size={24} color={theme.primary} /></Pressable>
              </View>
              <View style={styles.offsetReadout}><Text style={styles.miniLabel}>DESLOCAMENTO ATUAL</Text><Text style={[styles.offsetValue, { color: theme.primary }]}>X {Number(selectedFieldDelta?.xMm || 0).toFixed(1)} mm</Text><Text style={[styles.offsetValue, { color: theme.primary }]}>Y {Number(selectedFieldDelta?.yMm || 0).toFixed(1)} mm</Text><Text style={styles.helper}>{adjustScope === 'global' ? 'Aplicado em todas as posições do modelo.' : `Aplicado somente à plaqueta ${adjustSlotIndex + 1}.`}</Text></View>
            </View>
            <View style={styles.stepFooter}><View style={{ flex: 1 }}><Text style={styles.helper}>Posicione olhando a prévia. O valor será salvo neste aparelho e aplicado ao exportar.</Text></View><ActionButton title="Restaurar posição" onPress={resetTemplatePositions} secondary primaryColor={theme.primary} small /></View>
          </View>
        </>}

        <View style={[styles.exportBox, { borderColor: theme.surfaceVariant }]}>
          <Text style={styles.exportTitle}>Gerar arquivo</Text>
          <View style={styles.exportRow}><View style={{ flex: 1 }}><Text style={styles.helper}>Quantidade mínima: {pageLabel} · {capacity} plaquetas por folha.</Text></View><View style={styles.pageInputWrap}><Text style={styles.miniLabel}>FOLHAS</Text><TextInput keyboardType="number-pad" value={requestedPages} onChangeText={setRequestedPages} placeholder={String(requiredPages)} style={[styles.pageInput, { backgroundColor: theme.surfaceContainer, borderColor: theme.surfaceVariant }]} /></View></View>
          <View style={styles.exportButtons}><View style={{ flex: 1 }}><ActionButton title="Exportar PDF" onPress={() => exportFile('pdf')} icon="document-text-outline" primaryColor={theme.primary} small disabled={!!busy || !labels.length || !localReady} /></View><View style={{ flex: 1 }}><ActionButton title="Exportar PowerPoint" onPress={() => exportFile('pptx')} icon="file-tray-full-outline" primaryColor={theme.primary} small disabled={!!busy || !labels.length || !localReady} /></View></View>
          <Text style={styles.helper}>O PowerPoint preserva o modelo original. O PDF usa uma grade de impressão simplificada com os mesmos ajustes de posição.</Text>
          {lastExport && <Text style={[styles.successText, { color: theme.primary }]}><Ionicons name="checkmark-circle" size={14} color={theme.primary} /> Último arquivo: {lastExport.fileName} · {lastExport.pages} folha(s)</Text>}
        </View>
        <View style={styles.stepFooter}><ActionButton title="Voltar para os itens" onPress={() => setWorkflowStep('items')} secondary primaryColor={theme.primary} /><Text style={styles.helper}>Seu modelo e ajustes ficam salvos no aparelho.</Text></View>
      </Card>}

      <View style={styles.footer}><Text style={styles.footerText}>PLAQUETAS BY LIPPE · MATERIAL YOU</Text><Text style={styles.footerText}>Arquivos, base, preferências e ajustes ficam armazenados localmente.</Text></View>
    </ScrollView> : <ScrollView contentContainerStyle={[styles.content, { backgroundColor: theme.background }]} keyboardShouldPersistTaps="handled">
      <View style={[styles.heroRow, { backgroundColor: theme.primaryContainer, borderRadius: 24, padding: 18, marginBottom: 14 }]}><View style={{ flex: 1, gap: 5 }}><Text style={[styles.heroEyebrow, { color: theme.primary }]}>INVENTÁRIO · CONTAGEM</Text><Text style={[styles.heroTitle, { color: theme.onPrimaryContainer }]}>Conte sem perder o ritmo.</Text><Text style={[styles.heroText, { color: theme.onPrimaryContainer, opacity: 0.78 }]}>Leia o código do item e conte em unidades ou em lotes de dez. Os registros ficam neste aparelho.</Text></View><View style={[styles.heroIcon, { backgroundColor: `${theme.surface}B8` }]}><Ionicons name="scan-circle-outline" size={36} color={theme.primary} /></View></View>
      <View style={[styles.mainGrid, width < 960 && styles.mainGridMobile]}>
        <View style={[styles.leftColumn, width < 960 && styles.columnMobile]}>
          <Card style={[styles.materialCard, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}>
            <SectionTitle eyebrow="LEITOR" title="Identificar item" subtitle="Use a câmera para ler QR Code/código de barras ou digite o código." accent={theme.primary} />
            {scanEnabled ? <CameraScanner onScan={onScan} onClose={() => setScanEnabled(false)} accent={theme.primary} /> : <ActionButton title="Abrir câmera" onPress={() => setScanEnabled(true)} icon="camera-outline" primaryColor={theme.primary} />}
            <Text style={styles.fieldLabel}>Código do item</Text><TextInput value={scanCode} onChangeText={setScanCode} placeholder="Escaneie ou digite o código" style={[styles.textInput, { backgroundColor: theme.surfaceContainer, borderColor: theme.surfaceVariant }]} autoCapitalize="none" />
            {!!scanName && <Text style={[styles.scanName, { color: theme.primary }]}>{scanName}</Text>}
            <Text style={styles.fieldLabel}>Tamanho da contagem</Text><View style={styles.stepChoices}><Pressable onPress={() => setScanStep(1)} style={[styles.stepChoice, { backgroundColor: scanStep === 1 ? theme.primaryContainer : theme.surface, borderColor: scanStep === 1 ? theme.primary : theme.surfaceVariant }]}><Text style={[styles.stepChoiceText, { color: scanStep === 1 ? theme.primary : MUTED }]}>+1 unidade</Text></Pressable><Pressable onPress={() => setScanStep(10)} style={[styles.stepChoice, { backgroundColor: scanStep === 10 ? theme.primaryContainer : theme.surface, borderColor: scanStep === 10 ? theme.primary : theme.surfaceVariant }]}><Text style={[styles.stepChoiceText, { color: scanStep === 10 ? theme.primary : MUTED }]}>+10 unidades</Text></Pressable></View>
            <ActionButton title={`Adicionar ${scanStep} à contagem`} onPress={incrementCount} icon="add-circle-outline" primaryColor={theme.primary} disabled={!scanCode.trim() || !!busy} />
          </Card>
        </View>
        <View style={[styles.rightColumn, width < 960 && styles.columnMobile]}>
          <Card style={[styles.materialCard, { backgroundColor: theme.surface, borderColor: theme.surfaceVariant }]}>
            <SectionTitle eyebrow="RESULTADO" title={`Itens contados (${counts.length})`} subtitle="A quantidade é acumulada por código." accent={theme.primary} />
            {counts.length === 0 ? <View style={styles.emptyState}><Ionicons name="calculator-outline" size={30} color={theme.outline} /><Text style={styles.emptyTitle}>Nenhum item contado</Text><Text style={styles.helper}>Leia um código e use os botões +1 ou +10.</Text></View> : counts.map(item => <View key={item.code} style={[styles.countRow, { borderColor: theme.surfaceVariant }]}><View style={{ flex: 1 }}><Text style={styles.countName}>{item.name || 'Item sem descrição'}</Text><Text style={styles.helper}>{item.code}</Text></View><Text style={[styles.countValue, { color: theme.primary }]}>{item.quantity || 0}</Text></View>)}
            {counts.length > 0 && <View style={styles.rowActions}><ActionButton title="Atualizar" onPress={refreshCounts} secondary small primaryColor={theme.primary} /><ActionButton title="Zerar contagem" onPress={async () => { const confirmed = Platform.OS === 'web' ? window.confirm('Apagar todos os registros de contagem?') : true; if (confirmed) { await persistPatch({ counts: [] }); setCounts([]); } }} secondary small primaryColor={theme.primary} /> </View>}
          </Card>
          <Card style={[styles.materialCard, { backgroundColor: theme.primaryContainer, borderColor: theme.surfaceVariant, flexDirection: 'row', alignItems: 'flex-start' }]}><Ionicons name="information-circle-outline" size={20} color={theme.primary} /><Text style={[styles.noteText, { color: theme.onPrimaryContainer }]}>A leitura identifica o produto; a contagem só é somada quando você toca em “Adicionar”. Assim você pode contar 10 unidades fisicamente e registrar de uma vez com +10.</Text></Card>
        </View>
      </View>
    </ScrollView>}

    {!!busy && <View style={styles.busyOverlay}><View style={styles.busyCard}><ActivityIndicator color={theme.primary} size="small" /><Text style={styles.busyText}>{busy}</Text></View></View>}
  </View>;
}

class AppErrorBoundary extends React.Component<React.PropsWithChildren, { error: string | null }> {
  state = { error: null as string | null };

  static getDerivedStateFromError(error: unknown) {
    const message = error instanceof Error ? error.message : String(error || 'Erro inesperado');
    return { error: message };
  }

  componentDidCatch(error: Error, info: React.ErrorInfo) {
    console.error('Falha ao renderizar o Plaquetas:', error, info.componentStack);
  }

  render() {
    if (this.state.error) {
      return (
        <View style={{ flex: 1, backgroundColor: '#F4F7F4', justifyContent: 'center', padding: 24 }}>
          <Text style={{ color: '#17251F', fontSize: 22, fontWeight: '900', marginBottom: 10 }}>O Plaquetas encontrou um erro</Text>
          <Text style={{ color: '#53645B', fontSize: 14, lineHeight: 21, marginBottom: 18 }}>{this.state.error}</Text>
          <Pressable onPress={() => this.setState({ error: null })} style={{ backgroundColor: '#17734F', borderRadius: 10, padding: 14, alignItems: 'center' }}>
            <Text style={{ color: '#FFFFFF', fontWeight: '800' }}>Tentar abrir novamente</Text>
          </Pressable>
        </View>
      );
    }
    return this.props.children;
  }
}

export default function App() {
  return <AppErrorBoundary><AppScreen /></AppErrorBoundary>;
}

const styles = StyleSheet.create({
  screen: { flex: 1, backgroundColor: '#F4F7F4' },
  topBar: { minHeight: 70, paddingHorizontal: 22, paddingVertical: 12, flexDirection: 'row', alignItems: 'center', borderBottomWidth: 1, borderBottomColor: LINE, backgroundColor: '#FFFFFF', gap: 12 },
  brandMark: { width: 40, height: 40, borderRadius: 13, backgroundColor: GREEN, alignItems: 'center', justifyContent: 'center' },
  brandTitle: { color: INK, fontSize: 17, fontWeight: '900', letterSpacing: 1.1 }, brandMuted: { fontWeight: '500', color: MUTED, letterSpacing: 0 }, brandSubtitle: { color: MUTED, fontSize: 9, letterSpacing: 1.6, marginTop: 2, fontWeight: '700' },
  offlineTag: { flexDirection: 'row', alignItems: 'center', gap: 5, backgroundColor: '#EAF4EE', borderRadius: 8, paddingHorizontal: 8, paddingVertical: 6 }, offlineTagText: { color: GREEN, fontSize: 9, fontWeight: '900', letterSpacing: 0.7 },
  iconButton: { padding: 9, borderRadius: 10, backgroundColor: '#F2F6F3' }, connectionDot: { width: 8, height: 8, borderRadius: 9 }, dotOn: { backgroundColor: '#24A36B' }, dotOff: { backgroundColor: '#D59A52' },
  settingsPanel: { padding: 15, backgroundColor: '#FFFFFF', borderBottomWidth: 1, borderBottomColor: LINE, gap: 9 }, settingsTitle: { color: INK, fontSize: 13, fontWeight: '800' },
  tabBar: { flexDirection: 'row', backgroundColor: '#FFFFFF', paddingHorizontal: 18, paddingVertical: 8, borderBottomWidth: 1, borderBottomColor: LINE, gap: 8 }, tab: { flex: 1, flexDirection: 'row', alignItems: 'center', justifyContent: 'center', gap: 8, paddingVertical: 11, borderRadius: 16, borderBottomWidth: 0 }, tabActive: { borderBottomColor: GREEN }, tabText: { color: MUTED, fontSize: 13, fontWeight: '700' }, tabTextActive: { color: GREEN }, betaTag: { color: '#A36B20', backgroundColor: '#FFF3D8', fontSize: 8, fontWeight: '900', paddingHorizontal: 5, paddingVertical: 2, borderRadius: 5 },
  themeOpenButton: { width: 42, height: 42, borderRadius: 15, borderWidth: 1, alignItems: 'center', justifyContent: 'center' }, themePickerPanel: { flexDirection: 'row', alignItems: 'center', justifyContent: 'space-between', paddingHorizontal: 18, paddingVertical: 13, borderBottomWidth: 1, gap: 12 }, themeSwatches: { flexDirection: 'row', gap: 9, alignItems: 'center' }, themeSwatch: { width: 30, height: 30, borderRadius: 15, borderWidth: 2, alignItems: 'center', justifyContent: 'center' },
  workflowTabs: { flexDirection: 'row', padding: 5, borderRadius: 19, borderWidth: 1, gap: 4 }, workflowStep: { flex: 1, minWidth: 0, minHeight: 46, borderWidth: 1, borderRadius: 15, alignItems: 'center', justifyContent: 'center', flexDirection: 'row', gap: 7 }, workflowNumber: { width: 25, height: 25, borderRadius: 13, alignItems: 'center', justifyContent: 'center' }, workflowStepText: { fontSize: 12, fontWeight: '800' }, materialCard: { borderRadius: 24, padding: 18, borderWidth: 1, shadowColor: '#172033', shadowOpacity: 0.05, shadowRadius: 14, shadowOffset: { width: 0, height: 5 }, elevation: 2, gap: 16 }, materialFileRow: { flexDirection: 'row', alignItems: 'center', gap: 12, padding: 13, borderRadius: 18, borderWidth: 1 }, materialIconCircle: { width: 48, height: 48, borderRadius: 17, alignItems: 'center', justifyContent: 'center' }, inlineSuccess: { borderRadius: 14, padding: 12, flexDirection: 'row', alignItems: 'center', gap: 8 }, inlineSuccessText: { flex: 1, fontSize: 11, fontWeight: '700', lineHeight: 16 }, stepFooter: { flexDirection: 'row', alignItems: 'center', justifyContent: 'space-between', gap: 10, flexWrap: 'wrap', paddingTop: 6 }, listSummary: { padding: 14, borderRadius: 16, gap: 4 }, listSummaryTitle: { fontSize: 15, fontWeight: '900' }, previewHeader: { flexDirection: 'row', alignItems: 'center', gap: 10, paddingTop: 4 }, previewLiveTag: { flexDirection: 'row', alignItems: 'center', gap: 6, paddingHorizontal: 10, paddingVertical: 7, borderRadius: 20 }, previewLiveDot: { width: 7, height: 7, borderRadius: 4 }, previewLiveText: { fontSize: 9, fontWeight: '900', letterSpacing: 0.8 }, previewFallback: { minHeight: 120, borderRadius: 16, backgroundColor: '#F4F5F8', alignItems: 'center', justifyContent: 'center', gap: 8, padding: 18 }, adjustPanel: { borderRadius: 20, padding: 15, borderWidth: 1, gap: 13 }, adjustTitle: { fontSize: 16, fontWeight: '900' }, adjustSegmentRow: { flexDirection: 'row', gap: 8 }, adjustSegment: { minHeight: 42, flex: 1, borderWidth: 1, borderRadius: 14, alignItems: 'center', justifyContent: 'center', paddingHorizontal: 10 }, slotPickerRow: { flexDirection: 'row', alignItems: 'center', gap: 12, justifyContent: 'space-between' }, adjustStepRow: { flexDirection: 'row', alignItems: 'center', flexWrap: 'wrap', gap: 7 }, adjustStepChip: { borderWidth: 1, borderRadius: 18, paddingHorizontal: 10, paddingVertical: 7 }, dpadWrap: { flexDirection: 'row', alignItems: 'center', justifyContent: 'center', gap: 20, flexWrap: 'wrap' }, dpadColumn: { alignItems: 'center', gap: 6 }, dpadMiddle: { flexDirection: 'row', alignItems: 'center', gap: 6 }, dpadButton: { width: 46, height: 42, borderRadius: 14, borderWidth: 1, alignItems: 'center', justifyContent: 'center' }, dpadCenter: { width: 42, height: 42, borderRadius: 14, alignItems: 'center', justifyContent: 'center' }, offsetReadout: { minWidth: 120, flex: 1, backgroundColor: 'rgba(255,255,255,0.68)', padding: 13, borderRadius: 15, gap: 4 }, offsetValue: { fontSize: 15, fontWeight: '900', fontVariant: ['tabular-nums'] },
  apiWarning: { marginHorizontal: 22, marginTop: 12, backgroundColor: '#FFF7E9', borderColor: '#F2DEC0', borderWidth: 1, borderRadius: 12, padding: 12, flexDirection: 'row', alignItems: 'center', gap: 9 }, apiWarningTitle: { color: '#875018', fontSize: 12, fontWeight: '800' }, apiWarningText: { color: '#956F43', fontSize: 11, marginTop: 2 },
  content: { width: '100%', maxWidth: 1200, alignSelf: 'center', padding: 22, paddingBottom: 48, gap: 18 }, heroRow: { flexDirection: 'row', alignItems: 'center', paddingVertical: 5, gap: 14 }, heroEyebrow: { color: GREEN, fontWeight: '900', fontSize: 10, letterSpacing: 1.7, marginBottom: 6 }, heroTitle: { color: INK, fontSize: 28, fontWeight: '900', letterSpacing: -0.7 }, heroText: { color: MUTED, fontSize: 13, lineHeight: 19, marginTop: 5, maxWidth: 620 }, heroIcon: { width: 66, height: 66, borderRadius: 20, backgroundColor: '#E7F2EB', alignItems: 'center', justifyContent: 'center' },
  statRow: { flexDirection: 'row', gap: 11, flexWrap: 'wrap' }, statCard: { flex: 1, minWidth: 0, backgroundColor: '#FFFFFF', borderRadius: 14, padding: 14, borderColor: LINE, borderWidth: 1 }, statLabel: { fontSize: 9, color: MUTED, fontWeight: '900', letterSpacing: 1 }, statValue: { fontSize: 23, color: INK, fontWeight: '900', marginTop: 5 }, statFoot: { fontSize: 10, color: MUTED, marginTop: 2 },
  mainGrid: { flexDirection: 'row', alignItems: 'flex-start', gap: 16 }, mainGridMobile: { flexDirection: 'column' }, leftColumn: { flex: 1.2, minWidth: 0, gap: 15 }, rightColumn: { flex: 0.9, minWidth: 0, gap: 15 }, columnMobile: { width: '100%' },
  card: { backgroundColor: '#FFFFFF', borderRadius: 16, borderWidth: 1, borderColor: LINE, padding: 17, shadowColor: '#23382D', shadowOpacity: 0.03, shadowRadius: 10, shadowOffset: { width: 0, height: 3 }, elevation: 1, gap: 12 }, sectionTitle: { gap: 4, marginBottom: 2 }, eyebrow: { color: GREEN, fontSize: 9, fontWeight: '900', letterSpacing: 1.5 }, sectionHeading: { color: INK, fontSize: 18, fontWeight: '900' }, sectionSub: { color: MUTED, fontSize: 11, lineHeight: 16 },
  uploadRow: { flexDirection: 'row', alignItems: 'center', gap: 10, borderWidth: 1, borderColor: LINE, backgroundColor: '#FAFCFA', borderRadius: 12, padding: 10 }, uploadIcon: { width: 40, height: 40, borderRadius: 11, backgroundColor: '#E8F2EC', alignItems: 'center', justifyContent: 'center' }, uploadTitle: { color: INK, fontSize: 12, fontWeight: '800' }, helper: { color: MUTED, fontSize: 10, lineHeight: 15 },
  button: { minHeight: 42, flexDirection: 'row', alignItems: 'center', justifyContent: 'center', borderRadius: 10, paddingHorizontal: 14, paddingVertical: 9, gap: 7, alignSelf: 'flex-start' }, buttonPrimary: { backgroundColor: GREEN }, buttonSecondary: { backgroundColor: '#EFF6F1', borderWidth: 1, borderColor: '#D8E8DD' }, buttonText: { color: '#FFFFFF', fontWeight: '800', fontSize: 12 }, buttonTextSecondary: { color: GREEN }, buttonSmall: { minHeight: 34, paddingHorizontal: 10, paddingVertical: 7 }, buttonTextSmall: { fontSize: 10 }, buttonDisabled: { opacity: 0.45 }, buttonPressed: { opacity: 0.8 },
  mappingBlock: { gap: 10 }, fieldLabel: { color: INK, fontWeight: '800', fontSize: 11, marginTop: 2 }, miniLabel: { color: MUTED, fontWeight: '900', fontSize: 8, letterSpacing: 1.1, marginBottom: 4 }, pickerWrap: { borderWidth: 1, borderColor: LINE, backgroundColor: '#FAFCFA', borderRadius: 9, overflow: 'hidden', minHeight: 39, justifyContent: 'center' }, picker: { minHeight: 38, width: '100%', color: INK, fontSize: 11 }, mappingRow: { flexDirection: 'row', gap: 8 }, mappingRowMobile: { flexDirection: 'column' }, mappingField: { flex: 1, minWidth: 0, gap: 4 },
  previewTable: { borderWidth: 1, borderColor: LINE, borderRadius: 9, padding: 8, gap: 5, overflow: 'hidden' }, tableRow: { flexDirection: 'row', borderBottomWidth: 1, borderColor: '#EDF1EE' }, tableHead: { width: 112, color: GREEN, fontWeight: '900', fontSize: 9, padding: 5, backgroundColor: '#F2F7F3' }, tableCell: { width: 112, color: INK, fontSize: 9, padding: 5 },
  sizeChoices: { flexDirection: 'row', gap: 8 }, sizeChoice: { flex: 1, flexDirection: 'row', alignItems: 'center', gap: 7, borderWidth: 1, borderColor: LINE, borderRadius: 11, padding: 10, backgroundColor: '#FFFFFF' }, sizeChoiceActive: { backgroundColor: '#F0F7F2', borderColor: '#A9CBB5' }, sizeChoiceTitle: { color: INK, fontSize: 11, fontWeight: '800' }, sizeChoiceTitleActive: { color: GREEN }, sizeChoiceMeta: { color: MUTED, fontSize: 9, marginTop: 2 }, warningText: { color: '#9B6020', fontSize: 10, lineHeight: 15, backgroundColor: '#FFF7EA', padding: 8, borderRadius: 7 }, advancedMap: { gap: 7, borderTopWidth: 1, borderColor: LINE, paddingTop: 12 }, shapeList: { maxHeight: 170, borderWidth: 1, borderColor: LINE, borderRadius: 9, overflow: 'hidden' }, shapeRow: { flexDirection: 'row', alignItems: 'center', padding: 7, gap: 5, borderBottomWidth: 1, borderBottomColor: LINE }, shapeName: { color: INK, fontSize: 9, fontWeight: '800' }, shapeAssign: { borderRadius: 6, backgroundColor: '#EFF6F1', paddingHorizontal: 7, paddingVertical: 6 }, shapeAssignText: { color: GREEN, fontSize: 8, fontWeight: '900' },
  segmented: { flexDirection: 'row', backgroundColor: '#F1F5F2', borderRadius: 10, padding: 3, alignSelf: 'flex-start' }, segment: { paddingHorizontal: 17, paddingVertical: 8, borderRadius: 8 }, segmentActive: { backgroundColor: '#FFFFFF', shadowColor: '#000000', shadowOpacity: 0.06, shadowRadius: 3, elevation: 1 }, segmentText: { fontSize: 11, color: MUTED, fontWeight: '800' }, segmentTextActive: { color: GREEN }, textInput: { minHeight: 40, borderWidth: 1, borderColor: LINE, borderRadius: 9, paddingHorizontal: 11, paddingVertical: 9, color: INK, backgroundColor: '#FFFFFF', fontSize: 12 }, listInput: { minHeight: 112, textAlignVertical: 'top', lineHeight: 20 }, rowActions: { flexDirection: 'row', alignItems: 'center', gap: 9, flexWrap: 'wrap' }, suggestionPanel: { borderWidth: 1, borderColor: '#D9E9DE', borderRadius: 10, overflow: 'hidden', backgroundColor: '#FBFDFC' }, suggestionRow: { flexDirection: 'row', alignItems: 'center', padding: 10, borderBottomWidth: 1, borderBottomColor: LINE, gap: 8 }, suggestionName: { color: INK, fontSize: 11, fontWeight: '800' },
  previewControls: { flexDirection: 'row', alignSelf: 'flex-start', padding: 3, borderRadius: 8, backgroundColor: '#F0F4F1', gap: 3 }, previewToggle: { paddingHorizontal: 13, paddingVertical: 7, borderRadius: 6 }, previewToggleActive: { backgroundColor: '#FFFFFF' }, previewToggleText: { color: MUTED, fontSize: 10, fontWeight: '700' }, previewToggleTextActive: { color: GREEN }, labelStage: { backgroundColor: '#F6F8F6', padding: 12, borderRadius: 12, alignItems: 'center', gap: 8 }, labelPreview: { width: '100%', maxWidth: 400, minHeight: 225, backgroundColor: '#FFFFFF', borderWidth: 2, borderColor: '#111111' }, labelPreviewCompact: { maxWidth: 310, minHeight: 190 }, labelTop: { height: 54, flexDirection: 'row', borderBottomWidth: 2, borderColor: '#111111' }, logoCell: { width: '29%', borderRightWidth: 2, borderColor: '#111111', justifyContent: 'center', alignItems: 'center' }, logoText: { fontSize: 21, fontWeight: '900', letterSpacing: -1.7, color: '#1E2421' }, logoSub: { fontSize: 6, color: '#1E2421', fontWeight: '900', letterSpacing: 1 }, codeCell: { flex: 1, paddingHorizontal: 8, paddingVertical: 4, justifyContent: 'space-between' }, labelSmall: { fontSize: 9, color: '#111111', fontWeight: '800' }, labelCode: { textAlign: 'center', fontSize: 19, fontWeight: '900', color: '#111111', marginBottom: 2 }, itemCell: { flex: 1, paddingHorizontal: 7, paddingTop: 5, paddingBottom: 7, borderBottomWidth: 2, borderColor: '#111111', justifyContent: 'space-between' }, labelName: { fontSize: 17, lineHeight: 23, color: '#111111', fontWeight: '900', textAlign: 'center', marginHorizontal: 7, marginVertical: 8 }, labelBottom: { minHeight: 60, flexDirection: 'row' }, bottomCell: { flex: 1, padding: 4 }, labelNav: { flexDirection: 'row', justifyContent: 'center', alignItems: 'center', gap: 16 }, navCircle: { width: 30, height: 30, borderRadius: 15, backgroundColor: '#F1F5F2', alignItems: 'center', justifyContent: 'center' }, navCount: { color: INK, fontSize: 11, fontWeight: '800' },
  emptyState: { alignItems: 'center', paddingVertical: 22, gap: 6 }, emptyTitle: { color: INK, fontSize: 12, fontWeight: '800' }, itemList: { gap: 10 }, itemEditor: { borderWidth: 1, borderColor: LINE, borderRadius: 11, padding: 10, gap: 5, backgroundColor: '#FCFDFC' }, itemEditorTop: { flexDirection: 'row', alignItems: 'center', gap: 9, marginBottom: 3 }, itemNumber: { color: GREEN, fontSize: 10, fontWeight: '900' }, itemStock: { flex: 1, color: MUTED, fontSize: 9 }, editInput: { minHeight: 34, paddingHorizontal: 9, paddingVertical: 6, borderWidth: 1, borderColor: LINE, borderRadius: 7, color: INK, fontSize: 11, backgroundColor: '#FFFFFF' }, showPreviewLink: { flexDirection: 'row', alignItems: 'center', gap: 5, alignSelf: 'flex-start', paddingTop: 3 }, showPreviewText: { color: GREEN, fontSize: 9, fontWeight: '800' }, exportBox: { borderTopWidth: 1, borderColor: LINE, paddingTop: 13, marginTop: 2, gap: 10 }, exportTitle: { color: INK, fontSize: 13, fontWeight: '900' }, exportRow: { flexDirection: 'row', alignItems: 'center', gap: 12 }, pageInputWrap: { width: 78 }, pageInput: { borderWidth: 1, borderColor: LINE, borderRadius: 8, padding: 7, color: INK, fontWeight: '800', textAlign: 'center', minHeight: 34 }, exportButtons: { flexDirection: 'row', gap: 8 }, successText: { color: GREEN, fontSize: 9, fontWeight: '700' },
  footer: { alignItems: 'center', gap: 5, paddingTop: 12 }, footerText: { color: '#9BA79F', fontSize: 8, letterSpacing: 0.5 }, busyOverlay: { position: 'absolute', top: 0, bottom: 0, left: 0, right: 0, backgroundColor: 'rgba(244,247,244,0.55)', justifyContent: 'center', alignItems: 'center' }, busyCard: { flexDirection: 'row', alignItems: 'center', gap: 10, backgroundColor: '#FFFFFF', paddingHorizontal: 16, paddingVertical: 13, borderRadius: 12, borderWidth: 1, borderColor: LINE, elevation: 5 }, busyText: { color: INK, fontSize: 11, fontWeight: '800' },
  cameraContainer: { gap: 9 }, camera: { height: 230, width: '100%', borderRadius: 12, overflow: 'hidden' }, cameraPermission: { height: 160, borderRadius: 12, backgroundColor: '#F2F7F3', justifyContent: 'center', alignItems: 'center', gap: 10 }, scanName: { fontSize: 11, fontWeight: '700', color: GREEN, paddingVertical: 4 }, stepChoices: { flexDirection: 'row', gap: 8 }, stepChoice: { flex: 1, padding: 11, borderRadius: 9, borderWidth: 1, borderColor: LINE, alignItems: 'center' }, stepChoiceActive: { borderColor: '#8AB99B', backgroundColor: '#EFF7F1' }, stepChoiceText: { color: MUTED, fontSize: 11, fontWeight: '800' }, stepChoiceTextActive: { color: GREEN }, countRow: { flexDirection: 'row', alignItems: 'center', borderBottomWidth: 1, borderColor: LINE, paddingVertical: 10, gap: 10 }, countName: { fontSize: 11, color: INK, fontWeight: '800' }, countValue: { color: GREEN, fontWeight: '900', fontSize: 20, minWidth: 48, textAlign: 'right' }, noteCard: { flexDirection: 'row', alignItems: 'flex-start', backgroundColor: '#F0F7F2' }, noteText: { color: '#506C5D', fontSize: 10, lineHeight: 16, flex: 1 },
});
