import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  ActivityIndicator,
  Alert,
  Linking,
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
import * as Sharing from 'expo-sharing';
import { Ionicons } from '@expo/vector-icons';

type TemplateField = { id: string; shape: string; sample: string };
type TemplateSlide = {
  index: number;
  capacity: number;
  widthEmu: number;
  heightEmu: number;
  nameFields: TemplateField[];
  codeFields: TemplateField[];
  shapes: { id: string; shape: string; text: string; field: string | null }[];
};
type TemplateInfo = { templateId: string; name: string; slides: TemplateSlide[] };
type CatalogItem = { code: string; name: string; quantity?: string; sheetName?: string };
type LabelItem = CatalogItem & { localId: string; sourceInput?: string };
type WorkbookPreview = { headers: string[]; rows: string[][]; columns: { index: number; label: string }[]; suggested: { codeColumn: number; nameColumn: number; quantityColumn: number } };
type ExportResult = { fileName: string; format: string; downloadUrl: string; pages: number; capacity: number; printedItems: number; remainingItems: number };

const ENV_API = process.env.EXPO_PUBLIC_API_URL || '';
const WEB_LOCALHOST = Platform.OS === 'web' && typeof window !== 'undefined' && ['localhost', '127.0.0.1'].includes(window.location.hostname);
const DEFAULT_API = ENV_API || (Platform.OS === 'web' ? (WEB_LOCALHOST ? 'http://localhost:4000' : '') : 'https://plaquetas-by-lippe.onrender.com');
const GREEN = '#17734F';
const INK = '#17251F';
const MUTED = '#75827B';
const LINE = '#E2E9E4';

function showMessage(message: string) {
  if (Platform.OS === 'web') {
    // Alert.alert is supported by React Native Web, but window.alert keeps the message visible in browsers.
    if (typeof window !== 'undefined') window.alert(message);
    else Alert.alert('Plaquetas', message);
  } else Alert.alert('Plaquetas', message);
}
function makeId() { return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`; }
function splitValues(text: string) { return text.split(/[\n,;\t]+/).map(v => v.trim()).filter(Boolean); }

async function fetchJson<T>(base: string, endpoint: string, options?: RequestInit): Promise<T> {
  if (!base.trim() && Platform.OS !== 'web') throw new Error('Configure o endereço HTTPS do servidor nas configurações do aplicativo.');
  const response = await fetch(`${base.replace(/\/$/, '')}${endpoint}`, options);
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload.error || `Falha na comunicação (${response.status}).`);
  return payload as T;
}

async function chooseFile(type: string) {
  const result = await DocumentPicker.getDocumentAsync({ type, copyToCacheDirectory: true, multiple: false });
  if (result.canceled || !result.assets?.length) return null;
  return result.assets[0];
}
async function uploadPickedFile(base: string, endpoint: string, asset: any) {
  const form = new FormData();
  if (Platform.OS === 'web' && asset.file) form.append('file', asset.file);
  else form.append('file', { uri: asset.uri, name: asset.name || 'upload.xlsx', type: asset.mimeType || 'application/octet-stream' } as any);
  return fetchJson<any>(base, endpoint, { method: 'POST', body: form });
}

function SectionTitle({ eyebrow, title, subtitle }: { eyebrow: string; title: string; subtitle?: string }) {
  return <View style={styles.sectionTitle}>
    <Text style={styles.eyebrow}>{eyebrow}</Text>
    <Text style={styles.sectionHeading}>{title}</Text>
    {!!subtitle && <Text style={styles.sectionSub}>{subtitle}</Text>}
  </View>;
}
function ActionButton({ title, onPress, icon, secondary = false, disabled = false, small = false }: { title: string; onPress: () => void; icon?: any; secondary?: boolean; disabled?: boolean; small?: boolean }) {
  return <Pressable onPress={onPress} disabled={disabled} style={({ pressed }) => [styles.button, secondary ? styles.buttonSecondary : styles.buttonPrimary, small && styles.buttonSmall, disabled && styles.buttonDisabled, pressed && !disabled && styles.buttonPressed]}>
    {!!icon && <Ionicons name={icon} size={small ? 15 : 17} color={secondary ? GREEN : '#FFFFFF'} />}
    <Text style={[styles.buttonText, secondary && styles.buttonTextSecondary, small && styles.buttonTextSmall]}>{title}</Text>
  </Pressable>;
}
function Card({ children, style }: { children: React.ReactNode; style?: any }) {
  return <View style={[styles.card, style]}>{children}</View>;
}

function CameraScanner({ onScan, onClose }: { onScan: (data: string) => void; onClose: () => void }) {
  const [cameraPermission, requestCameraPermission] = useCameraPermissions();
  return (
    <View style={styles.cameraContainer}>
      {!cameraPermission?.granted ? (
        <View style={styles.cameraPermission}>
          <Ionicons name="camera-outline" size={28} color={GREEN} />
          <Text style={styles.helper}>A câmera precisa de permissão.</Text>
          <ActionButton title="Permitir câmera" onPress={async () => { const result = await requestCameraPermission(); if (!result.granted) showMessage('A permissão da câmera não foi concedida.'); }} />
        </View>
      ) : (
        <CameraView style={styles.camera} facing="back" onBarcodeScanned={({ data }) => onScan(data)} barcodeScannerSettings={{ barcodeTypes: ['qr', 'ean13', 'ean8', 'code128', 'code39', 'upc_a', 'upc_e'] }} />
      )}
      <ActionButton title="Fechar leitor" onPress={onClose} secondary small />
    </View>
  );
}

function AppScreen() {
  const { width } = useWindowDimensions();
  const [apiUrl, setApiUrl] = useState(DEFAULT_API);
  const [apiEdit, setApiEdit] = useState(DEFAULT_API);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [apiReady, setApiReady] = useState(false);
  const [busy, setBusy] = useState('');
  const [tab, setTab] = useState<'labels' | 'count'>('labels');

  const [database, setDatabase] = useState({ total: 0, sheetName: '', sourceFile: '', ready: false });
  const [databaseFile, setDatabaseFile] = useState('');
  const [uploadId, setUploadId] = useState('');
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [sheetName, setSheetName] = useState('');
  const [workbookPreview, setWorkbookPreview] = useState<WorkbookPreview | null>(null);
  const [codeColumn, setCodeColumn] = useState(1);
  const [nameColumn, setNameColumn] = useState(2);
  const [quantityColumn, setQuantityColumn] = useState(6);

  const [template, setTemplate] = useState<TemplateInfo | null>(null);
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

  const selectedSlide = template?.slides?.[templateSlide] || template?.slides?.[0];
  const manualCapacity = Math.min(customCodeIds.split(/[\s,;]+/).filter(Boolean).length, customNameIds.split(/[\s,;]+/).filter(Boolean).length);
  const capacity = selectedSlide ? (selectedSlide.capacity || manualCapacity || 0) : 0;
  const requiredPages = Math.max(1, Math.ceil(labels.length / Math.max(capacity, 1)));
  const activeLabel = labels[activeLabelIndex] || labels[0];
  const validCustomFields = customCodeIds.trim().length > 0 && customNameIds.trim().length > 0;

  const apiCall = useCallback(<T,>(endpoint: string, options?: RequestInit) => fetchJson<T>(apiUrl, endpoint, options), [apiUrl]);

  const refreshCounts = useCallback(async () => {
    const result = await fetchJson<any>(apiUrl, '/api/counts');
    setCounts(result.results || []);
  }, [apiUrl]);

  useEffect(() => {
    let alive = true;
    if (!apiUrl.trim()) {
      setApiReady(false);
      setTemplate(null);
      setDatabase({ total: 0, sheetName: '', sourceFile: '', ready: false });
      return () => { alive = false; };
    }
    (async () => {
      try {
        await fetchJson(apiUrl, '/api/health');
        const [info, status] = await Promise.all([
          fetchJson<TemplateInfo>(apiUrl, '/api/templates/default'),
          fetchJson<any>(apiUrl, '/api/database/status')
        ]);
        if (!alive) return;
        setApiReady(true);
        setTemplate(info);
        setDatabase(status);
      } catch (_error) {
        if (alive) setApiReady(false);
      }
    })();
    return () => { alive = false; };
  }, [apiUrl]);

  useEffect(() => {
    if (tab === 'count' && apiReady) refreshCounts().catch(() => {});
  }, [tab, apiReady, refreshCounts]);

  useEffect(() => {
    if (inputMode !== 'name' || inputText.trim().length < 3 || !apiReady) { setSuggestions([]); return; }
    let alive = true;
    const q = encodeURIComponent(inputText.split(/[\n,;]+/).pop()?.trim() || '');
    fetchJson<any>(apiUrl, `/api/database/search?mode=name&q=${q}`).then(result => {
      if (alive) setSuggestions(result.results || []);
    }).catch(() => { if (alive) setSuggestions([]); });
    return () => { alive = false; };
  }, [inputText, inputMode, apiReady, apiUrl]);

  const loadSheetPreview = useCallback(async (id: string, name: string) => {
    if (!id || !name) return;
    setBusy('Lendo cabeçalhos da aba...');
    try {
      const preview = await apiCall<WorkbookPreview>('/api/database/preview', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ uploadId: id, sheetName: name }) });
      setWorkbookPreview(preview);
      setCodeColumn(preview.suggested.codeColumn || 1);
      setNameColumn(preview.suggested.nameColumn || 2);
      setQuantityColumn(preview.suggested.quantityColumn || 0);
    } catch (error: any) { showMessage(error.message); }
    finally { setBusy(''); }
  }, [apiCall]);

  const onChooseDatabase = async () => {
    try {
      const asset = await chooseFile('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      if (!asset) return;
      setBusy('Analisando planilha...');
      const result = await uploadPickedFile(apiUrl, '/api/database/inspect', asset);
      setUploadId(result.uploadId);
      setDatabaseFile(result.fileName);
      setSheetNames(result.sheets || []);
      const first = result.sheets?.[0] || '';
      setSheetName(first);
      await loadSheetPreview(result.uploadId, first);
    } catch (error: any) { showMessage(error.message); }
    finally { setBusy(''); }
  };

  const onImportDatabase = async () => {
    if (!uploadId || !sheetName) return showMessage('Selecione uma planilha e uma aba primeiro.');
    setBusy('Importando base de referência...');
    try {
      const result = await apiCall<any>('/api/database/import', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ uploadId, sheetName, mapping: { codeColumn, nameColumn, quantityColumn } }) });
      setDatabase({ total: result.total || result.imported, ready: true, sheetName: result.sheetName, sourceFile: result.sourceFile });
      showMessage(`${result.imported.toLocaleString('pt-BR')} itens importados da aba ${result.sheetName}.`);
    } catch (error: any) { showMessage(error.message); }
    finally { setBusy(''); }
  };

  const onChooseTemplate = async () => {
    try {
      const asset = await chooseFile('application/vnd.openxmlformats-officedocument.presentationml.presentation');
      if (!asset) return;
      setBusy('Analisando modelo PowerPoint...');
      const result = await uploadPickedFile(apiUrl, '/api/templates/upload', asset);
      setTemplate(result);
      setTemplateSlide(0);
      setCustomCodeIds('');
      setCustomNameIds('');
      const firstSlide = result.slides?.[0];
      const detected = (firstSlide?.capacity || 0) > 0;
      setTemplateError(detected ? '' : 'Não identifiquei automaticamente as caixas de código e nome. Informe os IDs das caixas no mapeamento avançado.');
    } catch (error: any) { showMessage(error.message); }
    finally { setBusy(''); }
  };

  const addItems = async () => {
    const values = splitValues(inputText);
    if (!values.length) return showMessage('Cole ou digite ao menos um código ou nome.');
    if (values.length > 1000) return showMessage('Adicione até 1.000 itens por vez.');
    setBusy('Conferindo itens na base...');
    try {
      const result = await apiCall<any>('/api/database/resolve', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ values, mode: inputMode }) });
      const mapped: LabelItem[] = (result.entries || []).map((entry: any) => {
        const match = entry.item;
        return {
          localId: makeId(),
          code: match?.code || (inputMode === 'code' ? entry.input : ''),
          name: match?.name || (inputMode === 'name' ? entry.input : ''),
          quantity: match?.quantity || '',
          sheetName: match?.sheetName || database.sheetName,
          sourceInput: entry.input,
        };
      });
      setLabels(previous => [...previous, ...mapped]);
      setActiveLabelIndex(labels.length);
      setInputText('');
      setSuggestions([]);
      const missing = (result.entries || []).filter((entry: any) => !entry.item).length;
      if (missing) showMessage(`${mapped.length - missing} item(ns) encontrado(s). ${missing} precisa(m) de revisão manual.`);
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
    setTemplateError('Mapeamento manual aplicado. Confira a prévia antes de exportar.');
  };

  const exportFile = async (format: 'pptx' | 'pdf') => {
    if (!labels.length) return showMessage('Adicione itens antes de gerar as plaquetas.');
    if (!template) return showMessage('Nenhum modelo PowerPoint está configurado.');
    const ids = (value: string) => value.split(/[\s,;]+/).map(v => v.trim()).filter(Boolean);
    const fields = validCustomFields ? { codeShapeIds: ids(customCodeIds), nameShapeIds: ids(customNameIds) } : {};
    setBusy(format === 'pdf' ? 'Montando PDF de impressão...' : 'Montando PowerPoint...');
    try {
      const result = await apiCall<ExportResult>('/api/export', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          templateId: template.templateId || 'default', slideIndex: templateSlide,
          items: labels.map(({ code, name }) => ({ code, name })),
          requestedPages: requestedPages ? Number(requestedPages) : undefined,
          fields, format, fileName: 'plaquetas-by-lippe'
        })
      });
      setLastExport(result);
      const url = `${apiUrl.replace(/\/$/, '')}${result.downloadUrl}`;
      if (Platform.OS === 'web') {
        window.open(url, '_blank');
      } else {
        const localUri = `${FileSystem.cacheDirectory || ''}${result.fileName}`;
        const downloaded = await FileSystem.downloadAsync(url, localUri);
        if (await Sharing.isAvailableAsync()) await Sharing.shareAsync(downloaded.uri);
        else await Linking.openURL(url);
      }
      if (result.remainingItems > 0) showMessage(`${result.printedItems} plaquetas incluídas. ${result.remainingItems} ficaram fora do total de folhas escolhido.`);
    } catch (error: any) { showMessage(error.message); }
    finally { setBusy(''); }
  };

  const onScan = async (data: string) => {
    const now = Date.now();
    if (!data || (lastScan.current.code === data && now - lastScan.current.at < 1800)) return;
    lastScan.current = { code: data, at: now };
    setScanCode(data);
    try {
      const result = await apiCall<any>(`/api/database/search?mode=code&q=${encodeURIComponent(data)}`);
      const found = result.results?.find((item: CatalogItem) => item.code === data) || result.results?.[0];
      setScanName(found?.name || 'Item não encontrado na base');
    } catch { setScanName('Não foi possível consultar a base.'); }
  };
  const incrementCount = async () => {
    if (!scanCode.trim()) return showMessage('Leia um código ou digite o código do produto.');
    try {
      setBusy('Atualizando contagem...');
      await apiCall('/api/counts/increment', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ code: scanCode.trim(), name: scanName, amount: scanStep }) });
      await refreshCounts();
    } catch (error: any) { showMessage(error.message); }
    finally { setBusy(''); }
  };

  const openSettings = () => setSettingsOpen(value => !value);
  const applyApi = async () => {
    const value = apiEdit.trim().replace(/\/$/, '');
    if (value && !/^https?:\/\//i.test(value)) return showMessage('Informe uma URL começando com https:// ou http://.');
    setApiUrl(value);
    setSettingsOpen(false);
  };

  const pageLabel = useMemo(() => `${requiredPages} ${requiredPages === 1 ? 'folha' : 'folhas'}`, [requiredPages]);

  return <View style={styles.screen}>
    <StatusBar barStyle="dark-content" backgroundColor="#F4F7F4" />
    <View style={styles.topBar}>
      <View style={styles.brandMark}><Ionicons name="pricetag" size={22} color="#FFFFFF" /></View>
      <View style={{ flex: 1 }}>
        <Text style={styles.brandTitle}>PLAQUETAS <Text style={styles.brandMuted}>by lippe</Text></Text>
        <Text style={styles.brandSubtitle}>IDENTIFICAÇÃO · ESTOQUE</Text>
      </View>
      <Pressable onPress={openSettings} style={styles.iconButton}><Ionicons name="settings-outline" size={20} color={INK} /></Pressable>
      <View style={[styles.connectionDot, apiReady ? styles.dotOn : styles.dotOff]} />
    </View>

    {settingsOpen && <View style={styles.settingsPanel}>
      <Text style={styles.settingsTitle}>Conexão com o servidor</Text>
      <TextInput value={apiEdit} onChangeText={setApiEdit} placeholder="https://plaquetas-by-lippe.onrender.com" autoCapitalize="none" autoCorrect={false} keyboardType="url" style={styles.textInput} />
      <View style={styles.rowActions}><ActionButton title="Aplicar URL" onPress={applyApi} small /><ActionButton title="Fechar" onPress={openSettings} secondary small /></View>
      <Text style={styles.helper}>Use a URL HTTPS pública do serviço no Render. O endereço padrão já aponta para o serviço configurado neste projeto.</Text>
    </View>}

    <View style={styles.tabBar}>
      <Pressable onPress={() => setTab('labels')} style={[styles.tab, tab === 'labels' && styles.tabActive]}>
        <Ionicons name="pricetags-outline" size={17} color={tab === 'labels' ? GREEN : MUTED} />
        <Text style={[styles.tabText, tab === 'labels' && styles.tabTextActive]}>Gerar plaquetas</Text>
      </Pressable>
      <Pressable onPress={() => setTab('count')} style={[styles.tab, tab === 'count' && styles.tabActive]}>
        <Ionicons name="scan-outline" size={17} color={tab === 'count' ? GREEN : MUTED} />
        <Text style={[styles.tabText, tab === 'count' && styles.tabTextActive]}>Contagem</Text>
        <Text style={styles.betaTag}>BETA</Text>
      </Pressable>
    </View>

    {!apiReady && <View style={styles.apiWarning}>
      <Ionicons name="cloud-offline-outline" size={20} color="#A45B18" />
      <View style={{ flex: 1 }}><Text style={styles.apiWarningTitle}>Servidor não conectado</Text><Text style={styles.apiWarningText}>A interface abriu, mas a base, os modelos e as exportações dependem do servidor. Confira a URL HTTPS nas configurações.</Text></View>
      <ActionButton title="Configurar" onPress={() => setSettingsOpen(true)} secondary small />
    </View>}

    {tab === 'labels' ? <ScrollView contentContainerStyle={styles.content} keyboardShouldPersistTaps="handled">
      <View style={styles.heroRow}>
        <View style={{ flex: 1 }}>
          <Text style={styles.heroEyebrow}>CENTRAL DE ETIQUETAS</Text>
          <Text style={styles.heroTitle}>Prepare suas plaquetas.</Text>
          <Text style={styles.heroText}>Use sua planilha como referência, confira os dados e exporte no modelo original.</Text>
        </View>
        <View style={styles.heroIcon}><Ionicons name="albums-outline" size={34} color={GREEN} /></View>
      </View>

      <View style={styles.statRow}>
        <View style={styles.statCard}><Text style={styles.statLabel}>ITENS NA BASE</Text><Text style={styles.statValue}>{database.total.toLocaleString('pt-BR')}</Text><Text style={styles.statFoot}>{database.ready ? 'Base pronta' : 'Sem base importada'}</Text></View>
        <View style={styles.statCard}><Text style={styles.statLabel}>NA FILA</Text><Text style={styles.statValue}>{labels.length}</Text><Text style={styles.statFoot}>Plaquetas adicionadas</Text></View>
        <View style={styles.statCard}><Text style={styles.statLabel}>FOLHAS</Text><Text style={styles.statValue}>{labels.length ? requiredPages : '—'}</Text><Text style={styles.statFoot}>{capacity} por folha</Text></View>
      </View>

      <View style={[styles.mainGrid, width < 960 && styles.mainGridMobile]}>
        <View style={[styles.leftColumn, width < 960 && styles.columnMobile]}>
          <Card>
            <SectionTitle eyebrow="01 · REFERÊNCIA" title="Banco de dados" subtitle="Envie o Excel e escolha a aba que contém os itens." />
            <View style={styles.uploadRow}>
              <View style={styles.uploadIcon}><Ionicons name="file-tray-full-outline" size={22} color={GREEN} /></View>
              <View style={{ flex: 1 }}>
                <Text style={styles.uploadTitle}>{databaseFile || database.sourceFile || 'Nenhuma planilha selecionada'}</Text>
                <Text style={styles.helper}>{database.ready ? `${database.total.toLocaleString('pt-BR')} registros · ${database.sheetName}` : 'Formato aceito: .xlsx com uma ou mais abas'} · A importação substitui a base ativa</Text>
              </View>
              <ActionButton title="Enviar Excel" onPress={onChooseDatabase} secondary small disabled={!!busy} />
            </View>
            {sheetNames.length > 0 && <View style={styles.mappingBlock}>
              <Text style={styles.fieldLabel}>Aba do Excel</Text>
              <View style={styles.pickerWrap}><Picker selectedValue={sheetName} onValueChange={(value) => { setSheetName(value); loadSheetPreview(uploadId, value); }} style={styles.picker}><Picker.Item label="Selecione a aba" value="" />{sheetNames.map(name => <Picker.Item key={name} label={name} value={name} />)}</Picker></View>
              {workbookPreview && <>
                <Text style={styles.fieldLabel}>Mapeamento das colunas</Text>
                <View style={[styles.mappingRow, width < 550 && styles.mappingRowMobile]}>
                  <View style={styles.mappingField}><Text style={styles.miniLabel}>CÓDIGO</Text><View style={styles.pickerWrap}><Picker selectedValue={codeColumn} onValueChange={setCodeColumn} style={styles.picker}>{workbookPreview.columns.map(c => <Picker.Item key={c.index} label={`${c.index}. ${c.label}`} value={c.index} />)}</Picker></View></View>
                  <View style={styles.mappingField}><Text style={styles.miniLabel}>NOME / DESCRIÇÃO</Text><View style={styles.pickerWrap}><Picker selectedValue={nameColumn} onValueChange={setNameColumn} style={styles.picker}>{workbookPreview.columns.map(c => <Picker.Item key={c.index} label={`${c.index}. ${c.label}`} value={c.index} />)}</Picker></View></View>
                  <View style={styles.mappingField}><Text style={styles.miniLabel}>QUANTIDADE</Text><View style={styles.pickerWrap}><Picker selectedValue={quantityColumn} onValueChange={setQuantityColumn} style={styles.picker}><Picker.Item label="Não importar" value={0} />{workbookPreview.columns.map(c => <Picker.Item key={c.index} label={`${c.index}. ${c.label}`} value={c.index} />)}</Picker></View></View>
                </View>
                <View style={styles.previewTable}>
                  <Text style={styles.miniLabel}>AMOSTRA DA ABA SELECIONADA</Text>
                  <ScrollView horizontal showsHorizontalScrollIndicator>
                    <View>
                      <View style={styles.tableRow}>{workbookPreview.headers.slice(0, 7).map((h, i) => <Text key={i} style={styles.tableHead}>{h}</Text>)}</View>
                      {workbookPreview.rows.slice(0, 3).map((row, ri) => <View key={ri} style={styles.tableRow}>{row.slice(0, 7).map((cell, ci) => <Text key={ci} style={styles.tableCell}>{cell || '—'}</Text>)}</View>)}
                    </View>
                  </ScrollView>
                </View>
                <ActionButton title="Importar esta aba como referência" onPress={onImportDatabase} icon="cloud-upload-outline" disabled={!!busy} />
              </>}
            </View>}
          </Card>

          <Card>
            <SectionTitle eyebrow="02 · MODELO" title="Modelo da plaqueta" subtitle="O arquivo original é preservado. Só o código e o nome dos itens são trocados." />
            <View style={styles.uploadRow}>
              <View style={styles.uploadIcon}><Ionicons name="document-text-outline" size={22} color={GREEN} /></View>
              <View style={{ flex: 1 }}>
                <Text style={styles.uploadTitle}>{template?.name || 'Carregando modelo padrão...'}</Text>
                <Text style={styles.helper}>{template?.slides?.length || 0} modelo(s) detectado(s) · {capacity} plaquetas por folha</Text>
              </View>
              <ActionButton title="Trocar PPTX" onPress={onChooseTemplate} secondary small disabled={!!busy} />
            </View>
            {template?.slides?.length ? <>
              <Text style={styles.fieldLabel}>Tipo de plaqueta</Text>
              <View style={styles.sizeChoices}>{template.slides.map((slide, index) => <Pressable key={slide.index} onPress={() => { setTemplateSlide(index); setCompactPreview(index === 1); }} style={[styles.sizeChoice, templateSlide === index && styles.sizeChoiceActive]}>
                <Ionicons name={templateSlide === index ? 'radio-button-on' : 'radio-button-off'} color={templateSlide === index ? GREEN : MUTED} size={17} />
                <View style={{ flex: 1 }}><Text style={[styles.sizeChoiceTitle, templateSlide === index && styles.sizeChoiceTitleActive]}>{index === 0 ? 'Plaqueta maior' : index === 1 ? 'Plaqueta menor' : `Modelo ${index + 1}`}</Text><Text style={styles.sizeChoiceMeta}>{slide.capacity} por folha · {slide.orientation === 'portrait' ? 'Folha retrato' : 'Folha paisagem'}</Text></View>
              </Pressable>)}</View>
              {!!templateError && <Text style={styles.warningText}>{templateError}</Text>}
              {selectedSlide && selectedSlide.capacity === 0 && <View style={styles.advancedMap}>
                <Text style={styles.fieldLabel}>Mapeamento avançado de campos (IDs das caixas)</Text>
                <Text style={styles.helper}>Dê às caixas os nomes name.item e cod.item no PowerPoint, ou informe os IDs detectados abaixo, separados por vírgula.</Text>
                <Text style={styles.miniLabel}>IDs DO CÓDIGO</Text><TextInput value={customCodeIds} onChangeText={setCustomCodeIds} style={styles.textInput} placeholder="12, 26, 41..." />
                <Text style={styles.miniLabel}>IDs DO NOME</Text><TextInput value={customNameIds} onChangeText={setCustomNameIds} style={styles.textInput} placeholder="15, 29, 44..." />
                <Text style={styles.miniLabel}>CAIXAS DE TEXTO DETECTADAS · TOQUE PARA ATRIBUIR</Text>
                <View style={styles.shapeList}>
                  {selectedSlide.shapes.filter(shape => shape.text || shape.field).slice(0, 80).map(shape => <View key={shape.id} style={styles.shapeRow}>
                    <View style={{ flex: 1 }}><Text style={styles.shapeName}>#{shape.id} · {shape.shape}</Text><Text numberOfLines={1} style={styles.helper}>{shape.text || 'Caixa vazia'}</Text></View>
                    <Pressable onPress={() => addShapeId('code', shape.id)} style={styles.shapeAssign}><Text style={styles.shapeAssignText}>Código</Text></Pressable>
                    <Pressable onPress={() => addShapeId('name', shape.id)} style={styles.shapeAssign}><Text style={styles.shapeAssignText}>Nome</Text></Pressable>
                  </View>)}
                </View>
                <ActionButton title="Aplicar mapeamento" onPress={useFields} secondary small />
              </View>}
            </> : null}
          </Card>

          <Card>
            <SectionTitle eyebrow="03 · ITENS" title="Adicionar à lista" subtitle="Cole vários códigos ou nomes. Um item por linha, ou separados por vírgula." />
            <View style={styles.segmented}>
              <Pressable onPress={() => { setInputMode('code'); setSuggestions([]); }} style={[styles.segment, inputMode === 'code' && styles.segmentActive]}><Text style={[styles.segmentText, inputMode === 'code' && styles.segmentTextActive]}>Códigos</Text></Pressable>
              <Pressable onPress={() => { setInputMode('name'); setSuggestions([]); }} style={[styles.segment, inputMode === 'name' && styles.segmentActive]}><Text style={[styles.segmentText, inputMode === 'name' && styles.segmentTextActive]}>Nomes</Text></Pressable>
            </View>
            <TextInput value={inputText} onChangeText={setInputText} multiline textAlignVertical="top" placeholder={inputMode === 'code' ? 'Ex.:\n76120200487\n76150100651\n76100301045' : 'Digite o nome do item para pesquisar na base...'} style={[styles.textInput, styles.listInput]} />
            {inputMode === 'name' && suggestions.length > 0 && <View style={styles.suggestionPanel}>
              {suggestions.slice(0, 5).map((item, i) => <Pressable key={`${item.code}-${i}`} onPress={() => chooseSuggestion(item)} style={styles.suggestionRow}>
                <View style={{ flex: 1 }}><Text style={styles.suggestionName}>{item.name}</Text><Text style={styles.helper}>Código: {item.code} · Estoque: {item.quantity || '—'}</Text></View><Ionicons name="add-circle-outline" size={20} color={GREEN} />
              </Pressable>)}
            </View>}
            <View style={styles.rowActions}><ActionButton title="Adicionar itens" onPress={addItems} icon="add" disabled={!!busy || !apiReady} /><Text style={styles.helper}>Até 1.000 por envio</Text></View>
          </Card>
        </View>

        <View style={[styles.rightColumn, width < 960 && styles.columnMobile]}>
          <Card>
            <SectionTitle eyebrow="PRÉVIA INDIVIDUAL" title="Sua plaqueta" subtitle="A visualização mostra apenas uma unidade, não a folha inteira." />
            <View style={styles.previewControls}>
              <Pressable onPress={() => setCompactPreview(false)} style={[styles.previewToggle, !compactPreview && styles.previewToggleActive]}><Text style={[styles.previewToggleText, !compactPreview && styles.previewToggleTextActive]}>Maior</Text></Pressable>
              <Pressable onPress={() => setCompactPreview(true)} style={[styles.previewToggle, compactPreview && styles.previewToggleActive]}><Text style={[styles.previewToggleText, compactPreview && styles.previewToggleTextActive]}>Menor</Text></Pressable>
            </View>
            <View style={styles.labelStage}>
              <View style={[styles.labelPreview, compactPreview && styles.labelPreviewCompact]}>
                <View style={styles.labelTop}>
                  <View style={styles.logoCell}><Text style={styles.logoText}>CKTR</Text><Text style={styles.logoSub}>SERVIÇOS</Text></View>
                  <View style={styles.codeCell}><Text style={styles.labelSmall}>Código:</Text><Text numberOfLines={1} adjustsFontSizeToFit style={styles.labelCode}>{activeLabel?.code || '14.112.000'}</Text></View>
                </View>
                <View style={styles.itemCell}><Text style={styles.labelSmall}>Item:</Text><Text numberOfLines={compactPreview ? 3 : 4} adjustsFontSizeToFit style={styles.labelName}>{activeLabel?.name || 'NOME DO ITEM SELECIONADO'}</Text></View>
                <View style={styles.labelBottom}><View style={styles.bottomCell}><Text style={styles.labelSmall}>Quantidade:</Text></View><View style={styles.bottomCell}><Text style={styles.labelSmall}>Data de contagem:</Text></View></View>
              </View>
              <Text style={styles.helper}>Prévia ilustrativa. O arquivo exportado usa o layout exato do PowerPoint enviado.</Text>
            </View>
            <View style={styles.labelNav}>
              <Pressable onPress={() => setActiveLabelIndex(i => Math.max(0, i - 1))} disabled={activeLabelIndex <= 0} style={styles.navCircle}><Ionicons name="chevron-back" size={18} color={INK} /></Pressable>
              <Text style={styles.navCount}>{labels.length ? `${Math.min(activeLabelIndex + 1, labels.length)} de ${labels.length}` : 'Sem itens'}</Text>
              <Pressable onPress={() => setActiveLabelIndex(i => Math.min(labels.length - 1, i + 1))} disabled={activeLabelIndex >= labels.length - 1} style={styles.navCircle}><Ionicons name="chevron-forward" size={18} color={INK} /></Pressable>
            </View>
          </Card>

          <Card>
            <SectionTitle eyebrow="04 · REVISÃO" title={`Lista de plaquetas (${labels.length})`} subtitle="Você pode corrigir qualquer código ou nome antes de gerar o arquivo." />
            {labels.length === 0 ? <View style={styles.emptyState}><Ionicons name="file-tray-outline" size={29} color="#A7B4AC" /><Text style={styles.emptyTitle}>Sua lista está vazia</Text><Text style={styles.helper}>Adicione códigos ou nomes no painel ao lado.</Text></View> : <View style={styles.itemList}>
              {labels.map((item, index) => <View key={item.localId} style={styles.itemEditor}>
                <View style={styles.itemEditorTop}><Text style={styles.itemNumber}>#{String(index + 1).padStart(2, '0')}</Text><Text style={styles.itemStock}>Estoque: {item.quantity || '—'} · Aba: {item.sheetName || database.sheetName || '—'}</Text><Pressable onPress={() => removeLabel(item.localId)} hitSlop={8}><Ionicons name="trash-outline" size={17} color="#A95C54" /></Pressable></View>
                <Text style={styles.miniLabel}>CÓDIGO</Text><TextInput value={item.code} onChangeText={v => updateLabel(item.localId, 'code', v)} style={styles.editInput} />
                <Text style={styles.miniLabel}>NOME DO ITEM</Text><TextInput value={item.name} onChangeText={v => updateLabel(item.localId, 'name', v)} style={styles.editInput} />
                {!item.code || !item.name ? <Text style={styles.warningText}>Preencha o código e o nome antes de imprimir.</Text> : null}
                <Pressable onPress={() => setActiveLabelIndex(index)} style={styles.showPreviewLink}><Ionicons name="eye-outline" size={15} color={GREEN} /><Text style={styles.showPreviewText}>Visualizar esta plaqueta</Text></Pressable>
              </View>)}
            </View>}
            <View style={styles.exportBox}>
              <Text style={styles.exportTitle}>Folhas para impressão</Text>
              <View style={styles.exportRow}><View style={{ flex: 1 }}><Text style={styles.helper}>Mínimo automático: {pageLabel} · {capacity} plaquetas por folha.</Text></View><View style={styles.pageInputWrap}><Text style={styles.miniLabel}>FOLHAS</Text><TextInput keyboardType="number-pad" value={requestedPages} onChangeText={setRequestedPages} placeholder={String(requiredPages)} style={styles.pageInput} /></View></View>
              <ActionButton title="Visualizar arquivo completo (PDF)" onPress={() => exportFile('pdf')} icon="eye-outline" disabled={!!busy || !labels.length || !apiReady} />
              <View style={styles.exportButtons}><View style={{ flex: 1 }}><ActionButton title="Baixar PDF" onPress={() => exportFile('pdf')} secondary icon="download-outline" small disabled={!!busy || !labels.length || !apiReady} /></View><View style={{ flex: 1 }}><ActionButton title="Baixar PPTX" onPress={() => exportFile('pptx')} secondary icon="document-outline" small disabled={!!busy || !labels.length || !apiReady} /></View></View>
              {lastExport && <Text style={styles.successText}><Ionicons name="checkmark-circle" size={14} color={GREEN} /> Último arquivo: {lastExport.fileName} · {lastExport.pages} folha(s)</Text>}
            </View>
          </Card>
        </View>
      </View>
      <View style={styles.footer}><Text style={styles.footerText}>PLAQUETAS BY LIPPE · PROJETO INDEPENDENTE</Text><Text style={styles.footerText}>O PowerPoint original é mantido sem alterações de fonte, tamanho ou formatação.</Text></View>
    </ScrollView> : <ScrollView contentContainerStyle={styles.content} keyboardShouldPersistTaps="handled">
      <View style={styles.heroRow}><View style={{ flex: 1 }}><Text style={styles.heroEyebrow}>INVENTÁRIO · CONTAGEM</Text><Text style={styles.heroTitle}>Conte sem perder o ritmo.</Text><Text style={styles.heroText}>Leia o código do item e conte em unidades ou em lotes de dez.</Text></View><View style={styles.heroIcon}><Ionicons name="scan-circle-outline" size={36} color={GREEN} /></View></View>
      <View style={[styles.mainGrid, width < 960 && styles.mainGridMobile]}>
        <View style={[styles.leftColumn, width < 960 && styles.columnMobile]}>
          <Card>
            <SectionTitle eyebrow="LEITOR" title="Identificar item" subtitle="Use a câmera para ler QR Code/código de barras ou digite o código." />
            {scanEnabled ? <CameraScanner onScan={onScan} onClose={() => setScanEnabled(false)} /> : <ActionButton title="Abrir câmera" onPress={() => setScanEnabled(true)} icon="camera-outline" />}
            <Text style={styles.fieldLabel}>Código do item</Text><TextInput value={scanCode} onChangeText={setScanCode} placeholder="Escaneie ou digite o código" style={styles.textInput} autoCapitalize="none" />
            {!!scanName && <Text style={styles.scanName}>{scanName}</Text>}
            <Text style={styles.fieldLabel}>Tamanho da contagem</Text><View style={styles.stepChoices}><Pressable onPress={() => setScanStep(1)} style={[styles.stepChoice, scanStep === 1 && styles.stepChoiceActive]}><Text style={[styles.stepChoiceText, scanStep === 1 && styles.stepChoiceTextActive]}>+1 unidade</Text></Pressable><Pressable onPress={() => setScanStep(10)} style={[styles.stepChoice, scanStep === 10 && styles.stepChoiceActive]}><Text style={[styles.stepChoiceText, scanStep === 10 && styles.stepChoiceTextActive]}>+10 unidades</Text></Pressable></View>
            <ActionButton title={`Adicionar ${scanStep} à contagem`} onPress={incrementCount} icon="add-circle-outline" disabled={!scanCode.trim() || !!busy} />
          </Card>
        </View>
        <View style={[styles.rightColumn, width < 960 && styles.columnMobile]}>
          <Card>
            <SectionTitle eyebrow="RESULTADO" title={`Itens contados (${counts.length})`} subtitle="A quantidade é acumulada por código." />
            {counts.length === 0 ? <View style={styles.emptyState}><Ionicons name="calculator-outline" size={30} color="#A7B4AC" /><Text style={styles.emptyTitle}>Nenhum item contado</Text><Text style={styles.helper}>Leia um código e use os botões +1 ou +10.</Text></View> : counts.map(item => <View key={item.code} style={styles.countRow}><View style={{ flex: 1 }}><Text style={styles.countName}>{item.name || 'Item sem descrição'}</Text><Text style={styles.helper}>{item.code}</Text></View><Text style={styles.countValue}>{item.quantity || 0}</Text></View>)}
            {counts.length > 0 && <View style={styles.rowActions}><ActionButton title="Atualizar" onPress={refreshCounts} secondary small /><ActionButton title="Zerar contagem" onPress={async () => { if (Platform.OS === 'web' ? window.confirm('Apagar todos os registros de contagem?') : true) { await apiCall('/api/counts', { method: 'DELETE' }); await refreshCounts(); } }} secondary small /> </View>}
          </Card>
          <Card style={styles.noteCard}><Ionicons name="information-circle-outline" size={20} color={GREEN} /><Text style={styles.noteText}>Módulo experimental: a leitura identifica o produto; a contagem só é somada quando você toca em “Adicionar”. Assim você pode contar 10 unidades fisicamente e registrar de uma vez com +10.</Text></Card>
        </View>
      </View>
    </ScrollView>}

    {!!busy && <View style={styles.busyOverlay}><View style={styles.busyCard}><ActivityIndicator color={GREEN} size="small" /><Text style={styles.busyText}>{busy}</Text></View></View>}
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
  iconButton: { padding: 9, borderRadius: 10, backgroundColor: '#F2F6F3' }, connectionDot: { width: 8, height: 8, borderRadius: 9 }, dotOn: { backgroundColor: '#24A36B' }, dotOff: { backgroundColor: '#D59A52' },
  settingsPanel: { padding: 15, backgroundColor: '#FFFFFF', borderBottomWidth: 1, borderBottomColor: LINE, gap: 9 }, settingsTitle: { color: INK, fontSize: 13, fontWeight: '800' },
  tabBar: { flexDirection: 'row', backgroundColor: '#FFFFFF', paddingHorizontal: 22, borderBottomWidth: 1, borderBottomColor: LINE, gap: 22 }, tab: { flexDirection: 'row', alignItems: 'center', gap: 8, paddingVertical: 14, borderBottomWidth: 2, borderBottomColor: 'transparent' }, tabActive: { borderBottomColor: GREEN }, tabText: { color: MUTED, fontSize: 13, fontWeight: '700' }, tabTextActive: { color: GREEN }, betaTag: { color: '#A36B20', backgroundColor: '#FFF3D8', fontSize: 8, fontWeight: '900', paddingHorizontal: 5, paddingVertical: 2, borderRadius: 5 },
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
