const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const ts = require('typescript');
const JSZip = require('jszip');
const XLSX = require('xlsx');

function loadTypeScriptModule(relativePath) {
  const filename = path.resolve(__dirname, relativePath);
  const source = fs.readFileSync(filename, 'utf8');
  const compiled = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.CommonJS,
      target: ts.ScriptTarget.ES2020,
      esModuleInterop: true,
    },
  }).outputText;
  const loaded = new Module(filename, module);
  loaded.filename = filename;
  loaded.paths = Module._nodeModulePaths(path.dirname(filename));
  loaded._compile(compiled, filename);
  return loaded.exports;
}

const pptxTools = loadTypeScriptModule('../offline/pptx.ts');
const workbookTools = loadTypeScriptModule('../offline/workbook.ts');
const pdfTools = loadTypeScriptModule('../offline/pdf.ts');
const xlsxXmlTools = loadTypeScriptModule('../offline/xlsx-xml.ts');
const utf8Tools = loadTypeScriptModule('../offline/utf8-stream.ts');
const stockDetectionTools = loadTypeScriptModule('../offline/stock-detection.ts');
const templatePath = path.resolve(__dirname, '..', 'assets', 'default-template.pptx');

test('inspects the bundled template locally and finds both 8-label layouts', async () => {
  const base64 = fs.readFileSync(templatePath).toString('base64');
  const info = await pptxTools.inspectPptx(base64, 'Modelo padrão');
  assert.equal(info.slides.length, 2);
  assert.deepEqual(info.slides.map(slide => slide.capacity), [8, 8]);
  assert.equal(info.slides[0].orientation, 'portrait');
  assert.ok(info.slides[0].visuals.length > 0, 'the live preview must contain the original slide geometry');
  assert.ok(info.slides[0].visuals.some(item => item.kind === 'image' && item.imageBase64), 'template logos/images must be embedded in the preview data');
  assert.ok(info.slides[0].visuals.some(item => item.kind === 'image' && item.imageFit === 'cover'), 'PowerPoint crop instructions must be reflected in image preview sizing');
});

test('generates a valid multi-page PPTX locally without any server', async () => {
  const base64 = fs.readFileSync(templatePath).toString('base64');
  const items = Array.from({ length: 10 }, (_, index) => ({
    code: `76100000${index}`,
    name: `ITEM DE TESTE ${index}`,
  }));
  const result = await pptxTools.buildPptx(base64, 0, items);
  assert.equal(result.pages, 2);
  assert.equal(result.capacity, 8);
  assert.equal(result.printedItems, 10);
  const zip = await JSZip.loadAsync(result.base64, { base64: true });
  const presentation = await zip.file('ppt/presentation.xml').async('string');
  const slide = await zip.file('ppt/slides/slide101.xml').async('string');
  assert.equal((presentation.match(/<p:sldId\b/g) || []).length, 2);
  assert.match(slide, /761000000/);
  assert.match(slide, /ITEM DE TESTE 0/);
  assert.match(slide, /Quantidade:/);
});

test('manual field mapping is reflected in the live preview model', async () => {
  const base64 = fs.readFileSync(templatePath).toString('base64');
  const info = await pptxTools.inspectPptx(base64, 'Modelo padrão');
  const slide = info.slides[0];
  const codeId = slide.shapes.find(shape => shape.text.includes('14.112.000'))?.id;
  const nameId = slide.shapes.find(shape => shape.text.includes('BLABLA'))?.id;
  assert.ok(codeId && nameId);
  const mapped = pptxTools.applyTemplateFieldMapping(slide, [codeId], [nameId]);
  assert.equal(mapped.capacity, 1);
  assert.equal(mapped.visuals.find(item => item.id === codeId).field, 'code');
  assert.equal(mapped.visuals.find(item => item.id === nameId).field, 'name');
});

test('global and per-position offsets are applied to the original PPTX text boxes', async () => {
  const base64 = fs.readFileSync(templatePath).toString('base64');
  const info = await pptxTools.inspectPptx(base64, 'Modelo padrão');
  const codeId = info.slides[0].codeFields[0].id;
  const nameId = info.slides[0].nameFields[0].id;
  const originalZip = await JSZip.loadAsync(base64, { base64: true });
  const originalXml = await originalZip.file('ppt/slides/slide1.xml').async('string');
  const adjusted = await pptxTools.buildPptx(base64, 0, [{ code: '123', name: 'ITEM TESTE' }], 1, {}, {
    global: { name: { xMm: 1, yMm: -2 } },
    slots: { '0': { code: { xMm: 0.5, yMm: 0.5 } } },
  });
  const adjustedZip = await JSZip.loadAsync(adjusted.base64, { base64: true });
  const adjustedXml = await adjustedZip.file('ppt/slides/slide101.xml').async('string');
  const readOffset = (xml, id) => {
    const block = xml.match(/<p:sp\b[\s\S]*?<\/p:sp>/g).find(part => {
      const attrs = part.match(/<p:cNvPr\b([^>]*)/);
      return attrs && new RegExp(`\\bid=\"${id}\"`).test(attrs[1]);
    });
    assert.ok(block, `shape ${id} should exist`);
    const off = block.match(/<a:off\b([^>]*)/);
    return { x: Number(off[1].match(/\bx=\"(-?\d+)/)[1]), y: Number(off[1].match(/\by=\"(-?\d+)/)[1]) };
  };
  const beforeCode = readOffset(originalXml, codeId);
  const afterCode = readOffset(adjustedXml, codeId);
  assert.equal(afterCode.x - beforeCode.x, 18000);
  assert.equal(afterCode.y - beforeCode.y, 18000);
  const beforeName = readOffset(originalXml, nameId);
  const afterName = readOffset(adjustedXml, nameId);
  assert.equal(afterName.x - beforeName.x, 36000);
  assert.equal(afterName.y - beforeName.y, -72000);
});

test('imports Excel locally and suggests code, description and quantity columns', () => {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Código', 'Descrição do item', 'Depósito', 'Unidade', 'Status', 'Quantidade em estoque'],
    ['00123', 'MOTOR ELÉTRICO', 'A', 'UN', 'ATIVO', 17],
    ['00456', 'ROLAMENTO', 'B', 'UN', 'ATIVO', 9],
  ]);
  XLSX.utils.book_append_sheet(workbook, sheet, 'ESTOQUE SAP');
  const base64 = XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });
  const parsed = workbookTools.parseWorkbookBase64(base64);
  const byteParsed = workbookTools.parseWorkbookBytes(Buffer.from(base64, 'base64'));
  assert.deepEqual(byteParsed.sheetNames, parsed.sheetNames, 'byte-based import should discover the same sheets');
  const bytePreview = workbookTools.previewWorkbookSheet(byteParsed, 'ESTOQUE SAP');
  assert.equal(bytePreview.rows[0][1], 'MOTOR ELÉTRICO');
  const preview = workbookTools.previewWorkbookSheet(parsed, 'ESTOQUE SAP');
  assert.equal(preview.suggested.codeColumn, 1);
  assert.equal(preview.suggested.nameColumn, 2);
  assert.equal(preview.suggested.quantityColumn, 6);
  const result = workbookTools.importWorkbookSheet(parsed, 'ESTOQUE SAP', {
    codeColumn: 1, nameColumn: 2, quantityColumn: 6,
  }, 'estoque.xlsx');
  assert.equal(result.items.length, 2);
  assert.deepEqual(result.items[0], {
    code: '00123', name: 'MOTOR ELÉTRICO', quantity: '17', sheetName: 'ESTOQUE SAP',
  });
});

async function* textChunks(text, chunkSize = 23) {
  for (let index = 0; index < text.length; index += chunkSize) yield text.slice(index, index + chunkSize);
}

test('streams XLSX workbook metadata and resolves worksheet relationship paths', () => {
  const workbook = `<?xml version="1.0"?><workbook xmlns:r="urn:r"><sheets><sheet name="ESTOQUE SAP" sheetId="1" r:id="rId1"/></sheets></workbook>`;
  const rels = `<?xml version="1.0"?><Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`;
  const sheets = xlsxXmlTools.parseWorkbookXml(workbook);
  const relationships = xlsxXmlTools.parseWorkbookRelationshipsXml(rels);
  assert.deepEqual(sheets, [{ name: 'ESTOQUE SAP', relationId: 'rId1' }]);
  assert.equal(xlsxXmlTools.resolveWorksheetPath(`xl/${relationships.rId1}`), 'xl/worksheets/sheet1.xml');
});

test('detects the default stock tab and columns by semantic header names', () => {
  assert.equal(stockDetectionTools.chooseDefaultStockSheet(['Cadastro', 'Estoque SAP', 'Resumo']), 'Estoque SAP');
  assert.equal(stockDetectionTools.chooseDefaultStockSheet(['Produtos', 'Inventário']), 'Produtos');
  const detected = stockDetectionTools.detectStockColumnMapping([
    ['Relatório do almoxarifado'],
    ['Atualizado em 10/11/2026'],
    ['Número do produto', 'Descrição do produto', 'Unidade', 'Quantidade em estoque'],
    ['0012', 'CALÇA JEANS', 'UN', '25'],
  ], [1, 2, 4, 5]);
  assert.deepEqual(detected, { codeColumn: 1, nameColumn: 2, quantityColumn: 4, headerRow: 4 });
  const alternate = stockDetectionTools.detectStockColumnMapping([
    ['Código', 'Nome do item', 'Qtd.'],
    ['0099', 'MOTOR', '14'],
  ], [1, 2]);
  assert.deepEqual(alternate, { codeColumn: 1, nameColumn: 2, quantityColumn: 3, headerRow: 1 });
  const lateHeader = stockDetectionTools.detectStockColumnMapping([
    ['Relatório'], ['Emitido em 10/11'], ['Empresa X'], ['Depósito central'], [''],
    [''], [''], [''], [''], [''], ['Nº do material', 'Descrição', 'Unidade', 'Qtd. estoque'],
    ['0007', 'CORREIA', 'UN', '12'],
  ], Array.from({ length: 12 }, (_, index) => index + 1));
  assert.deepEqual(lateHeader, { codeColumn: 1, nameColumn: 2, quantityColumn: 4, headerRow: 11 });
});

test('reads only selected XLSX columns A, B and D in bounded chunks', async () => {
  const xml = `<?xml version="1.0"?><worksheet><dimension ref="A1:D4"/><sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>Código</t></is></c><c r="B1" t="inlineStr"><is><t>Descrição</t></is></c><c r="D1" t="inlineStr"><is><t>Quantidade</t></is></c></row>
    <row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2" t="s"><v>1</v></c><c r="C2"><v>IGNORAR</v></c><c r="D2"><v>17</v></c></row>
    <row r="3"><c r="A3" t="s"><v>2</v></c><c r="B3" t="s"><v>3</v></c><c r="D3"><v>9</v></c></row>
  </sheetData></worksheet>`;
  const received = [];
  const result = await xlsxXmlTools.parseWorksheetToImportRows(
    textChunks(xml, 19),
    { codeColumn: 1, nameColumn: 2, quantityColumn: 4 },
    async batch => received.push(...batch),
  );
  assert.equal(result.rowsRead, 2);
  assert.equal(result.totalRows, 4);
  assert.deepEqual(received.map(row => [row.rowNumber, row.codeRaw, row.nameRaw, row.quantityRaw]), [
    [2, '0', '1', '17'], [3, '2', '3', '9'],
  ]);
  assert.ok(received.every(row => !Object.values(row).includes('IGNORAR')));
});

test('streaming importer skips title rows and the detected header row before saving stock items', async () => {
  const xml = `<worksheet><dimension ref="A1:D5"/><sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>RELATÓRIO</t></is></c></row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>Estoque geral</t></is></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>Número</t></is></c><c r="B3" t="inlineStr"><is><t>Descrição</t></is></c><c r="D3" t="inlineStr"><is><t>Quantidade</t></is></c></row>
    <row r="4"><c r="A4"><v>123</v></c><c r="B4" t="inlineStr"><is><t>CALÇA JEANS</t></is></c><c r="D4"><v>7</v></c></row>
    <row r="5"><c r="A5"><v>124</v></c><c r="B5" t="inlineStr"><is><t>CAMISA</t></is></c><c r="D5"><v>3</v></c></row>
  </sheetData></worksheet>`;
  const received = [];
  const result = await xlsxXmlTools.parseWorksheetToImportRows(
    textChunks(xml, 37),
    { codeColumn: 1, nameColumn: 2, quantityColumn: 4, headerRow: 3 },
    async batch => received.push(...batch),
  );
  assert.equal(result.rowsRead, 2);
  assert.deepEqual(received.map(row => [row.rowNumber, row.codeRaw, row.nameRaw, row.quantityRaw]), [
    [4, '123', 'CALÇA JEANS', '7'],
    [5, '124', 'CAMISA', '3'],
  ]);
});

test('preserves Excel numeric formats for identifiers with leading zeroes', () => {
  const styles = `<styleSheet><numFmts><numFmt numFmtId="164" formatCode="000000"/><numFmt numFmtId="165" formatCode="000-000"/></numFmts><cellXfs><xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="165"/></cellXfs></styleSheet>`;
  const formats = xlsxXmlTools.parseStylesXml(styles);
  assert.equal(xlsxXmlTools.formatExcelValue('123', 'n:1', formats[1]), '000123');
  assert.equal(xlsxXmlTools.formatExcelValue('123', 'n:2', formats[2]), '000-123');
  assert.equal(xlsxXmlTools.formatExcelValue('17', 'n:0', formats[0]), '17');
});

test('decodes UTF-8 correctly when a multi-byte character crosses a 256KB chunk boundary', () => {
  const { Utf8StreamDecoder } = utf8Tools;
  const decoder = new Utf8StreamDecoder();
  const encoder = new TextEncoder();
  const source = `A`.repeat(256 * 1024 - 2) + 'ação — estoque';
  const bytes = encoder.encode(source);
  const first = decoder.decode(bytes.slice(0, 256 * 1024), false);
  const second = decoder.decode(bytes.slice(256 * 1024), false);
  const tail = decoder.decode(new Uint8Array(0), true);
  assert.equal(first + second + tail, source);
});

test('streams shared strings in batches, selects only referenced IDs and preview stops after the sample rows', async () => {
  const shared = `<sst><si><t>00123</t></si><si><t>MOTOR ELÉTRICO</t></si><si><t>00456</t></si><si><t>ROLAMENTO</t></si><si><t>STRING NÃO UTILIZADA</t></si></sst>`;
  const imported = [];
  const parsed = await xlsxXmlTools.parseSharedStrings(textChunks(shared, 13), async batch => imported.push(...batch));
  assert.equal(parsed.count, 5);
  assert.deepEqual(imported.map(item => item.value), ['00123', 'MOTOR ELÉTRICO', '00456', 'ROLAMENTO', 'STRING NÃO UTILIZADA']);

  const selected = [];
  const selectedResult = await xlsxXmlTools.parseSharedStrings(
    textChunks(shared, 11), async batch => selected.push(...batch), undefined, undefined, [0, 3],
  );
  assert.equal(selectedResult.count, 2);
  assert.deepEqual(selected.map(item => [item.id, item.value]), [[0, '00123'], [3, 'ROLAMENTO']]);

  const previewXml = `<worksheet><dimension ref="A1:D9999"/><sheetData><row r="1"><c r="A1"><v>header</v></c></row><row r="2"><c r="A2"><v>one</v></c></row><row r="3"><c r="A3"><v>not-read</v></c></row></sheetData></worksheet>`;
  const preview = await xlsxXmlTools.readPreviewRows(textChunks(previewXml, 92), 2);
  assert.equal(preview.totalRows, 9999);
  assert.equal(preview.rows.length, 2);
  assert.equal(preview.rows[1].cells.get(1).value, 'one');
});

test('builds a local print document and escapes imported text', () => {
  const html = pdfTools.buildLabelsPdfHtml([
    { code: '001', name: 'PEÇA <ESPECIAL> & TESTE' },
  ], { pages: 1, capacity: 8, orientation: 'portrait' });
  assert.match(html, /PEÇA &lt;ESPECIAL&gt; &amp; TESTE/);
  assert.match(html, /@page \{ size: A4 portrait/);
  const adjusted = pdfTools.buildLabelsPdfHtml([{ code: '001', name: 'ITEM' }], { pages: 1, capacity: 8, orientation: 'portrait', adjustments: { global: { name: { xMm: 1, yMm: -2 } }, slots: {} } });
  assert.match(adjusted, /translate\(1mm,-2mm\)/);
});

test('mobile app contains no backend API calls and blocks Android network permissions', () => {
  const root = path.resolve(__dirname, '..');
  const app = fs.readFileSync(path.join(root, 'App.tsx'), 'utf8');
  const config = JSON.parse(fs.readFileSync(path.join(root, 'app.json'), 'utf8'));
  assert.doesNotMatch(app, /EXPO_PUBLIC_API_URL|\/api\//);
  assert.match(app, /OFFLINE/);
  assert.match(app, /readPickedFileBytes/);
  assert.match(app, /bytesToBase64\(await readPickedFileBytes\(asset\)\)/, 'custom PowerPoint templates must use the byte reader before Base64 encoding');
  assert.doesNotMatch(app, /maxImportBytes|50 MB|250 mil linhas/, 'the native importer must not impose an arbitrary file/row size cap');
  assert.match(app, /inspectXlsxFile/);
  assert.match(app, /importXlsxSheetToCatalog/);
  assert.match(app, /etaSeconds/);
  assert.match(app, /setInterval\(\(\) =>/ , 'elapsed time and ETA must refresh while a long import stage is processing');
  assert.doesNotMatch(app, /readPickedFileBase64\(asset\)/, 'Excel import must not cross the native bridge as Base64 text');
  const importer = fs.readFileSync(path.join(root, 'offline', 'streaming-xlsx.ts'), 'utf8');
  assert.match(importer, /import \{ File as ExpoFile \} from 'expo-file-system';/, 'use the Expo SDK 54 File export');
  assert.match(importer, /const handle = file\.open\(\);/, 'SDK 54 File.open takes no mode argument');
  assert.doesNotMatch(importer, /FileMode|file\.open\(\s*[^)]/, 'do not use FileMode or pass arguments to File.open in the pinned SDK');
  assert.match(importer, /detectStockColumnMapping/, 'stock columns must be detected by header names');
  assert.match(importer, /getNeededSharedStringIds/);
  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  assert.ok(pkg.dependencies['expo-sqlite']);
  assert.ok(pkg.dependencies['react-native-zip-archive']);
  assert.ok(pkg.dependencies.saxes);
  assert.ok(pkg.dependencies['expo-file-system']);
  assert.ok(pkg.devDependencies.xlsx, 'SheetJS must not be bundled in the production app runtime');
  assert.ok(config.expo.plugins.includes('react-native-zip-archive'));
  assert.equal(config.expo.plugins.find(entry => Array.isArray(entry) && entry[0] === 'expo-build-properties')[1].ios.deploymentTarget, '15.5');
  assert.ok(config.expo.android.blockedPermissions.includes('android.permission.INTERNET'));
  assert.ok(config.expo.android.blockedPermissions.includes('android.permission.ACCESS_NETWORK_STATE'));
});

test('app exposes the real template preview, global/individual adjustments and Material You palettes', () => {
  const root = path.resolve(__dirname, '..');
  const app = fs.readFileSync(path.join(root, 'App.tsx'), 'utf8');
  const preview = fs.readFileSync(path.join(root, 'offline', 'TemplatePreview.tsx'), 'utf8');
  for (const text of ['Todas as plaquetas', 'Só uma plaqueta', 'Precisão do movimento', 'Ajustar posição do texto', 'Decorrido:', 'linhas/s']) assert.ok(app.includes(text));
  for (const theme of ["id: 'azul'", "id: 'verde'", "id: 'roxo'", "id: 'laranja'", "id: 'turquesa'"]) assert.ok(app.includes(theme));
  assert.match(app, /<TemplatePreview/);
  assert.match(preview, /visual\.imageBase64/);
});


test('catalog browser filters the complete imported stock database, not the selected label list', () => {
  const catalogSource = fs.readFileSync(path.resolve(__dirname, '..', 'offline', 'catalog-db.ts'), 'utf8');
  const appSource = fs.readFileSync(path.resolve(__dirname, '..', 'App.tsx'), 'utf8');
  const importer = fs.readFileSync(path.resolve(__dirname, '..', 'offline', 'streaming-xlsx.ts'), 'utf8');
  assert.match(catalogSource, /export async function searchCatalogSuggestions\([\s\S]*?field: 'code' \| 'name'/);
  assert.match(catalogSource, /export async function searchStockCatalogPage\(/);
  assert.match(catalogSource, /export async function countStockCatalogMatches\(/);
  assert.match(catalogSource, /name_norm LIKE \?/);
  assert.match(appSource, /searchCatalogSuggestions\(query, inputMode, 10\)/);
  assert.match(appSource, /quickSearchQuery/);
  assert.match(appSource, /Buscar um produto com preenchimento automático/);
  assert.match(appSource, /Adicionar produto selecionado/);
  assert.match(appSource, /Ou adicione vários produtos de uma vez/);
  assert.match(appSource, /setQuickSelectedItem\(item\)/);
  assert.match(appSource, /searchStockCatalogPage\(catalogSearch, 60, 0\)/);
  assert.match(appSource, /Busca pesquisa o estoque inteiro|filtro pesquisa o estoque inteiro/i);
  assert.match(appSource, /Adicionar \$\{selectedCatalogRows\.length\} selecionado/);
  assert.match(appSource, /accessibilityRole="checkbox"/);
  assert.match(appSource, /Carregar mais produtos/);
  assert.match(catalogSource, /FROM catalog WHERE name_norm LIKE/);
  assert.match(appSource, /stockWorkbookUri/);
  assert.match(appSource, /chooseDefaultStockSheet/);
  assert.match(appSource, /stockColumnMappings/);
  assert.match(appSource, /selectedSheetName/);
  assert.match(importer, /readPreviewRows\([\s\S]*?, 30,/);
  assert.match(appSource, /stockColumnMappings/);
  assert.match(appSource, /selectedSheetName/);
  assert.match(appSource, /Restaurar a última planilha de estoque|Restore the last stock workbook/);
});

test('live PPTX preview crops embedded square logo artwork to the actual wide shape bounds', () => {
  const previewSource = fs.readFileSync(path.resolve(__dirname, '..', 'offline', 'TemplatePreview.tsx'), 'utf8');
  assert.match(previewSource, /resizeMode=\{visual\.imageFit \|\| 'contain'\}/);
});
