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
const templatePath = path.resolve(__dirname, '..', 'assets', 'default-template.pptx');

test('inspects the bundled template locally and finds both 8-label layouts', async () => {
  const base64 = fs.readFileSync(templatePath).toString('base64');
  const info = await pptxTools.inspectPptx(base64, 'Modelo padrão');
  assert.equal(info.slides.length, 2);
  assert.deepEqual(info.slides.map(slide => slide.capacity), [8, 8]);
  assert.equal(info.slides[0].orientation, 'portrait');
  assert.ok(info.slides[0].visuals.length > 0, 'the live preview must contain the original slide geometry');
  assert.ok(info.slides[0].visuals.some(item => item.kind === 'image' && item.imageBase64), 'template logos/images must be embedded in the preview data');
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
  assert.match(app, /maxImportBytes/);
  assert.match(app, /50 MB/);
  assert.doesNotMatch(app, /readPickedFileBase64\(asset\)/, 'Excel import must not cross the native bridge as Base64 text');
  assert.ok(config.expo.android.blockedPermissions.includes('android.permission.INTERNET'));
  assert.ok(config.expo.android.blockedPermissions.includes('android.permission.ACCESS_NETWORK_STATE'));
});

test('app exposes the real template preview, global/individual adjustments and Material You palettes', () => {
  const root = path.resolve(__dirname, '..');
  const app = fs.readFileSync(path.join(root, 'App.tsx'), 'utf8');
  const preview = fs.readFileSync(path.join(root, 'offline', 'TemplatePreview.tsx'), 'utf8');
  for (const text of ['Todas as plaquetas', 'Só uma plaqueta', 'Precisão do movimento', 'Ajustar posição do texto']) assert.ok(app.includes(text));
  for (const theme of ["id: 'azul'", "id: 'verde'", "id: 'roxo'", "id: 'laranja'", "id: 'turquesa'"]) assert.ok(app.includes(theme));
  assert.match(app, /<TemplatePreview/);
  assert.match(preview, /visual\.imageBase64/);
});
