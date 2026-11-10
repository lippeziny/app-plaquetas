require('dotenv').config();
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const morgan = require('morgan');
const path = require('node:path');
const fs = require('node:fs');
const crypto = require('node:crypto');
const { execFile } = require('node:child_process');
const { promisify } = require('node:util');
const db = require('./db.cjs');
const { getWorkbookSheets, previewSheet, importSheet } = require('./lib/workbook.cjs');
const { inspectTemplate, buildPptx } = require('./lib/pptx.cjs');

const execFileAsync = promisify(execFile);
const ROOT = __dirname;
const DATA_DIR = path.resolve(process.env.APP_DATA_DIR || path.join(ROOT, 'data'));
const UPLOAD_DIR = path.join(ROOT, 'uploads');
const TEMPLATE_DIR = path.join(DATA_DIR, 'templates');
const OUTPUT_DIR = path.join(ROOT, 'outputs');
const DEFAULT_TEMPLATE = path.join(ROOT, 'assets', 'default-template.pptx');
const DIST_DIR = path.join(__dirname, '..', 'dist');
for (const dir of [DATA_DIR, UPLOAD_DIR, TEMPLATE_DIR, OUTPUT_DIR]) fs.mkdirSync(dir, { recursive: true });

const app = express();
app.use(cors());
app.use(express.json({ limit: '12mb' }));
app.use(morgan('tiny'));

const storage = multer.diskStorage({
  destination: (_req, _file, cb) => cb(null, UPLOAD_DIR),
  filename: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    cb(null, `${Date.now()}-${crypto.randomUUID()}${ext}`);
  }
});
const upload = multer({ storage, limits: { fileSize: 120 * 1024 * 1024 } });
const uploadSessions = new Map();
const templatePathFor = id => {
  if (!id || id === 'default') return DEFAULT_TEMPLATE;
  const safe = String(id).replace(/[^a-zA-Z0-9_-]/g, '');
  const candidate = path.join(TEMPLATE_DIR, `${safe}.pptx`);
  if (!fs.existsSync(candidate)) throw new Error('O modelo enviado não foi encontrado. Envie-o novamente.');
  return candidate;
};
function safeName(name) {
  return String(name || 'plaquetas').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 45) || 'plaquetas';
}

app.get('/api/health', (_req, res) => res.json({ ok: true, app: 'Plaquetas by Lippe', version: '1.0.0' }));

app.get('/api/templates/default', (_req, res, next) => {
  try {
    res.json({ templateId: 'default', name: 'PLAQUETAS_BY_FILLIPE.PPTX', ...inspectTemplate(DEFAULT_TEMPLATE) });
  } catch (error) { next(error); }
});

app.post('/api/templates/upload', upload.single('file'), (req, res, next) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Selecione um arquivo PowerPoint (.pptx).' });
    if (path.extname(req.file.originalname).toLowerCase() !== '.pptx') {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: 'O modelo precisa estar no formato .pptx.' });
    }
    const templateId = crypto.randomUUID().replace(/-/g, '');
    const target = path.join(TEMPLATE_DIR, `${templateId}.pptx`);
    fs.renameSync(req.file.path, target);
    const inspection = inspectTemplate(target);
    res.json({ templateId, name: req.file.originalname, ...inspection });
  } catch (error) { next(error); }
});

app.post('/api/database/inspect', upload.single('file'), (req, res, next) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Selecione uma planilha Excel (.xlsx).' });
    if (path.extname(req.file.originalname).toLowerCase() !== '.xlsx') {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: 'A base de dados precisa estar no formato .xlsx.' });
    }
    const sheets = getWorkbookSheets(req.file.path);
    const uploadId = crypto.randomUUID();
    uploadSessions.set(uploadId, { path: req.file.path, originalName: req.file.originalname, sheets, createdAt: Date.now() });
    res.json({ uploadId, fileName: req.file.originalname, sheets: sheets.map(s => s.name) });
  } catch (error) { next(error); }
});

app.post('/api/database/preview', async (req, res, next) => {
  try {
    const { uploadId, sheetName } = req.body || {};
    const session = uploadSessions.get(uploadId);
    if (!session) return res.status(404).json({ error: 'Envie a planilha novamente para continuar.' });
    const result = await previewSheet(session.path, sheetName, 5);
    res.json(result);
  } catch (error) { next(error); }
});

app.post('/api/database/import', async (req, res, next) => {
  try {
    const { uploadId, sheetName, mapping } = req.body || {};
    const session = uploadSessions.get(uploadId);
    if (!session) return res.status(404).json({ error: 'Envie a planilha novamente para continuar.' });
    const result = await importSheet(session.path, sheetName, mapping || {}, db, session.originalName);
    uploadSessions.delete(uploadId);
    try { fs.unlinkSync(session.path); } catch {}
    res.json({ ...result, total: db.prepare('SELECT COUNT(*) AS count FROM items').get().count });
  } catch (error) { next(error); }
});

app.get('/api/database/status', (_req, res) => {
  const total = db.prepare('SELECT COUNT(*) AS count FROM items').get().count;
  const source = db.prepare('SELECT sheet_name, source_file FROM items ORDER BY id DESC LIMIT 1').get() || {};
  res.json({ total, sheetName: source.sheet_name || '', sourceFile: source.source_file || '', ready: total > 0 });
});

app.get('/api/database/search', (req, res) => {
  const q = String(req.query.q || '').trim();
  const mode = req.query.mode === 'code' ? 'code' : 'name';
  if (!q) return res.json({ results: [] });
  const column = mode === 'code' ? 'code' : 'name';
  const results = db.prepare(`SELECT code, name, quantity, sheet_name AS sheetName FROM items WHERE ${column} LIKE ? COLLATE NOCASE ORDER BY CASE WHEN ${column} = ? COLLATE NOCASE THEN 0 ELSE 1 END, name LIMIT 16`).all(`%${q}%`, q);
  res.json({ results });
});

app.post('/api/database/resolve', (req, res) => {
  const { values = [], mode = 'code' } = req.body || {};
  if (!Array.isArray(values)) return res.status(400).json({ error: 'A lista enviada não é válida.' });
  const byCode = db.prepare('SELECT code, name, quantity, sheet_name AS sheetName FROM items WHERE code = ? COLLATE NOCASE LIMIT 1');
  const byName = db.prepare('SELECT code, name, quantity, sheet_name AS sheetName FROM items WHERE name = ? COLLATE NOCASE ORDER BY code LIMIT 1');
  const suggestionsByCode = db.prepare('SELECT code, name, quantity, sheet_name AS sheetName FROM items WHERE code LIKE ? COLLATE NOCASE ORDER BY code LIMIT 5');
  const suggestionsByName = db.prepare('SELECT code, name, quantity, sheet_name AS sheetName FROM items WHERE name LIKE ? COLLATE NOCASE ORDER BY CASE WHEN name LIKE ? COLLATE NOCASE THEN 0 ELSE 1 END, name LIMIT 5');
  const entries = values.slice(0, 1000).map(value => {
    const input = String(value || '').trim();
    if (!input) return null;
    const item = mode === 'name' ? byName.get(input) : byCode.get(input);
    const suggestions = item ? [] : (mode === 'name'
      ? suggestionsByName.all(`%${input}%`, `${input}%`)
      : suggestionsByCode.all(`%${input}%`));
    return { input, item: item || null, suggestions, matched: Boolean(item) };
  }).filter(Boolean);
  res.json({ entries });
});

app.get('/api/counts', (_req, res) => {
  const results = db.prepare('SELECT code, name, quantity, updated_at AS updatedAt FROM counts ORDER BY name COLLATE NOCASE').all();
  res.json({ results });
});
app.post('/api/counts/increment', (req, res) => {
  const code = String(req.body?.code || '').trim();
  const name = String(req.body?.name || '').trim();
  const amount = Number(req.body?.amount || 1);
  if (!code) return res.status(400).json({ error: 'Informe o código do item.' });
  if (![1, 10].includes(amount)) return res.status(400).json({ error: 'O incremento permitido é 1 ou 10.' });
  db.prepare(`INSERT INTO counts (code, name, quantity) VALUES (?, ?, ?)
    ON CONFLICT(code) DO UPDATE SET quantity = quantity + excluded.quantity,
    name = CASE WHEN excluded.name <> '' THEN excluded.name ELSE counts.name END,
    updated_at = CURRENT_TIMESTAMP`).run(code, name, amount);
  res.json({ item: db.prepare('SELECT code, name, quantity, updated_at AS updatedAt FROM counts WHERE code = ?').get(code) });
});
app.delete('/api/counts', (_req, res) => {
  db.exec('DELETE FROM counts');
  res.json({ ok: true });
});

app.post('/api/export', async (req, res, next) => {
  let tempPptx = null;
  try {
    const { templateId = 'default', slideIndex = 0, items = [], requestedPages, fields = {}, format = 'pptx', fileName } = req.body || {};
    if (!Array.isArray(items) || items.length === 0) return res.status(400).json({ error: 'Adicione pelo menos uma plaqueta antes de exportar.' });
    if (items.length > 5000) return res.status(400).json({ error: 'O limite desta versão é de 5.000 plaquetas por arquivo.' });
    const templatePath = templatePathFor(templateId);
    const result = buildPptx({ templatePath, slideIndex, items, requestedPages, fields });
    const outputId = `${Date.now()}-${crypto.randomUUID().slice(0, 8)}`;
    const baseName = safeName(fileName || 'plaquetas');
    tempPptx = path.join(OUTPUT_DIR, `${baseName}-${outputId}.pptx`);
    fs.writeFileSync(tempPptx, result.buffer);

    if (format === 'pdf') {
      try {
        await execFileAsync('libreoffice', [`-env:UserInstallation=file:///tmp/lo-${outputId}`, '--headless', '--convert-to', 'pdf', '--outdir', OUTPUT_DIR, tempPptx], { timeout: 240000 });
      } catch (error) {
        return res.status(501).json({ error: 'A conversão para PDF não está disponível neste servidor. Baixe o PPTX ou instale o LibreOffice no ambiente.', detail: error.message });
      }
      const pdfPath = tempPptx.replace(/\.pptx$/i, '.pdf');
      if (!fs.existsSync(pdfPath)) throw new Error('O PowerPoint foi criado, mas a conversão para PDF não gerou o arquivo esperado.');
      return res.json({
        fileName: path.basename(pdfPath),
        format: 'pdf',
        downloadUrl: `/files/${path.basename(pdfPath)}`,
        pages: result.pages,
        capacity: result.capacity,
        printedItems: result.printedItems,
        remainingItems: result.remainingItems
      });
    }
    res.json({
      fileName: path.basename(tempPptx),
      format: 'pptx',
      downloadUrl: `/files/${path.basename(tempPptx)}`,
      pages: result.pages,
      capacity: result.capacity,
      printedItems: result.printedItems,
      remainingItems: result.remainingItems
    });
  } catch (error) { next(error); }
});

app.use('/files', express.static(OUTPUT_DIR, { fallthrough: false, setHeaders(res) { res.setHeader('Cache-Control', 'no-store'); } }));
if (fs.existsSync(path.join(DIST_DIR, 'index.html'))) {
  app.use(express.static(DIST_DIR));
  app.get('*', (req, res, next) => req.path.startsWith('/api/') || req.path.startsWith('/files/')
    ? next() : res.sendFile(path.join(DIST_DIR, 'index.html')));
}
app.use((err, _req, res, _next) => {
  console.error(err);
  const message = err?.message || 'Erro interno do servidor.';
  res.status(err.status || 500).json({ error: message });
});

const port = Number(process.env.PORT || 4000);
app.listen(port, '0.0.0.0', () => console.log(`Plaquetas API disponível na porta ${port}`));
