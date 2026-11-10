const ExcelJS = require('exceljs');
const AdmZip = require('adm-zip');

function attr(xmlAttrs, name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = xmlAttrs.match(new RegExp(`(?:^|\\s)${escaped}="([^"]*)"`));
  return match ? match[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>') : '';
}
function getWorkbookSheets(filePath) {
  const zip = new AdmZip(filePath);
  const workbookXml = zip.getEntry('xl/workbook.xml')?.getData().toString('utf8');
  const relsXml = zip.getEntry('xl/_rels/workbook.xml.rels')?.getData().toString('utf8');
  if (!workbookXml || !relsXml) throw new Error('O arquivo não parece ser uma planilha Excel .xlsx válida.');
  const targets = new Map();
  for (const m of relsXml.matchAll(/<Relationship\b([^>]*)\/?\s*>/g)) {
    targets.set(attr(m[1], 'Id'), attr(m[1], 'Target'));
  }
  const sheets = [];
  for (const m of workbookXml.matchAll(/<sheet\b([^>]*)\/?\s*>/g)) {
    const a = m[1];
    const name = attr(a, 'name');
    const relId = attr(a, 'r:id');
    const target = targets.get(relId) || '';
    if (name) sheets.push({ name, relId, target });
  }
  if (!sheets.length) throw new Error('Nenhuma aba foi encontrada no arquivo Excel.');
  return sheets;
}
function plainCell(value) {
  if (value == null) return '';
  if (typeof value === 'object') {
    if (Object.prototype.hasOwnProperty.call(value, 'result')) return plainCell(value.result);
    if (Array.isArray(value.richText)) return value.richText.map(x => x.text || '').join('');
    if (value.text) return String(value.text);
    if (value.hyperlink) return String(value.text || value.hyperlink);
    if (value instanceof Date) return value.toISOString().slice(0, 10);
  }
  return String(value).trim();
}
function suggestColumn(headers, terms, fallback = 0) {
  const normalized = headers.map(h => String(h || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase());
  const index = normalized.findIndex(h => terms.some(term => h.includes(term)));
  return index >= 0 ? index + 1 : fallback;
}
async function withStreamingSheet(filePath, sheetName, handler) {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, {
    worksheets: 'emit',
    sharedStrings: 'cache',
    hyperlinks: 'ignore',
    styles: 'ignore',
    entries: 'ignore'
  });
  let found = false;
  for await (const worksheet of reader) {
    const isTarget = worksheet.name === sheetName;
    if (!isTarget) {
      for await (const _row of worksheet) { /* consume non-selected sheet */ }
      continue;
    }
    found = true;
    await handler(worksheet);
    break;
  }
  if (!found) throw new Error(`A aba "${sheetName}" não foi localizada na planilha.`);
}
async function previewSheet(filePath, sheetName, limit = 6) {
  let headers = [];
  const rows = [];
  await withStreamingSheet(filePath, sheetName, async worksheet => {
    let first = true;
    for await (const row of worksheet) {
      const values = (row.values || []).slice(1).map(plainCell);
      if (first) {
        headers = values.map((v, i) => v || `Coluna ${i + 1}`);
        first = false;
      } else if (values.some(Boolean)) {
        rows.push(values);
      }
      if (rows.length >= limit) break;
    }
  });
  const normalized = headers.map(h => String(h).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase());
  return {
    headers,
    rows,
    suggested: {
      codeColumn: suggestColumn(normalized, ['n do item', 'no item', 'codigo', 'código', 'material', 'item'], 1),
      nameColumn: suggestColumn(normalized, ['descricao', 'descrição', 'nome', 'item'], 2),
      quantityColumn: suggestColumn(normalized, ['em estoque', 'estoque', 'quantidade', 'qde'], 6)
    },
    columns: headers.map((label, index) => ({ index: index + 1, label }))
  };
}
async function importSheet(filePath, sheetName, mapping, db, sourceFile = '') {
  const codeColumn = Number(mapping.codeColumn);
  const nameColumn = Number(mapping.nameColumn);
  const quantityColumn = Number(mapping.quantityColumn) || 0;
  if (!codeColumn || !nameColumn) throw new Error('Selecione as colunas de código e nome antes de importar.');
  const insert = db.prepare('INSERT INTO items_stage (code, name, quantity, sheet_name, source_file) VALUES (?, ?, ?, ?, ?)');
  db.exec('DELETE FROM items_stage');
  let imported = 0;
  let skipped = 0;
  db.exec('BEGIN IMMEDIATE');
  try {
    await withStreamingSheet(filePath, sheetName, async worksheet => {
      let headerSkipped = false;
      for await (const row of worksheet) {
        if (!headerSkipped) { headerSkipped = true; continue; }
        const code = plainCell(row.getCell(codeColumn).value).replace(/\s+/g, ' ').trim();
        const name = plainCell(row.getCell(nameColumn).value).replace(/\s+/g, ' ').trim();
        const quantity = quantityColumn ? plainCell(row.getCell(quantityColumn).value) : '';
        if (!code || !name) { skipped++; continue; }
        insert.run(code, name, quantity, sheetName, sourceFile);
        imported++;
      }
    });
    db.exec('COMMIT');
  } catch (error) {
    db.exec('ROLLBACK');
    db.exec('DELETE FROM items_stage');
    throw error;
  }
  const replace = db.transaction(() => {
    db.exec('DELETE FROM items');
    db.exec('INSERT INTO items (code, name, quantity, sheet_name, source_file) SELECT code, name, quantity, sheet_name, source_file FROM items_stage');
    db.exec('DELETE FROM items_stage');
  });
  replace();
  return { imported, skipped, sheetName, sourceFile };
}

module.exports = { getWorkbookSheets, previewSheet, importSheet, plainCell };
