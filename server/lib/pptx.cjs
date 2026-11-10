const AdmZip = require('adm-zip');
const path = require('node:path');

function attr(xmlAttrs, name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const m = xmlAttrs.match(new RegExp(`(?:^|\\s)${escaped}="([^"]*)"`));
  return m ? decodeXml(m[1]) : '';
}
function decodeXml(value = '') {
  return value.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}
function encodeXml(value = '') {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}
function getZipText(zip, file) {
  const entry = zip.getEntry(file);
  if (!entry) throw new Error(`Arquivo interno do PowerPoint não encontrado: ${file}`);
  return entry.getData().toString('utf8');
}
function setZipText(zip, file, content) {
  const existing = zip.getEntry(file);
  if (existing) zip.updateFile(file, Buffer.from(content, 'utf8'));
  else zip.addFile(file, Buffer.from(content, 'utf8'));
}
function parseSlideShapes(slideXml) {
  const shapes = [];
  for (const match of slideXml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)) {
    const block = match[0];
    const nv = block.match(/<p:cNvPr\b([^>]*)\/?\s*>/);
    if (!nv) continue;
    const id = attr(nv[1], 'id');
    const name = attr(nv[1], 'name');
    const text = [...block.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)]
      .map(m => decodeXml(m[1])).join('');
    const semantic = name.toLowerCase();
    let field = null;
    if (/cod\.item|item\.cod|code\.item|item\.code/.test(semantic)) field = 'code';
    if (/name\.item|item\.name|nome\.item|item\.nome/.test(semantic)) field = 'name';
    if (!field && /blabla|nome\s+do\s+item|name\s+item|descri[cç][aã]o\s+do\s+item/i.test(text)) field = 'name';
    if (!field && (/14\.112\.000/.test(text) || /^\s*\d{2,3}(?:\.\d{3}){1,3}\s*$/.test(text))) field = 'code';
    const off = block.match(/<a:xfrm[\s\S]*?<a:off\b([^>]*)\/?\s*>[\s\S]*?<a:ext\b([^>]*)\/?\s*>/);
    const x = off ? Number(attr(off[1], 'x')) : null;
    const y = off ? Number(attr(off[1], 'y')) : null;
    const cx = off ? Number(attr(off[2], 'cx')) : null;
    const cy = off ? Number(attr(off[2], 'cy')) : null;
    shapes.push({ id, name, text, field, x, y, cx, cy });
  }
  return shapes;
}
function slideEntries(zip) {
  const presentation = getZipText(zip, 'ppt/presentation.xml');
  const relationships = getZipText(zip, 'ppt/_rels/presentation.xml.rels');
  const rels = new Map();
  for (const m of relationships.matchAll(/<Relationship\b([^>]*)\/?\s*>/g)) {
    const a = m[1];
    rels.set(attr(a, 'Id'), { target: attr(a, 'Target'), type: attr(a, 'Type') });
  }
  const list = presentation.match(/<p:sldIdLst\b[^>]*>([\s\S]*?)<\/p:sldIdLst>/);
  if (!list) throw new Error('A apresentação não contém uma lista de slides válida.');
  const slides = [];
  for (const m of list[1].matchAll(/<p:sldId\b([^>]*)\/?\s*>/g)) {
    const a = m[1];
    const rid = attr(a, 'r:id');
    const rel = rels.get(rid);
    if (!rel || !rel.target) continue;
    const target = path.posix.normalize(path.posix.join('ppt', rel.target));
    slides.push({ rid, target, id: attr(a, 'id') });
  }
  return { presentation, relationships, slides };
}
function inspectTemplate(filePathOrBuffer) {
  const zip = Buffer.isBuffer(filePathOrBuffer)
    ? new AdmZip(filePathOrBuffer)
    : new AdmZip(filePathOrBuffer);
  const { slides } = slideEntries(zip);
  const pres = getZipText(zip, 'ppt/presentation.xml');
  const size = pres.match(/<p:sldSz\b([^>]*)\/?\s*>/);
  const width = size ? Number(attr(size[1], 'cx')) : null;
  const height = size ? Number(attr(size[1], 'cy')) : null;
  const result = slides.map((slide, index) => {
    const xml = getZipText(zip, slide.target);
    const shapes = parseSlideShapes(xml);
    const nameFields = shapes.filter(s => s.field === 'name');
    const codeFields = shapes.filter(s => s.field === 'code');
    const isPortrait = height && width ? height > width : true;
    const capacity = Math.min(nameFields.length, codeFields.length);
    return {
      index,
      slideFile: slide.target,
      widthEmu: width,
      heightEmu: height,
      orientation: isPortrait ? 'portrait' : 'landscape',
      capacity,
      nameFields: nameFields.map(s => ({ id: s.id, shape: s.name, sample: s.text })),
      codeFields: codeFields.map(s => ({ id: s.id, shape: s.name, sample: s.text })),
      shapes: shapes.map(s => ({ id: s.id, shape: s.name, text: s.text, field: s.field }))
    };
  });
  return { slides: result, totalSlides: result.length, source: 'pptx' };
}
function replaceShapeText(slideXml, shapeId, value) {
  let replaced = false;
  const out = slideXml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, block => {
    const nv = block.match(/<p:cNvPr\b([^>]*)\/?\s*>/);
    if (!nv || attr(nv[1], 'id') !== String(shapeId)) return block;
    let seen = 0;
    let updated = block.replace(/(<a:t(?:\s[^>]*)?>)([\s\S]*?)(<\/a:t>)/g, (whole, open, _content, close) => {
      const next = seen++ === 0 ? encodeXml(value) : '';
      return `${open}${next}${close}`;
    });
    if (seen === 0) {
      updated = block.replace(/<a:p(\s[^>]*)?>([\s\S]*?)<\/a:p>/, (paragraph, attrs = '', inner) => {
        const endProps = inner.match(/<a:endParaRPr\b([^>]*)\/?\s*>/);
        const runProps = endProps ? `<a:rPr${endProps[1]}/>` : '<a:rPr lang="pt-BR"/>';
        return `<a:p${attrs}><a:r>${runProps}<a:t>${encodeXml(value)}</a:t></a:r>${inner}</a:p>`;
      });
      replaced = updated !== block;
    } else {
      replaced = true;
    }
    return updated;
  });
  if (!replaced) throw new Error(`Não foi possível substituir o campo de texto do PowerPoint (shape ${shapeId}).`);
  return out;
}
function cleanIdList(value, fallback) {
  if (!Array.isArray(value) || value.length === 0) return fallback;
  return value.map(v => String(v));
}
function buildPptx({ templatePath, templateBuffer, slideIndex = 0, items = [], requestedPages, fields = {} }) {
  const zip = templateBuffer ? new AdmZip(templateBuffer) : new AdmZip(templatePath);
  const meta = slideEntries(zip);
  const selected = meta.slides[Number(slideIndex)];
  if (!selected) throw new Error('O slide/modelo escolhido não existe no arquivo.');
  const baseXml = getZipText(zip, selected.target);
  const shapes = parseSlideShapes(baseXml);
  const detectedNames = shapes.filter(s => s.field === 'name').map(s => s.id);
  const detectedCodes = shapes.filter(s => s.field === 'code').map(s => s.id);
  const nameIds = cleanIdList(fields.nameShapeIds, detectedNames);
  const codeIds = cleanIdList(fields.codeShapeIds, detectedCodes);
  const capacity = Math.min(nameIds.length, codeIds.length);
  if (!capacity) throw new Error('O modelo não tem campos de código/nome reconhecidos. Nomeie as caixas como name.item e cod.item ou informe os IDs dos campos.');

  const minimumPages = Math.max(1, Math.ceil(items.length / capacity));
  const pages = Math.max(minimumPages, Math.min(100, Number(requestedPages) || minimumPages));
  const sourceName = path.posix.basename(selected.target);
  const sourceRelsPath = `ppt/slides/_rels/${sourceName}.rels`;
  const sourceRels = zip.getEntry(sourceRelsPath) ? getZipText(zip, sourceRelsPath) : null;
  const generatedSlides = [];
  const firstGeneratedNumber = 101;

  for (let page = 0; page < pages; page++) {
    let xml = baseXml;
    for (let slot = 0; slot < capacity; slot++) {
      const item = items[page * capacity + slot] || {};
      xml = replaceShapeText(xml, codeIds[slot], item.code || '');
      xml = replaceShapeText(xml, nameIds[slot], item.name || '');
    }
    const fileNumber = firstGeneratedNumber + page;
    const filename = `slide${fileNumber}.xml`;
    const slidePath = `ppt/slides/${filename}`;
    setZipText(zip, slidePath, xml);
    if (sourceRels) setZipText(zip, `ppt/slides/_rels/${filename}.rels`, sourceRels);
    generatedSlides.push({ slidePath, rid: `rIdPlaquetas${page + 1}`, id: 256 + page });
  }

  const presXml = getZipText(zip, 'ppt/presentation.xml');
  const listXml = `<p:sldIdLst>${generatedSlides.map(s => `<p:sldId id="${s.id}" r:id="${s.rid}"/>`).join('')}</p:sldIdLst>`;
  const patchedPresentation = presXml.replace(/<p:sldIdLst\b[^>]*>[\s\S]*?<\/p:sldIdLst>/, listXml);
  setZipText(zip, 'ppt/presentation.xml', patchedPresentation);

  let relXml = getZipText(zip, 'ppt/_rels/presentation.xml.rels');
  relXml = relXml.replace(/<Relationship\b(?=[^>]*Type="[^"]*\/slide")([^>]*)\/?\s*>/g, '');
  const newRels = generatedSlides.map(s => `<Relationship Id="${s.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/${path.posix.basename(s.slidePath)}"/>`).join('');
  relXml = relXml.replace('</Relationships>', `${newRels}</Relationships>`);
  setZipText(zip, 'ppt/_rels/presentation.xml.rels', relXml);

  let typesXml = getZipText(zip, '[Content_Types].xml');
  const overrides = generatedSlides.map(s => `<Override PartName="/${s.slidePath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('');
  typesXml = typesXml.replace('</Types>', `${overrides}</Types>`);
  setZipText(zip, '[Content_Types].xml', typesXml);

  return { buffer: zip.toBuffer(), pages, capacity, printedItems: Math.min(items.length, pages * capacity), remainingItems: Math.max(0, items.length - pages * capacity) };
}

module.exports = { inspectTemplate, buildPptx, parseSlideShapes, replaceShapeText };
