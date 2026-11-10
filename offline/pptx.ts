import JSZip from 'jszip';

export type TemplateField = { id: string; shape: string; sample: string };
export type TemplateVisual = {
  id: string;
  name: string;
  text: string;
  field: string | null;
  x: number;
  y: number;
  width: number;
  height: number;
  kind: 'text' | 'rect' | 'line' | 'image';
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
  fontSizePt?: number;
  color?: string;
  bold?: boolean;
  align?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'center' | 'bottom';
  imageBase64?: string;
  imageMime?: string;
};
export type TemplateSlide = {
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
export type TemplateInfo = { templateId: string; name: string; slides: TemplateSlide[] };

/** Applies manually selected code/name shape IDs to the slide model used by the live preview. */
export function applyTemplateFieldMapping(slide: TemplateSlide, codeIds: string[], nameIds: string[]): TemplateSlide {
  const shapeById = new Map(slide.shapes.map(shape => [String(shape.id), shape]));
  const uniqueKnown = (ids: string[]) => [...new Set(ids.map(String).map(id => id.trim()).filter(id => id && shapeById.has(id)))];
  const codes = uniqueKnown(codeIds);
  const names = uniqueKnown(nameIds);
  const codeSet = new Set(codes);
  const nameSet = new Set(names.filter(id => !codeSet.has(id)));
  const toField = (ids: string[]) => ids.map(id => {
    const shape = shapeById.get(id)!;
    return { id, shape: shape.shape, sample: shape.text };
  });
  return {
    ...slide,
    capacity: Math.min(codes.length, nameSet.size),
    codeFields: toField(codes),
    nameFields: toField([...nameSet]),
    visuals: slide.visuals.map(visual => ({
      ...visual,
      field: codeSet.has(String(visual.id)) ? 'code' : nameSet.has(String(visual.id)) ? 'name' : null,
    })),
  };
}
export type PptxItem = { code: string; name: string };
export type PptxFields = { codeShapeIds?: string[]; nameShapeIds?: string[] };
export type PositionDelta = { xMm: number; yMm: number };
export type FieldPositionDelta = { code?: PositionDelta; name?: PositionDelta };
export type SlidePositionSettings = {
  global: FieldPositionDelta;
  slots: Record<string, FieldPositionDelta>;
};
export type PptxBuildResult = {
  base64: string;
  pages: number;
  capacity: number;
  printedItems: number;
  remainingItems: number;
};

type Shape = { id: string; name: string; text: string; field: string | null };

function attr(xmlAttrs: string, name: string): string {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = xmlAttrs.match(new RegExp(`(?:^|\\s)${escaped}="([^"]*)"`));
  return match ? decodeXml(match[1]) : '';
}
function decodeXml(value = ''): string {
  return value.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}
function encodeXml(value = ''): string {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}
async function zipText(zip: JSZip, file: string): Promise<string> {
  const entry = zip.file(file);
  if (!entry) throw new Error(`Arquivo interno do PowerPoint não encontrado: ${file}`);
  return entry.async('string');
}
function parseSlideShapes(slideXml: string): Shape[] {
  const shapes: Shape[] = [];
  for (const match of slideXml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)) {
    const block = match[0];
    const nv = block.match(/<p:cNvPr\b([^>]*)\/?\s*>/);
    if (!nv) continue;
    const id = attr(nv[1], 'id');
    const name = attr(nv[1], 'name');
    const text = [...block.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)]
      .map(item => decodeXml(item[1])).join('');
    const semantic = name.toLowerCase();
    let field: string | null = null;
    if (/cod\.item|item\.cod|code\.item|item\.code/.test(semantic)) field = 'code';
    if (/name\.item|item\.name|nome\.item|item\.nome/.test(semantic)) field = 'name';
    if (!field && /blabla|nome\s+do\s+item|name\s+item|descri[cç][aã]o\s+do\s+item/i.test(text)) field = 'name';
    if (!field && (/14\.112\.000/.test(text) || /^\s*\d{2,3}(?:\.\d{3}){1,3}\s*$/.test(text))) field = 'code';
    shapes.push({ id, name, text, field });
  }
  return shapes;
}
async function getSlidePaths(zip: JSZip) {
  const presentation = await zipText(zip, 'ppt/presentation.xml');
  const relationships = await zipText(zip, 'ppt/_rels/presentation.xml.rels');
  const rels = new Map<string, { target: string; type: string }>();
  for (const match of relationships.matchAll(/<Relationship\b([^>]*)\/?\s*>/g)) {
    const attrs = match[1];
    rels.set(attr(attrs, 'Id'), { target: attr(attrs, 'Target'), type: attr(attrs, 'Type') });
  }
  const list = presentation.match(/<p:sldIdLst\b[^>]*>([\s\S]*?)<\/p:sldIdLst>/);
  if (!list) throw new Error('A apresentação não contém uma lista de slides válida.');
  const slides: { target: string }[] = [];
  for (const match of list[1].matchAll(/<p:sldId\b([^>]*)\/?\s*>/g)) {
    const rid = attr(match[1], 'r:id');
    const relationship = rels.get(rid);
    if (!relationship?.target) continue;
    const target = relationship.target.startsWith('/')
      ? relationship.target.slice(1)
      : `ppt/${relationship.target}`.replace(/\/\.\//g, '/').replace(/[^/]+\/\.\.\//g, '');
    slides.push({ target });
  }
  return { presentation, relationships, slides };
}

type Transform = { sx: number; sy: number; tx: number; ty: number };
const IDENTITY: Transform = { sx: 1, sy: 1, tx: 0, ty: 0 };
function xmlAttr(attrs: string, key: string): string {
  const match = attrs.match(new RegExp(`(?:^|\\s)${key}="([^"]*)"`));
  return match ? match[1] : '';
}
function tagBlock(source: string, tag: string, start: number): { block: string; end: number } | null {
  const open = new RegExp(`<${tag}\\b[^>]*>`, 'g');
  open.lastIndex = start;
  const first = open.exec(source);
  if (!first) return null;
  const closePattern = new RegExp(`<\\/?${tag}\\b[^>]*>`, 'g');
  closePattern.lastIndex = first.index;
  let depth = 0;
  let token: RegExpExecArray | null;
  while ((token = closePattern.exec(source))) {
    const closing = token[0].startsWith('</');
    const selfClosing = /\/>$/.test(token[0]);
    if (closing) depth--;
    else if (!selfClosing) depth++;
    if (depth === 0) return { block: source.slice(first.index, closePattern.lastIndex), end: closePattern.lastIndex };
  }
  return null;
}
function readXfrm(block: string, parentTag: 'p:grpSpPr' | 'p:spPr'): { offX: number; offY: number; extX: number; extY: number; childX: number; childY: number; childW: number; childH: number } | null {
  const parent = block.match(new RegExp(`<${parentTag}\\b[\\s\\S]*?<\\/${parentTag}>`))?.[0] || '';
  const xfrm = parent.match(/<a:xfrm\b[^>]*>([\s\S]*?)<\/a:xfrm>/)?.[1] || '';
  if (!xfrm) return null;
  const off = xfrm.match(/<a:off\b([^>]*)\/?\s*>/);
  const ext = xfrm.match(/<a:ext\b([^>]*)\/?\s*>/);
  if (!off || !ext) return null;
  const childOff = xfrm.match(/<a:chOff\b([^>]*)\/?\s*>/)?.[1] || '';
  const childExt = xfrm.match(/<a:chExt\b([^>]*)\/?\s*>/)?.[1] || '';
  return {
    offX: Number(xmlAttr(off[1], 'x')) || 0,
    offY: Number(xmlAttr(off[1], 'y')) || 0,
    extX: Number(xmlAttr(ext[1], 'cx')) || 0,
    extY: Number(xmlAttr(ext[1], 'cy')) || 0,
    childX: Number(xmlAttr(childOff, 'x')) || 0,
    childY: Number(xmlAttr(childOff, 'y')) || 0,
    childW: Number(xmlAttr(childExt, 'cx')) || 0,
    childH: Number(xmlAttr(childExt, 'cy')) || 0,
  };
}
function composeTransform(parent: Transform, group: NonNullable<ReturnType<typeof readXfrm>>): Transform {
  const scaleX = group.childW ? group.extX / group.childW : 1;
  const scaleY = group.childH ? group.extY / group.childH : 1;
  return {
    sx: parent.sx * scaleX,
    sy: parent.sy * scaleY,
    tx: parent.tx + parent.sx * (group.offX - group.childX * scaleX),
    ty: parent.ty + parent.sy * (group.offY - group.childY * scaleY),
  };
}
function slideObjectChildren(container: string): { tag: 'p:grpSp' | 'p:sp' | 'p:pic'; block: string }[] {
  const result: { tag: 'p:grpSp' | 'p:sp' | 'p:pic'; block: string }[] = [];
  const open = /<(p:grpSp|p:sp|p:pic)\b[^>]*>/g;
  let match: RegExpExecArray | null;
  let cursor = 0;
  while ((match = open.exec(container))) {
    if (match.index < cursor) continue;
    const tag = match[1] as 'p:grpSp' | 'p:sp' | 'p:pic';
    const found = tagBlock(container, tag, match.index);
    if (!found) continue;
    result.push({ tag, block: found.block });
    cursor = found.end;
    open.lastIndex = cursor;
  }
  return result;
}
function slideShapeId(block: string): string {
  const attrs = block.match(/<p:cNvPr\b([^>]*)\/?\s*>/)?.[1] || '';
  return xmlAttr(attrs, 'id');
}
function slideShapeName(block: string): string {
  const attrs = block.match(/<p:cNvPr\b([^>]*)\/?\s*>/)?.[1] || '';
  return decodeXml(xmlAttr(attrs, 'name'));
}
function firstHex(source: string): string | undefined {
  const value = source.match(/<a:srgbClr\b[^>]*val="([0-9A-Fa-f]{6})"/);
  return value?.[1] ? `#${value[1]}` : undefined;
}
function parseTextStyle(block: string) {
  const txBody = block.match(/<p:txBody\b[\s\S]*?<\/p:txBody>/)?.[0] || '';
  const run = txBody.match(/<a:rPr\b([^>]*)/)?.[1] || txBody.match(/<a:defRPr\b([^>]*)/)?.[1] || '';
  const bodyPr = txBody.match(/<a:bodyPr\b([^>]*)/)?.[1] || '';
  const paragraph = txBody.match(/<a:pPr\b([^>]*)/)?.[1] || '';
  const alignRaw = xmlAttr(paragraph, 'algn');
  const anchor = xmlAttr(bodyPr, 'anchor');
  const size = Number(xmlAttr(run, 'sz')) / 100;
  const bold = xmlAttr(run, 'b') === 'true' || xmlAttr(run, 'b') === '1';
  const runColor = txBody.match(/<a:rPr\b[\s\S]*?<a:solidFill>([\s\S]*?)<\/a:solidFill>/)?.[1] || '';
  return {
    fontSizePt: size > 0 ? size : undefined,
    color: firstHex(runColor),
    bold,
    align: alignRaw === 'ctr' ? 'center' as const : alignRaw === 'r' ? 'right' as const : 'left' as const,
    verticalAlign: anchor === 'ctr' ? 'center' as const : anchor === 'b' ? 'bottom' as const : 'top' as const,
  };
}
function normalizeZipTarget(basePath: string, target: string): string {
  const stack = basePath.split('/').slice(0, -1);
  for (const part of target.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') stack.pop(); else stack.push(part);
  }
  return stack.join('/');
}
async function parseSlideVisuals(zip: JSZip, slidePath: string, slideXml: string, fields: Shape[]): Promise<TemplateVisual[]> {
  const relationshipPath = `${slidePath.slice(0, slidePath.lastIndexOf('/'))}/_rels/${slidePath.split('/').pop()}.rels`;
  const relationships = zip.file(relationshipPath) ? await zipText(zip, relationshipPath) : '';
  const media = new Map<string, { path: string; mime: string }>();
  for (const match of relationships.matchAll(/<Relationship\b([^>]*)\/?\s*>/g)) {
    const attrs = match[1];
    const id = xmlAttr(attrs, 'Id');
    const target = xmlAttr(attrs, 'Target');
    if (!id || !target) continue;
    const path = normalizeZipTarget(slidePath, target);
    const lower = path.toLowerCase();
    const mime = lower.endsWith('.png') ? 'image/png' : lower.endsWith('.jpg') || lower.endsWith('.jpeg') ? 'image/jpeg' : lower.endsWith('.gif') ? 'image/gif' : lower.endsWith('.svg') ? 'image/svg+xml' : '';
    if (mime) media.set(id, { path, mime });
  }
  const fieldMap = new Map(fields.map(field => [field.id, field]));
  const visuals: TemplateVisual[] = [];
  const walk = async (container: string, transform: Transform): Promise<void> => {
    for (const object of slideObjectChildren(container)) {
      if (object.tag === 'p:grpSp') {
        const groupXfrm = readXfrm(object.block, 'p:grpSpPr');
        const innerGroup = object.block.replace(/^<p:grpSp\b[^>]*>/, '').replace(/<\/p:grpSp>\s*$/, '');
        await walk(innerGroup, groupXfrm ? composeTransform(transform, groupXfrm) : transform);
        continue;
      }
      const id = slideShapeId(object.block);
      const name = slideShapeName(object.block);
      const spPr = object.block.match(/<p:spPr\b[\s\S]*?<\/p:spPr>/)?.[0] || '';
      const xfrm = readXfrm(object.block, 'p:spPr');
      if (!xfrm) continue;
      const x = transform.sx * xfrm.offX + transform.tx;
      const y = transform.sy * xfrm.offY + transform.ty;
      const width = Math.max(0, transform.sx * xfrm.extX);
      const height = Math.max(0, transform.sy * xfrm.extY);
      const field = fieldMap.get(id);
      const text = [...object.block.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)].map(item => decodeXml(item[1])).join('');
      const imageRelId = object.block.match(/<a:blip\b[^>]*r:embed="([^"]+)"/)?.[1];
      const imageInfo = imageRelId ? media.get(imageRelId) : undefined;
      let imageBase64: string | undefined;
      if (imageInfo && imageInfo.mime !== 'image/svg+xml') {
        const file = zip.file(imageInfo.path);
        if (file) imageBase64 = await file.async('base64');
      }
      const isLine = /<a:prstGeom[^>]*prst="line"/.test(spPr) || width === 0 || height === 0;
      const fillSection = spPr.match(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/)?.[1] || '';
      const lineSection = spPr.match(/<a:ln\b[\s\S]*?<\/a:ln>/)?.[0] || '';
      const textStyle = parseTextStyle(object.block);
      const borderWidth = Number(lineSection.match(/<a:ln\b([^>]*)/)?.[1] && xmlAttr(lineSection.match(/<a:ln\b([^>]*)/)![1], 'w')) / 12700;
      visuals.push({
        id, name, text, field: field?.field || null, x, y, width, height,
        kind: imageBase64 ? 'image' : (text ? 'text' : (isLine ? 'line' : 'rect')),
        fill: firstHex(fillSection), stroke: firstHex(lineSection),
        strokeWidth: Number.isFinite(borderWidth) && borderWidth > 0 ? borderWidth : undefined,
        ...textStyle,
        imageBase64, imageMime: imageInfo?.mime,
      });
    }
  };
  const tree = slideXml.match(/<p:spTree\b[\s\S]*?<\/p:spTree>/)?.[0] || slideXml;
  await walk(tree, IDENTITY);
  return visuals;
}

export async function inspectPptx(base64: string, name = 'Modelo PowerPoint'): Promise<TemplateInfo> {
  const zip = await JSZip.loadAsync(base64, { base64: true, checkCRC32: false });
  const meta = await getSlidePaths(zip);
  const presentation = meta.presentation;
  const size = presentation.match(/<p:sldSz\b([^>]*)\/?\s*>/);
  const width = size ? Number(attr(size[1], 'cx')) : 0;
  const height = size ? Number(attr(size[1], 'cy')) : 0;
  const slides: TemplateSlide[] = [];
  for (let index = 0; index < meta.slides.length; index++) {
    const xml = await zipText(zip, meta.slides[index].target);
    const shapes = parseSlideShapes(xml);
    const visuals = await parseSlideVisuals(zip, meta.slides[index].target, xml, shapes);
    const names = shapes.filter(shape => shape.field === 'name');
    const codes = shapes.filter(shape => shape.field === 'code');
    slides.push({
      index,
      capacity: Math.min(names.length, codes.length),
      widthEmu: width,
      heightEmu: height,
      orientation: height > width ? 'portrait' : 'landscape',
      nameFields: names.map(shape => ({ id: shape.id, shape: shape.name, sample: shape.text })),
      codeFields: codes.map(shape => ({ id: shape.id, shape: shape.name, sample: shape.text })),
      shapes: shapes.map(shape => ({ id: shape.id, shape: shape.name, text: shape.text, field: shape.field })),
      visuals,
    });
  }
  if (!slides.length) throw new Error('O PowerPoint não contém slides.');
  return { templateId: 'local', name, slides };
}

function replaceShapeText(slideXml: string, shapeId: string, value: string): string {
  let replaced = false;
  const output = slideXml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, block => {
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
  return output;
}
function cleanIdList(value: string[] | undefined, fallback: string[]): string[] {
  return Array.isArray(value) && value.length ? value.map(String) : fallback;
}

function shiftShapePosition(slideXml: string, shapeId: string, delta: PositionDelta): string {
  if (!delta.xMm && !delta.yMm) return slideXml;
  let shifted = false;
  return slideXml.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, block => {
    const nv = block.match(/<p:cNvPr\b([^>]*)\/?\s*>/);
    if (!nv || attr(nv[1], 'id') !== String(shapeId)) return block;
    const dx = Math.round(Number(delta.xMm || 0) * 36000);
    const dy = Math.round(Number(delta.yMm || 0) * 36000);
    const updated = block.replace(/(<a:off\b[^>]*x=")(-?\d+)("[^>]*y=")(-?\d+)("[^>]*\/>)/, (_all, a, x, b, y, c) => {
      shifted = true;
      return `${a}${Number(x) + dx}${b}${Number(y) + dy}${c}`;
    });
    return updated;
  });
}

export async function buildPptx(
  templateBase64: string,
  slideIndex: number,
  items: PptxItem[],
  requestedPages?: number,
  fields: PptxFields = {},
  adjustments?: SlidePositionSettings,
): Promise<PptxBuildResult> {
  const zip = await JSZip.loadAsync(templateBase64, { base64: true, checkCRC32: false });
  const meta = await getSlidePaths(zip);
  const selected = meta.slides[Number(slideIndex)];
  if (!selected) throw new Error('O slide/modelo escolhido não existe no arquivo.');
  const baseXml = await zipText(zip, selected.target);
  const shapes = parseSlideShapes(baseXml);
  const detectedNames = shapes.filter(shape => shape.field === 'name').map(shape => shape.id);
  const detectedCodes = shapes.filter(shape => shape.field === 'code').map(shape => shape.id);
  const nameIds = cleanIdList(fields.nameShapeIds, detectedNames);
  const codeIds = cleanIdList(fields.codeShapeIds, detectedCodes);
  const availableIds = new Set(shapes.map(shape => String(shape.id)));
  const invalidIds = [...codeIds, ...nameIds].filter(id => !availableIds.has(String(id)));
  if (invalidIds.length) throw new Error(`O mapeamento contém ID(s) que não existem como caixa de texto neste slide: ${[...new Set(invalidIds)].join(', ')}.`);
  if (new Set(codeIds).size !== codeIds.length || new Set(nameIds).size !== nameIds.length) throw new Error('O mapeamento contém IDs repetidos. Remova os IDs duplicados.');
  if (codeIds.some(id => nameIds.includes(id))) throw new Error('A mesma caixa de texto não pode ser usada como código e nome.');
  const capacity = Math.min(nameIds.length, codeIds.length);
  if (!capacity) throw new Error('O modelo não tem campos de código/nome reconhecidos. Nomeie as caixas como name.item e cod.item ou informe os IDs manualmente.');

  const minimumPages = Math.max(1, Math.ceil(items.length / capacity));
  const pages = Math.max(minimumPages, Math.min(100, Number(requestedPages) || minimumPages));
  const sourceName = selected.target.split('/').pop() || 'slide1.xml';
  const sourceRelsPath = `ppt/slides/_rels/${sourceName}.rels`;
  const sourceRels = zip.file(sourceRelsPath) ? await zipText(zip, sourceRelsPath) : null;
  const generatedSlides: { slidePath: string; rid: string; id: number }[] = [];

  for (let page = 0; page < pages; page++) {
    let xml = baseXml;
    for (let slot = 0; slot < capacity; slot++) {
      const item = items[page * capacity + slot] || { code: '', name: '' };
      xml = replaceShapeText(xml, codeIds[slot], item.code || '');
      xml = replaceShapeText(xml, nameIds[slot], item.name || '');
      const globalCode = adjustments?.global?.code || { xMm: 0, yMm: 0 };
      const globalName = adjustments?.global?.name || { xMm: 0, yMm: 0 };
      const slotAdjustment = adjustments?.slots?.[String(slot)] || {};
      const slotCode = slotAdjustment.code || { xMm: 0, yMm: 0 };
      const slotName = slotAdjustment.name || { xMm: 0, yMm: 0 };
      xml = shiftShapePosition(xml, codeIds[slot], { xMm: globalCode.xMm + slotCode.xMm, yMm: globalCode.yMm + slotCode.yMm });
      xml = shiftShapePosition(xml, nameIds[slot], { xMm: globalName.xMm + slotName.xMm, yMm: globalName.yMm + slotName.yMm });
    }
    const fileNumber = 101 + page;
    const filename = `slide${fileNumber}.xml`;
    const slidePath = `ppt/slides/${filename}`;
    zip.file(slidePath, xml);
    if (sourceRels) zip.file(`ppt/slides/_rels/${filename}.rels`, sourceRels);
    generatedSlides.push({ slidePath, rid: `rIdPlaquetas${page + 1}`, id: 256 + page });
  }

  const listXml = `<p:sldIdLst>${generatedSlides.map(slide => `<p:sldId id="${slide.id}" r:id="${slide.rid}"/>`).join('')}</p:sldIdLst>`;
  const patchedPresentation = meta.presentation.replace(/<p:sldIdLst\b[^>]*>[\s\S]*?<\/p:sldIdLst>/, listXml);
  zip.file('ppt/presentation.xml', patchedPresentation);

  let relXml = meta.relationships;
  relXml = relXml.replace(/<Relationship\b(?=[^>]*Type="[^"]*\/slide")([^>]*)\/?\s*>/g, '');
  const newRels = generatedSlides.map(slide => `<Relationship Id="${slide.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/${slide.slidePath.split('/').pop()}"/>`).join('');
  zip.file('ppt/_rels/presentation.xml.rels', relXml.replace('</Relationships>', `${newRels}</Relationships>`));

  let typesXml = await zipText(zip, '[Content_Types].xml');
  const overrides = generatedSlides.map(slide => `<Override PartName="/${slide.slidePath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('');
  typesXml = typesXml.replace('</Types>', `${overrides}</Types>`);
  zip.file('[Content_Types].xml', typesXml);

  const output = await zip.generateAsync({ type: 'base64', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  return {
    base64: output,
    pages,
    capacity,
    printedItems: Math.min(items.length, pages * capacity),
    remainingItems: Math.max(0, items.length - pages * capacity),
  };
}
