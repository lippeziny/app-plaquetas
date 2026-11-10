import type { SlidePositionSettings } from './pptx';

export type PdfLabelItem = { code: string; name: string };

function escapeHtml(value: string): string {
  return String(value || '').replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

export function buildLabelsPdfHtml(
  items: PdfLabelItem[],
  options: { pages: number; capacity: number; orientation: 'portrait' | 'landscape'; fileTitle?: string; adjustments?: SlidePositionSettings },
): string {
  const capacity = Math.max(1, options.capacity || 8);
  const pages = Math.max(1, options.pages || 1);
  const columns = capacity === 1 ? 1 : 2;
  const rows = Math.ceil(capacity / columns);
  const pageMarkup: string[] = [];
  for (let page = 0; page < pages; page++) {
    const cards = [];
    for (let slot = 0; slot < capacity; slot++) {
      const item = items[page * capacity + slot];
      const globalCode = options.adjustments?.global?.code || { xMm: 0, yMm: 0 };
      const globalName = options.adjustments?.global?.name || { xMm: 0, yMm: 0 };
      const slotAdjustment = options.adjustments?.slots?.[String(slot)] || {};
      const code = slotAdjustment.code || { xMm: 0, yMm: 0 };
      const name = slotAdjustment.name || { xMm: 0, yMm: 0 };
      const codeTransform = `transform:translate(${globalCode.xMm + code.xMm}mm,${globalCode.yMm + code.yMm}mm);`;
      const nameTransform = `transform:translate(${globalName.xMm + name.xMm}mm,${globalName.yMm + name.yMm}mm);`;
      cards.push(`
        <article class="label">
          <header class="label-head">
            <div class="brand"><strong>C<span>K</span>TR</strong><small>SERVIÇOS</small></div>
            <div class="code"><b>Código:</b><strong style="${codeTransform}">${escapeHtml(item?.code || '')}</strong></div>
          </header>
          <section class="item"><b>Item:</b><div class="item-name" style="${nameTransform}">${escapeHtml(item?.name || '')}</div></section>
          <footer class="label-foot"><div><b>Quantidade:</b></div><div><b>Data de contagem:</b></div></footer>
        </article>`);
    }
    pageMarkup.push(`<main class="sheet"><div class="grid">${cards.join('')}</div></main>`);
  }
  const landscape = options.orientation === 'landscape';
  return `<!doctype html><html><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
    <title>${escapeHtml(options.fileTitle || 'Plaquetas by Lippe')}</title>
    <style>
      @page { size: A4 ${landscape ? 'landscape' : 'portrait'}; margin: 7mm; }
      * { box-sizing: border-box; }
      html,body { margin:0; padding:0; font-family: Arial, Helvetica, sans-serif; color:#080808; }
      .sheet { width:100%; height:${landscape ? '196mm' : '283mm'}; page-break-after:always; break-after:page; overflow:hidden; }
      .sheet:last-child { page-break-after:auto; break-after:auto; }
      .grid { width:100%; height:100%; display:grid; grid-template-columns:repeat(${columns},minmax(0,1fr)); grid-template-rows:repeat(${rows},minmax(0,1fr)); column-gap:2.2mm; row-gap:6.5mm; }
      .label { min-width:0; min-height:0; border:0.55mm solid #111; display:flex; flex-direction:column; overflow:hidden; }
      .label-head { display:flex; height:19%; min-height:15mm; border-bottom:0.55mm solid #111; }
      .brand { width:30%; border-right:0.55mm solid #111; display:flex; flex-direction:column; justify-content:center; align-items:center; line-height:1; }
      .brand strong { font-size:clamp(15px,3.5vw,29px); letter-spacing:-1.5px; }
      .brand strong span { color:#397ab9; }
      .brand small { font-size:6px; font-weight:900; letter-spacing:1.2px; margin-left:25%; }
      .code { flex:1; padding:1.2mm; display:flex; flex-direction:column; justify-content:space-between; }
      .code b,.item b,.label-foot b { font-size:clamp(8px,1.6vw,11px); }
      .code strong { text-align:center; font-size:clamp(13px,2.9vw,21px); line-height:1.1; overflow-wrap:anywhere; }
      .item { flex:1; min-height:0; display:flex; flex-direction:column; padding:1.2mm; border-bottom:0.55mm solid #111; }
      .item-name { flex:1; display:flex; justify-content:center; align-items:center; text-align:center; font-size:clamp(13px,2.8vw,20px); line-height:1.18; font-weight:800; padding:1mm 3mm; overflow-wrap:anywhere; }
      .label-foot { height:21%; min-height:12mm; display:grid; grid-template-columns:1fr 1fr; }
      .label-foot>div { padding:1mm; }
      .label-foot>div+div { border-left:0.55mm solid #111; }
      @media print { .sheet { margin:0; } }
    </style></head><body>${pageMarkup.join('')}</body></html>`;
}
