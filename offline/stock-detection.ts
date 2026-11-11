export type StockColumnMapping = {
  codeColumn: number;
  nameColumn: number;
  quantityColumn: number;
  headerRow: number;
};

function compact(value: string): string {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

const CODE_HEADERS = [
  'codigo', 'codigoproduto', 'codproduto', 'codmaterial', 'numerodoproduto',
  'numeroproduto', 'numerodomaterial', 'numeromaterial', 'numero', 'nitem',
  'numitem', 'nproduto', 'nmaterial', 'nitemproduto', 'ndoproduto', 'ndomaterial',
  'nrproduto', 'nrmaterial', 'numproduto', 'nummaterial', 'sku', 'referencia', 'ref', 'material', 'idproduto',
];
const NAME_HEADERS = [
  'descricao', 'descricaodoproduto', 'descricaodomaterial', 'descricaodoitem',
  'nome', 'nomedoproduto', 'nomedomaterial', 'nomeitem', 'item', 'produto', 'materialdescricao',
];
const QUANTITY_HEADERS = [
  'quantidade', 'quantidadeestoque', 'qtd', 'qtdestoque', 'qtdestoqueatual', 'qtddisponivel', 'qtde', 'qte', 'qnt', 'qntd', 'qty', 'quant',
  'estoque', 'saldo', 'saldoestoque', 'disponivel', 'quantidadedisponivel',
  'existencia', 'saldoatual',
];

function headerScore(value: string, aliases: string[]): number {
  const header = compact(value);
  if (!header) return 0;
  if (aliases.includes(header)) return 100;
  let best = 0;
  for (const alias of aliases) {
    if (alias.length >= 3 && header.includes(alias)) best = Math.max(best, 70 + Math.min(alias.length, 20));
    else if (alias.length >= 4 && alias.includes(header)) best = Math.max(best, 35 + header.length);
  }
  return best;
}

function bestColumn(headers: string[], aliases: string[], excluded: Set<number> = new Set()): { column: number; score: number } {
  let bestColumn = 0;
  let bestScore = 0;
  headers.forEach((header, index) => {
    const column = index + 1;
    if (excluded.has(column)) return;
    const score = headerScore(header, aliases);
    if (score > bestScore) { bestScore = score; bestColumn = column; }
  });
  return { column: bestColumn, score: bestScore };
}

/** Chooses the most likely stock tab, preferring an exact/leading "Estoque" name. */
export function chooseDefaultStockSheet(sheetNames: string[]): string {
  if (!sheetNames.length) return '';
  const normalized = sheetNames.map(name => ({ name, key: compact(name) }));
  return normalized.find(item => item.key === 'estoque')?.name
    || normalized.find(item => item.key.startsWith('estoque'))?.name
    || normalized.find(item => item.key.includes('estoque'))?.name
    || sheetNames[0];
}

/** Detects the header row and semantic columns from a small sample of worksheet rows.
 * Falls back to the legacy A/B/D mapping when the workbook does not provide useful headers. */
export function detectStockColumnMapping(rows: string[][], rowNumbers: number[] = []): StockColumnMapping {
  let best: { mapping: StockColumnMapping; score: number; hits: number } | null = null;
  const candidates = rows.slice(0, 30);
  for (let rowIndex = 0; rowIndex < candidates.length; rowIndex++) {
    const headers = candidates[rowIndex];
    const code = bestColumn(headers, CODE_HEADERS);
    const name = bestColumn(headers, NAME_HEADERS, new Set([code.column]));
    const quantity = bestColumn(headers, QUANTITY_HEADERS, new Set([code.column, name.column]));
    const hits = Number(code.score > 0) + Number(name.score > 0) + Number(quantity.score > 0);
    // Code + description are mandatory to identify a header row. Quantity is optional.
    if (code.score === 0 || name.score === 0) continue;
    const score = code.score + name.score + quantity.score + hits * 100 - rowIndex;
    if (best === null || score > best.score) {
      best = {
        score,
        hits,
        mapping: {
          codeColumn: code.column,
          nameColumn: name.column,
          quantityColumn: quantity.column || (headers.length >= 4 ? 4 : 0),
          headerRow: rowNumbers[rowIndex] || rowIndex + 1,
        },
      };
    }
  }
  return best === null ? { codeColumn: 1, nameColumn: 2, quantityColumn: 4, headerRow: rowNumbers[0] || 1 } : best.mapping;
}
