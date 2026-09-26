import type { Block, Doc, ParseContext } from '../types.js';
import { OoxmlPackage } from '../core/ooxml.js';
import { child, children, strAttr, textOf, type XmlNode } from '../core/xml.js';
import { normalizeText } from '../normalize.js';

interface Sheet {
  name: string;
  part: string;
  state?: string;
}

/** XLSX 解析器。 */
export class XlsxParser {
  private sharedStrings: string[] = [];
  /** 单元格样式 → numFmtId */
  private cellXfs: number[] = [];
  private numFmts = new Map<number, string>();

  constructor(
    private readonly pkg: OoxmlPackage,
    private readonly ctx: ParseContext,
  ) {}

  parse(): Doc {
    const sheets = this.sheetList();
    const metaBase = this.pkg.readCoreProps();
    const meta: Doc['meta'] = {
      title: metaBase.title,
      author: metaBase.author,
      created: metaBase.created,
      modified: metaBase.modified,
      sheets: sheets.map((s) => s.name),
    };

    this.loadSharedStrings();
    this.loadStyles();

    const blocks: Block[] = [];
    let truncated = false;
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const table = this.parseSheet(sheet);
      if (!table) continue;
      const { rows, header, more } = table;
      const name = normalizeText(sheet.name);
      if (this.ctx.options.xlsxSheetTitles !== false) {
        blocks.push({
          type: 'section',
          title: name,
          level: 2,
          meta: { 工作表: String(i + 1) },
          blocks: [{ type: 'table', header, rows, align: undefined }],
        });
      } else {
        blocks.push({ type: 'table', header, rows, align: undefined });
      }
      if (more) truncated = true;
    }

    if (truncated) {
      this.ctx.warn('部分工作表因行数超出 maxRows 被截断');
    }

    return { format: 'xlsx', meta, blocks };
  }

  private sheetList(): Sheet[] {
    if (!this.pkg.has('xl/workbook.xml')) {
      this.ctx.warn('XLSX 未找到 xl/workbook.xml');
      return [];
    }
    const root = this.pkg.readXml('xl/workbook.xml') as XmlNode;
    const sheets: Sheet[] = [];
    for (const s of children(child(root, 'sheets'), 'sheet')) {
      const rid = strAttr(s, 'r:id') ?? strAttr(s, 'id');
      const rel = rid ? this.pkg.getRel('xl/workbook.xml', rid) : undefined;
      const name = strAttr(s, 'name') ?? `Sheet${sheets.length + 1}`;
      sheets.push({ name, part: rel?.targetPart ?? '', state: strAttr(s, 'state') });
    }
    // 找不到时兜底枚举
    if (!sheets.length) {
      const found = [...this.pkg.entries.keys()].filter((k) => /^xl\/worksheets\/sheet\d+\.xml$/.test(k));
      found.sort((a, b) => sheetNum(a) - sheetNum(b));
      for (let i = 0; i < found.length; i++) {
        sheets.push({ name: `Sheet${i + 1}`, part: found[i] });
      }
    }
    return sheets.filter((s) => s.part);
  }

  private loadSharedStrings(): void {
    if (!this.pkg.has('xl/sharedStrings.xml')) return;
    const root = this.pkg.readXml('xl/sharedStrings.xml') as XmlNode;
    for (const si of children(root, 'si')) {
      const t = child(si, 't');
      if (t !== undefined) {
        this.sharedStrings.push(textOf(t));
        continue;
      }
      // rich text runs
      const parts: string[] = [];
      for (const r of children(si, 'r')) {
        const rt = child(r, 't');
        if (rt !== undefined) parts.push(textOf(rt));
      }
      this.sharedStrings.push(parts.join(''));
    }
  }

  private loadStyles(): void {
    if (!this.pkg.has('xl/styles.xml')) return;
    const root = this.pkg.readXml('xl/styles.xml') as XmlNode;
    for (const nf of children(child(root, 'numFmts'), 'numFmt')) {
      const id = Number(strAttr(nf, 'numFmtId'));
      const code = strAttr(nf, 'formatCode') ?? '';
      if (Number.isFinite(id)) this.numFmts.set(id, code);
    }
    const cellXfs = child(child(root, 'cellXfs'), 'xf');
    if (cellXfs !== undefined) {
      const arr = Array.isArray(cellXfs) ? cellXfs : [cellXfs];
      for (const xf of arr) {
        const numFmtId = Number(strAttr(xf, 'numFmtId') ?? 0);
        this.cellXfs.push(Number.isFinite(numFmtId) ? numFmtId : 0);
      }
    }
  }

  private parseSheet(sheet: Sheet): { header: string[]; rows: string[][]; more: boolean } | undefined {
    if (!this.pkg.has(sheet.part)) return undefined;
    const root = this.pkg.readXml(sheet.part) as XmlNode;
    const sheetData = child(root, 'sheetData');
    if (sheetData === undefined) return { header: [], rows: [], more: false };

    const rowsOut: string[][] = [];
    let maxCols = 0;
    for (const row of children(sheetData, 'row')) {
      const cells = new Map<number, string>();
      let rowMax = 0;
      for (const c of children(row, 'c')) {
        const ref = strAttr(c, 'r') ?? '';
        const colIdx = ref ? colFromRef(ref) : rowMax;
        const value = this.cellValue(c);
        cells.set(colIdx, value);
        if (colIdx >= rowMax) rowMax = colIdx + 1;
      }
      const arr: string[] = [];
      for (let i = 0; i < rowMax; i++) arr.push(cells.get(i) ?? '');
      maxCols = Math.max(maxCols, arr.length);
      rowsOut.push(arr);
    }

    // 合并单元格：非左上角置空
    const merges = children(child(root, 'mergeCells'), 'mergeCell');
    const mergeSet = new Set<string>();
    const topLeftSet = new Set<string>();
    for (const mc of merges) {
      const ref = strAttr(mc, 'ref');
      if (!ref) continue;
      for (const cell of expandRange(ref)) mergeSet.add(cell);
      topLeftSet.add(ref.split(':')[0]);
    }

    // 应用合并：按行号打标记
    const finalGrid = rowsOut.map((row, ri) => {
      const rowCells = row.slice();
      for (let ci = 0; ci < rowCells.length; ci++) {
        const ref = `${colName(ci + 1)}${ri + 1}`;
        if (mergeSet.has(ref) && !topLeftSet.has(ref)) rowCells[ci] = '';
      }
      return rowCells;
    });

    const trimmed = trimGrid(finalGrid);
    if (!trimmed.length) return { header: [], rows: [], more: false };

    const header = trimmed[0] ?? [];
    let data = trimmed.slice(1);
    let more = false;
    const maxRows = this.ctx.options.maxRows ?? 2000;
    if (data.length > maxRows) {
      more = true;
      data = data.slice(0, maxRows);
    }
    return { header, rows: data, more };
  }

  private cellValue(c: XmlNode): string {
    const t = strAttr(c, 't');
    const v = child(c, 'v');
    const raw = v !== undefined ? textOf(v) : undefined;
    if (t === 's') {
      const idx = Number(raw ?? '0');
      return this.sharedStrings[idx] ?? '';
    }
    if (t === 'inlineStr') {
      const is = child(c, 'is');
      const tEl = child(is, 't');
      if (tEl !== undefined) return textOf(tEl);
      const parts: string[] = [];
      for (const r of children(is, 'r')) {
        const rt = child(r, 't');
        if (rt !== undefined) parts.push(textOf(rt));
      }
      return parts.join('');
    }
    if (t === 'b') {
      return raw === '1' ? 'TRUE' : 'FALSE';
    }
    if (t === 'e') return raw ?? '';
    if (raw === undefined || raw === '') return '';
    if (t === 'd') return raw;
    // 数值/字符串
    const s = Number(strAttr(c, 's') ?? 0);
    const numFmtId = this.cellXfs[s] ?? 0;
    const numFmt = this.numFmts.get(numFmtId);
    return formatCell(raw, numFmtId, numFmt);
  }
}

function sheetNum(path: string): number {
  const m = /sheet(\d+)\.xml$/.exec(path);
  return m ? Number(m[1]) : 0;
}

function colFromRef(ref: string): number {
  const m = /^([A-Z]+)/.exec(ref);
  if (!m) return 0;
  let n = 0;
  for (const ch of m[1]) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function colName(idx: number): string {
  let n = idx;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 展开 "A1:B2" 为单元格集合。 */
function expandRange(ref: string): string[] {
  const [a, b] = ref.split(':');
  if (!b) return [a];
  const am = /^([A-Z]+)(\d+)$/.exec(a);
  const bm = /^([A-Z]+)(\d+)$/.exec(b);
  if (!am || !bm) return [a];
  const c1 = colFromRef(`${am[1]}1`);
  const c2 = colFromRef(`${bm[1]}1`);
  const r1 = Number(am[2]);
  const r2 = Number(bm[2]);
  const out: string[] = [];
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      out.push(`${colName(c + 1)}${r}`);
    }
  }
  return out;
}

function trimGrid(grid: string[][]): string[][] {
  // 去掉全空行
  const rows = grid.filter((r) => r.some((c) => c.trim() !== ''));
  if (!rows.length) return [];
  let maxCols = 0;
  for (const r of rows) maxCols = Math.max(maxCols, r.length);
  // 去掉右侧全空列
  let rightCut = maxCols;
  for (let c = maxCols - 1; c >= 0; c--) {
    if (rows.some((r) => (r[c] ?? '').trim() !== '')) break;
    rightCut = c;
  }
  return rows.map((r) => {
    const out = r.slice(0, rightCut);
    while (out.length < rightCut) out.push('');
    return out;
  });
}

const BUILTIN_DATE_FMT: Record<number, boolean> = {
  14: true, 15: true, 16: true, 17: true, 18: true, 19: true, 20: true, 21: true, 22: true,
  45: true, 46: true, 47: true,
};

const BUILTIN_PERCENT_FMT: Record<number, boolean> = { 9: true, 10: true };

function formatCell(raw: string, numFmtId: number, numFmt: string | undefined): string {
  const n = Number(raw);
  if (!Number.isFinite(n)) return raw;

  if (BUILTIN_DATE_FMT[numFmtId]) {
    return excelDate(n);
  }
  if (numFmt) {
    const stripped = numFmt.replace(/"[^"]*"/g, '');
    const lower = stripped.toLowerCase();
    if (/[ymdhs]/.test(lower) && !lower.includes('e')) {
      return excelDate(n);
    }
    if (stripped.includes('%')) {
      return `${trimNumber(n * 100)}%`;
    }
  }
  if (BUILTIN_PERCENT_FMT[numFmtId]) {
    return `${trimNumber(n * 100)}%`;
  }
  return trimNumber(n);
}

function trimNumber(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  // 避免浮点误差尾巴，如 0.1+0.2
  const rounded = Number(n.toPrecision(12));
  return String(rounded);
}

function excelDate(serial: number): string {
  if (!Number.isFinite(serial) || serial < 1 || serial > 2958465) return String(serial);
  // Excel 序号起点 1899-12-30 → Unix epoch 偏移 25569 天
  const ms = Math.round((serial - 25569) * 86400 * 1000);
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return String(serial);
  const iso = d.toISOString();
  const hasTime = Math.abs(serial - Math.round(serial)) > 1e-9;
  return hasTime ? iso.replace('T', ' ').replace(/\.\d+Z$/, '') : iso.slice(0, 10);
}

/** 对外入口。 */
export function parseXlsx(pkg: OoxmlPackage, ctx: ParseContext): Doc {
  return new XlsxParser(pkg, ctx).parse();
}
