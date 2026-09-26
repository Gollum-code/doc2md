import type { Block, Doc, ParseContext } from '../types.js';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

/** 从 PDF 抽取的文本项（带版面坐标，y 为页面顶部坐标，单位 pt）。 */
export interface PdfTextItem {
  str: string;
  x: number;
  y: number;
  w: number;
  h: number;
}

/** 用 pdf.js 抽取全文文本项，按页分组。 */
export async function extractPdfPages(data: Uint8Array): Promise<PdfTextItem[][]> {
  const task = getDocument({ data });
  const pdf = await task.promise;
  const pages: PdfTextItem[][] = [];
  try {
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const content = await page.getTextContent();
      const items: PdfTextItem[] = [];
      for (const item of content.items) {
        if (typeof item.str !== 'string' || !item.str.trim()) continue;
        items.push({
          str: item.str,
          // transform = [a b c d e f]；e=x，f=y（左下原点）
          x: item.transform[4],
          y: item.transform[5],
          w: item.width ?? 0,
          h: item.height ?? 0,
        });
      }
      pages.push(items);
      page.cleanup();
    }
  } finally {
    await pdf.destroy();
  }
  return pages;
}

/**
 * 阅读顺序：按 y 聚类成行（容差随字号缩放），行内按 x 排序。
 * 输出：自上而下、自左而右的文本行。
 */
export function orderTextItems(items: PdfTextItem[]): PdfTextItem[][] {
  const sorted = [...items].sort((a, b) => b.y - a.y); // 页面 y 越大越靠上
  const lines: PdfTextItem[][] = [];
  for (const it of sorted) {
    let placed = false;
    for (const line of lines) {
      const ref = line[0];
      if (Math.abs(it.y - ref.y) <= Math.max(ref.h * 0.9, 2)) {
        line.push(it);
        placed = true;
        break;
      }
    }
    if (!placed) lines.push([it]);
  }
  for (const line of lines) line.sort((a, b) => a.x - b.x);
  return lines;
}

export interface ColumnInfo {
  /** 各列 x 中心（升序）。 */
  x: number[];
  /** 该列模式出现的行数。 */
  support: number;
}

/**
 * 列对齐启发式：统计每一行的第一个文本项 x 位置，找跨多行重复出现的 x。
 * 当 ≥2 个列位置在 ≥2 行中重复 → 判定为表格。
 */
export function detectColumns(lines: PdfTextItem[][], opts: { minSupport?: number; tolerance?: number } = {}): ColumnInfo | undefined {
  const minSupport = opts.minSupport ?? 2;
  const tolerance = opts.tolerance ?? 1.5;
  // 统计所有文本项的 x 位置（不只看行首），跨行重复出现 ≥minSupport 次的 x 即候选列
  const allXs: number[] = [];
  for (const line of lines) {
    for (const it of line) allXs.push(it.x);
  }
  if (allXs.length < 2) return undefined;

  const clusters: number[] = [];
  const support: number[] = [];
  for (const x of allXs) {
    const idx = clusters.findIndex((c) => Math.abs(c - x) <= tolerance);
    if (idx >= 0) {
      support[idx] += 1;
    } else {
      clusters.push(x);
      support.push(1);
    }
  }
  const strong = clusters
    .map((x, i) => ({ x, support: support[i] }))
    .filter((c) => c.support >= minSupport)
    .sort((a, b) => a.x - b.x);
  if (strong.length >= 2) {
    return { x: strong.map((c) => c.x), support: strong[0].support };
  }
  return undefined;
}

/** 行拼接文本（保留词间距）。 */
export function lineToText(line: PdfTextItem[]): string {
  let out = '';
  let prevRight = -Infinity;
  for (const it of line) {
    if (prevRight > -Infinity && it.x - prevRight > Math.max(it.h * 0.15, 0.6)) {
      out += ' ';
    }
    out += it.str;
    prevRight = it.x + it.w;
  }
  return out.trim();
}

/**
 * 根据检测到的列把文本行组装成网格：每个文本项按 x 归入最近列。
 * 只保留 ≥2 个非空单元格的行（单列正文段落不进入表格），至少 2 行才判定为表格。
 */
export function linesToTable(lines: PdfTextItem[][], columns: ColumnInfo): string[][] | undefined {
  const grid: string[][] = [];
  for (const line of lines) {
    const row: string[] = columns.x.map(() => '');
    for (const it of line) {
      const colIdx = nearestCol(it.x, columns.x);
      if (colIdx < 0) continue;
      row[colIdx] = (row[colIdx] ? row[colIdx] + ' ' : '') + it.str;
    }
    // 单列内容视为正文段落，不是表格行
    if (row.filter((c) => c.trim()).length >= 2) grid.push(row);
  }
  if (grid.length >= 2 && columns.x.length >= 2) {
    return grid;
  }
  return undefined;
}

function nearestCol(x: number, cols: number[]): number {
  let best = -1;
  let bestDist = Infinity;
  for (let i = 0; i < cols.length; i++) {
    const d = Math.abs(x - cols[i]);
    if (d < bestDist) {
      bestDist = d;
      best = i;
    }
  }
  return bestDist <= 12 ? best : -1;
}

/** PDF → Doc（异步）。 */
export async function parsePdf(data: Uint8Array, ctx: ParseContext): Promise<Doc> {
  let pages: PdfTextItem[][];
  try {
    pages = await extractPdfPages(data);
  } catch (err) {
    ctx.warn(`PDF 文本抽取失败：${(err as Error).message}（文件可能被加密或是扫描件）`);
    return { format: 'pdf', meta: {}, blocks: [] };
  }
  if (pages.length === 0) {
    ctx.warn('PDF 未抽取到文本（可能是扫描件/图片型 PDF，OCR 暂未支持）');
  }

  const blocks: Block[] = [];
  for (let i = 0; i < pages.length; i++) {
    const lines = orderTextItems(pages[i]);
    if (!lines.length) continue;
    const columns = detectColumns(lines, { minSupport: 2 });
    const inner: Block[] = [];
    const grid = columns ? linesToTable(lines, columns) : undefined;
    if (grid) {
      inner.push({ type: 'table', header: grid[0] ?? [], rows: grid.slice(1) });
    } else {
      for (const line of lines) {
        const text = lineToText(line);
        if (text) inner.push({ type: 'paragraph', children: [{ kind: 'text', value: text }] });
      }
    }
    blocks.push({
      type: 'section',
      title: `第 ${i + 1} 页`,
      level: 2,
      meta: { 页码: String(i + 1) },
      blocks: inner,
    });
  }

  return { format: 'pdf', meta: { pages: pages.length }, blocks };
}