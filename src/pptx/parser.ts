import type { Block, Doc, Inline, ParseContext } from '../types.js';
import { OoxmlPackage, REL } from '../core/ooxml.js';
import { child, children, boolAttr, numAttr, strAttr, textOf, type XmlNode } from '../core/xml.js';
import { ListTreeBuilder } from '../core/lists.js';
import { normalizeText } from '../normalize.js';

/** EMU 转像素（96dpi）。 */
const EMU_PER_PX = 9525;

interface Run {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  link?: string;
  code?: boolean;
  sz?: number;
}

interface Para {
  runs: Run[];
  level: number;
  bullet: boolean;
  align?: string;
  empty: boolean;
}

interface Shape {
  name: string;
  phType?: string;
  phIdx?: string;
  kind: 'text' | 'table' | 'picture' | 'chart' | 'other';
  x: number;
  y: number;
  w: number;
  h: number;
  paras: Para[];
  table?: TableLike;
  image?: { target: string; alt?: string };
  chart?: { target: string; title?: string };
}

interface TableLike {
  grid: string[][];
  header: boolean;
  align: Array<'left' | 'center' | 'right'>;
}

interface SlideData {
  index: number;
  title: string;
  hidden: boolean;
  shapes: Shape[];
  notes: string;
}

const SLIDE_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster';

/** PPTX 解析器（按实例持有包与上下文，避免全局状态）。 */
export class PptxParser {
  private readonly layoutCache = new Map<string, string>();
  private readonly masterCache = new Map<string, string>();

  constructor(
    private readonly pkg: OoxmlPackage,
    private readonly ctx: ParseContext,
  ) {}

  parse(): Doc {
    const slidePaths = this.slideOrder();
    if (slidePaths.length === 0) {
      this.ctx.warn('PPTX 未找到任何幻灯片');
    }

    const metaBase = this.pkg.readCoreProps();
    const meta: Doc['meta'] = {
      title: metaBase.title,
      author: metaBase.author,
      created: metaBase.created,
      modified: metaBase.modified,
      lastModifiedBy: metaBase.lastModifiedBy,
      keywords: metaBase.keywordsRaw ? metaBase.keywordsRaw.split(',').map((s) => s.trim()).filter(Boolean) : undefined,
      slides: slidePaths.length,
    };

    const slides: SlideData[] = [];
    for (let i = 0; i < slidePaths.length; i++) {
      slides.push(this.parseSlide(slidePaths[i], i + 1));
    }

    return {
      format: 'pptx',
      meta,
      blocks: this.buildBlocks(slides),
    };
  }

  // ---- 包级结构 ----

  private slideOrder(): string[] {
    const presentation = 'ppt/presentation.xml';
    if (!this.pkg.has(presentation)) return [];
    const presRoot = this.pkg.readXml(presentation) as XmlNode;
    const sldIdLst = child(presRoot, 'p:sldIdLst');
    const ids = children(sldIdLst, 'p:sldId');
    if (ids.length) {
      const out: string[] = [];
      for (const id of ids) {
        const rid = strAttr(id, 'r:id');
        const rel = rid ? this.pkg.getRel(presentation, rid) : undefined;
        if (rel && !rel.external && rel.targetPart) out.push(rel.targetPart);
        else this.ctx.warn(`幻灯片条目缺少有效 r:id`);
      }
      return out;
    }
    // 兜底：按文件名数字排序
    const slides = [...this.pkg.entries.keys()].filter((p) => /^ppt\/slides\/slide\d+\.xml$/i.test(p));
    slides.sort((a, b) => slideNum(a) - slideNum(b));
    return slides;
  }

  private slideLayout(slidePath: string): string | undefined {
    const hit = this.layoutCache.get(slidePath);
    if (hit !== undefined) return hit || undefined;
    const rel = this.pkg.getRelsByType(slidePath, REL.slideLayout)[0];
    const path = rel?.targetPart;
    this.layoutCache.set(slidePath, path ?? '');
    return path;
  }

  private slideMaster(layoutPath: string): string | undefined {
    const hit = this.masterCache.get(layoutPath);
    if (hit !== undefined) return hit || undefined;
    const rel =
      this.pkg.getRelsByType(layoutPath, SLIDE_MASTER_REL)[0] ?? this.pkg.getRelsByType(layoutPath, REL.slideLayout)[0];
    const path = rel?.targetPart;
    this.masterCache.set(layoutPath, path ?? '');
    return path;
  }

  // ---- 单页 ----

  private parseSlide(slidePath: string, index: number): SlideData {
    const root = this.pkg.readXml(slidePath) as XmlNode;
    const cSld = child(root, 'p:cSld') ?? root;
    const spTree = child(cSld, 'p:spTree') ?? cSld;
    const hidden = strAttr(root, 'show') === '0';
    const shapes = this.flattenShapeTree(spTree, 0, 0, slidePath);

    const ordered = orderShapes(shapes);
    const titleShape = findTitleShape(ordered);
    const title = titleShape ? paraText(firstNonEmpty(titleShape.paras)).trim() : '';

    const notes = this.parseNotes(slidePath);

    return { index, title, hidden, shapes: ordered, notes };
  }

  /** 展平 spTree：支持分组偏移。 */
  private flattenShapeTree(spTree: XmlNode, offX: number, offY: number, slidePath: string): Shape[] {
    const out: Shape[] = [];
    for (const sp of children(spTree, 'p:sp')) {
      const s = this.parseTextShape(sp, offX, offY, slidePath);
      if (s) out.push(s);
    }
    for (const pic of children(spTree, 'p:pic')) out.push(this.parsePicture(pic, offX, offY, slidePath));
    for (const fr of children(spTree, 'p:graphicFrame')) out.push(this.parseGraphicFrame(fr, offX, offY, slidePath));
    for (const grp of children(spTree, 'p:grpSp')) {
      const g = this.groupOffset(grp);
      const subTree = child(grp, 'p:spTree') ?? grp;
      out.push(...this.flattenShapeTree(subTree, offX + g.x, offY + g.y, slidePath));
    }
    return out;
  }

  private groupOffset(group: XmlNode): { x: number; y: number } {
    const xfrm = child(group, 'p:grpSpPr') ?? child(group, 'p:xfrm');
    const off = child(xfrm, 'a:off');
    const chOff = child(xfrm, 'a:chOff');
    if (off && chOff) {
      return { x: numAttr(off, 'x') - numAttr(chOff, 'x'), y: numAttr(off, 'y') - numAttr(chOff, 'y') };
    }
    return { x: 0, y: 0 };
  }

  private shapeGeom(spPr: XmlNode): { x: number; y: number; w: number; h: number } {
    const xfrm = child(spPr, 'a:xfrm');
    return geomFromXfrm(xfrm);
  }

  private frameGeom(fr: XmlNode): { x: number; y: number; w: number; h: number } {
    return geomFromXfrm(child(fr, 'p:xfrm'));
  }

  private parseTextShape(sp: XmlNode, offX: number, offY: number, slidePath: string): Shape | undefined {
    const nvSpPr = child(sp, 'p:nvSpPr');
    const cNvPr = child(nvSpPr, 'p:cNvPr');
    const nvPr = child(nvSpPr, 'p:nvPr');
    const ph = child(nvPr, 'p:ph');
    const g = this.shapeGeom(child(sp, 'p:spPr'));
    const txBody = child(sp, 'p:txBody');
    const paras = txBody ? this.parseParas(txBody, sp, slidePath) : [];
    return {
      name: strAttr(cNvPr, 'name') ?? '',
      phType: strAttr(ph, 'type'),
      phIdx: strAttr(ph, 'idx'),
      kind: 'text',
      x: g.x + offX,
      y: g.y + offY,
      w: g.w,
      h: g.h,
      paras,
    };
  }

  private parsePicture(pic: XmlNode, offX: number, offY: number, slidePath: string): Shape {
    const cNvPr = child(child(pic, 'p:nvPicPr'), 'p:cNvPr');
    const blip = child(child(pic, 'p:blipFill'), 'a:blip');
    const g = this.shapeGeom(child(pic, 'p:spPr'));
    const embed = strAttr(blip, 'r:embed');
    let target = '';
    if (embed) {
      const rel = this.pkg.getRel(slidePath, embed);
      target = rel?.targetPart ?? '';
    }
    return {
      name: strAttr(cNvPr, 'name') ?? '',
      kind: 'picture',
      x: g.x + offX,
      y: g.y + offY,
      w: g.w,
      h: g.h,
      paras: [],
      image: { target, alt: strAttr(cNvPr, 'descr') },
    };
  }

  private parseGraphicFrame(fr: XmlNode, offX: number, offY: number, slidePath: string): Shape {
    const cNvPr = child(child(fr, 'p:nvGraphicFramePr'), 'p:cNvPr');
    const g = this.frameGeom(fr);
    const graphicData = child(child(fr, 'a:graphic'), 'a:graphicData');
    const uri = strAttr(graphicData, 'uri') ?? '';
    const name = strAttr(cNvPr, 'name') ?? '';
    if (uri.includes('chart')) {
      const chart = child(graphicData, 'c:chart');
      const rid = strAttr(chart, 'r:id');
      const rel = rid ? this.pkg.getRel(slidePath, rid) : undefined;
      return {
        name,
        kind: 'chart',
        x: g.x + offX,
        y: g.y + offY,
        w: g.w,
        h: g.h,
        paras: [],
        chart: { target: rel?.targetPart ?? '', title: name },
      };
    }
    const tbl = child(graphicData, 'a:tbl');
    if (tbl) {
      return {
        name,
        kind: 'table',
        x: g.x + offX,
        y: g.y + offY,
        w: g.w,
        h: g.h,
        paras: [],
        table: parseTable(tbl),
      };
    }
    return { name, kind: 'other', x: g.x + offX, y: g.y + offY, w: g.w, h: g.h, paras: [] };
  }

  private parseParas(txBody: XmlNode, sp: XmlNode, slidePath: string): Para[] {
    const bulletDefault = this.resolveBulletDefault(sp, slidePath);
    const out: Para[] = [];
    for (const p of children(txBody, 'a:p')) {
      const pPr = child(p, 'a:pPr');
      const level = numAttr(pPr, 'lvl', 0);
      const align = strAttr(pPr, 'algn');
      let bullet = bulletDefault;
      if (child(pPr, 'a:buNone') !== undefined) bullet = false;
      else if (child(pPr, 'a:buChar') !== undefined || child(pPr, 'a:buAutoNum') !== undefined) bullet = true;
      const runs = this.parseRuns(p, slidePath);
      const text = runs.map((r) => r.text).join('');
      out.push({ runs, level, bullet, align, empty: !text.trim() });
    }
    return out;
  }

  private parseRuns(p: XmlNode, slidePath: string): Run[] {
    const runs: Run[] = [];
    for (const r of children(p, 'a:r')) {
      const rPr = child(r, 'a:rPr');
      const t = textOf(child(r, 'a:t')).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
      const link = this.hyperlinkTarget(rPr, slidePath);
      runs.push({
        text: t,
        bold: boolAttr(rPr, 'b', false) || undefined,
        italic: boolAttr(rPr, 'i', false) || undefined,
        underline: strAttr(rPr, 'u') !== undefined ? true : undefined,
        strike: strAttr(rPr, 'strike') !== undefined && strAttr(rPr, 'strike') !== 'noStrike' ? true : undefined,
        link,
        code: isCodeFont(rPr),
        sz: numAttr(rPr, 'sz', 0),
      });
    }
    for (const _br of children(p, 'a:br')) runs.push({ text: '\n' });
    for (const fld of children(p, 'a:fld')) {
      const t = textOf(child(fld, 'a:t')).replace(/^‹#›$/, '').trim();
      if (t) runs.push({ text: t });
    }
    for (const _tab of children(p, 'a:tab')) runs.push({ text: ' ' });
    return runs;
  }

  private hyperlinkTarget(rPr: XmlNode, slidePath: string): string | undefined {
    const h = child(rPr, 'a:hlinkClick');
    const rid = strAttr(h, 'r:id');
    if (!rid) return undefined;
    const rel = this.pkg.getRel(slidePath, rid);
    return rel ? (rel.external ? rel.target : rel.targetPart) : undefined;
  }

  private resolveBulletDefault(sp: XmlNode, slidePath: string): boolean {
    const ph = child(child(child(sp, 'p:nvSpPr'), 'p:nvPr'), 'p:ph');
    if (!ph) return false;
    const phType = strAttr(ph, 'type');
    const phIdx = strAttr(ph, 'idx');
    if (phType === 'title' || phType === 'ctrTitle' || phType === 'subTitle') return false;

    const layout = this.slideLayout(slidePath);
    const master = layout ? this.slideMaster(layout) : undefined;
    const layoutBullet = this.findPlaceholderBullet(layout, phType, phIdx);
    if (layoutBullet !== undefined) return layoutBullet;
    const masterBullet = this.findPlaceholderBullet(master, phType, phIdx);
    if (masterBullet !== undefined) return masterBullet;
    return phType === 'body' || phType === undefined;
  }

  private findPlaceholderBullet(
    partPath: string | undefined,
    phType: string | undefined,
    phIdx: string | undefined,
  ): boolean | undefined {
    if (!partPath || !this.pkg.has(partPath)) return undefined;
    const root = this.pkg.readXml(partPath) as XmlNode;
    const cSld = child(root, 'p:cSld') ?? root;
    const spTree = child(cSld, 'p:spTree') ?? cSld;
    for (const sp of children(spTree, 'p:sp')) {
      const ph = child(child(child(sp, 'p:nvSpPr'), 'p:nvPr'), 'p:ph');
      if (!ph) continue;
      const type = strAttr(ph, 'type');
      const idx = strAttr(ph, 'idx');
      if (phType && type === phType) return placeholderBullet(sp);
      if (!phType && idx !== undefined && idx === phIdx) return placeholderBullet(sp);
    }
    return undefined;
  }

  private parseNotes(slidePath: string): string {
    const rel = this.pkg.getRelsByType(slidePath, REL.notesSlide)[0];
    if (!rel?.targetPart || !this.pkg.has(rel.targetPart)) return '';
    const root = this.pkg.readXml(rel.targetPart) as XmlNode;
    const cSld = child(root, 'p:cSld') ?? root;
    const spTree = child(cSld, 'p:spTree') ?? cSld;
    const parts: string[] = [];
    for (const sp of children(spTree, 'p:sp')) {
      const ph = child(child(child(sp, 'p:nvSpPr'), 'p:nvPr'), 'p:ph');
      if (!ph) continue;
      const type = strAttr(ph, 'type');
      if (type === 'sldImg' || type === 'sldNum' || type === 'hdr' || type === 'ftr' || type === 'dt') continue;
      const txBody = child(sp, 'p:txBody');
      for (const p of children(txBody, 'a:p')) {
        const t = textOf(p).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').trim();
        if (t) parts.push(t);
      }
    }
    return parts.join('\n');
  }

  // ---- 块构建 ----

  private buildBlocks(slides: SlideData[]): Block[] {
    const blocks: Block[] = [];
    const visible = this.ctx.options.hiddenSlides ? slides : slides.filter((s) => !s.hidden);
    let first = true;
    for (const slide of visible) {
      if (!first && this.ctx.options.slideSeparator !== false) {
        blocks.push({ type: 'hr' });
      }
      first = false;
      const title = slide.title || (this.ctx.options.slideNumbers ? `第 ${slide.index} 页` : `Slide ${slide.index}`);
      const displayTitle =
        this.ctx.options.slideNumbers && slide.title ? `第 ${slide.index} 页 · ${slide.title}` : title;
      blocks.push({
        type: 'section',
        title: displayTitle,
        level: 2,
        meta: { 页码: String(slide.index) },
        blocks: this.slideBlocks(slide),
      });
    }
    return blocks;
  }

  private slideBlocks(slide: SlideData): Block[] {
    const out: Block[] = [];
    const titleShape = slide.shapes.find((s) => s.kind === 'text' && (s.phType === 'title' || s.phType === 'ctrTitle'));
    for (const shape of slide.shapes) {
      if (shape === titleShape) continue;
      switch (shape.kind) {
        case 'text':
          out.push(...this.textShapeBlocks(shape));
          break;
        case 'table':
          if (shape.table) out.push(...tableBlocks(shape.table));
          break;
        case 'picture':
          out.push(...this.pictureBlocks(shape));
          break;
        case 'chart':
          out.push(...this.chartBlocks(shape));
          break;
        default:
          break;
      }
    }
    if (slide.notes && this.ctx.options.notes !== false) {
      const lines = slide.notes.split('\n').filter(Boolean);
      out.push({
        type: 'quote',
        children: [
          {
            type: 'paragraph',
            children: [
              { kind: 'strong', children: [{ kind: 'text', value: '演讲备注' }] },
              { kind: 'text', value: `：${lines.join(' ')}` },
            ],
          },
        ],
      });
    }
    return out;
  }

  private textShapeBlocks(shape: Shape): Block[] {
    const blocks: Block[] = [];
    const builder = new ListTreeBuilder(blocks);

    for (const para of shape.paras) {
      if (para.empty && !para.bullet) {
        builder.flush();
        continue;
      }
      if (para.bullet) {
        builder.add({
          level: para.level,
          ordered: false,
          start: 1,
          content: [{ type: 'paragraph', children: runsToInline(para.runs) }],
        });
      } else {
        builder.flush();
        const children = runsToInline(para.runs);
        const text = para.runs.map((r) => r.text).join('');
        if (text.trim()) {
          blocks.push({ type: 'paragraph', children, align: paraAlign(para) });
        }
      }
    }
    builder.flush();
    return blocks;
  }

  private pictureBlocks(shape: Shape): Block[] {
    if (!shape.image?.target) return [];
    const bytes = this.pkg.readBlob(shape.image.target);
    if (!bytes) return [];
    const ref = this.ctx.emitAsset(shape.image.target, bytes);
    if (!ref) return [];
    const alt = shape.image.alt ? normalizeText(shape.image.alt) : undefined;
    return [{ type: 'image', src: ref, alt }];
  }

  private chartBlocks(shape: Shape): Block[] {
    if (!shape.chart?.target) return [];
    const table = parseChart(this.pkg, shape.chart.target);
    if (!table) return [];
    const blocks: Block[] = [];
    if (shape.chart.title) {
      blocks.push({
        type: 'paragraph',
        children: [{ kind: 'strong', children: [{ kind: 'text', value: shape.chart.title }] }],
      });
    }
    blocks.push({ type: 'table', header: table.header, rows: table.rows });
    return blocks;
  }
}

function slideNum(path: string): number {
  const m = /slide(\d+)\.xml$/i.exec(path);
  return m ? Number(m[1]) : 0;
}

/** 从 a:xfrm 读取几何。 */
function geomFromXfrm(xfrm: XmlNode | undefined): { x: number; y: number; w: number; h: number } {
  const off = child(xfrm, 'a:off');
  const ext = child(xfrm, 'a:ext');
  return {
    x: numAttr(off, 'x'),
    y: numAttr(off, 'y'),
    w: numAttr(ext, 'cx'),
    h: numAttr(ext, 'cy'),
  };
}

function placeholderBullet(sp: XmlNode): boolean {
  const txBody = child(sp, 'p:txBody');
  const lstStyle = child(txBody, 'a:lstStyle');
  const lvl = child(lstStyle, 'a:lvl1pPr') ?? lstStyle;
  if (child(lvl, 'a:buNone') !== undefined) return false;
  if (child(lvl, 'a:buChar') !== undefined || child(lvl, 'a:buAutoNum') !== undefined) return true;
  return false;
}

function isCodeFont(rPr: XmlNode): boolean {
  const typeface = strAttr(child(rPr, 'a:latin'), 'typeface');
  if (!typeface) return false;
  const t = typeface.toLowerCase();
  return ['consolas', 'courier', 'monospace', 'monaco', 'menlo', 'sf mono', 'cascadia', 'jetbrains mono', 'source code pro'].some((f) =>
    t.includes(f),
  );
}

/** 找到标题形状：占位符 title/ctrTitle；否则顶部最大字号文本。 */
function findTitleShape(shapes: Shape[]): Shape | undefined {
  const ph = shapes.find((s) => s.kind === 'text' && (s.phType === 'title' || s.phType === 'ctrTitle'));
  if (ph) return ph;
  let best: Shape | undefined;
  let bestSz = 0;
  let bestY = Infinity;
  for (const s of shapes) {
    if (s.kind !== 'text' || s.phType === 'subTitle') continue;
    const para = firstNonEmpty(s.paras);
    if (!para) continue;
    const sz = maxRunSize(para);
    if (s.y / EMU_PER_PX < 200 && sz > 0) {
      if (sz > bestSz || (sz === bestSz && s.y < bestY)) {
        best = s;
        bestSz = sz;
        bestY = s.y;
      }
    }
  }
  return best;
}

function firstNonEmpty(paras: Para[]): Para | undefined {
  return paras.find((p) => !p.empty);
}

function paraText(para: Para | undefined): string {
  if (!para) return '';
  return para.runs.map((r) => r.text).join('');
}

function maxRunSize(para: Para): number {
  let max = 0;
  for (const r of para.runs) if ((r.sz ?? 0) > max) max = r.sz ?? 0;
  return max;
}

/** 阅读顺序：按行聚类（y），行内按 x。 */
function orderShapes(shapes: Shape[]): Shape[] {
  const sorted = [...shapes].filter((s) => s.kind !== 'other').sort((a, b) => a.y - b.y);
  const rows: Shape[][] = [];
  for (const s of sorted) {
    const last = rows[rows.length - 1];
    if (last && last.length && Math.abs(s.y - last[0].y) <= Math.max(s.h, last[0].h) * 0.6) {
      last.push(s);
    } else {
      rows.push([s]);
    }
  }
  const out: Shape[] = [];
  for (const row of rows) {
    row.sort((a, b) => a.x - b.x);
    out.push(...row);
  }
  return out;
}

function parseTable(tbl: XmlNode): TableLike {
  const tblPr = child(tbl, 'a:tblPr');
  const grid = children(child(tbl, 'a:tblGrid'), 'a:gridCol');
  const colCount = grid.length ? grid.length : 0;
  const trs = children(tbl, 'a:tr');
  const rawRows: { cells: { text: string; span: number }[] }[] = [];
  let maxCols = colCount;
  for (const tr of trs) {
    const cells = children(tr, 'a:tc');
    const row: { text: string; span: number }[] = [];
    for (const tc of cells) {
      const tcPr = child(tc, 'a:tcPr');
      const span = Math.max(1, numAttr(tcPr, 'gridSpan', 1));
      const text = textOf(tc)
        .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '')
        .split('\n')
        .map((s) => s.trim())
        .filter(Boolean)
        .join(' ');
      row.push({ text, span });
    }
    rawRows.push({ cells: row });
    const total = row.reduce((acc, c) => acc + c.span, 0);
    if (total > maxCols) maxCols = total;
  }
  if (maxCols === 0) maxCols = 1;

  const gridRows: string[][] = [];
  for (const r of rawRows) {
    const outRow: string[] = [];
    for (const c of r.cells) {
      for (let i = 0; i < c.span; i++) outRow.push(i === 0 ? c.text : '');
    }
    while (outRow.length < maxCols) outRow.push('');
    gridRows.push(outRow);
  }

  const header = boolAttr(tblPr, 'firstRow', true);
  const align = grid.map(() => 'left' as const);
  return { grid: gridRows, header, align };
}

function paraAlign(para: Para): 'left' | 'center' | 'right' | undefined {
  if (para.align === 'ctr') return 'center';
  if (para.align === 'r') return 'right';
  return undefined;
}

function runsToInline(runs: Run[]): Inline[] {
  const out: Inline[] = [];
  for (const r of runs) {
    if (r.text === '\n') {
      out.push({ kind: 'break' });
      continue;
    }
    if (!r.text) continue;
    let node: Inline = { kind: 'text', value: r.text };
    if (r.code) node = { kind: 'code', value: r.text };
    if (r.strike) node = { kind: 'del', children: [node] };
    if (r.italic) node = { kind: 'em', children: [node] };
    if (r.underline) node = { kind: 'em', children: [node] };
    if (r.bold) node = { kind: 'strong', children: [node] };
    if (r.link) node = { kind: 'link', href: r.link, children: [node] };
    out.push(node);
  }
  return out;
}

function tableBlocks(t: TableLike): Block[] {
  const header = t.header ? t.grid[0] ?? [] : [];
  const rows = t.header ? t.grid.slice(1) : t.grid;
  if (!rows.length) return [];
  if (!rows.some((r) => r.some((c) => c.trim()))) return [];
  return [{ type: 'table', header, rows, align: t.align }];
}

/** 图表 → Markdown 表格。 */
export function parseChart(pkg: OoxmlPackage, chartPath: string): { header: string[]; rows: string[][] } | undefined {
  if (!pkg.has(chartPath)) return undefined;
  const root = pkg.readXml(chartPath) as XmlNode;
  const chartSpace = child(root, 'c:chartSpace') ?? root;
  const chart = child(chartSpace, 'c:chart');
  const plotArea = child(chart, 'c:plotArea');
  const ser = findSeries(plotArea);
  if (!ser) return undefined;
  const name = seriesName(ser);
  const cats = seriesCategories(ser);
  const vals = seriesValues(ser);
  const header = cats.length ? ['类别', name || '值'] : [name || '值'];
  const rows: string[][] = [];
  const n = Math.max(cats.length, vals.length);
  for (let i = 0; i < n; i++) {
    rows.push([cats[i] ?? '', vals[i] ?? '']);
  }
  return { header, rows };
}

function findSeries(plotArea: XmlNode | undefined): XmlNode | undefined {
  if (!plotArea || typeof plotArea !== 'object') return undefined;
  for (const key of Object.keys(plotArea)) {
    if (key.startsWith('@') || key === '#text') continue;
    const v = plotArea[key];
    const arr = Array.isArray(v) ? v : [v];
    for (const node of arr) {
      if (node && typeof node === 'object') {
        const ser = child(node, 'c:ser');
        if (ser) return Array.isArray(ser) ? ser[0] : ser;
      }
    }
  }
  return undefined;
}

function seriesName(ser: XmlNode): string {
  return findFirstV(child(ser, 'c:tx')) ?? '';
}

function seriesCategories(ser: XmlNode): string[] {
  const cat = child(ser, 'c:cat');
  const strPts = children(child(child(cat, 'c:strRef'), 'c:strCache'), 'c:pt');
  if (strPts.length) return strPts.map((pt) => textOf(child(pt, 'c:v')).trim());
  const numPts = children(child(child(cat, 'c:numRef'), 'c:numCache'), 'c:pt');
  return numPts.map((pt) => textOf(child(pt, 'c:v')).trim());
}

function seriesValues(ser: XmlNode): string[] {
  const val = child(ser, 'c:val');
  const numPts = children(child(child(val, 'c:numRef'), 'c:numCache'), 'c:pt');
  if (numPts.length) return numPts.map((pt) => textOf(child(pt, 'c:v')).trim());
  const strPts = children(child(child(val, 'c:strRef'), 'c:strCache'), 'c:pt');
  return strPts.map((pt) => textOf(child(pt, 'c:v')).trim());
}

function findFirstV(node: XmlNode | undefined): string | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const direct = child(node, 'c:v');
  if (direct) return textOf(direct).trim() || undefined;
  for (const key of Object.keys(node)) {
    if (key.startsWith('@')) continue;
    const v = node[key];
    const arr = Array.isArray(v) ? v : [v];
    for (const item of arr) {
      const found = findFirstV(item);
      if (found !== undefined) return found;
    }
  }
  return undefined;
}

/** 对外入口：包 → Doc。 */
export function parsePptx(pkg: OoxmlPackage, ctx: ParseContext): Doc {
  return new PptxParser(pkg, ctx).parse();
}
