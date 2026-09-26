import type { Block, Doc, Inline, ParseContext } from '../types.js';
import { OoxmlPackage } from '../core/ooxml.js';
import { child, children, numAttr, strAttr, textOf, type XmlNode } from '../core/xml.js';
import { ListTreeBuilder } from '../core/lists.js';
import { normalizeText } from '../normalize.js';

const MAIN = 'word/document.xml';

interface StyleInfo {
  id: string;
  name: string;
  type: string;
  outlineLvl?: number;
  numId?: string;
  isCode: boolean;
  isQuote: boolean;
}

interface NumLevel {
  fmt: string;
  text: string;
  start: number;
}

interface CellLike {
  text: string;
  span: number;
}

/** DOCX 解析器。 */
export class DocxParser {
  private readonly styles = new Map<string, StyleInfo>();
  private readonly numbering = new Map<string, NumLevel[]>();
  private readonly footnotes = new Map<string, Block[]>();
  private readonly footnoteOrder: string[] = [];

  constructor(
    private readonly pkg: OoxmlPackage,
    private readonly ctx: ParseContext,
  ) {}

  parse(): Doc {
    this.loadStyles();
    this.loadNumbering();
    this.loadFootnotes();

    const metaBase = this.pkg.readCoreProps();
    const meta: Doc['meta'] = {
      title: metaBase.title,
      author: metaBase.author,
      subject: metaBase.subject,
      description: metaBase.description,
      created: metaBase.created,
      modified: metaBase.modified,
      lastModifiedBy: metaBase.lastModifiedBy,
      revision: metaBase.revision,
      company: metaBase.company,
      keywords: metaBase.keywordsRaw ? metaBase.keywordsRaw.split(',').map((s) => s.trim()).filter(Boolean) : undefined,
    };

    const blocks: Block[] = [];
    if (this.pkg.has(MAIN)) {
      const root = this.pkg.readXml(MAIN) as XmlNode;
      const body = child(root, 'w:body') ?? root;
      this.parseBlocks(body, blocks, MAIN);
    } else {
      this.ctx.warn('DOCX 未找到 word/document.xml');
    }

    if (this.footnoteOrder.length) {
      const items = this.footnoteOrder
        .map((id) => ({ label: id, blocks: this.footnotes.get(id) ?? [] }))
        .filter((x) => x.blocks.length > 0);
      if (items.length) {
        blocks.push({ type: 'footnotes', title: '脚注', items });
      }
    }

    return { format: 'docx', meta, blocks };
  }

  // ---- 样式 / 编号 / 脚注 ----

  private loadStyles(): void {
    if (!this.pkg.has('word/styles.xml')) return;
    const root = this.pkg.readXml('word/styles.xml') as XmlNode;
    for (const st of children(root, 'w:style')) {
      const id = strAttr(st, 'w:styleId') ?? '';
      if (!id) continue;
      const name = strAttr(child(st, 'w:name'), 'w:val') ?? '';
      const pPr = child(st, 'w:pPr');
      const outlineLvlRaw = strAttr(child(pPr, 'w:outlineLvl'), 'w:val');
      const numId = strAttr(child(child(pPr, 'w:numPr'), 'w:numId'), 'w:val');
      const lowerName = name.toLowerCase();
      this.styles.set(id, {
        id,
        name,
        type: strAttr(st, 'w:type') ?? 'paragraph',
        outlineLvl: outlineLvlRaw !== undefined ? Number(outlineLvlRaw) : headingFromName(lowerName),
        numId,
        isCode: lowerName.includes('code') || lowerName.includes('源') || lowerName.includes('代码'),
        isQuote: lowerName.includes('quote') || lowerName.includes('引用'),
      });
    }
  }

  private loadNumbering(): void {
    if (!this.pkg.has('word/numbering.xml')) return;
    const root = this.pkg.readXml('word/numbering.xml') as XmlNode;
    const abstracts = new Map<string, NumLevel[]>();
    for (const an of children(root, 'w:abstractNum')) {
      const id = strAttr(an, 'w:abstractNumId') ?? '';
      const lvls: NumLevel[] = [];
      for (const lvl of children(an, 'w:lvl')) {
        const ilvl = numAttr(lvl, 'w:ilvl', 0);
        const fmt = strAttr(child(lvl, 'w:numFmt'), 'w:val') ?? 'bullet';
        const text = strAttr(child(lvl, 'w:lvlText'), 'w:val') ?? '';
        const start = numAttr(child(lvl, 'w:start'), 'w:val', 1);
        const rFonts = strAttr(child(child(lvl, 'w:rPr'), 'w:rFonts'), 'w:ascii') ?? '';
        lvls[ilvl] = { fmt, text: bulletCharFor(fmt, text, rFonts), start };
      }
      abstracts.set(id, lvls);
    }
    for (const num of children(root, 'w:num')) {
      const numId = strAttr(num, 'w:numId') ?? '';
      const absId = strAttr(child(num, 'w:abstractNumId'), 'w:val') ?? '';
      this.numbering.set(numId, abstracts.get(absId) ?? []);
    }
  }

  private loadFootnotes(): void {
    if (!this.pkg.has('word/footnotes.xml')) return;
    const root = this.pkg.readXml('word/footnotes.xml') as XmlNode;
    for (const fn of children(root, 'w:footnote')) {
      const id = strAttr(fn, 'w:id') ?? '';
      if (!id) continue;
      const n = Number(id);
      if (!Number.isFinite(n) || n < 0) continue;
      const blocks: Block[] = [];
      this.parseBlocks(fn, blocks, 'word/footnotes.xml');
      this.footnotes.set(id, blocks.filter((b) => b.type !== 'hr'));
    }
  }

  // ---- 块级解析 ----

  private parseBlocks(container: XmlNode, out: Block[], part: string): void {
    const list = new ListTreeBuilder(out);
    // 保持文档顺序：遍历对象键
    for (const key of Object.keys(container)) {
      if (key.startsWith('@') || key === '#text') continue;
      const v = container[key];
      const arr = Array.isArray(v) ? v : [v];
      if (key === 'w:p') {
        for (const p of arr) this.parseParagraph(p, out, list, part);
      } else if (key === 'w:tbl') {
        for (const t of arr) {
          list.flush();
          this.parseTable(t, out, part);
        }
      } else if (key === 'w:sdt') {
        for (const sdt of arr) {
          const content = child(sdt, 'w:sdtContent');
          if (content) this.parseBlocks(content, out, part);
        }
      } else if (key === 'w:customXml') {
        for (const custom of arr) this.parseBlocks(custom, out, part);
      }
      // w:sectPr / w:bookmarkStart / w:altChunk 忽略
    }
    list.flush();
  }

  private parseParagraph(p: XmlNode, out: Block[], list: ListTreeBuilder, part: string): void {
    const pPr = child(p, 'w:pPr');
    const styleId = strAttr(child(pPr, 'w:pStyle'), 'w:val');
    const style = styleId ? this.styles.get(styleId) : undefined;

    let headingLevel: number | undefined;
    const directLvl = strAttr(child(pPr, 'w:outlineLvl'), 'w:val');
    if (directLvl !== undefined) {
      const n = Number(directLvl);
      if (Number.isFinite(n) && n >= 0 && n <= 8) headingLevel = Math.min(6, n + 1);
    }
    if (headingLevel === undefined && style?.outlineLvl !== undefined && style.outlineLvl >= 0) {
      headingLevel = Math.min(6, style.outlineLvl + 1);
    }

    const numPr = child(pPr, 'w:numPr');
    let numId = strAttr(child(numPr, 'w:numId'), 'w:val');
    const ilvl = numAttr(child(numPr, 'w:ilvl'), 'w:val', 0);
    if (!numId && style?.numId) numId = style.numId;
    const lvls = numId ? this.numbering.get(numId) : undefined;
    const level = lvls?.[ilvl];

    const inlines = this.parseInline(p, part);
    const text = inlinePlain(inlines).trim();
    const images = inlines.filter((i): i is Extract<Inline, { kind: 'image' }> => i.kind === 'image');
    const isEmpty = !text && images.length === 0;

    if (isEmpty) return;

    if (images.length && !text) {
      for (const img of images) {
        out.push({ type: 'image', src: img.src, alt: img.alt });
      }
      return;
    }

    const align = alignFromJc(strAttr(child(pPr, 'w:jc'), 'w:val'));

    if (headingLevel !== undefined) {
      list.flush();
      out.push({ type: 'heading', level: headingLevel, children: inlines });
      return;
    }

    if (level !== undefined) {
      const ordered = level.fmt !== 'bullet';
      list.add({
        level: ilvl,
        ordered,
        start: level.start || 1,
        content: [{ type: 'paragraph', children: inlines }],
      });
      return;
    }

    if (style?.isQuote) {
      list.flush();
      out.push({ type: 'quote', children: [{ type: 'paragraph', children: inlines }] });
      return;
    }

    list.flush();
    out.push({ type: 'paragraph', children: inlines, align });
  }

  // ---- 内联解析（保持文档顺序） ----

  private parseInline(p: XmlNode, part: string): Inline[] {
    const out: Inline[] = [];
    let fieldDepth = 0;
    let fieldEmitting = false;

    const pushRun = (r: XmlNode): void => {
      const fldChar = strAttr(child(r, 'w:fldChar'), 'w:fldCharType');
      if (fldChar === 'begin') {
        fieldDepth += 1;
        fieldEmitting = false;
        return;
      }
      if (fldChar === 'separate') {
        fieldEmitting = true;
        return;
      }
      if (fldChar === 'end') {
        fieldDepth = Math.max(0, fieldDepth - 1);
        fieldEmitting = fieldDepth === 0;
        return;
      }
      if (child(r, 'w:instrText') !== undefined) return;

      const rPr = child(r, 'w:rPr');
      const textParts: string[] = [];
      for (const t of children(r, 'w:t')) textParts.push(textOf(t));
      for (const _tab of children(r, 'w:tab')) textParts.push(' ');
      for (const _nb of children(r, 'w:noBreakHyphen')) textParts.push('-');

      for (const br of children(r, 'w:br')) {
        const type = strAttr(br, 'w:type');
        out.push(type === 'page' ? { kind: 'pageBreak' } : { kind: 'break' });
      }
      for (const dr of children(r, 'w:drawing')) {
        const img = this.extractDrawing(dr, part);
        if (img) out.push({ kind: 'image', src: img.src, alt: img.alt });
      }
      for (const pict of children(r, 'w:pict')) {
        const img = this.extractPict(pict, part);
        if (img) out.push({ kind: 'image', src: img.src, alt: img.alt });
      }
      for (const fn of children(r, 'w:footnoteReference')) {
        const id = strAttr(fn, 'w:id');
        if (id && this.footnotes.has(id)) {
          if (!this.footnoteOrder.includes(id)) this.footnoteOrder.push(id);
          out.push({ kind: 'text', value: `[^${id}]` });
        }
      }

      let text = textParts.join('');
      if (fieldDepth > 0 && !fieldEmitting) text = '';
      if (text) out.push(this.styleRun(rPr, text));
    };

    if (!p || typeof p !== 'object') return [];
    for (const key of Object.keys(p)) {
      if (key.startsWith('@') || key === '#text') continue;
      const v = p[key];
      const arr = Array.isArray(v) ? v : [v];
      if (key === 'w:r') {
        for (const r of arr) pushRun(r);
      } else if (key === 'w:hyperlink') {
        for (const h of arr) {
          const inner: Inline[] = [];
          for (const r of children(h, 'w:r')) {
            const before = out.length;
            pushRun(r);
            inner.push(...out.slice(before));
          }
          const rid = strAttr(h, 'r:id');
          const anchor = strAttr(h, 'w:anchor');
          let href: string | undefined;
          if (rid) {
            const rel = this.pkg.getRel(part, rid);
            if (rel && rel.external) href = rel.target;
          }
          if (!href && anchor && !anchor.startsWith('_Toc') && !anchor.startsWith('_Ref')) href = `#${anchor}`;
          if (href && inner.length) {
            // 把已推入 out 的内容换成语义链接（链接文本不变）
            out.splice(out.length - inner.length, inner.length, { kind: 'link', href, children: inner });
          }
        }
      } else if (key === 'w:fldSimple') {
        for (const fld of arr) {
          const instr = strAttr(fld, 'w:instr') ?? '';
          const url = /HYPERLINK\s+"([^"]+)"/i.exec(instr)?.[1];
          const inner: Inline[] = [];
          for (const r of children(fld, 'w:r')) {
            const before = out.length;
            pushRun(r);
            inner.push(...out.slice(before));
          }
          if (url && inner.length) {
            out.splice(out.length - inner.length, inner.length, { kind: 'link', href: url, children: inner });
          }
        }
      } else if (key === 'w:proofErr' || key === 'w:bookmarkStart' || key === 'w:bookmarkEnd') {
        continue;
      }
    }
    return out;
  }

  private styleRun(rPr: XmlNode, text: string): Inline {
    const get = (name: string): boolean => {
      const el = child(rPr, `w:${name}`);
      if (el === undefined) return false;
      const v = strAttr(el, 'w:val');
      return v === undefined ? true : v !== '0' && v !== 'false' && v !== 'none';
    };
    const rStyle = strAttr(child(rPr, 'w:rStyle'), 'w:val');
    const styleInfo = rStyle ? this.styles.get(rStyle) : undefined;
    const styleName = styleInfo?.name.toLowerCase() ?? '';

    const bold = get('b') || styleName.includes('strong') || styleName.includes('粗体');
    const italic = get('i') || styleName.includes('emphasis') || styleName.includes('斜体');
    const underline = get('u') || styleName.includes('underline') || styleName.includes('下划线');
    const strike = get('strike');
    const va = strAttr(child(rPr, 'w:vertAlign'), 'w:val');
    const isCode = !!styleInfo?.isCode || isMonoFont(rPr);

    let node: Inline = { kind: 'text', value: text };
    if (isCode) node = { kind: 'code', value: text };
    if (strike) node = { kind: 'del', children: [node] };
    if (italic) node = { kind: 'em', children: [node] };
    if (underline) node = { kind: 'em', children: [node] };
    if (bold) node = { kind: 'strong', children: [node] };
    if (va === 'superscript') node = { kind: 'sup', children: [node] };
    if (va === 'subscript') node = { kind: 'sub', children: [node] };
    return node;
  }

  private extractDrawing(drawing: XmlNode, part: string): { src: string; alt?: string } | undefined {
    const inline = child(drawing, 'wp:inline') ?? child(drawing, 'wp:anchor');
    if (!inline) return undefined;
    const docPr = child(inline, 'wp:docPr');
    const graphicData = child(child(inline, 'a:graphic'), 'a:graphicData');
    const blipEl = child(child(child(graphicData, 'pic:pic'), 'pic:blipFill'), 'a:blip') ?? child(graphicData, 'a:blip');
    const rid = strAttr(blipEl, 'r:embed') ?? strAttr(blipEl, 'r:link');
    if (!rid) return undefined;
    const rel = this.pkg.getRel(part, rid);
    if (!rel?.targetPart) return undefined;
    const bytes = this.pkg.readBlob(rel.targetPart);
    if (!bytes) return undefined;
    const src = this.ctx.emitAsset(rel.targetPart, bytes);
    if (!src) return undefined;
    const alt = strAttr(docPr, 'descr') || strAttr(docPr, 'name');
    return { src, alt: alt ? normalizeText(alt) : undefined };
  }

  private extractPict(pict: XmlNode, part: string): { src: string; alt?: string } | undefined {
    const data = child(child(pict, 'v:shape'), 'v:imagedata') ?? child(pict, 'v:imagedata');
    const rid = strAttr(data, 'r:id');
    if (!rid) return undefined;
    const rel = this.pkg.getRel(part, rid);
    if (!rel?.targetPart) return undefined;
    const bytes = this.pkg.readBlob(rel.targetPart);
    if (!bytes) return undefined;
    const src = this.ctx.emitAsset(rel.targetPart, bytes);
    if (!src) return undefined;
    return { src };
  }

  // ---- 表格 ----

  private parseTable(tbl: XmlNode, out: Block[], part: string): void {
    const look = child(child(tbl, 'w:tblPr'), 'w:tblLook');
    const firstRowHeader = look ? strAttr(look, 'w:firstRow') === '1' : false;

    const rows: CellLike[][] = [];
    for (const tr of children(tbl, 'w:tr')) {
      const cells: CellLike[] = [];
      for (const tc of children(tr, 'w:tc')) {
        const tcPr = child(tc, 'w:tcPr');
        const span = Math.max(1, numAttr(child(tcPr, 'w:gridSpan'), 'w:val', 1));
        const text = this.cellText(tc, part);
        cells.push({ text, span });
      }
      rows.push(cells);
    }
    if (!rows.length) return;

    let maxCols = 0;
    for (const r of rows) maxCols = Math.max(maxCols, r.reduce((a, c) => a + c.span, 0));
    if (maxCols === 0) return;

    const grid: string[][] = [];
    for (const r of rows) {
      const row: string[] = [];
      for (const c of r) {
        for (let i = 0; i < c.span; i++) row.push(i === 0 ? c.text : '');
      }
      while (row.length < maxCols) row.push('');
      grid.push(row);
    }

    if (!grid.some((r) => r.some((c) => c.trim()))) return;

    const header = firstRowHeader ? grid[0] : [];
    const dataRows = firstRowHeader ? grid.slice(1) : grid;
    out.push({ type: 'table', header, rows: dataRows });
  }

  private cellText(tc: XmlNode, part: string): string {
    const parts: string[] = [];
    for (const p of children(tc, 'w:p')) {
      const t = inlinePlain(this.parseInline(p, part)).replace(/\s+/g, ' ').trim();
      if (t) parts.push(t);
    }
    return parts.join(' ');
  }
}

function headingFromName(lowerName: string): number | undefined {
  const m = /^heading\s*([1-9])$/.exec(lowerName);
  if (m) return Number(m[1]) - 1;
  if (lowerName === 'title') return 0;
  if (lowerName === 'subtitle') return 1;
  return undefined;
}

function bulletCharFor(fmt: string, text: string, rFonts: string): string {
  if (fmt !== 'bullet') return text;
  const font = rFonts.toLowerCase();
  if (font === 'symbol' || font === 'wingdings' || font === 'wingdings2' || font === 'wingdings3') {
    const map: Record<string, string> = {
      '\uf0b7': '•',
      '\uf0b8': '•',
      '\uf0a7': '▪',
      '\uf0a8': '•',
      '\uf0b2': '◆',
      '\uf0d8': '✦',
      '\uf0fc': '✓',
      '\uf0fe': '☑',
      '\uf0a0': '◻',
    };
    return map[text] ?? '•';
  }
  return text || '•';
}

function isMonoFont(rPr: XmlNode): boolean {
  const fonts = child(rPr, 'w:rFonts');
  const ascii = strAttr(fonts, 'w:ascii') ?? strAttr(fonts, 'w:hAnsi');
  if (!ascii) return false;
  const t = ascii.toLowerCase();
  return ['consolas', 'courier', 'monaco', 'menlo', 'cascadia', 'jetbrains mono', 'source code pro', 'monospace'].some((f) =>
    t.includes(f),
  );
}

function alignFromJc(jc: string | undefined): 'left' | 'center' | 'right' | undefined {
  if (jc === 'center') return 'center';
  if (jc === 'right' || jc === 'end') return 'right';
  return undefined;
}

function inlinePlain(inlines: Inline[]): string {
  let s = '';
  for (const i of inlines) s += inlineTextRec(i);
  return s;
}

function inlineTextRec(i: Inline): string {
  switch (i.kind) {
    case 'text':
    case 'code':
      return i.value;
    case 'strong':
    case 'em':
    case 'del':
    case 'sup':
    case 'sub':
    case 'link':
      return i.children.map(inlineTextRec).join('');
    case 'image':
      return '';
    case 'break':
      return '\n';
    case 'pageBreak':
      return '\n';
  }
}

/** 对外入口。 */
export function parseDocx(pkg: OoxmlPackage, ctx: ParseContext): Doc {
  return new DocxParser(pkg, ctx).parse();
}
